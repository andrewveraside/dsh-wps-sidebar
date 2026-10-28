// Minimal client for the office-helper.exe JSON-lines protocol.
// The plugin itself talks to the same helper; this is the copy the smoke test uses
// so that a broken protocol is caught before a user ever opens the sidebar.
'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, execFileSync } = require('child_process');

/** sha256 of helper/OfficeHelper.cs, first 12 lowercase hex chars. */
function sourceHash(repoRoot) {
  const src = path.join(repoRoot, 'helper', 'OfficeHelper.cs');
  const data = fs.readFileSync(src);
  return crypto.createHash('sha256').update(data).digest('hex').slice(0, 12);
}

/** Where the plugin caches the compiled helper: %LOCALAPPDATA%\dsh-office\helper-<hash>\office-helper.exe */
function exePath(repoRoot) {
  const hash = sourceHash(repoRoot);
  const dir = path.join(
    process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'),
    'dsh-office',
    `helper-${hash}`,
  );
  return { hash, exe: path.join(dir, 'office-helper.exe'), dir };
}

/**
 * stdio of the helper.
 *
 * Answers come back through FILES: the helper writes JSON lines into files that we
 * hand it as its stdout/stderr and poll ourselves (same handles, opened `w+`).
 * Piped stdout needs a console the test may not have, and reopening a file for
 * reading after the child started does not see its bytes here; one shared
 * read-write handle is what works.
 *
 * Commands go IN through a real PIPE -- `RedirectStandardInput` -- which the
 * sandbox allows, because only *named* pipes are refused there. Do NOT hand the
 * child a file as stdin (2026-10-28/29, all probed):
 *   - the child gets EOF at once and exits 0 after printing {"event":"ready"}
 *     without answering anything -- wherever the `node` process was launched from
 *     a shell that shares its console (`cmd /c "helper < in > out"` behaves the
 *     same way);
 *   - a writable handle instead makes ReadLine throw UnauthorizedAccessException;
 *   - the only route that did deliver bytes was a read handle on an already filled
 *     file started by `node` as a direct child of the shell -- too narrow to rely on.
 * A pipe has none of those problems: the helper reads from it exactly as the plugin
 * feeds it in production.
 */
class Helper {
  constructor(exe, opts = {}) {
    this.exe = exe;
    this.timeoutMs = opts.timeoutMs || 60000;
    this.workDir = opts.workDir || fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-helper-'));
    this.proc = null;
    this.buf = '';
    this.waiters = [];
    this.stderr = '';
    this.exited = null;
    this._outPos = 0;
  }

  async start() {
    this.outFile = path.join(this.workDir, 'helper-out.jsonl');
    this.errFile = path.join(this.workDir, 'helper-err.log');
    // Both are opened READ-WRITE (`w+`) and stay open: the child writes into them
    // through this very handle, and the same handle serves the polling reads.
    // Reopening them for reading after the child started does not see the child's
    // bytes in this environment.
    this.outFd = fs.openSync(this.outFile, 'w+');
    this.errFd = fs.openSync(this.errFile, 'w+');

    const ready = new Promise((resolve, reject) => {
      this._readyResolve = resolve;
      this._readyReject = reject;
    });

    this.proc = spawn(this.exe, [], {
      stdio: ['pipe', this.outFd, this.errFd],
      windowsHide: true,
      // detach: the child runs in its own process group, so a Ctrl+C or console
      // teardown aimed at this test never lands on it.
      detached: true,
    });
    this.proc.on('error', (e) => {
      if (this._readyReject) this._readyReject(e);
    });
    this.proc.stdin.on('error', () => {
      // EPIPE after the helper died: the exit handler reports the real failure.
    });
    this.exited = new Promise((resolve) => {
      this.proc.on('exit', (code, signal) => {
        this._pump();
        const err = new Error(`helper exited (code=${code}, signal=${signal})`);
        for (const w of this.waiters) w(err);
        this.waiters = [];
        resolve({ code, signal });
      });
    });

    this._poll = setInterval(() => this._pump(), 25);
    const kill = setTimeout(() => this._readyReject(new Error('helper did not report ready in 20s')), 20000);
    try {
      return await ready;
    } finally {
      clearTimeout(kill);
    }
  }

  _pump() {
    let chunk;
    try {
      const size = fs.fstatSync(this.outFd).size;
      if (size === undefined || size <= this._outPos) return;
      const buf = Buffer.alloc(size - this._outPos);
      fs.readSync(this.outFd, buf, 0, buf.length, this._outPos);
      this._outPos += buf.length;
      chunk = buf.toString('utf8');
    } catch {
      return;
    }
    this.buf += chunk;
    let i;
    while ((i = this.buf.indexOf('\n')) >= 0) {
      const line = this.buf.slice(0, i).trim();
      this.buf = this.buf.slice(i + 1);
      if (!line) continue;
      let msg;
      try {
        msg = JSON.parse(line);
      } catch {
        continue;
      }
      if (msg.event === 'ready') {
        if (this._readyResolve) {
          this._readyResolve(msg);
          this._readyResolve = null;
        }
        continue;
      }
      const w = this.waiters.shift();
      if (w) w(null, msg);
    }
  }

  send(payload) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        const idx = this.waiters.indexOf(wrapped);
        if (idx >= 0) this.waiters.splice(idx, 1);
        reject(new Error(`helper timeout after ${this.timeoutMs}ms: ${JSON.stringify(payload).slice(0, 120)}`));
      }, this.timeoutMs);
      const wrapped = (err, msg) => {
        clearTimeout(timer);
        if (err) reject(err);
        else resolve(msg);
      };
      this.waiters.push(wrapped);
      this.proc.stdin.write(JSON.stringify(payload) + '\n');
    });
  }

  stop() {
    clearInterval(this._poll);
    if (!this.proc || this.proc.exitCode !== null) {
      this._closeFds();
      return Promise.resolve();
    }
    // Ending stdin is what tells the helper to finish.
    try {
      this.proc.stdin.end();
    } catch {}
    const pid = this.proc.pid;
    const force = setTimeout(() => {
      // The helper ignores EOF (or is stuck in a COM call): stop waiting for it,
      // and record that the stop was not clean so the test can print it.
      try {
        execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { windowsHide: true, timeout: 5000 });
      } catch {}
      this.stopNote = 'helper did not exit after stdin closed; killed it (taskkill /F)';
    }, 4000);
    return this.exited.then(() => {
      clearTimeout(force);
      this._closeFds();
    });
  }

  _closeFds() {
    for (const k of ['outFd', 'errFd']) {
      try {
        if (this[k] != null) fs.closeSync(this[k]);
      } catch {}
      this[k] = null;
    }
  }

  readStderr() {
    try {
      const size = fs.fstatSync(this.errFd).size;
      if (!size) return '';
      const buf = Buffer.alloc(size);
      fs.readSync(this.errFd, buf, 0, size, 0);
      return buf.toString('utf8');
    } catch {
      return '';
    }
  }
}

module.exports = { Helper, exePath, sourceHash };
