#!/usr/bin/env node
// End-to-end smoke test for the WPS helper.
//
//   node scripts/office-smoke.mjs [--keep] [--timeout 90000]
//
// What it proves: a freshly built office-helper.exe can drive a real WPS window
// through the whole lifecycle a user hits   open -> read -> edit -> render ->
// save -> versions -> revert -> close  -- and that every step answers ok:true.
//
// Prerequisites: WPS installed, and the helper compiled for the CURRENT
// helper/OfficeHelper.cs (run scripts/build-helper.ps1 first). Exit code 0 = pass.
//
// Never touches the user's documents: the fixture is generated in a temp dir and
// removed unless --keep is given.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '..');
const { Helper, exePath } = require('./lib/helper.cjs');
const { makeDocx } = require('./lib/docx.cjs');

const args = process.argv.slice(2);
const keep = args.includes('--keep');
const timeoutIdx = args.indexOf('--timeout');
const timeoutMs = timeoutIdx >= 0 ? Number(args[timeoutIdx + 1]) : 90000;

const green = (s) => `\x1b[32m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const dim = (s) => `\x1b[2m${s}\x1b[0m`;

let stepNo = 0;
const results = [];
function pass(name, detail) {
  stepNo += 1;
  results.push({ name, ok: true });
  console.log(`  ${green('PASS')} ${stepNo}. ${name}${detail ? dim(' — ' + detail) : ''}`);
}
function fail(name, detail) {
  stepNo += 1;
  results.push({ name, ok: false, detail });
  console.log(`  ${red('FAIL')} ${stepNo}. ${name}`);
  if (detail) console.log('       ' + String(detail).split('\n').join('\n       '));
}

function need(res, name) {
  if (!res || res.ok !== true) {
    const msg = res && res.error ? res.error : JSON.stringify(res).slice(0, 400);
    throw new Error(`${name} failed: ${msg}`);
  }
}

/** The helper answers {ok, result, id}: the payload everything else reads is `result`. */
const body = (res) => (res && res.result !== undefined ? res.result : res);

/** Paragraph texts of a `read` answer. Word reports the count in `paragraphs` and the text in `items`. */
function paragraphTexts(res) {
  const r = body(res) || {};
  const list = [].concat(r.items || [], r.paragraphs || [], r.blocks || []);
  return list.filter((p) => p && typeof p.text === 'string').map((p) => p.text);
}

async function main() {
  console.log('DSH WPS helper smoke test');
  console.log('');

  // --- environment ---------------------------------------------------------
  const { hash, exe } = exePath(repoRoot);
  console.log(dim(`source hash : ${hash}`));
  console.log(dim(`helper exe  : ${exe}`));

  if (!fs.existsSync(exe)) {
    fail('helper exe exists', `not found.\nThe helper must be compiled for the current OfficeHelper.cs.\nRun:  powershell -ExecutionPolicy Bypass -File scripts/build-helper.ps1\nthen re-run:  npm run smoke`);
    return 1;
  }
  pass('helper exe exists', `${Math.round(fs.statSync(exe).size / 1024)} KB`);

  // --- fixture -------------------------------------------------------------
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-smoke-'));
  const docPath = path.join(tmp, 'smoke.docx');
  fs.writeFileSync(docPath, makeDocx({ title: 'DSH smoke', heading: 'Smoke', paragraphs: ['smoke paragraph one', 'smoke paragraph two'] }));
  pass('fixture .docx written', docPath);

  const h = new Helper(exe, { timeoutMs });
  let opened = false;

  try {
    await h.start();
    pass('helper starts and reports ready');

    const st = await h.send({ cmd: 'status' });
    need(st, 'status');
    pass('status', 'helper is alive');

    const op = await h.send({ cmd: 'open', app: 'word', path: docPath, show: false });
    need(op, 'open');
    opened = true;
    // `created` means the app made a blank document instead of reading our file;
    // say so here, where the cause is visible, instead of failing at the read.
    const how = body(op).how;
    if (how !== 'opened') throw new Error(`open did not take the fixture: how=${how}`);
    pass('open', 'how=opened');

    const rd = await h.send({ cmd: 'read', app: 'word', doc: docPath });
    need(rd, 'read');
    const texts = paragraphTexts(rd);
    if (!texts.some((t) => t.includes('smoke paragraph one'))) {
      throw new Error('read did not return the fixture text; got: ' + JSON.stringify(texts).slice(0, 300));
    }
    pass('read', `${texts.length} paragraph(s)`);

    const ed = await h.send({
      cmd: 'edit',
      app: 'word',
      doc: docPath,
      ops: [{ op: 'insert_paragraphs', text: 'smoke edit ok' }],
    });
    need(ed, 'edit');
    pass('edit', 'insert_paragraphs');

    const rd2 = await h.send({ cmd: 'read', app: 'word', doc: docPath });
    need(rd2, 'read (after edit)');
    const texts2 = paragraphTexts(rd2);
    if (!texts2.some((t) => t.includes('smoke edit ok'))) {
      throw new Error('edited paragraph missing after re-read: ' + JSON.stringify(texts2).slice(0, 300));
    }
    pass('read after edit', 'edit is visible');

    // render needs an explicit output path; it writes nothing otherwise.
    const renderOut = path.join(tmp, 'page1.png');
    const rn = await h.send({ cmd: 'render', app: 'word', doc: docPath, page: 1, out: renderOut });
    need(rn, 'render');
    if (!fs.existsSync(renderOut) || fs.statSync(renderOut).size < 1000) {
      throw new Error('render wrote no usable image: ' + renderOut);
    }
    pass('render', `page 1 -> ${Math.round(fs.statSync(renderOut).size / 1024)} KB PNG`);

    const sv = await h.send({ cmd: 'save', app: 'word', doc: docPath });
    need(sv, 'save');
    pass('save');

    const vs = await h.send({ cmd: 'versions', app: 'word', path: docPath });
    need(vs, 'versions');
    const vbody = body(vs) || {};
    const vlist = vbody.items || vbody.versions || [];
    if (!vlist.length) {
      throw new Error(
        `save produced no version snapshot — expected files in ${vbody.dir || '<unknown>'}\n` +
          '(a snapshot is the file as it stood BEFORE the save; every save keeps one)',
      );
    }
    // Newest first (the helper's contract). The newest snapshot is the file before
    // this save, i.e. without the edit -- which is what step 13 verifies was restored.
    const target = vlist[0];
    pass('versions', `${vlist.length} snapshot(s), newest "${target.reason || '?'}"`);
    const ver = target.version || target.stamp || target.name;
    const rv = await h.send({ cmd: 'revert', app: 'word', path: docPath, version: ver, keep: false });
    need(rv, 'revert');
    pass('revert', `to ${String(ver).slice(0, 20)}`);

    const rd3 = await h.send({ cmd: 'read', app: 'word', doc: docPath });
    need(rd3, 'read (after revert)');
    const texts3 = paragraphTexts(rd3);
    if (texts3.some((t) => t.includes('smoke edit ok'))) {
      throw new Error('revert did not undo the edit: edited paragraph still present');
    }
    pass('read after revert', 'edit was rolled back');
  } catch (e) {
    fail('SMOKE ABORTED', e && e.message ? e.message : String(e));
    if (h.readStderr && h.readStderr()) {
      console.log(dim('helper stderr: ' + h.readStderr().slice(0, 800)));
    }
  } finally {
    if (opened) {
      try {
        const cl = await h.send({ cmd: 'close', app: 'word', doc: docPath, save: false });
        need(cl, 'close');
        pass('close');
      } catch (e) {
        fail('close', e && e.message ? e.message : String(e));
      }
    }
    await h.stop();
    if (h.stopNote) console.log(dim('note: ' + h.stopNote));
    if (!keep) {
      try {
        fs.rmSync(tmp, { recursive: true, force: true });
      } catch {}
    } else {
      console.log(dim('kept: ' + tmp));
    }
  }

  const bad = results.filter((r) => !r.ok);
  console.log('');
  if (bad.length) {
    console.log(red(`SMOKE FAILED: ${bad.length}/${results.length} step(s) failed`));
    return 1;
  }
  console.log(green(`SMOKE PASSED: ${results.length}/${results.length} steps`));
  return 0;
}

main().then(
  (code) => process.exit(code),
  (e) => {
    console.error(red('SMOKE ERROR: ' + (e && e.stack ? e.stack : e)));
    process.exit(1);
  },
);
