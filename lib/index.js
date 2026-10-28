import { createRequire } from "node:module";
import z from "@deepseek-ai/schemastery";
import { execFile, spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync } from "node:fs";
import { access, copyFile, mkdir, readFile, rename, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, win32 } from "node:path";
import { fileURLToPath } from "node:url";
import { Context, Service } from "@deepseek-ai/cordis";
//#region lib/types/helper-client.js
/**
* Builds and drives the Windows helper (helper/OfficeHelper.cs): compiled once
* with the in-box .NET Framework csc.exe into a per-source-hash cache folder,
* then kept running and spoken to as JSON lines over stdio.
*/
const REFERENCES = [
	"System.dll",
	"System.Core.dll",
	"Microsoft.CSharp.dll",
	"System.Drawing.dll",
	"System.Windows.Forms.dll",
	"System.Web.Extensions.dll",
	"System.IO.Compression.dll",
	"System.IO.Compression.FileSystem.dll"
];
/** A file in helper/ at the package root: one level up from lib/ or src/, two from lib/types/. */
function helperAssetPath(name) {
	const candidates = [`../helper/${name}`, `../../helper/${name}`].map((rel) => fileURLToPath(new URL(rel, import.meta.url)));
	return candidates.find((path) => existsSync(path)) ?? candidates[0];
}
function helperSourcePath() {
	return helperAssetPath("OfficeHelper.cs");
}
function frameworkDir() {
	const windir = process.env.SystemRoot ?? process.env.windir ?? "C:\\Windows";
	for (const arch of ["Framework64", "Framework"]) {
		const dir = join(windir, "Microsoft.NET", arch, "v4.0.30319");
		if (existsSync(join(dir, "csc.exe"))) return dir;
	}
	throw new Error("The .NET Framework 4 C# compiler (csc.exe) was not found. It ships with Windows 10/11; enable \".NET Framework 4.8\" in Windows Features.");
}
function cacheRoot() {
	return join(process.env.LOCALAPPDATA ?? tmpdir(), "dsh-office");
}
let building;
/** Compile the helper if this source version has no cached exe yet; returns the exe path. */
async function ensureHelperExe(source = helperSourcePath()) {
	building ??= (async () => {
		const code = await readFile(source);
		const hash = createHash("sha256").update(code).digest("hex").slice(0, 12);
		const debug = process.env.DSH_OFFICE_DEBUG === "1";
		const dir = join(cacheRoot(), `helper-${hash}${debug ? "-debug" : ""}`);
		const exe = join(dir, "office-helper.exe");
		if (existsSync(exe)) return exe;
		await mkdir(dir, { recursive: true });
		const fw = frameworkDir();
		const temp = join(dir, `build-${process.pid}-${Date.now()}.exe`);
		const args = [
			"-nologo",
			...debug ? ["-debug:pdbonly"] : ["-optimize+"],
			"-target:exe",
			`-out:${temp}`,
			...REFERENCES.map((ref) => `-r:${ref}`),
			source
		];
		await new Promise((resolve, reject) => {
			execFile(join(fw, "csc.exe"), args, {
				windowsHide: true,
				timeout: 12e4
			}, (error, stdout) => {
				if (error) reject(/* @__PURE__ */ new Error(`Compiling the office helper failed: ${String(stdout).trim() || error.message}`));
				else resolve();
			});
		});
		try {
			await rename(temp, exe);
		} catch (error) {
			await rm(temp, { force: true });
			if (!existsSync(exe)) throw error;
		}
		return exe;
	})();
	try {
		return await building;
	} catch (error) {
		building = void 0;
		throw error;
	}
}
var HelperClient = class {
	log;
	child;
	starting;
	pending = /* @__PURE__ */ new Map();
	listeners = /* @__PURE__ */ new Set();
	nextId = 0;
	disposed = false;
	constructor(log = () => {}) {
		this.log = log;
	}
	onEvent(listener) {
		this.listeners.add(listener);
		return () => {
			this.listeners.delete(listener);
		};
	}
	async call(cmd, args = {}, timeoutMs = 2e4) {
		if (this.disposed) throw new Error("office helper is shut down");
		const child = await this.ensure();
		const id = ++this.nextId;
		return new Promise((resolve, reject) => {
			const timer = setTimeout(() => {
				this.pending.delete(id);
				reject(/* @__PURE__ */ new Error(`office helper timed out on "${cmd}"`));
				this.kill();
			}, timeoutMs);
			this.pending.set(id, {
				resolve,
				reject,
				timer
			});
			child.stdin.write(`${JSON.stringify({
				...args,
				id,
				cmd
			})}\n`);
		});
	}
	dispose() {
		this.disposed = true;
		this.kill();
	}
	kill() {
		const child = this.child;
		this.child = void 0;
		this.starting = void 0;
		if (child && child.exitCode === null) {
			try {
				child.stdin.end();
			} catch {}
			setTimeout(() => {
				if (child.exitCode === null) child.kill();
			}, 800).unref();
		}
	}
	ensure() {
		if (this.child && this.child.exitCode === null) return Promise.resolve(this.child);
		this.starting ??= this.start().catch((error) => {
			this.starting = void 0;
			throw error;
		});
		return this.starting;
	}
	async start() {
		const exe = await ensureHelperExe();
		const child = spawn(exe, [], {
			stdio: [
				"pipe",
				"pipe",
				"pipe"
			],
			windowsHide: true
		});
		let buffer = "";
		let ready;
		const readyPromise = new Promise((resolve, reject) => {
			ready = resolve;
			child.once("error", reject);
			setTimeout(() => reject(/* @__PURE__ */ new Error("office helper did not start")), 15e3).unref();
		});
		child.stdout.setEncoding("utf8");
		child.stdout.on("data", (chunk) => {
			buffer += chunk;
			let index;
			while ((index = buffer.indexOf("\n")) >= 0) {
				const line = buffer.slice(0, index).trim();
				buffer = buffer.slice(index + 1);
				if (line) this.handle(line, () => ready?.());
			}
		});
		child.stderr.setEncoding("utf8");
		child.stderr.on("data", (chunk) => this.log(`helper stderr: ${chunk.trim()}`));
		child.on("exit", (code) => {
			if (this.child === child) this.child = void 0;
			this.starting = void 0;
			for (const [id, entry] of this.pending) {
				clearTimeout(entry.timer);
				entry.reject(/* @__PURE__ */ new Error(`office helper exited (code ${String(code)})`));
				this.pending.delete(id);
			}
		});
		await readyPromise;
		this.child = child;
		return child;
	}
	handle(line, onReady) {
		let message;
		try {
			message = JSON.parse(line);
		} catch {
			this.log(`helper sent a non-JSON line: ${line.slice(0, 200)}`);
			return;
		}
		if (typeof message.event === "string") {
			if (message.event === "ready") {
				onReady();
				return;
			}
			for (const listener of this.listeners) try {
				listener({
					event: message.event,
					...message.reason === void 0 ? {} : { reason: message.reason }
				});
			} catch {}
			return;
		}
		if (typeof message.id !== "number") return;
		const entry = this.pending.get(message.id);
		if (!entry) return;
		this.pending.delete(message.id);
		clearTimeout(entry.timer);
		if (message.ok) entry.resolve(message.result ?? null);
		else entry.reject(Object.assign(new Error(message.error ?? "helper error"), { code: message.code ?? "ERROR" }));
	}
};
//#endregion
//#region lib/types/prompt.js
/**
* The host's file-based Office skills this plugin stands in for, and the app of each. office-xlsx is left out on
* purpose: spreadsheet work is mostly data analysis, which code does better, so that skill loads as usual.
*/
const FILE_SKILLS = {
	"office-docx": "word",
	"office-pptx": "ppt"
};
const APPS = {
	word: "Microsoft Word",
	ppt: "Microsoft PowerPoint"
};
/**
* What the model reads the first time it loads one of those skills: this plugin's own short skill in its place, as
* the result of the call (not a refusal: nothing went wrong, and the user should not see an error).
*/
function redirectText(skill, app) {
	return [
		`Skill ${skill} on this computer: ${APPS[app]} (or WPS Office) is installed, so this work is done inside the app itself with the office_ tools, not by building the file with a script.`,
		"Call office_open with the absolute path (a path that does not exist yet creates the document; the app is started for you if it is not running), then office_read, office_edit, office_render, office_save. The user watches the document take shape in the window and can edit alongside; formulas, captions and layout are done by the app itself.",
		`Only if the task truly cannot be done that way (converting many files in bulk, a file the app cannot open, no document to show), call skill "${skill}" again and it will load.`
	].join("\n");
}
/** System-prompt section: when to reach for the office_ tools, in as few tokens as will do. */
function promptText() {
	return [
		"# Office documents (Word, Excel, PowerPoint)",
		"The office_ tools work inside the real Office apps on this computer (Microsoft Office, or WPS Office where that is what the user has — the tools and operations are the same): the document is open in its window, every change you make appears there at once, and the user can read and edit alongside you.",
		"- For any Word or PowerPoint work, creating a new document included, start with office_open (it starts the app if it is not running, and a path that does not exist yet creates the file), then office_read and office_edit. Do not load the office-docx / office-pptx skills and do not write the file with a script: an open file cannot be overwritten, and the user would not see the work.",
		"- Spreadsheets are the exception. Data analysis, and reading or producing .xlsx / .csv files, is done with code (the office-xlsx skill, pandas and the like), not in the Excel window. Use the office_ tools on a workbook only when the user asks for the work to be done in Excel or WPS Spreadsheets. If a workbook is open there and cannot be written, write the result to a copy; that also keeps the original data intact.",
		"- Put all the changes you can into one office_edit call; it runs them in order. What it can do in an app (the operations and their fields) is listed with the result of your first office_open or office_read there, and by office_help; it covers formulas, tables, pictures, SmartArt diagrams, shapes, styles and page layout, so look there before reaching for anything else.",
		"- The user may edit while you work. Read again before editing a part you have not just read, and pass \"expect\" on Word paragraphs.",
		"- Use these tools, not python-docx / python-pptx or other file libraries, whenever the document is to be worked on in the user's Office app or is already open there; and check the result with office_render, not by converting the file with another renderer, which lays it out differently from Word.",
		"- Finish with a polish pass, it is part of the job: office_render with overview: true for the whole document at a glance, then the pages that need a closer look; look at the pictures the way a careful reader would, fix what is off with office_edit, and render again until it is right. Look for: content outside the frame or section it belongs to; fonts, sizes and line spacing that differ from the surrounding text or from what the template asks for; formulas left as plain text; tables that are too wide, uneven or uncentred; captions not next to their table or figure; leftover template hints and sample text; runs of blank paragraphs and half-empty pages (in Word, fill_gaps closes the ones before figures and tables); pictures that are too small for their lettering to be read, or that leave most of a slide empty. Then office_save. Nothing is saved until you call office_save; do not close the user's documents.",
		"- In Word, formulas are LaTeX between dollar signs ($...$ inline, $$...$$ on its own line); they become native equations. Citations are \\cite{n} plus insert_references.",
		"- Set the look through styles before writing a long document (style_format for Normal and the headings, page_setup); all of that is done with office_edit, not with mouse and keyboard in Word's dialogs.",
		"- When filling in a template, keep its structure: put each part where its heading is, replace or delete the sample text and hints, and match the formatting the template asks for.",
		"- If a call reports the app is busy, the user is typing or has a dialog open: wait a moment and retry.",
		"- If you also have mouse / keyboard computer tools, use them on an Office window only for what office_edit cannot do, such as a setting that exists only in a dialog."
	].join("\n");
}
//#endregion
//#region lib/types/settings.js
/**
* Plugin settings shared by the host and the settings page: defaults and the
* tolerant reader for the Loader entry config. No schemastery here so the
* browser bundle stays small.
*/
/** The Loader entry id (settings namespace); see cordis.patch.yml. */
const ENTRY_ID = "andrewveraside-wps-sidebar";
const DEFAULTS = {
	suite: "auto",
	silent: false,
	follow: true,
	typing: true,
	card: true,
	finalCheck: true,
	preferLive: true,
	showOnOpen: true,
	renderWidth: 1100
};
/** Unwrap `.volatile()` refs (`{ get() }`) and fall back to defaults for bad values. */
function resolveConfig(raw) {
	const out = {};
	if (raw !== null && typeof raw === "object") for (const [key, value] of Object.entries(raw)) out[key] = value !== null && typeof value === "object" && typeof value.get === "function" ? value.get() : value;
	const width = out.renderWidth;
	return {
		suite: out.suite === "office" || out.suite === "wps" ? out.suite : DEFAULTS.suite,
		silent: typeof out.silent === "boolean" ? out.silent : DEFAULTS.silent,
		follow: typeof out.follow === "boolean" ? out.follow : DEFAULTS.follow,
		typing: typeof out.typing === "boolean" ? out.typing : DEFAULTS.typing,
		card: typeof out.card === "boolean" ? out.card : DEFAULTS.card,
		finalCheck: typeof out.finalCheck === "boolean" ? out.finalCheck : DEFAULTS.finalCheck,
		preferLive: typeof out.preferLive === "boolean" ? out.preferLive : DEFAULTS.preferLive,
		showOnOpen: typeof out.showOnOpen === "boolean" ? out.showOnOpen : DEFAULTS.showOnOpen,
		renderWidth: typeof width === "number" && Number.isFinite(width) ? Math.min(2e3, Math.max(600, Math.round(width))) : DEFAULTS.renderWidth
	};
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-scope@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4_@deepseek-ai+dsh-invariants_0d6b19188928ab894325adc32e18b0b5/node_modules/@deepseek-ai/dsh-scope/lib/index.js
/**
* Shared insertion-ordered storage and effect ownership for scope-aware registries.
*
* @module @deepseek-ai/dsh-scope
*/
/**
* Insertion-ordered named entries with caller-owned duplicate diagnostics.
*
* Values are borrowed. Iterators are live within one nonempty table
* generation; draining the table detaches them from later insertions. Each
* successful insertion returns an idempotent undo for that exact entry.
*/
var NamedEntries = class {
	duplicateError;
	data = /* @__PURE__ */ new Map();
	constructor(duplicateError) {
		this.duplicateError = duplicateError;
	}
	/**
	* Insert one unique name.
	* @param name - name unique within this table.
	* @param value - borrowed value to retain.
	* @returns an idempotent undo that removes only this insertion.
	*/
	insert(name, value) {
		const data = this.data;
		if (data.has(name)) throw this.duplicateError(name);
		data.set(name, value);
		let active = true;
		return () => {
			if (!active) return;
			active = false;
			data.delete(name);
			if (data.size === 0 && this.data === data) this.data = /* @__PURE__ */ new Map();
		};
	}
	/**
	* Read one named value.
	* @param name - name to resolve.
	* @returns the retained value, or `undefined` when absent.
	*/
	get(name) {
		return this.data.get(name);
	}
	/**
	* Test one name for membership.
	* @param name - name to test.
	* @returns whether the table contains that name.
	*/
	has(name) {
		return this.data.has(name);
	}
	/**
	* Iterate live names in insertion order.
	* @returns the native live key iterator.
	*/
	keys() {
		return this.data.keys();
	}
	/**
	* Iterate live entries in insertion order.
	* @returns the native live entry iterator.
	*/
	entries() {
		return this.data.entries();
	}
	/**
	* Iterate live values in insertion order.
	* @returns the native live value iterator.
	*/
	values() {
		return this.data.values();
	}
	/**
	* Test whether this table has no entries.
	* @returns whether the table is empty.
	*/
	isEmpty() {
		return this.data.size === 0;
	}
};
/**
* Insertion-ordered anonymous entries with independent registration identity.
*
* Equal values remain separate registrations. Values are borrowed, and
* iterators are live within one nonempty table generation; draining the table
* detaches them from later appends.
*/
var AnonymousEntries = class {
	data = /* @__PURE__ */ new Map();
	/**
	* Append one independently owned value.
	* @param value - borrowed value to retain.
	* @returns an idempotent undo for this exact append.
	*/
	append(value) {
		const data = this.data;
		const key = Symbol();
		data.set(key, value);
		let active = true;
		return () => {
			if (!active) return;
			active = false;
			data.delete(key);
			if (data.size === 0 && this.data === data) this.data = /* @__PURE__ */ new Map();
		};
	}
	/**
	* Iterate live values in insertion order.
	* @returns the native live value iterator.
	*/
	values() {
		return this.data.values();
	}
	/**
	* Test whether this table has no entries.
	* @returns whether the table is empty.
	*/
	isEmpty() {
		return this.data.size === 0;
	}
};
/**
* Own the global and exact-scope layers for one registry.
*
* Reads never create scoped layers. Registrations derive both visibility and
* effect ownership from the supplied Cordis context, collect undo before
* notification, and reclaim only a completely empty aggregate layer.
*/
var ScopedLayers = class {
	createLayer;
	onChange;
	/** The eagerly constructed context-global layer. */
	global;
	scoped = /* @__PURE__ */ new Map();
	constructor(createLayer, onChange) {
		this.createLayer = createLayer;
		this.onChange = onChange;
		this.global = createLayer(void 0);
	}
	/**
	* Read an existing exact-scope overlay. Deliberately chain-blind: callers
	* addressing one scope's OWN contributions (its restrictions, its guards)
	* must not silently pick up an ancestor's — use {@link chainLayers} where
	* inheritance is the point.
	* @param scope - exact scope key; `undefined` denotes no overlay.
	* @returns the existing scoped layer, or `undefined` without creating one.
	*/
	peek(scope) {
		if (scope === void 0) return void 0;
		return this.scoped.get(scope);
	}
	/**
	* Existing overlays along the scope's parent chain ({@link scopeChainOf}),
	* farthest ancestor first and the exact scope last, so a caller layering
	* them in order gives the nearest scope the final word.
	* @param scope - viewing scope, or `undefined` for no overlays.
	* @returns the existing layers, nearest last; absent overlays are skipped.
	*/
	chainLayers(scope) {
		const layers = [];
		for (const key of scopeChainOf(scope).reverse()) {
			const layer = this.scoped.get(key);
			if (layer !== void 0) layers.push(layer);
		}
		return layers;
	}
	/**
	* Materialize global named entries followed by scope-chain shadows,
	* farthest ancestor first, so the nearest scope's entry wins a name.
	* @param scope - viewing scope, or `undefined` for the global view.
	* @param pick - select the named table from a layer.
	* @returns an insertion-ordered effective map.
	*/
	merge(scope, pick) {
		const merged = new Map(pick(this.global).entries());
		for (const layer of this.chainLayers(scope)) for (const [name, value] of pick(layer).entries()) merged.set(name, value);
		return merged;
	}
	/**
	* Attach one synchronous layer mutation to its registration context.
	* @param ctx - context that determines both scope visibility and effect ownership.
	* @param action - atomic mutation returning its synchronous undo.
	* @param options - Cordis effect label and optional change notification.
	* @returns the exact disposer returned by `ctx.effect()`.
	*/
	effect(ctx, action, options) {
		const scope = scopeOf(ctx);
		const notify = options.notify ?? true;
		return ctx.effect(function* () {
			let layer;
			let created = false;
			if (scope === void 0) layer = this.global;
			else {
				const existing = this.scoped.get(scope);
				if (existing === void 0) {
					layer = this.createLayer(scope);
					this.scoped.set(scope, layer);
					created = true;
				} else layer = existing;
			}
			let undo;
			try {
				undo = action(layer);
			} catch (error) {
				if (scope !== void 0 && created && layer.isEmpty()) this.scoped.delete(scope);
				throw error;
			}
			yield () => {
				undo();
				if (scope !== void 0 && layer.isEmpty()) this.scoped.delete(scope);
				if (notify) this.onChange();
			};
			if (notify) this.onChange();
		}.bind(this), options.label);
	}
};
/**
* Scoped-context primitive: mint a Cordis context that tags registrations with
* an opaque identity and build routing-only event carriers for that identity.
*
* @module @deepseek-ai/dsh-scope
*/
/** Context tag written by {@link createScope}. */
const kScope = Symbol("dsh.scope");
/** The key associated with each carrier. Presence distinguishes an unkeyed carrier from a non-carrier. */
const carrierKeys = /* @__PURE__ */ new WeakMap();
/**
* The enclosing scope of each key. One relation powers both directions of
* scope nesting: registration views inherit DOWN the chain (a child scope
* sees its ancestors' layers — {@link ScopedLayers}), and event admission
* extends UP it (a listener tagged with an ancestor receives events dispatched
* to a descendant key — {@link scopeTarget}).
*/
const scopeParents = /* @__PURE__ */ new WeakMap();
/**
* The chain from a key to its root ancestor.
* @param key - the starting key, or `undefined` for the empty chain.
* @returns keys nearest-first: `[key, parent, grandparent, …]`.
*/
function scopeChainOf(key) {
	const chain = [];
	for (let cursor = key; cursor !== void 0; cursor = scopeParents.get(cursor)) chain.push(cursor);
	return chain;
}
/**
* Read the nearest scope tag inherited by a context.
* @param ctx - context to inspect.
* @returns its scope key, or `undefined` for an unscoped context.
*/
function scopeOf(ctx) {
	return ctx[kScope];
}
/**
* Build an opaque receiver that preserves the base filter, admits untagged
* listeners globally, and admits tagged listeners for a matching key or any
* of its ancestors ({@link bindScopeParent}): a listener owned by an enclosing
* scope receives every descendant scope's events, which is what lets one
* standing composition observe each of the agents composed under it. A tag
* BELOW the dispatch key stays excluded — events flow up the chain, never
* down.
* @param base - subject or service whose existing Cordis filter is preserved.
* @param key - routed scope identity, or `undefined` for an unscoped subject.
* @returns a carrier whose subject remains available only through event arguments.
*/
function scopeTarget(base, key) {
	const baseFilter = base[Context.filter];
	const carrier = { [Context.filter](ctx) {
		if (baseFilter !== void 0 && !baseFilter.call(base, ctx)) return false;
		const tag = scopeOf(ctx);
		if (tag === void 0) return true;
		for (let cursor = key; cursor !== void 0; cursor = scopeParents.get(cursor)) if (cursor === tag) return true;
		return false;
	} };
	carrierKeys.set(carrier, key);
	return carrier;
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-typert-protocol@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-typert-protocol/lib/index.js
/** The one Remote failure class shared by owners, the Gateway, and consumers. */
/**
* One Remote call failure: a real Error carrying its stable code and typed
* details. Owners throw it at the failure point; the Host Gateway encodes it
* onto the wire unchanged; the Client face rebuilds an instance for the
* `RemoteResult` error branch, so `throw result.error` keeps throw semantics.
* Discrimination is always by `code`, never by instanceof.
*/
var RemoteError = class extends Error {
	code;
	details;
	/** Structural marker: cross-realm/bundle identification never uses instanceof. */
	isDSHRemoteError = true;
	/**
	* @param code - stable failure code declared in {@link RemoteErrorDetailsMap}.
	* @param message - human diagnostic carried across the wire.
	* @param details - structured payload typed by the code.
	* @param options - standard Error options (`cause` survives in-process only).
	*/
	constructor(code, message, details, options) {
		super(message, options);
		this.code = code;
		this.details = details;
		this.name = "RemoteError";
	}
};
/**
* Remote decorators and explicit Gateway bindings backed by versioned
* descriptors carried on decorated class prototypes. Strict reflection
* remains a Typert compiler responsibility.
* @module @deepseek-ai/dsh-typert-protocol
*/
const TYPERT_REMOTE_SEGMENT_PATTERN = /^[A-Za-z0-9_$.-]+$/;
/**
* Test one generated Remote name against the Connection endpoint grammar.
* @param value - namespace, method, lookup, or Context segment.
* @returns whether the value can cross the shared RPC carrier unchanged.
*/
function isTypertRemoteSegment(value) {
	return value !== "." && value !== ".." && TYPERT_REMOTE_SEGMENT_PATTERN.test(value);
}
const REMOTE_METHOD_DESCRIPTOR = "@deepseek-ai/dsh-typert-protocol/remote-methods";
/**
* Bind one visible Service field to a Cordis key and Remote namespace. A
* service that owns a Cordis Context also gives its tree `ctx.invocation`,
* `undefined` outside a Remote call, so no `TypertRemoteService` is needed for
* a Host composition to read it.
* @param service - owning Service instance, normally `this`.
* @param serviceKey - exact Cordis service key.
* @param options - optional distinct wire namespace.
* @returns a frozen, inspectable binding with no compiler-injected metadata.
*/
function bindTypertRemote(service, serviceKey, options = {}) {
	validateName("service key", serviceKey);
	const namespace = options.namespace ?? serviceKey;
	validateName("namespace", namespace);
	const ctx = Reflect.get(service, "ctx");
	if (ctx instanceof Context) provideInvocationAccessor(ctx);
	return Object.freeze({
		service,
		serviceKey,
		namespace
	});
}
/** Cordis Service base that exposes its registered name through Typert Gateway. */
var TypertRemoteService = class extends Service {
	/** Visible binding consumed by the Gateway's source-mode discovery. */
	typertRemote;
	/**
	* Register the Service and bind the same key to Typert Gateway.
	* @param ctx - owning Cordis Context.
	* @param serviceKey - exact Cordis service key and default wire namespace.
	* @param options - optional distinct wire namespace.
	*/
	constructor(ctx, serviceKey, options = {}) {
		super(ctx, serviceKey);
		this.typertRemote = bindTypertRemote(this, this.name, options);
	}
};
/**
* Make `ctx.invocation` read as `undefined` outside a Remote call instead of the
* reflect service's "cannot get property" error; a call-derived Context shadows
* the accessor with its own property. The first Remote Service constructed in a
* tree registers it on the root, where it outlives any one Service.
*/
function provideInvocationAccessor(ctx) {
	if (Object.hasOwn(ctx.root.reflect.props, "invocation")) return;
	ctx.root.accessor("invocation", { get: () => void 0 });
}
function Remote(methodExportOrOptions, context) {
	if (typeof methodExportOrOptions === "string") {
		validateName("Remote export name", methodExportOrOptions);
		return remoteDecorator({ kind: "direct" }, void 0, methodExportOrOptions);
	}
	if (typeof methodExportOrOptions === "object") {
		if (remoteOptionMode(methodExportOrOptions) !== "stream" || Reflect.ownKeys(methodExportOrOptions).length !== 1) throw new TypeError("typert-protocol: Remote options must contain exactly mode: \"stream\"");
		return remoteDecorator({ kind: "direct" }, "stream");
	}
	if (context === void 0) throw new TypeError("typert-protocol: Remote decorator context is missing");
	addMarkerInitializer(context, { kind: "direct" });
}
function remoteOptionMode(options) {
	return Reflect.get(options, "mode");
}
function remoteDecorator(invocation, mode, exportName) {
	return function(_method, context) {
		addMarkerInitializer(context, invocation, mode, exportName);
	};
}
function readRemoteMethodDescriptor(prototype) {
	const property = Object.getOwnPropertyDescriptor(prototype, REMOTE_METHOD_DESCRIPTOR);
	if (property === void 0) return void 0;
	const descriptor = property.value;
	if (descriptor === null || typeof descriptor !== "object") throw new TypeError("typert-protocol: Remote method descriptor must be an object");
	const version = Reflect.get(descriptor, "version");
	if (version !== 1) throw new TypeError(`typert-protocol: unsupported Remote method descriptor version ${String(version)}`);
	const methods = Reflect.get(descriptor, "methods");
	if (!Array.isArray(methods)) throw new TypeError("typert-protocol: Remote method descriptor methods must be an array");
	return descriptor;
}
function addMarkerInitializer(context, invocation, mode, exportName) {
	if (context.private || context.static || typeof context.name !== "string") throw new TypeError("typert-protocol: Remote decorators require a public instance method with a string name");
	const method = context.name;
	context.addInitializer(function() {
		const prototype = Object.getPrototypeOf(this);
		if (prototype === null) throw new TypeError(`typert-protocol: cannot mark Remote method "${method}" on an object without a prototype`);
		mark(prototype, method, invocation, mode, exportName);
	});
}
function mark(prototype, method, invocation, mode, exportName) {
	const descriptor = readRemoteMethodDescriptor(prototype);
	const marker = Object.freeze({
		method,
		...exportName === void 0 || exportName === method ? {} : { exportName },
		...mode === void 0 ? {} : { mode },
		invocation: Object.freeze(invocation)
	});
	const current = descriptor?.methods.find((candidate) => candidate.method === method);
	if (current !== void 0) {
		if (current.exportName === marker.exportName && current.mode === marker.mode && sameInvocation(current.invocation, invocation)) return;
		throw new Error(`typert-protocol: Remote method "${method}" has conflicting invocation markers`);
	}
	Object.defineProperty(prototype, REMOTE_METHOD_DESCRIPTOR, {
		configurable: true,
		value: Object.freeze({
			version: 1,
			methods: Object.freeze([...descriptor?.methods ?? [], marker])
		})
	});
}
function sameInvocation(left, right) {
	if (left.kind === "direct") return right.kind === "direct";
	if (right.kind === "direct") return false;
	return left.context === right.context;
}
function validateName(subject, value) {
	if (!isTypertRemoteSegment(value)) throw new TypeError(`typert-protocol: ${subject} must contain only RPC endpoint segment characters`);
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-util-values@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-util-values/lib/index.js
/** Duplicate-install-safe JSON and immutable-value helpers. @module @deepseek-ai/dsh-util-values */
/**
* Mark an unreachable closed-union branch.
* @param value - impossible value; an unhandled typed variant fails at the call site.
* @param context - optional switch-site label included in the failure message.
* @returns never; a runtime value that escaped its type always throws.
*/
function assertNever(value, context) {
	const rendered = JSON.stringify(value) ?? String(value);
	throw new Error(`unreachable variant${context ? ` in ${context}` : ""}: ${rendered}`);
}
/** Whether a realm-owned intrinsic prototype has a native constructor matching this engine's representation. */
function hasIntrinsicConstructor$1(prototype, name) {
	const constructor = Object.getOwnPropertyDescriptor(prototype, "constructor")?.value;
	if (typeof constructor !== "function") return false;
	try {
		return constructor.name === name && constructor.prototype === prototype && Function.prototype.toString.call(constructor) === Function.prototype.toString.call(name === "Array" ? Array : Object);
	} catch {
		return false;
	}
}
/** Whether a candidate is one realm's intrinsic `Object.prototype`. */
function isIntrinsicObjectPrototype$1(value) {
	return Object.getPrototypeOf(value) === null && hasIntrinsicConstructor$1(value, "Object");
}
/** Whether an array uses one realm's intrinsic `Array.prototype`, not a subclass or forged prototype. */
function hasPlainArrayPrototype$1(value) {
	const prototype = Object.getPrototypeOf(value);
	if (!Array.isArray(prototype) || !hasIntrinsicConstructor$1(prototype, "Array")) return false;
	const objectPrototype = Object.getPrototypeOf(prototype);
	return typeof objectPrototype === "object" && objectPrototype !== null && isIntrinsicObjectPrototype$1(objectPrototype);
}
/** Whether an object is a plain or null-prototype record from any JavaScript realm. */
function hasPlainObjectPrototype(value) {
	const prototype = Object.getPrototypeOf(value);
	return prototype === null || typeof prototype === "object" && isIntrinsicObjectPrototype$1(prototype);
}
/** Return every JSON-visible object key, or reject own data JSON would discard. */
function enumerableStringKeys(value) {
	const keys = Reflect.ownKeys(value);
	if (keys.some((key) => typeof key !== "string" || !Object.prototype.propertyIsEnumerable.call(value, key))) return void 0;
	return keys;
}
/** Validate lossless JSON iteratively, optionally materializing a detached snapshot. */
function walkJsonValue(value, detach) {
	const ancestors = /* @__PURE__ */ new Set();
	let root;
	const assign = (destination, item) => {
		if (destination === void 0) return;
		if (destination.kind === "root") root = item;
		else if (destination.kind === "array") destination.target[destination.index] = item;
		else Object.defineProperty(destination.target, destination.key, {
			value: item,
			enumerable: true,
			configurable: true,
			writable: true
		});
	};
	const tasks = [{
		kind: "visit",
		value,
		...detach ? { destination: { kind: "root" } } : {}
	}];
	for (let task = tasks.pop(); task !== void 0; task = tasks.pop()) {
		if (task.kind === "leave") {
			ancestors.delete(task.source);
			continue;
		}
		if (task.kind === "array-item") {
			if (!Object.prototype.hasOwnProperty.call(task.source, task.index)) return void 0;
			tasks.push({
				kind: "visit",
				value: task.source[task.index],
				...task.target === void 0 ? {} : { destination: {
					kind: "array",
					target: task.target,
					index: task.index
				} }
			});
			continue;
		}
		if (task.kind === "object-property") {
			tasks.push({
				kind: "visit",
				value: task.source[task.key],
				...task.target === void 0 ? {} : { destination: {
					kind: "object",
					target: task.target,
					key: task.key
				} }
			});
			continue;
		}
		const current = task.value;
		if (current === null) {
			assign(task.destination, null);
			continue;
		}
		if (typeof current === "boolean" || typeof current === "string") {
			assign(task.destination, current);
			continue;
		}
		if (typeof current === "number") {
			if (!Number.isFinite(current) || Object.is(current, -0)) return void 0;
			assign(task.destination, current);
			continue;
		}
		if (typeof current !== "object") return void 0;
		if (ancestors.has(current)) return void 0;
		if (Array.isArray(current)) {
			if (!hasPlainArrayPrototype$1(current)) return void 0;
			const length = current.length;
			if (Reflect.ownKeys(current).length !== length + 1) return void 0;
			const target = detach ? [] : void 0;
			if (target !== void 0) assign(task.destination, target);
			ancestors.add(current);
			tasks.push({
				kind: "leave",
				source: current
			});
			for (let index = length - 1; index >= 0; index--) tasks.push({
				kind: "array-item",
				source: current,
				index,
				...target === void 0 ? {} : { target }
			});
			continue;
		}
		if (!hasPlainObjectPrototype(current)) return void 0;
		const keys = enumerableStringKeys(current);
		if (keys === void 0) return void 0;
		const target = detach ? {} : void 0;
		if (target !== void 0) assign(task.destination, target);
		ancestors.add(current);
		tasks.push({
			kind: "leave",
			source: current
		});
		for (let index = keys.length - 1; index >= 0; index--) {
			const key = keys[index];
			/* v8 ignore next -- the loop is bounded by the captured key count. */
			if (key === void 0) return void 0;
			tasks.push({
				kind: "object-property",
				source: current,
				key,
				...target === void 0 ? {} : { target }
			});
		}
	}
	return detach ? root : true;
}
/**
* Validate and detach lossless JSON in one read per property.
* @param value - candidate value to validate and detach.
* @returns the detached snapshot, or `undefined` when the value is not losslessly JSON-serializable.
*/
function snapshotJsonValue(value) {
	return walkJsonValue(value, true);
}
/**
* Test the same lossless JSON rules as {@link snapshotJsonValue} without detaching the value.
* @param value - candidate value to test.
* @returns whether the value survives a JSON round trip without loss.
*/
function isJsonValue(value) {
	return walkJsonValue(value, false) === true;
}
/**
* Deep-freeze an object graph in place while leaving live AbortSignal objects mutable.
* @param value - value to freeze.
* @returns the same value after every reachable enumerable child is frozen.
*/
function deepFreeze(value) {
	const seen = /* @__PURE__ */ new WeakSet();
	const pending = [{
		kind: "visit",
		node: value
	}];
	while (pending.length > 0) {
		const task = pending.pop();
		/* v8 ignore next -- the loop condition guarantees one pending task. */
		if (task === void 0) continue;
		if (task.kind === "property") {
			pending.push({
				kind: "visit",
				node: task.source[task.key]
			});
			continue;
		}
		const node = task.node;
		if (node === null || typeof node !== "object") continue;
		if (node instanceof AbortSignal) continue;
		if (seen.has(node)) continue;
		seen.add(node);
		Object.freeze(node);
		const keys = Object.keys(node);
		for (let index = keys.length - 1; index >= 0; index--) {
			const key = keys[index];
			/* v8 ignore next -- the loop is bounded by the captured key count. */
			if (key === void 0) continue;
			pending.push({
				kind: "property",
				source: node,
				key
			});
		}
	}
	return value;
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-util-crypto@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-util-crypto/lib/index.js
/**
* Random v4 UUID, minted from `crypto.getRandomValues`.
* @returns the UUID string.
*/
function randomUUID() {
	const bytes = globalThis.crypto.getRandomValues(/* @__PURE__ */ new Uint8Array(16));
	const hex = Array.from(bytes, (byte, index) => {
		return (index === 6 ? byte & 15 | 64 : index === 8 ? byte & 63 | 128 : byte).toString(16).padStart(2, "0");
	}).join("");
	return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-brand@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-brand/lib/index.js
/**
* Duplicate-install-safe nominal primitive helpers.
*
* A brand makes structurally identical strings or numbers non-interchangeable
* at the type level: a `SessionId` cannot be passed where a `ToolCallId` is
* expected, and an event sequence cannot be passed as a log offset. Comparison,
* logging, and serialization retain the underlying primitive behavior.
*
* This package owns no concrete domain value and keeps no runtime identity or mutable
* state, so independently installed copies produce interchangeable values.
*
* @module @deepseek-ai/dsh-brand
*/
/**
* Apply a compile-time string brand without changing the value.
* @param value - string admitted by the domain that owns the target brand.
* @returns the same string with the requested compile-time brand.
*/
function brandString(value) {
	return value;
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-timeout@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-timeout/lib/index.js
/** Largest delay Node schedules without clamping it to one millisecond. */
const MAX_TIMER_DELAY_MS = 2147483647;
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-llm@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4/node_modules/@deepseek-ai/dsh-llm/lib/index.js
/**
* Detach and deep-freeze a message whose identity already exists.
* @param message - complete message, including its stable identity.
* @returns an immutable snapshot that preserves the identity.
*/
function freezeMessage(message) {
	return deepFreeze(structuredClone(message));
}
/**
* Create one identified message and freeze it before publication.
* @param input - complete role, content, and source for a new message.
* @returns an immutable message with a fresh stable identity.
*/
function createMessage(input) {
	return deepFreeze(structuredClone({
		...input,
		id: brandString(randomUUID())
	}));
}
/**
* Create one identified user-role message and freeze it before publication.
* @param input - complete content and source for a new user message.
* @returns an immutable user message with a fresh stable identity.
*/
function createUserMessage(input) {
	return createMessage({
		...input,
		role: "user"
	});
}
/**
* Harness error base with a stable machine-routable code and chained cause.
* Package errors extend it so tool results and replay can retain failure class.
* @module @deepseek-ai/dsh-llm/error
*/
/**
* Base class for all harness errors. Carries a `code` (stable, programmatic —
* e.g. `NO_ADAPTER`, `INVALID_ARGS`, `INVARIANT`) distinct from the
* human-readable `message`, and supports `cause` chaining via the standard
* `ErrorOptions`. `name` defaults to the subclass constructor name.
*/
var HarnessError = class extends Error {
	/** Stable machine-routable failure class (e.g. `RATE_LIMIT`); route on this, never by parsing `message`. */
	code;
	constructor(message, code, options) {
		super(message, options);
		this.code = code;
		this.name = new.target.name;
	}
};
/**
* Canonical provider-neutral code for a response that completed normally but
* carried no content blocks at all. Providers occasionally emit a degenerate
* completion (a terminal stop with zero output); adapters classify it as this
* failure instead of yielding an empty assistant message, because an empty
* message silently ends the turn with nothing for the user or the loop to act
* on. The attempt produced nothing durable, so retry policy treats it as safe
* to repeat.
*/
const EMPTY_RESPONSE_CODE = "EMPTY_RESPONSE";
new RegExp(String.raw`(?:^|[^a-z0-9])context[\s_-](?:length|window)[\s_-]` + String.raw`(?:exceed(?:ed|s)?|overflow(?:ed)?|limit[\s_-]exceeded)(?:$|[^a-z0-9])`, "i");
new RegExp(String.raw`\b(?:request|prompt|input|messages?)\s+(?:is\s+|are\s+)?` + String.raw`too\s+(?:large|long)\s+for\s+(?:(?:this|the)\s+)?` + String.raw`(?:model(?:'s)?\s+)?context(?:\s+window)?\b`, "i");
new RegExp(String.raw`\b(?:input|prompt|request|messages?)\b.{0,40}` + String.raw`\b(?:exceed(?:s|ed)?|overflows?|is\s+larger\s+than)\b.{0,40}` + String.raw`\b(?:the\s+)?(?:model(?:'s)?\s+)?context(?:\s+(?:length|window))?\b`, "i");
/**
* Provider-owned request-retry policy configuration and resolution.
*
* Adapters expose one resolved policy per registered provider route; the
* optional dsh-llm-retry plugin executes it on the agent's failed-step extension point.
*
* @module @deepseek-ai/dsh-llm/retry-policy
*/
const DEFAULT_MAX_RETRIES = 5;
const DEFAULT_INITIAL_DELAY_MS = 500;
const DEFAULT_MAX_DELAY_MS = 1e4;
const DEFAULT_JITTER_RATIO = .1;
const DEFAULT_RETRYABLE_CODES = Object.freeze([
	EMPTY_RESPONSE_CODE,
	"RATE_LIMIT",
	"SERVER",
	"TIMEOUT",
	"TRANSPORT"
]);
const backoffSchema = z.object({
	initialDelayMs: z.number().max(MAX_TIMER_DELAY_MS).default(DEFAULT_INITIAL_DELAY_MS),
	maxDelayMs: z.number().max(MAX_TIMER_DELAY_MS).default(DEFAULT_MAX_DELAY_MS),
	jitterRatio: z.number().min(0).max(1).default(DEFAULT_JITTER_RATIO)
});
const normalPolicySchema = z.object({
	mode: z.const("normal").required(),
	maxRetries: z.number().step(1).min(0).max(Number.MAX_SAFE_INTEGER).default(DEFAULT_MAX_RETRIES),
	retryableCodes: z.array(z.string()).default([...DEFAULT_RETRYABLE_CODES]),
	backoff: backoffSchema
});
const alwaysPolicySchema = z.object({
	mode: z.const("always").required(),
	backoff: backoffSchema
});
z.union([normalPolicySchema, alwaysPolicySchema]);
const NORMAL_POLICY_KEYS = /* @__PURE__ */ new Set([
	"mode",
	"maxRetries",
	"retryableCodes",
	"backoff"
]);
const ALWAYS_POLICY_KEYS = /* @__PURE__ */ new Set([
	"mode",
	"maxRetries",
	"retryableCodes",
	"backoff"
]);
const BACKOFF_KEYS = /* @__PURE__ */ new Set([
	"initialDelayMs",
	"maxDelayMs",
	"jitterRatio"
]);
function validateKeys(value, allowed, path) {
	for (const key of Object.keys(value)) if (!allowed.has(key)) throw new Error(`${path}: unknown key "${key}"`);
}
function resolveBackoff(config, path) {
	if (config !== void 0) validateKeys(config, BACKOFF_KEYS, path);
	const initialDelayMs = config?.initialDelayMs ?? DEFAULT_INITIAL_DELAY_MS;
	const maxDelayMs = config?.maxDelayMs ?? DEFAULT_MAX_DELAY_MS;
	const jitterRatio = config?.jitterRatio ?? DEFAULT_JITTER_RATIO;
	if (!Number.isFinite(initialDelayMs) || initialDelayMs <= 0 || initialDelayMs > 2147483647) throw new Error(`${path}.initialDelayMs must be a positive finite number no greater than ${MAX_TIMER_DELAY_MS}`);
	if (!Number.isFinite(maxDelayMs) || maxDelayMs <= 0 || maxDelayMs > 2147483647) throw new Error(`${path}.maxDelayMs must be a positive finite number no greater than ${MAX_TIMER_DELAY_MS}`);
	if (initialDelayMs > maxDelayMs) throw new Error(`${path}.initialDelayMs must be less than or equal to maxDelayMs`);
	if (!Number.isFinite(jitterRatio) || jitterRatio < 0 || jitterRatio > 1) throw new Error(`${path}.jitterRatio must be between 0 and 1`);
	return Object.freeze({
		initialDelayMs,
		maxDelayMs,
		jitterRatio
	});
}
/**
* Validate, default, and detach one provider-owned retry policy.
* @param config - optional provider configuration; omission selects normal defaults.
* @param path - diagnostic path naming the provider config that owns the value.
* @returns an immutable policy safe to capture in provider registration state.
*/
function resolveRetryPolicy(config, path) {
	if (config === void 0) return Object.freeze({
		mode: "normal",
		maxRetries: DEFAULT_MAX_RETRIES,
		retryableCodes: DEFAULT_RETRYABLE_CODES,
		...resolveBackoff(void 0, `${path}.backoff`)
	});
	switch (config.mode) {
		case "normal": {
			validateKeys(config, NORMAL_POLICY_KEYS, path);
			const maxRetries = config.maxRetries ?? DEFAULT_MAX_RETRIES;
			const retryableCodes = config.retryableCodes ?? [...DEFAULT_RETRYABLE_CODES];
			if (!Number.isSafeInteger(maxRetries) || maxRetries < 0) throw new Error(`${path}.maxRetries must be a non-negative safe integer`);
			if (retryableCodes.length === 0) throw new Error(`${path}.retryableCodes must not be empty`);
			if (retryableCodes.some((code) => typeof code !== "string" || code.length === 0)) throw new Error(`${path}.retryableCodes must contain only non-empty strings`);
			if (new Set(retryableCodes).size !== retryableCodes.length) throw new Error(`${path}.retryableCodes must not contain duplicates`);
			return Object.freeze({
				mode: "normal",
				maxRetries,
				retryableCodes: Object.freeze([...retryableCodes]),
				...resolveBackoff(config.backoff, `${path}.backoff`)
			});
		}
		case "always":
			validateKeys(config, ALWAYS_POLICY_KEYS, path);
			return Object.freeze({
				mode: "always",
				...resolveBackoff(config.backoff, `${path}.backoff`)
			});
		default: throw new Error(`${path}.mode must be "normal" or "always"`);
	}
}
/**
* Field-wise equality over {@link LlmCallConfig} — the comparison a caller
* runs to decide whether a proposed configuration is a real change (worth a
* logged header snapshot) or the held one restated.
* @param a - one configuration.
* @param b - the other.
* @returns whether every field (including the `stop` list, element-wise) matches.
*/
function callConfigEquals(a, b) {
	if (a.provider !== b.provider || a.model !== b.model || a.reasoningEffort !== b.reasoningEffort || a.temperature !== b.temperature || a.maxTokens !== b.maxTokens) return false;
	if (a.stop === void 0 || b.stop === void 0) return a.stop === b.stop;
	return a.stop.length === b.stop.length && a.stop.every((s, i) => s === b.stop?.[i]);
}
/**
* Normalization for values thrown by a final LLM adapter boundary.
*
* @module @deepseek-ai/dsh-llm/adapter-failure
*/
/**
* Detach serializable provider facts from a value thrown by an adapter.
* @param value - arbitrary value thrown during adapter dispatch or iteration.
* @returns immutable provider-neutral facts suitable for a terminal finish chunk.
* @internal
*/
function normalizeLlmFailure(value) {
	const error = value instanceof Error ? value : new HarnessError(thrownMessage(value), "UNKNOWN", { cause: value });
	const carried = ownFailureSnapshot(error);
	if (carried !== void 0 && carried.code === ownErrorCode(error)) return carried;
	return Object.freeze({
		message: errorMessage$1(error),
		code: harnessErrorCode(error)
	});
}
/** Render a non-Error throw without letting hostile coercion escape normalization. */
function thrownMessage(value) {
	try {
		const message = String(value);
		return message.length > 0 ? message : "LLM adapter failed";
	} catch (_hostileThrownValue) {
		return "LLM adapter failed";
	}
}
/** Read a foreign error's own data-backed `code` without invoking accessors. */
function ownErrorCode(error) {
	try {
		const descriptor = Object.getOwnPropertyDescriptor(error, "code");
		return descriptor !== void 0 && "value" in descriptor ? descriptor.value : void 0;
	} catch (_sdkPropertyTrap) {
		return;
	}
}
/** Snapshot an own data property without invoking an SDK-defined accessor. */
function ownFailureSnapshot(error) {
	try {
		const descriptor = Object.getOwnPropertyDescriptor(error, "failure");
		return descriptor !== void 0 && "value" in descriptor ? failureSnapshot(descriptor.value) : void 0;
	} catch (_sdkPropertyTrap) {
		return;
	}
}
/** Validate and detach an arbitrary serializable failure payload. */
function failureSnapshot(value) {
	if (typeof value !== "object" || value === null) return void 0;
	try {
		const candidate = value;
		const message = candidate.message;
		const code = candidate.code;
		const status = candidate.status;
		const providerRetryAfterMs = candidate.providerRetryAfterMs;
		const requestId = candidate.requestId;
		const offloadImages = candidate.offloadImages;
		if (typeof message !== "string" || message.length === 0 || typeof code !== "string" || code.length === 0 || status !== void 0 && (!Number.isInteger(status) || status < 100 || status > 599) || providerRetryAfterMs !== void 0 && (!Number.isFinite(providerRetryAfterMs) || providerRetryAfterMs <= 0) || requestId !== void 0 && (typeof requestId !== "string" || requestId.length === 0) || offloadImages !== void 0 && (!Number.isSafeInteger(offloadImages) || offloadImages <= 0)) return void 0;
		return Object.freeze({
			message,
			code,
			...status === void 0 ? {} : { status },
			...providerRetryAfterMs === void 0 ? {} : { providerRetryAfterMs },
			...requestId === void 0 ? {} : { requestId },
			...offloadImages === void 0 ? {} : { offloadImages }
		});
	} catch (_sdkFailureGetter) {
		return;
	}
}
/** Read an SDK error message without letting an accessor replace the primary failure. */
function errorMessage$1(error) {
	try {
		const message = error.message;
		if (typeof message === "string" && message.length > 0) return message;
	} catch (_sdkMessageGetter) {}
	return "LLM adapter failed";
}
/** Trust only Harness-owned codes; third-party SDK codes are not our taxonomy. */
function harnessErrorCode(error) {
	return error instanceof HarnessError ? error.code : "UNKNOWN";
}
function quoted(value) {
	return JSON.stringify(value);
}
/**
* Stable text shown to a model that cannot accept one durable image reference.
* @param ref - durable normalized attachment omitted from the request.
* @returns deterministic text-only placeholder.
*/
function textOnlyImageText(ref) {
	return `[image omitted because this model accepts text only; attachment sha256:${String(ref.attachmentId).slice(7, 15)}]`;
}
/**
* True when typed model content contains an image block. This is the one image
* walk shared by every image policy (capability gating, text-only
* serialization, compaction survey), so a consumer cannot silently diverge.
* @param content - typed model content blocks.
* @returns whether any block is an image.
*/
function contentHasImage(content) {
	return content.some((block) => block.type === "image");
}
/**
* True when typed model content contains a file block.
* Reads current content on every call without retaining scan results.
* @param content - typed model content blocks.
* @returns whether any block is a file.
*/
function contentHasFile(content) {
	for (const block of content) if (block.type === "file") return true;
	return false;
}
/**
* Stable model-facing handle for one durable file reference: the address of
* the verbatim stored copy and the instruction to read it on demand. This is
* the only representation a provider ever receives for a file.
* @param ref - durable verbatim file reference.
* @param readonlyPath - execution-world path of the stored copy, when resolvable.
* @returns deterministic handle text naming the file, its size, and its address.
*/
function fileHandleText(ref, readonlyPath) {
	const digest = String(ref.attachmentId).slice(7, 15);
	const identity = `File ${quoted(ref.name)} (${ref.bytes} bytes, sha256:${digest})`;
	if (readonlyPath === void 0) return `[${identity} was uploaded, but the current execution environment cannot access a readable path. Report that limitation if its contents are needed; do not claim to have read it.]`;
	return `[${identity}: verbatim read-only copy saved at ${quoted(readonlyPath)}. Read that path with your file tools when its contents are needed; copy it to a writable location before modifying it. When delegating file work, include this saved path in the delegation prompt; only subagents sharing this execution environment can read it.]`;
}
/** Replace every file occurrence with handle text. */
function replaceFilesWithHandles(blocks, resolvePath) {
	let next;
	for (const [index, block] of blocks.entries()) {
		if (block.type === "file") {
			next ??= blocks.slice(0, index);
			next.push({
				type: "text",
				text: fileHandleText(block.attachment, resolvePath(block.attachment))
			});
			continue;
		}
		next?.push(block);
	}
	return next ?? blocks;
}
function projectFilesToText(messages, resolvePath) {
	if (!messages.some((message) => contentHasFile(message.content))) return messages;
	return messages.map((message) => {
		const content = replaceFilesWithHandles(message.content, resolvePath);
		return content === message.content ? message : {
			...message,
			content
		};
	});
}
/** Replace every image occurrence for a text-only model. */
function replaceImagesForTextModel(blocks) {
	let next;
	for (const [index, block] of blocks.entries()) {
		if (block.type === "image") {
			next ??= blocks.slice(0, index);
			next.push({
				type: "text",
				text: textOnlyImageText(block.attachment)
			});
			continue;
		}
		next?.push(block);
	}
	return next ?? blocks;
}
function projectImagesForTextModel(messages) {
	if (!messages.some((message) => contentHasImage(message.content))) return messages;
	return messages.map((message) => {
		const content = replaceImagesForTextModel(message.content);
		return content === message.content ? message : {
			...message,
			content
		};
	});
}
function withoutDeveloperMessages(messages) {
	const retained = messages.filter((message) => message.role !== "developer");
	return retained.length === messages.length ? messages : retained;
}
function toolDeclarations(tools, mode, history) {
	const declarations = new Map(history.tools.map((tool) => [tool.name, tool]));
	for (const update of history.updates) for (const tool of update.additions) if (!declarations.has(tool.name)) declarations.set(tool.name, {
		...tool,
		deferLoading: true
	});
	switch (mode) {
		case "in-history": return declarations;
		case "addition-only": {
			const activeNames = new Set(tools?.map((tool) => tool.name));
			for (const name of declarations.keys()) if (!activeNames.has(name)) declarations.delete(name);
			return declarations;
		}
		/* v8 ignore next 2 -- closed-union exhaustiveness guard */
		default: return assertNever(mode);
	}
}
/**
* Construct provider declarations from session-folded history without changing logged active tools.
* Unsupported routes and incomplete history use current declarations without developer updates.
* Explicitly deferred baseline tools become available only after their first retained addition.
* @param messages - complete request inputs, or the prefix selected for an auxiliary call.
* @param tools - currently active tool schemas.
* @param toolUpdate - the resolved route's update mode.
* @param history - immutable state folded from committed headers and developer messages.
* @returns provider declarations and the corresponding filtered history.
*/
function projectToolUpdates(messages, tools, toolUpdate, history) {
	if (toolUpdate === void 0) {
		let immediateTools = tools;
		if (tools?.some((tool) => tool.deferLoading === true)) immediateTools = tools.map(({ deferLoading: _loading, ...tool }) => tool);
		return {
			messages: withoutDeveloperMessages(messages),
			tools: immediateTools
		};
	}
	if (history === void 0) return {
		messages: withoutDeveloperMessages(messages),
		tools
	};
	const messageIds = new Set(messages.flatMap((message) => message.role === "developer" ? [message.id] : []));
	if (history.updates.some((update) => !messageIds.has(update.messageId))) return {
		messages: withoutDeveloperMessages(messages),
		tools
	};
	const declarations = toolDeclarations(tools, toolUpdate, history);
	const updateIds = new Set(history.updates.map((update) => update.messageId));
	const offered = new Set(history.tools.filter((tool) => !tool.deferLoading).map((tool) => tool.name));
	const projectedMessages = [];
	for (const message of messages) {
		if (message.role !== "developer") {
			projectedMessages.push(message);
			continue;
		}
		if (!updateIds.has(message.id)) continue;
		const content = message.content.filter((block) => {
			switch (block.type) {
				case "tool-addition":
					if (!declarations.has(block.toolName) || offered.has(block.toolName)) return false;
					offered.add(block.toolName);
					return true;
				case "tool-removal":
					if (toolUpdate !== "in-history") return false;
					return offered.delete(block.toolName);
				default: return true;
			}
		});
		if (content.length === 0) continue;
		if (content.length === message.content.length) projectedMessages.push(message);
		else projectedMessages.push({
			...message,
			content
		});
	}
	return {
		messages: projectedMessages.length === messages.length && projectedMessages.every((message, index) => message === messages[index]) ? messages : projectedMessages,
		tools: [...declarations.values()]
	};
}
/**
* Centralize the non-secret product identity every provider request sends as `User-Agent`, keeping
* adapters from drifting. See
* `.agents/notes/implemented/architecture/2026-06-21-mandatory-app-attribution-headers.md`.
*
* App-attribution vocabulary for provider requests.
* @module @deepseek-ai/dsh-llm/attribution
*/
const { version: version$1 } = createRequire(import.meta.url)("../package.json");
/**
* LLM service: adapter registry with a waterfall-interceptable streaming call
* API. Exports the `LlmRuntime` default, the abstract `LlmAdapter` for
* provider backends, and `BlockAssembler` for chunk assembly.
*
* @module @deepseek-ai/dsh-llm
*/
var __runInitializers = function(thisArg, initializers, value) {
	var useValue = arguments.length > 2;
	for (var i = 0; i < initializers.length; i++) value = useValue ? initializers[i].call(thisArg, value) : initializers[i].call(thisArg);
	return useValue ? value : void 0;
};
var __esDecorate = function(ctor, descriptorIn, decorators, contextIn, initializers, extraInitializers) {
	function accept(f) {
		if (f !== void 0 && typeof f !== "function") throw new TypeError("Function expected");
		return f;
	}
	var kind = contextIn.kind, key = kind === "getter" ? "get" : kind === "setter" ? "set" : "value";
	var target = !descriptorIn && ctor ? contextIn["static"] ? ctor : ctor.prototype : null;
	var descriptor = descriptorIn || (target ? Object.getOwnPropertyDescriptor(target, contextIn.name) : {});
	var _, done = false;
	for (var i = decorators.length - 1; i >= 0; i--) {
		var context = {};
		for (var p in contextIn) context[p] = p === "access" ? {} : contextIn[p];
		for (var p in contextIn.access) context.access[p] = contextIn.access[p];
		context.addInitializer = function(f) {
			if (done) throw new TypeError("Cannot add initializers after decoration has completed");
			extraInitializers.push(accept(f || null));
		};
		var result = (0, decorators[i])(kind === "accessor" ? {
			get: descriptor.get,
			set: descriptor.set
		} : descriptor[key], context);
		if (kind === "accessor") {
			if (result === void 0) continue;
			if (result === null || typeof result !== "object") throw new TypeError("Object expected");
			if (_ = accept(result.get)) descriptor.get = _;
			if (_ = accept(result.set)) descriptor.set = _;
			if (_ = accept(result.init)) initializers.unshift(_);
		} else if (_ = accept(result)) if (kind === "field") initializers.unshift(_);
		else descriptor[key] = _;
	}
	if (target) Object.defineProperty(target, contextIn.name, descriptor);
	done = true;
};
/**
* Typed error for LLM-related failures. Extends {@link HarnessError}, so the
* `code` string (e.g. `AUTH`, `RATE_LIMIT`, `NO_ADAPTER`) is shared taxonomy.
*/
var LlmError = class extends HarnessError {
	/** Serializable facts retained beside this live Error. */
	failure;
	/**
	* @param message - non-empty human-readable failure summary.
	* @param code - non-empty stable provider-neutral machine code.
	* @param options - optional cause and validated serializable provider facts.
	*/
	constructor(message, code, options) {
		if (typeof message !== "string" || message.length === 0) throw new Error("LlmError message must be a non-empty string");
		if (typeof code !== "string" || code.length === 0) throw new Error("LlmError code must be a non-empty string");
		if (options?.status !== void 0 && (!Number.isInteger(options.status) || options.status < 100 || options.status > 599)) throw new Error("LlmError status must be an integer from 100 through 599");
		if (options?.providerRetryAfterMs !== void 0 && (!Number.isFinite(options.providerRetryAfterMs) || options.providerRetryAfterMs <= 0)) throw new Error("LlmError providerRetryAfterMs must be a positive finite number");
		if (options?.requestId !== void 0 && (typeof options.requestId !== "string" || options.requestId.length === 0)) throw new Error("LlmError requestId must be a non-empty string");
		super(message, code, options);
		this.name = "LlmError";
		this.failure = Object.freeze({
			message,
			code,
			...options?.status === void 0 ? {} : { status: options.status },
			...options?.providerRetryAfterMs === void 0 ? {} : { providerRetryAfterMs: options.providerRetryAfterMs },
			...options?.requestId === void 0 ? {} : { requestId: options.requestId },
			...options?.offloadImages === void 0 ? {} : { offloadImages: options.offloadImages }
		});
	}
};
(() => {
	let _classSuper = TypertRemoteService;
	let _instanceExtraInitializers = [];
	let _listProviders_decorators;
	let _listConfigurableProviders_decorators;
	let _remoteDiscoverModels_decorators;
	return class LlmRuntime extends _classSuper {
		static {
			const _metadata = typeof Symbol === "function" && Symbol.metadata ? Object.create(_classSuper[Symbol.metadata] ?? null) : void 0;
			_listProviders_decorators = [Remote];
			_listConfigurableProviders_decorators = [Remote];
			_remoteDiscoverModels_decorators = [Remote("discoverModels")];
			__esDecorate(this, null, _listProviders_decorators, {
				kind: "method",
				name: "listProviders",
				static: false,
				private: false,
				access: {
					has: (obj) => "listProviders" in obj,
					get: (obj) => obj.listProviders
				},
				metadata: _metadata
			}, null, _instanceExtraInitializers);
			__esDecorate(this, null, _listConfigurableProviders_decorators, {
				kind: "method",
				name: "listConfigurableProviders",
				static: false,
				private: false,
				access: {
					has: (obj) => "listConfigurableProviders" in obj,
					get: (obj) => obj.listConfigurableProviders
				},
				metadata: _metadata
			}, null, _instanceExtraInitializers);
			__esDecorate(this, null, _remoteDiscoverModels_decorators, {
				kind: "method",
				name: "remoteDiscoverModels",
				static: false,
				private: false,
				access: {
					has: (obj) => "remoteDiscoverModels" in obj,
					get: (obj) => obj.remoteDiscoverModels
				},
				metadata: _metadata
			}, null, _instanceExtraInitializers);
			if (_metadata) Object.defineProperty(this, Symbol.metadata, {
				enumerable: true,
				configurable: true,
				writable: true,
				value: _metadata
			});
		}
		adapters = (__runInitializers(this, _instanceExtraInitializers), /* @__PURE__ */ new Map());
		directory = /* @__PURE__ */ new Map();
		discoveries = /* @__PURE__ */ new Map();
		constructor(ctx) {
			super(ctx, "llm");
		}
		/** Notify topology observers without letting one broken listener veto the commit. */
		emitAdaptersUpdated() {
			let invariantFailure;
			for (const listener of this.ctx.events.dispatch("emit", ["llm/adapters-updated"])) try {
				const returned = listener();
				if (returned != null && typeof returned.then === "function") Promise.resolve(returned).then(void 0, (error) => {
					this.warnAdaptersListenerFailure(error);
				});
			} catch (error) {
				if (error?.code === "INVARIANT") {
					invariantFailure ??= error;
					continue;
				}
				this.warnAdaptersListenerFailure(error);
			}
			if (invariantFailure !== void 0) throw invariantFailure;
		}
		/** Contained-listener diagnostic shared by the sync and async failure paths. */
		warnAdaptersListenerFailure(error) {
			this.ctx.logger.warn("llm: an llm/adapters-updated listener failed");
			this.ctx.logger.warn(error);
		}
		/**
		* Register an adapter for the given provider routes. Throws `LlmError` with code
		* `DUPLICATE_ADAPTER` if any provider already has an adapter (all-or-nothing).
		* Disposed with the fiber.
		* @param providers - every provider route this adapter should serve.
		* @param adapter - the adapter that streams calls for those providers.
		* @returns the disposer, carrying {@link AdapterRegistrationHandle.replace}.
		*/
		registerAdapter(providers, adapter) {
			const owned = /* @__PURE__ */ new Set();
			let released = false;
			const dispose = this.ctx.effect(function* () {
				if (providers.length === 0) throw new LlmError("an adapter must register at least one provider", "INVALID_ADAPTER");
				this.commitRoutes(owned, this.prepareRoutes(providers, adapter, owned));
				yield () => {
					released = true;
					for (const provider of owned) this.adapters.delete(provider);
					owned.clear();
					this.emitAdaptersUpdated();
				};
			}.bind(this), "llm.registerAdapter()");
			const handle = (() => void dispose());
			handle.replace = (next) => {
				if (released) throw new LlmError("a disposed adapter registration cannot replace its routes", "REGISTRATION_DISPOSED");
				this.commitRoutes(owned, this.prepareRoutes(next, adapter, owned));
			};
			return handle;
		}
		/**
		* Validate one candidate route set for `adapter`, treating routes this
		* registration already holds as available. Nothing is mutated: a rejected
		* candidate leaves the registry exactly as it was.
		*/
		prepareRoutes(providers, adapter, owned) {
			const unique = /* @__PURE__ */ new Set();
			const registrations = [];
			for (const provider of providers) {
				if (provider.length === 0) throw new LlmError("adapter provider names must be non-empty", "INVALID_ADAPTER");
				if (unique.has(provider) || this.adapters.has(provider) && !owned.has(provider)) throw new LlmError(`an adapter for provider "${provider}" is already registered`, "DUPLICATE_ADAPTER");
				const info = adapter.providerInfo(provider);
				if (typeof info.id !== "string" || info.id !== provider || typeof info.name !== "string" || info.name.length === 0) throw new LlmError(`adapter metadata for provider "${provider}" must preserve its id and have a non-empty name`, "INVALID_ADAPTER");
				unique.add(provider);
				const retryPolicy = adapter.providerRetryPolicy(provider) ?? resolveRetryPolicy(void 0, `llm: provider "${provider}" retryPolicy`);
				registrations.push({
					adapter,
					provider: {
						id: info.id,
						name: info.name
					},
					retryPolicy
				});
			}
			return registrations;
		}
		/**
		* Swap this registration's routes for the prepared ones in one synchronous
		* section, so no observer can see the registry between the release and the
		* re-registration. The route set's one mutation point is also where
		* `llm/adapters-updated` is published, so a `replace` announces itself
		* exactly like a first registration.
		*/
		commitRoutes(owned, registrations) {
			for (const provider of owned) this.adapters.delete(provider);
			owned.clear();
			for (const registration of registrations) {
				this.adapters.set(registration.provider.id, registration);
				owned.add(registration.provider.id);
			}
			this.emitAdaptersUpdated();
		}
		/**
		* Describe provider routes with a registered adapter.
		* @returns detached provider metadata in registration order.
		*/
		listProviders() {
			return [...this.adapters.values()].map(({ provider }) => ({ ...provider }));
		}
		/**
		* Declare provider routes an adapter plugin can activate through
		* configuration. Registration is all-or-nothing: an empty list, invalid
		* entry, or a provider already declared by any registration throws
		* `LlmError` without registering the rest. Disposed with the fiber.
		* @param entries - every configurable provider this plugin owns.
		* @returns a handle that withdraws all of them, and can atomically replace them.
		*/
		registerConfigurableProviders(entries) {
			let held = [];
			let disposed = false;
			/**
			* Validate a candidate set in full against everything this registration
			* does not already hold, then publish it. Nothing is written until the
			* whole set passes, so a refused candidate leaves the current entries in
			* place — the property that makes `replace` a swap rather than a
			* delete-then-add that can strand the directory empty.
			*/
			const commit = (candidates) => {
				const detached = [];
				const own = new Set(held.map((entry) => entry.provider));
				for (const entry of candidates) {
					if (entry.provider.length === 0 || entry.displayName.length === 0 || entry.settingsNs.length === 0) throw new LlmError("configurable providers need a non-empty provider, displayName, and settingsNs", "INVALID_DIRECTORY");
					if (entry.settingsPath.some((segment) => segment.length === 0)) throw new LlmError(`configurable provider "${entry.provider}" has an empty settingsPath segment`, "INVALID_DIRECTORY");
					if (this.directory.has(entry.provider) && !own.has(entry.provider) || detached.some((seen) => seen.provider === entry.provider)) throw new LlmError(`configurable provider "${entry.provider}" is already declared`, "DUPLICATE_DIRECTORY");
					detached.push({
						...entry,
						settingsPath: [...entry.settingsPath]
					});
				}
				for (const entry of held) this.directory.delete(entry.provider);
				for (const entry of detached) this.directory.set(entry.provider, entry);
				held = detached;
				this.emitAdaptersUpdated();
			};
			const dispose = this.ctx.effect(function* () {
				if (entries.length === 0) throw new LlmError("a configurable-provider registration must declare at least one provider", "INVALID_DIRECTORY");
				commit(entries);
				yield () => {
					disposed = true;
					for (const entry of held) this.directory.delete(entry.provider);
					held = [];
					this.emitAdaptersUpdated();
				};
			}.bind(this), "llm.registerConfigurableProviders()");
			const handle = (() => void dispose());
			handle.replace = (next) => {
				if (disposed) throw new LlmError("this configurable-provider registration was disposed", "REGISTRATION_DISPOSED");
				commit(next);
			};
			return handle;
		}
		/**
		* List every declared configurable provider, registered or dormant.
		* @returns detached directory entries in declaration order.
		*/
		listConfigurableProviders() {
			return [...this.directory.values()].map((entry) => ({
				...entry,
				settingsPath: [...entry.settingsPath]
			}));
		}
		/**
		* Offer to interrogate provider endpoints on behalf of the settings
		* namespace this plugin owns. The namespace is the key because that is what
		* a configuration surface already holds from the configurable-provider
		* directory, and because a provider being *added* has no route to name yet.
		* Disposed with the fiber.
		* @param settingsNs - the namespace whose profiles this discovery serves.
		* @param discover - interrogates one endpoint and must honor the supplied signal.
		* @returns the disposer that withdraws the offer.
		*/
		registerModelDiscovery(settingsNs, discover) {
			const dispose = this.ctx.effect(function* () {
				if (settingsNs.length === 0) throw new LlmError("model discovery needs a non-empty settings namespace", "INVALID_DISCOVERY");
				if (this.discoveries.has(settingsNs)) throw new LlmError(`model discovery for "${settingsNs}" is already registered`, "DUPLICATE_DISCOVERY");
				this.discoveries.set(settingsNs, discover);
				yield () => {
					this.discoveries.delete(settingsNs);
				};
			}.bind(this), "llm.registerModelDiscovery()");
			return () => void dispose();
		}
		/**
		* Interrogate one provider endpoint for the models it advertises. The
		* request describes a draft, not a stored route, so nothing here reads or
		* writes settings or credentials — the caller owns both, and the reply is
		* candidate metadata a surface may offer for adoption.
		* @param settingsNs - namespace whose registered discovery serves this draft.
		* @param request - the endpoint, protocol, and one-shot credential to use.
		* @param signal - caller cancellation.
		* @returns the advertised models, deduplicated in endpoint order.
		*/
		async discoverModels(settingsNs, request, signal) {
			const discover = this.discoveries.get(settingsNs);
			if (discover === void 0) throw new LlmError(`no model discovery is registered for "${settingsNs}"`, "NO_DISCOVERY");
			if ((request.provider ?? "").length === 0 && (request.baseURL ?? "").length === 0) throw new LlmError("model discovery needs a provider route or a baseURL", "INVALID_DISCOVERY");
			const discovered = signal === void 0 ? await discover(request) : await discover(request, signal);
			const seen = /* @__PURE__ */ new Set();
			const models = [];
			for (const model of discovered) {
				if (typeof model.id !== "string" || model.id.length === 0 || seen.has(model.id)) continue;
				seen.add(model.id);
				models.push({
					id: model.id,
					...model.name === void 0 ? {} : { name: model.name },
					...model.contextWindow === void 0 ? {} : { contextWindow: model.contextWindow },
					...model.maxTokens === void 0 ? {} : { maxTokens: model.maxTokens },
					...model.inputModalities === void 0 ? {} : { inputModalities: [...model.inputModalities] }
				});
			}
			return models;
		}
		/**
		* Remote adapter for one draft provider interrogation.
		* @param settingsNs - namespace whose registered discovery serves this draft.
		* @param request - endpoint, protocol, and one-shot credential to use.
		* @param signal - caller cancellation supplied by the Remote carrier.
		* @returns advertised models in endpoint order.
		* @throws RemoteError with `llm/model-discovery-rejected` when discovery refuses or fails.
		*/
		async remoteDiscoverModels(settingsNs, request, signal) {
			try {
				return await this.discoverModels(settingsNs, request, signal);
			} catch (error) {
				throw new RemoteError("llm/model-discovery-rejected", error instanceof Error ? error.message : String(error), {
					settingsNs,
					...request.baseURL === void 0 ? {} : { baseURL: request.baseURL }
				}, { cause: error });
			}
		}
		/**
		* Resolve the retry policy captured when one provider route was registered.
		* @param provider - registered provider route to inspect.
		* @returns the provider-owned policy, with normal defaults already resolved.
		*/
		providerRetryPolicy(provider) {
			return this.registration(provider).retryPolicy;
		}
		/**
		* Resolve provider-side request-image pricing for one exact route, or
		* `undefined` when the provider is unregistered or declares none. Unknown
		* providers degrade to `undefined` rather than throwing because callers
		* price durable history whose route may no longer be mounted.
		* @param provider - provider route named by a request header.
		* @param model - exact model id named by the same header.
		* @returns the owning adapter's image pricing for the route, when declared.
		*/
		imageRequestPricing(provider, model) {
			return this.adapters.get(provider)?.adapter.imageRequestPricing(provider, model);
		}
		/**
		* Resolve the exact text one durable file occurrence contributes to every
		* provider request in the current execution environment.
		* @param ref - durable verbatim file reference from model history.
		* @returns the same deterministic handle text used at adapter dispatch.
		*/
		fileRequestText(ref) {
			return fileHandleText(ref, this.fileReadPath(ref));
		}
		/** Detach typed adapter-owned modality metadata. */
		detachedModalities(modalities) {
			return modalities === void 0 ? void 0 : [...modalities];
		}
		/**
		* Discover models advertised by one registered provider. Catalog membership
		* does not constrain core routing. Catalog-driven entry points may restrict
		* selection and submission to the advertised models.
		* @param provider - registered provider route to inspect.
		* @returns detached model metadata in adapter-preferred order.
		*/
		async listModels(provider) {
			const models = await this.registration(provider).adapter.listModels(provider);
			const seen = /* @__PURE__ */ new Set();
			return models.map((model) => {
				if (typeof model.provider !== "string" || model.provider !== provider || typeof model.id !== "string" || model.id.length === 0 || typeof model.name !== "string" || model.name.length === 0 || model.description !== void 0 && typeof model.description !== "string" || seen.has(model.id)) throw new LlmError(`adapter returned invalid or duplicate model metadata for provider "${provider}"`, "INVALID_CATALOG");
				seen.add(model.id);
				const inputModalities = this.detachedModalities(model.inputModalities);
				return {
					provider: model.provider,
					id: model.id,
					name: model.name,
					...model.description === void 0 ? {} : { description: model.description },
					...inputModalities === void 0 ? {} : { inputModalities }
				};
			});
		}
		/**
		* Resolve and validate all metadata from the adapter that owns one exact
		* route. The result is detached from adapter-owned objects; catalog
		* membership remains advisory and does not control request routing.
		* @param provider - registered provider route to inspect.
		* @param model - exact model id passed to the adapter.
		* @param signal - optional cancellation for adapter-owned asynchronous lookup.
		* @returns exact model identity plus available context and reasoning metadata.
		*/
		async resolveModelInfo(provider, model, signal) {
			return this.resolveModelInfoFor(this.registration(provider), model, signal);
		}
		async resolveModelInfoFor(registration, model, signal) {
			const resolved = await registration.adapter.resolveModel(registration.provider.id, model, signal);
			return this.normalizeModelInfo(registration, model, resolved);
		}
		/** Validate and detach one adapter-returned exact model result. */
		normalizeModelInfo(registration, model, resolved) {
			const provider = registration.provider.id;
			if (typeof resolved.provider !== "string" || resolved.provider !== provider || typeof resolved.id !== "string" || resolved.id !== model || typeof resolved.name !== "string" || resolved.name.length === 0 || resolved.description !== void 0 && typeof resolved.description !== "string") throw new LlmError(`adapter returned invalid exact model metadata for provider "${provider}" model "${model}"`, "INVALID_MODEL_INFO");
			const context = resolved.context;
			if (context !== void 0 && (!Number.isInteger(context.contextWindow) || context.contextWindow <= 0)) throw new LlmError(`adapter returned invalid context metadata for provider "${provider}" model "${model}"`, "INVALID_MODEL_CONTEXT");
			const inputModalities = this.detachedModalities(resolved.inputModalities);
			const systemPromptUpdate = resolved.systemPromptUpdate;
			if (systemPromptUpdate !== void 0 && systemPromptUpdate !== "in-history") throw new LlmError(`adapter returned invalid system prompt update mode for provider "${provider}" model "${model}"`, "INVALID_MODEL_INFO");
			const toolUpdate = resolved.toolUpdate;
			if (toolUpdate !== void 0 && toolUpdate !== "in-history" && toolUpdate !== "addition-only") throw new LlmError(`adapter returned invalid tool update mode for provider "${provider}" model "${model}"`, "INVALID_MODEL_INFO");
			const defaultMaxTokens = resolved.defaultMaxTokens;
			if (defaultMaxTokens !== void 0 && (!Number.isSafeInteger(defaultMaxTokens) || defaultMaxTokens <= 0)) throw new LlmError(`adapter returned invalid default maxTokens for provider "${provider}" model "${model}"`, "INVALID_MODEL_MAX_TOKENS");
			const info = {
				provider,
				id: model,
				name: resolved.name,
				...resolved.description === void 0 ? {} : { description: resolved.description },
				...inputModalities === void 0 ? {} : { inputModalities },
				...context === void 0 ? {} : { context: { contextWindow: context.contextWindow } },
				...defaultMaxTokens === void 0 ? {} : { defaultMaxTokens },
				...resolved.systemPromptUpdate === void 0 ? {} : { systemPromptUpdate: resolved.systemPromptUpdate },
				...resolved.toolUpdate === void 0 ? {} : { toolUpdate: resolved.toolUpdate }
			};
			const reasoning = resolved.reasoning;
			if (reasoning === void 0) return info;
			if (reasoning.efforts.length === 0) throw new LlmError(`adapter returned invalid reasoning metadata for provider "${provider}" model "${model}"`, "INVALID_MODEL_REASONING");
			const seen = /* @__PURE__ */ new Set();
			const efforts = reasoning.efforts.map((effort) => {
				if (typeof effort.id !== "string" || effort.id.length === 0 || typeof effort.name !== "string" || effort.name.length === 0 || effort.description !== void 0 && typeof effort.description !== "string" || seen.has(effort.id)) throw new LlmError(`adapter returned invalid or duplicate reasoning effort metadata for provider "${provider}" model "${model}"`, "INVALID_MODEL_REASONING");
				seen.add(effort.id);
				return {
					id: effort.id,
					name: effort.name,
					...effort.description === void 0 ? {} : { description: effort.description }
				};
			});
			if (reasoning.defaultEffort !== void 0 && !seen.has(reasoning.defaultEffort)) throw new LlmError(`adapter returned an unknown default reasoning effort for provider "${provider}" model "${model}"`, "INVALID_MODEL_REASONING");
			return {
				...info,
				reasoning: {
					efforts,
					...reasoning.defaultEffort === void 0 ? {} : { defaultEffort: reasoning.defaultEffort }
				}
			};
		}
		/**
		* Validate a conversation call config against its exact model capability and
		* materialize adapter-configured defaults. Unsupported explicit efforts
		* reject before provider I/O; no clamping or aliasing is performed. This
		* standalone query does not bind a later dispatch; use {@link prepareCall}
		* when logging and streaming must share one adapter registration.
		* @param config - provider/model route and optional request controls.
		* @param signal - optional cancellation for adapter-owned capability lookup.
		* @returns a detached config only when a default must be materialized.
		*/
		async resolveCallConfig(config, signal) {
			return (await this.resolveCallFor(this.registration(config.provider), config, signal)).config;
		}
		async resolveCallFor(registration, config, signal) {
			const info = await this.resolveModelInfoFor(registration, config.model, signal);
			return this.resolveCallWithInfo(config, info);
		}
		/** Validate request controls against one already-bound exact model result. */
		resolveCallWithInfo(config, info) {
			const defaulted = config.maxTokens === void 0 && info.defaultMaxTokens !== void 0 ? {
				...config,
				maxTokens: info.defaultMaxTokens
			} : config;
			const reasoning = info.reasoning;
			const requested = defaulted.reasoningEffort;
			let resolvedConfig = defaulted;
			if (reasoning === void 0) {
				if (requested !== void 0) throw new LlmError(`provider "${config.provider}" model "${config.model}" does not support reasoning effort "${requested}"`, "UNSUPPORTED_REASONING_EFFORT");
			} else {
				const effective = requested ?? reasoning.defaultEffort;
				if (effective !== void 0) {
					if (!reasoning.efforts.some((effort) => effort.id === effective)) throw new LlmError(`provider "${config.provider}" model "${config.model}" does not support reasoning effort "${effective}"`, "UNSUPPORTED_REASONING_EFFORT");
					if (requested !== effective) resolvedConfig = {
						...defaulted,
						reasoningEffort: effective
					};
				}
			}
			return {
				config: resolvedConfig,
				...info.context === void 0 ? {} : { context: info.context },
				modelInfo: info
			};
		}
		/**
		* Resolve one call under its current adapter registration. The returned
		* one-shot handle keeps that registration across header logging and dispatch,
		* so HMR cannot combine one adapter's capability result with another adapter.
		* @param config - provider/model route and optional request controls.
		* @param signal - optional cancellation for adapter-owned capability lookup.
		* @returns a prepared config and its registration-bound stream entry point.
		*/
		async prepareCall(config, signal) {
			const registration = this.registration(config.provider);
			const adapterCall = await registration.adapter.prepareCall(config.provider, config.model, signal);
			const modelInfo = this.normalizeModelInfo(registration, config.model, adapterCall.model);
			const resolved = this.resolveCallWithInfo(config, modelInfo);
			const resolvedConfig = deepFreeze(structuredClone(resolved.config));
			const context = resolved.context === void 0 ? void 0 : deepFreeze(structuredClone(resolved.context));
			const adapterDefaults = deepFreeze({
				...config.reasoningEffort === void 0 && resolvedConfig.reasoningEffort !== void 0 ? { reasoningEffort: true } : {},
				...config.maxTokens === void 0 && resolvedConfig.maxTokens !== void 0 ? { maxTokens: true } : {}
			});
			let dispatched = false;
			return Object.freeze({
				config: resolvedConfig,
				retryPolicy: registration.retryPolicy,
				adapterDefaults,
				...context === void 0 ? {} : { context },
				...modelInfo.inputModalities === void 0 ? {} : { inputModalities: Object.freeze([...modelInfo.inputModalities]) },
				...modelInfo.systemPromptUpdate === void 0 ? {} : { systemPromptUpdate: modelInfo.systemPromptUpdate },
				...modelInfo.toolUpdate === void 0 ? {} : { toolUpdate: modelInfo.toolUpdate },
				stream: (options) => {
					if (dispatched) throw new LlmError("a prepared LLM call can only be dispatched once", "INVALID_PREPARED_CALL");
					if (!callConfigEquals(options, resolvedConfig)) throw new LlmError("prepared LLM call config changed before adapter dispatch", "INVALID_PREPARED_CALL");
					dispatched = true;
					return this.streamWithRegistration(options, {
						registration,
						config: resolvedConfig,
						modelInfo,
						dispatch: (options) => adapterCall.stream(options)
					});
				}
			});
		}
		registration(provider) {
			const registration = this.adapters.get(provider);
			if (!registration) throw new LlmError(`no adapter registered for provider "${provider}"`, "NO_ADAPTER");
			return registration;
		}
		/** Remove replay state whose historical route is owned by another adapter. */
		forAdapter(options, adapter) {
			const messages = options.messages.map((message) => {
				if (message.role !== "assistant") return message;
				const source = message.source;
				if (source.replayState === void 0) return message;
				if (this.adapters.get(source.provider)?.adapter === adapter) return message;
				return freezeMessage({
					...message,
					source: {
						kind: "model",
						provider: source.provider,
						model: source.model
					}
				});
			});
			if (messages.every((message, index) => message === options.messages[index])) return options;
			const filtered = {
				...options,
				messages
			};
			return Object.isFrozen(options) ? deepFreeze(filtered) : filtered;
		}
		/**
		* Resolve the current execution-world read path of one durable file
		* reference through the mounted attachment and filesystem providers.
		*/
		fileReadPath(ref) {
			let hostPath;
			try {
				hostPath = this.ctx.get("attachments")?.fileHostPath(ref);
			} catch {
				return;
			}
			if (hostPath === void 0) return void 0;
			return this.ctx.get("fs")?.processPathFromHostPath(hostPath);
		}
		/**
		* Final adapter boundary. Adapter selection, dispatch, iterator construction,
		* and iteration failures become one terminal failure chunk. Middleware and
		* downstream consumer failures remain thrown plugin or consumer errors.
		*/
		async *adapterStream(options, prepared) {
			let iterator;
			try {
				const registration = prepared?.registration ?? this.registration(options.provider);
				const adapter = registration.adapter;
				let modelInfo;
				let resolvedConfig;
				let dispatch;
				if (prepared === void 0) {
					const adapterCall = await adapter.prepareCall(options.provider, options.model, options.signal);
					modelInfo = this.normalizeModelInfo(registration, options.model, adapterCall.model);
					resolvedConfig = this.resolveCallWithInfo(options, modelInfo).config;
					dispatch = (options) => adapterCall.stream(options);
				} else {
					modelInfo = prepared.modelInfo;
					resolvedConfig = prepared.config;
					dispatch = prepared.dispatch;
				}
				if (prepared !== void 0 && !callConfigEquals(options, resolvedConfig)) throw new LlmError("prepared LLM call config changed before adapter dispatch", "INVALID_PREPARED_CALL");
				const resolvedOptions = callConfigEquals(options, resolvedConfig) ? options : Object.isFrozen(options) ? deepFreeze({
					...options,
					...resolvedConfig
				}) : {
					...options,
					...resolvedConfig
				};
				let projectedMessages = resolvedOptions.messages;
				if (projectedMessages.some((message) => contentHasFile(message.content))) projectedMessages = projectFilesToText(projectedMessages, (ref) => this.fileReadPath(ref));
				if (modelInfo.inputModalities !== void 0 && !modelInfo.inputModalities.includes("image") && projectedMessages.some((message) => contentHasImage(message.content))) projectedMessages = projectImagesForTextModel(projectedMessages);
				const projectedTools = projectToolUpdates(projectedMessages, resolvedOptions.tools, modelInfo.toolUpdate, resolvedOptions.toolHistory);
				projectedMessages = projectedTools.messages;
				let projectedOptions = resolvedOptions;
				if (projectedMessages !== resolvedOptions.messages || projectedTools.tools !== resolvedOptions.tools) {
					projectedOptions = {
						...resolvedOptions,
						messages: projectedMessages,
						...projectedTools.tools === void 0 ? {} : { tools: projectedTools.tools }
					};
					if (Object.isFrozen(resolvedOptions)) deepFreeze(projectedOptions);
				}
				iterator = dispatch(this.forAdapter(projectedOptions, adapter))[Symbol.asyncIterator]();
			} catch (error) {
				yield adapterFailureChunk(error, options.signal);
				return;
			}
			let completed = false;
			try {
				while (true) {
					let item;
					try {
						const next = await iterator.next();
						item = next.done ? { done: true } : {
							done: false,
							value: next.value
						};
					} catch (error) {
						completed = true;
						yield adapterFailureChunk(error, options.signal);
						return;
					}
					if (item.done) {
						completed = true;
						return;
					}
					yield item.value;
				}
			} finally {
				if (!completed) {
					const close = iterator.return?.bind(iterator);
					if (close) await close();
				}
			}
		}
		/**
		* Stream one model call as raw chunks (token-level deltas). Replay state is
		* retained only when the same adapter instance owns its historical provider
		* and the target provider. Final adapter selection remains fixed through
		* asynchronous exact-model resolution and dispatch. Adapter selection,
		* dispatch, and iteration failures become terminal `error` or `aborted`
		* finish chunks; middleware, nested-call, cleanup, and consumer failures
		* remain thrown.
		* @param options - the full request; `options.provider` selects the adapter.
		* @returns the chunk stream, possibly wrapped by `llm/stream` listeners.
		*/
		stream(options) {
			return this.streamWithRegistration(options);
		}
		streamWithRegistration(options, prepared) {
			return this.ctx.waterfall(this, "llm/stream", options, () => this.adapterStream(options, prepared));
		}
	};
})();
/** Convert one adapter throw into the stream protocol's terminal outcome. */
function adapterFailureChunk(error, signal) {
	const failure = normalizeLlmFailure(error);
	return {
		type: "finish",
		reason: signal?.aborted || failure.code === "ABORTED" ? {
			kind: "aborted",
			failure
		} : {
			kind: "error",
			failure
		}
	};
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-sandbox@0.2.0-rc.2_@deepseek-ai+cordis@4.0.4_@deepseek-ai+dsh-llm@0.2._2cc6ce11d95db65b014013a78bf96eaa/node_modules/@deepseek-ai/dsh-sandbox/lib/index.js
/**
* The escalation vocabulary and choreography shared by every sandbox-enforcing
* tool family (`@deepseek-ai/dsh-tool-bash`, `@deepseek-ai/dsh-tool-fs`): the
* strictly-wider ladder, the argument-pairing validation, the model-facing
* denial/hint markers, and {@link approveEscalation} — the ordered fail-closed
* sequence that resolves a `sandbox_permissions` request through a
* user-approval channel BEFORE anything executes. One home keeps the two
* families' approval ordering and verbatim error texts from drifting apart.
*
* The channel is a minimal STRUCTURAL function shape ({@link EscalationAsk}),
* not the approval service type: the tool layer — which owns the agent, the
* call id, and the tool name — closes over `ctx.approval.request(...)` and
* hands the closure down, so this package never depends on the approval or
* agent packages.
*
* @module dsh-sandbox/escalation
*/
/**
* The strictly-wider table: what a call whose effective mode is the key may
* escalate TO. Checked at EXECUTION, never baked into a tool schema — the
* schema's enum is {@link ESCALATION_TARGETS}, because schemas are
* registry-global while the effective mode is per-call truth.
*/
const WIDER_MODES = {
	"read-only": ["workspace-write", "danger-full-access"],
	"workspace-write": ["danger-full-access"]
};
/**
* The closed escalation-target vocabulary — every mode a call could ever
* escalate TO (`read-only` is the floor; nothing escalates to it). Advertised
* whenever the mounted capability confines: cutting the enum down to the modes
* wider than the composition's DEFAULT would strand a session whose effective
* mode sits below it (a `danger-full-access` default would advertise nothing
* while a narrower-switched session stays confined with no lever).
*/
const ESCALATION_TARGETS = ["workspace-write", "danger-full-access"];
/**
* Validate the escalation argument pairing a tool schema cannot express:
* `sandbox_permissions` and `justification` travel together — an approval
* prompt without a reason, or a reason driving nothing, is a malformed ask —
* and the justification must be a non-empty sentence.
* @param sandboxPermissions - the raw `sandbox_permissions` argument, if given.
* @param justification - the raw `justification` argument, if given.
*/
function validateEscalationArgs(sandboxPermissions, justification) {
	if (sandboxPermissions !== void 0 && justification === void 0) throw new Error("invalid escalation: sandbox_permissions requires a justification");
	if (justification !== void 0 && sandboxPermissions === void 0) throw new Error("invalid escalation: justification is only valid together with sandbox_permissions");
	if (justification !== void 0 && justification.trim().length === 0) throw new Error("invalid justification: expected a non-empty sentence");
}
/**
* Resolve a sandbox permission request before execution. Repeating the call's
* effective mode returns it without approval. A strictly wider mode requires
* approval and applies only to this call. Narrower or unsupported targets,
* missing approval services or agents for widening, and non-grant outcomes
* throw before execution.
* @param request - the escalation to judge (see {@link EscalationRequest}).
* @param approval - the approval ingredients the tool holds (see {@link EscalationApproval}).
* @returns the granted mode, consumed by the one call that asked.
*/
async function approveEscalation(request, approval) {
	const { requestedMode: mode, effectiveMode, justification, subject } = request;
	if (mode === effectiveMode) return effectiveMode;
	if (!(WIDER_MODES[effectiveMode] ?? []).includes(mode)) throw new Error(`sandbox escalation to "${mode}" is not strictly wider than this call's current "${effectiveMode}" mode`);
	if (approval.approver === void 0) throw new Error(`sandbox escalation to "${mode}" requires approval, but no approval service is composed`);
	if (approval.agent === void 0) throw new Error(`sandbox escalation to "${mode}" requires approval, but the call has no agent to route it through`);
	const outcome = await approval.approver.request({
		agent: approval.agent,
		toolName: approval.toolName,
		callId: approval.callId,
		reason: `escalate sandbox to ${mode}: ${justification}`,
		displayReason: {
			en: `Allow this operation with ${mode} permissions: ${justification}`,
			zh: `允许本次操作使用 ${mode} 权限：${justification}`
		},
		...approval.signal ? { signal: approval.signal } : {}
	});
	switch (outcome) {
		case "allowed-once": return mode;
		case "rejected": throw new Error(`the user rejected escalating this ${subject} to "${mode}"; it stays denied, so stop and explain instead of working around it`);
		case "cancelled": throw new Error(`approval for escalating to "${mode}" was cancelled`);
		case "unavailable": throw new Error(`sandbox escalation to "${mode}" requires approval, but no approval channel is available`);
		default: return assertNever(outcome, "EscalationOutcome");
	}
}
//#endregion
//#region node_modules/.pnpm/@deepseek-ai+dsh-tools@0.2.0-rc.2_c6e351a6a6cfd1c2a664558bffcd410d/node_modules/@deepseek-ai/dsh-tools/lib/index.js
/**
* Enforced JSON Schema subset shared by tool outputs, generated PTC mode
* types, subagents, and workflows. The subset accepts any JSON root, an
* annotation-only schema for unconstrained JSON, one scalar `type`, object
* `properties`/`required`/boolean `additionalProperties`, array `items`,
* type-correct scalar `enum`/`const`, and exact-one `oneOf`.
*
* Unsupported or misplaced keywords reject rather than being accepted without
* enforcement. Consumers that require an object root apply
* {@link assertObjectJsonSchema} before accepting input.
* @module dsh-tools/json-schema
*/
/**
* Thrown when a raw schema falls outside the enforced subset. `violations`
* lists every offending path instead of stopping at the first author error.
*/
var JsonSchemaError = class extends HarnessError {
	/** Individual schema violations in walk order. */
	violations;
	constructor(violations) {
		super(`unsupported JSON schema: ${violations.join("; ")}`, "UNSUPPORTED_SCHEMA");
		this.name = "JsonSchemaError";
		this.violations = violations;
	}
};
const CONSTRAINT_KEYWORDS = /* @__PURE__ */ new Set([
	"type",
	"oneOf",
	"properties",
	"required",
	"additionalProperties",
	"items",
	"enum",
	"const"
]);
const ANNOTATION_KEYWORDS = /* @__PURE__ */ new Set([
	"description",
	"title",
	"default",
	"examples"
]);
const SCHEMA_TYPES = [
	"object",
	"array",
	"string",
	"number",
	"integer",
	"boolean",
	"null"
];
/** Whether a realm-owned intrinsic prototype is backed by its native constructor. */
function hasIntrinsicConstructor(prototype, name) {
	const constructor = Object.getOwnPropertyDescriptor(prototype, "constructor")?.value;
	if (typeof constructor !== "function") return false;
	try {
		return constructor.name === name && constructor.prototype === prototype && Function.prototype.toString.call(constructor) === `function ${name}() { [native code] }`;
	} catch {
		return false;
	}
}
/** Whether a candidate is one realm's intrinsic `Object.prototype`. */
function isIntrinsicObjectPrototype(value) {
	return Object.getPrototypeOf(value) === null && hasIntrinsicConstructor(value, "Object");
}
/**
* Test for a realm-agnostic plain JSON record without accepting arrays or
* exotic objects.
* @param value - candidate record from any JavaScript realm.
* @returns Whether the value has a plain-object prototype chain.
*/
function isPlainJsonRecord(value) {
	if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
	try {
		const prototype = Object.getPrototypeOf(value);
		return prototype === null || typeof prototype === "object" && isIntrinsicObjectPrototype(prototype);
	} catch {
		return false;
	}
}
/** Whether an array uses one realm's intrinsic `Array.prototype`. */
function hasPlainArrayPrototype(value) {
	const prototype = Object.getPrototypeOf(value);
	if (!Array.isArray(prototype) || !hasIntrinsicConstructor(prototype, "Array")) return false;
	const objectPrototype = Object.getPrototypeOf(prototype);
	return typeof objectPrototype === "object" && objectPrototype !== null && isIntrinsicObjectPrototype(objectPrototype);
}
/** Return whether a record contains only own enumerable string keys. */
function hasOnlyEnumerableStringKeys(value) {
	try {
		return Reflect.ownKeys(value).every((key) => typeof key === "string" && Object.prototype.propertyIsEnumerable.call(value, key));
	} catch {
		return false;
	}
}
/**
* Test for an ordinary schema record whose keys survive JSON projection.
* @param value - candidate record from any JavaScript realm.
* @returns Whether the record has an intrinsic prototype and only own enumerable string keys.
*/
function isJsonSchemaRecord(value) {
	return isPlainJsonRecord(value) && hasOnlyEnumerableStringKeys(value);
}
/**
* Test for a dense ordinary array with no JSON-invisible decorations.
* @param value - candidate array from any JavaScript realm.
* @returns Whether the array is intrinsic, dense, and undecorated.
*/
function isPlainJsonArray(value) {
	if (!Array.isArray(value)) return false;
	try {
		if (!hasPlainArrayPrototype(value) || Reflect.ownKeys(value).length !== value.length + 1) return false;
		for (let index = 0; index < value.length; index++) if (!Object.hasOwn(value, index)) return false;
		return true;
	} catch {
		return false;
	}
}
/** Lossless finite JSON number, excluding negative zero. */
function isJsonNumber(value) {
	return typeof value === "number" && Number.isFinite(value) && !Object.is(value, -0);
}
/** Whether a scalar is valid for one declared schema type. */
function scalarMatches(type, value) {
	switch (type) {
		case "string": return typeof value === "string";
		case "number": return isJsonNumber(value);
		case "integer": return isJsonNumber(value) && Number.isInteger(value);
		case "boolean": return typeof value === "boolean";
		case "null": return value === null;
		/* v8 ignore next -- JsonSchemaScalarType is closed; this retains compile-time exhaustiveness. */
		default: return assertNever(type, "JsonSchemaType");
	}
}
/** Keywords that are invalid beside `oneOf`. */
const ONE_OF_SIBLING_KEYWORDS = [
	"properties",
	"required",
	"additionalProperties",
	"items",
	"enum",
	"const"
];
/** Validate object-only fields after its property schemas have been visited. */
function checkObjectSchemaTail(node, path, properties, violations) {
	const hasRequired = Object.hasOwn(node, "required");
	const required = hasRequired ? node.required : void 0;
	if (hasRequired) if (!isPlainJsonArray(required) || required.some((entry) => typeof entry !== "string")) violations.push(`${path}.required must be an array of strings`);
	else {
		const declared = isJsonSchemaRecord(properties) ? properties : {};
		for (const key of required) if (!Object.hasOwn(declared, key)) violations.push(`${path}.required names "${key}" which is not in properties`);
	}
	if (Object.hasOwn(node, "additionalProperties") && typeof node.additionalProperties !== "boolean") violations.push(`${path}.additionalProperties must be a boolean`);
}
/** Collect every violation for one raw schema tree without using the JavaScript call stack. */
function checkSchemaNode(root, rootPath, violations, seen) {
	const tasks = [{
		kind: "enter",
		node: root,
		path: rootPath
	}];
	for (let task = tasks.pop(); task !== void 0; task = tasks.pop()) {
		if (task.kind === "leave") {
			seen.delete(task.node);
			continue;
		}
		if (task.kind === "one-of-tail") {
			for (const key of ONE_OF_SIBLING_KEYWORDS) if (Object.hasOwn(task.node, key)) violations.push(`${task.path}.${key} is not supported beside oneOf`);
			continue;
		}
		if (task.kind === "object-tail") {
			checkObjectSchemaTail(task.node, task.path, task.properties, violations);
			continue;
		}
		const { node, path } = task;
		if (!isJsonSchemaRecord(node)) {
			violations.push(`${path} must be a schema object`);
			continue;
		}
		if (seen.has(node)) {
			violations.push(`${path} is circular`);
			continue;
		}
		seen.add(node);
		tasks.push({
			kind: "leave",
			node
		});
		for (const key of Object.keys(node)) {
			if (CONSTRAINT_KEYWORDS.has(key)) continue;
			if (ANNOTATION_KEYWORDS.has(key)) {
				try {
					if (!isJsonValue(node[key])) violations.push(`${path}.${key} annotation must be lossless JSON data`);
				} catch {
					violations.push(`${path}.${key} annotation must be lossless JSON data`);
				}
				continue;
			}
			violations.push(`${path}.${key} is not a supported keyword (subset: type/oneOf/properties/required/additionalProperties/items/enum/const + annotations)`);
		}
		if (Object.hasOwn(node, "description") && typeof node.description !== "string") violations.push(`${path}.description must be a string`);
		if (Object.hasOwn(node, "title") && typeof node.title !== "string") violations.push(`${path}.title must be a string`);
		const hasType = Object.hasOwn(node, "type");
		const hasOneOf = Object.hasOwn(node, "oneOf");
		if (hasType && hasOneOf) {
			violations.push(`${path} cannot declare both type and oneOf`);
			continue;
		}
		if (!hasType && !hasOneOf) {
			for (const key of ONE_OF_SIBLING_KEYWORDS) if (Object.hasOwn(node, key)) violations.push(`${path}.${key} requires type or oneOf`);
			continue;
		}
		if (hasOneOf) {
			const oneOf = node.oneOf;
			tasks.push({
				kind: "one-of-tail",
				node,
				path
			});
			if (!isPlainJsonArray(oneOf) || oneOf.length < 2) violations.push(`${path}.oneOf must be an array of at least two schemas`);
			else for (let index = oneOf.length - 1; index >= 0; index--) tasks.push({
				kind: "enter",
				node: oneOf[index],
				path: `${path}.oneOf[${index}]`
			});
			continue;
		}
		const type = node.type;
		if (typeof type !== "string" || !SCHEMA_TYPES.includes(type)) {
			violations.push(Array.isArray(type) ? `${path}.type must be a single type string (type arrays are not supported)` : `${path}.type must be one of ${SCHEMA_TYPES.join("/")}`);
			continue;
		}
		const schemaType = type;
		for (const [key, types] of Object.entries({
			properties: ["object"],
			required: ["object"],
			additionalProperties: ["object"],
			items: ["array"],
			enum: [
				"string",
				"number",
				"integer",
				"boolean",
				"null"
			],
			const: [
				"string",
				"number",
				"integer",
				"boolean",
				"null"
			]
		})) if (Object.hasOwn(node, key) && !types.includes(schemaType)) violations.push(`${path}.${key} is not supported on type "${schemaType}"`);
		switch (schemaType) {
			case "object": {
				const properties = Object.hasOwn(node, "properties") ? node.properties : void 0;
				tasks.push({
					kind: "object-tail",
					node,
					path,
					properties
				});
				if (Object.hasOwn(node, "properties")) if (!isJsonSchemaRecord(properties)) violations.push(`${path}.properties must be an object of schemas`);
				else {
					const entries = Object.entries(properties);
					for (let index = entries.length - 1; index >= 0; index--) {
						const entry = entries[index];
						/* v8 ignore next -- the loop is bounded by the captured entry count. */
						if (entry === void 0) continue;
						tasks.push({
							kind: "enter",
							node: entry[1],
							path: `${path}.properties.${entry[0]}`
						});
					}
				}
				break;
			}
			case "array":
				if (Object.hasOwn(node, "items")) tasks.push({
					kind: "enter",
					node: node.items,
					path: `${path}.items`
				});
				break;
			case "string":
			case "number":
			case "integer":
			case "boolean":
			case "null": {
				const hasEnum = Object.hasOwn(node, "enum");
				const allowed = hasEnum ? node.enum : void 0;
				const enumValid = isPlainJsonArray(allowed) && allowed.length > 0 && allowed.every((entry) => scalarMatches(schemaType, entry));
				if (hasEnum && !enumValid) violations.push(`${path}.enum must be a non-empty array of ${schemaType} values`);
				const hasConst = Object.hasOwn(node, "const");
				const declaredConst = hasConst ? node.const : void 0;
				const constValid = scalarMatches(schemaType, declaredConst);
				if (hasConst) {
					if (!constValid) violations.push(`${path}.const must be a ${schemaType} value`);
					else if (enumValid && !allowed.includes(declaredConst)) violations.push(`${path}.const must be one of ${path}.enum when both are declared`);
				}
				break;
			}
			/* v8 ignore next -- schemaType was narrowed from the closed SCHEMA_TYPES table above. */
			default: assertNever(schemaType, "JsonSchemaType");
		}
	}
}
/**
* Assert that an arbitrary raw schema uses only the enforced subset.
* Annotation-only schemas are accepted as the standard unconstrained-JSON
* form; callers that require an object root use {@link assertObjectJsonSchema}.
* @param schema - untrusted raw JSON Schema.
* @returns Assertion that the schema belongs to the supported subset.
*/
function assertSupportedJsonSchema(schema) {
	const violations = [];
	checkSchemaNode(schema, "schema", violations, /* @__PURE__ */ new Set());
	if (violations.length > 0) throw new JsonSchemaError(violations);
}
/** Safely test the lossless JSON boundary when a getter may throw. */
function safelyIsJsonValue(value) {
	try {
		return isJsonValue(value);
	} catch {
		return false;
	}
}
/** Root-aware diagnostic path for the parameter validator's empty sentinel. */
function diagnosticPath(path) {
	return path === "" ? "arguments" : path;
}
/** Append one object property without a leading dot at an implicit root. */
function propertyPath(path, key) {
	return path === "" ? key : `${path}.${key}`;
}
/** The generic exception-containment diagnostic owned by one valid schema node. */
function losslessValueViolation(path) {
	return [`"${diagnosticPath(path)}" must be a lossless JSON value`];
}
/** Append diagnostics without spreading a potentially wide child result as call arguments. */
function appendViolations(target, source) {
	for (const violation of source) target.push(violation);
}
/** Initialize one validation frame with empty aggregation state. */
function valueFrame(node, value, path) {
	return {
		node,
		value,
		path,
		catches: false,
		phase: "start",
		children: [],
		childIndex: 0,
		violations: [],
		tailViolations: [],
		matches: 0
	};
}
/** Validate one scalar node after its primitive type check. */
function checkScalarValue(node, value, path) {
	const allowed = Object.hasOwn(node, "enum") ? node.enum : void 0;
	if (allowed !== void 0 && !allowed.includes(value)) return [`"${diagnosticPath(path)}" must be one of ${JSON.stringify(allowed)}`];
	if (Object.hasOwn(node, "const") && value !== node.const) return [`"${diagnosticPath(path)}" must be ${JSON.stringify(node.const)}`];
	return [];
}
/** Validate one trusted schema/value pair with explicit frames rather than recursive calls. */
function checkValue(schema, value, path) {
	const frames = [valueFrame(schema, value, path)];
	let rootResult;
	const receive = (result) => {
		const parent = frames.at(-1);
		if (parent === void 0) {
			rootResult = result;
			return;
		}
		if (parent.kind === "oneOf") {
			if (result.length === 0) parent.matches++;
		} else appendViolations(parent.violations, result);
	};
	const finish = (result) => {
		frames.pop();
		receive(result);
	};
	while (frames.length > 0) {
		const frame = frames.at(-1);
		/* v8 ignore next -- the loop condition guarantees a current frame. */
		if (frame === void 0) break;
		try {
			if (frame.phase === "children") {
				if (frame.childIndex < frame.children.length) {
					const child = frame.children[frame.childIndex];
					/* v8 ignore next -- childIndex is bounded by children.length. */
					if (child === void 0) throw new Error("missing schema-value child frame");
					frame.childIndex++;
					frames.push(valueFrame(child.node, child.value, child.path));
					continue;
				}
				if (frame.kind === "oneOf") {
					finish(frame.matches === 1 ? [] : [`"${diagnosticPath(frame.path)}" must match exactly one oneOf branch (matched ${frame.matches})`]);
					continue;
				}
				appendViolations(frame.violations, frame.tailViolations);
				if (frame.violations.length > 0) finish(frame.violations);
				else if (frame.kind === "object") finish(safelyIsJsonValue(frame.value) ? [] : [`"${diagnosticPath(frame.path)}" must be a lossless JSON object`]);
				else finish(safelyIsJsonValue(frame.value) ? [] : [`"${diagnosticPath(frame.path)}" must be a dense lossless JSON array`]);
				continue;
			}
			const nodeType = Object.hasOwn(frame.node, "type") ? frame.node.type : void 0;
			frame.catches = !(nodeType !== void 0 && !SCHEMA_TYPES.includes(nodeType));
			const oneOf = Object.hasOwn(frame.node, "oneOf") ? frame.node.oneOf : void 0;
			if (oneOf !== void 0) {
				frame.kind = "oneOf";
				frame.children = Array.from(oneOf, (branch) => ({
					node: branch,
					value: frame.value,
					path: frame.path
				}));
				frame.childIndex = 0;
				frame.matches = 0;
				frame.phase = "children";
				continue;
			}
			if (nodeType === void 0) {
				finish(safelyIsJsonValue(frame.value) ? [] : losslessValueViolation(frame.path));
				continue;
			}
			switch (nodeType) {
				case "object": {
					if (!isPlainJsonRecord(frame.value)) {
						finish([`"${diagnosticPath(frame.path)}" must be an object`]);
						break;
					}
					const properties = Object.hasOwn(frame.node, "properties") ? frame.node.properties ?? {} : {};
					const violations = [];
					const required = Object.hasOwn(frame.node, "required") ? frame.node.required ?? [] : [];
					for (const key of required) if (!Object.hasOwn(frame.value, key) || frame.value[key] === void 0) violations.push(`missing required property "${propertyPath(frame.path, key)}"`);
					const children = [];
					for (const [key, child] of Object.entries(properties)) {
						if (!Object.hasOwn(frame.value, key) || frame.value[key] === void 0) continue;
						children.push({
							node: child,
							value: frame.value[key],
							path: propertyPath(frame.path, key)
						});
					}
					const tailViolations = [];
					if (Object.hasOwn(frame.node, "additionalProperties") && frame.node.additionalProperties === false) {
						for (const key of Object.keys(frame.value)) if (!Object.hasOwn(properties, key)) tailViolations.push(`"${propertyPath(frame.path, key)}" is not a declared property (additionalProperties: false)`);
					}
					frame.kind = "object";
					frame.children = children;
					frame.childIndex = 0;
					frame.violations = violations;
					frame.tailViolations = tailViolations;
					frame.phase = "children";
					break;
				}
				case "array": {
					if (!Array.isArray(frame.value)) {
						finish([`"${diagnosticPath(frame.path)}" must be an array`]);
						break;
					}
					const items = Object.hasOwn(frame.node, "items") ? frame.node.items : void 0;
					const children = items === void 0 ? [] : frame.value.flatMap((entry, index) => [{
						node: items,
						value: entry,
						path: `${frame.path}[${index}]`
					}]);
					frame.kind = "array";
					frame.children = children;
					frame.childIndex = 0;
					frame.violations = [];
					frame.phase = "children";
					break;
				}
				case "string":
					finish(typeof frame.value === "string" ? checkScalarValue(frame.node, frame.value, frame.path) : [`"${diagnosticPath(frame.path)}" must be a string`]);
					break;
				case "number":
					finish(typeof frame.value !== "number" ? [`"${diagnosticPath(frame.path)}" must be a number`] : !isJsonNumber(frame.value) ? [`"${diagnosticPath(frame.path)}" must be a finite JSON number`] : checkScalarValue(frame.node, frame.value, frame.path));
					break;
				case "integer":
					finish(!isJsonNumber(frame.value) || !Number.isInteger(frame.value) ? [`"${diagnosticPath(frame.path)}" must be an integer`] : checkScalarValue(frame.node, frame.value, frame.path));
					break;
				case "boolean":
					finish(typeof frame.value === "boolean" ? checkScalarValue(frame.node, frame.value, frame.path) : [`"${diagnosticPath(frame.path)}" must be a boolean`]);
					break;
				case "null":
					finish(frame.value === null ? checkScalarValue(frame.node, frame.value, frame.path) : [`"${diagnosticPath(frame.path)}" must be null`]);
					break;
				default: finish(assertNever(nodeType, "JsonSchemaType"));
			}
		} catch (error) {
			let failed = frames.pop();
			while (failed !== void 0 && !failed.catches) failed = frames.pop();
			if (failed === void 0) throw error;
			receive(losslessValueViolation(failed.path));
		}
	}
	/* v8 ignore next -- every root frame finishes or throws. */
	return rootResult ?? losslessValueViolation(path);
}
/**
* Validate a candidate value against an asserted raw schema. The function is
* total for arbitrary values and returns path-qualified violations.
* @param schema - a schema accepted by {@link assertSupportedJsonSchema}.
* @param value - the candidate JSON value.
* @param path - root label used in diagnostics.
* @returns All violations in walk order; empty means valid.
*/
function validateJsonSchemaValue(schema, value, path = "value") {
	return checkValue(schema, value, path);
}
/** Unified JSON-value schema DSL, inference, compilation, and typed tool helper. @module dsh-tools/schema */
const ANNOTATION_KEYS = [
	"description",
	"title",
	"default",
	"examples"
];
/** Throw one author-schema violation through the shared schema error type. */
function authorError(message) {
	throw new JsonSchemaError([message]);
}
/** Copy own annotation fields for validation by the raw-schema boundary. */
function copyAnnotations(source, target) {
	if (Object.hasOwn(source, "description")) target.description = source.description;
	if (Object.hasOwn(source, "title")) target.title = source.title;
	if (Object.hasOwn(source, "default")) target.default = source.default;
	if (Object.hasOwn(source, "examples")) target.examples = source.examples;
}
/** Reject author-only keys outside one node's declared vocabulary. */
function assertAuthorKeys(source, path, allowed) {
	for (const key of Object.keys(source)) if (!allowed.includes(key)) authorError(`${path}.${key} is not supported by the value schema DSL`);
}
/** Install a compiled node without giving `__proto__` assignment semantics. */
function assignCompiledNode(destination, node) {
	switch (destination.kind) {
		case "root":
			destination.holder.value = node;
			break;
		case "property":
			Object.defineProperty(destination.target, destination.key, {
				value: node,
				enumerable: true,
				configurable: true,
				writable: true
			});
			break;
		case "item":
			destination.target.items = node;
			break;
		case "one-of": destination.target[destination.index] = node;
	}
}
/** Install a compiled property map at its root or containing object node. */
function assignCompiledPropertyMap(destination, compiled) {
	if (destination.kind === "root") destination.holder.value = compiled;
	else destination.target.properties = compiled.properties;
}
/** Execute an author-schema compilation task graph without recursive descent. */
function runSchemaCompiler(initial) {
	const seen = /* @__PURE__ */ new Set();
	const tasks = [initial];
	for (let task = tasks.pop(); task !== void 0; task = tasks.pop()) {
		if (task.kind === "leave") {
			seen.delete(task.input);
			continue;
		}
		if (task.kind === "property-map-tail") {
			if (task.required.length > 0) {
				task.compiled.required = task.required;
				if (task.destination.kind === "object") task.destination.target.required = task.required;
			}
			continue;
		}
		if (task.kind === "property") {
			if (!isJsonSchemaRecord(task.property)) authorError(`${task.path} must be a value schema object`);
			if (Object.hasOwn(task.property, "required") && task.property.required !== true) authorError(`${task.path}.required must be true when present`);
			if (Object.hasOwn(task.property, "required") && task.property.required === true) task.required.push(task.key);
			tasks.push({
				kind: "value",
				input: task.property,
				path: task.path,
				allowRequired: true,
				destination: {
					kind: "property",
					target: task.properties,
					key: task.key
				}
			});
			continue;
		}
		if (task.kind === "property-map") {
			if (!isJsonSchemaRecord(task.input)) authorError(`${task.path} must be an object of value schemas`);
			if (seen.has(task.input)) authorError(`${task.path} is circular`);
			seen.add(task.input);
			const compiled = { properties: {} };
			const required = [];
			assignCompiledPropertyMap(task.destination, compiled);
			tasks.push({
				kind: "leave",
				input: task.input
			});
			tasks.push({
				kind: "property-map-tail",
				compiled,
				required,
				destination: task.destination
			});
			const entries = Object.entries(task.input);
			for (let index = entries.length - 1; index >= 0; index--) {
				const entry = entries[index];
				/* v8 ignore next -- the loop is bounded by the captured entry count. */
				if (entry === void 0) continue;
				tasks.push({
					kind: "property",
					property: entry[1],
					path: `${task.path}.${entry[0]}`,
					key: entry[0],
					properties: compiled.properties,
					required
				});
			}
			continue;
		}
		const { input, path } = task;
		if (!isJsonSchemaRecord(input)) authorError(`${path} must be a value schema object`);
		if (seen.has(input)) authorError(`${path} is circular`);
		seen.add(input);
		const authorKeys = [...ANNOTATION_KEYS, ...task.allowRequired ? ["required"] : []];
		const node = {};
		assignCompiledNode(task.destination, node);
		tasks.push({
			kind: "leave",
			input
		});
		if (Object.hasOwn(input, "oneOf")) {
			assertAuthorKeys(input, path, [
				...authorKeys,
				"oneOf",
				"type"
			]);
			if (Object.hasOwn(input, "type")) authorError(`${path} cannot declare both type and oneOf`);
			if (!isPlainJsonArray(input.oneOf)) authorError(`${path}.oneOf must be an array of at least two value schemas`);
			const branches = [];
			node.oneOf = branches;
			copyAnnotations(input, node);
			for (let index = input.oneOf.length - 1; index >= 0; index--) tasks.push({
				kind: "value",
				input: input.oneOf[index],
				path: `${path}.oneOf[${index}]`,
				allowRequired: false,
				destination: {
					kind: "one-of",
					target: branches,
					index
				}
			});
			continue;
		}
		const inputType = Object.hasOwn(input, "type") ? input.type : void 0;
		switch (inputType) {
			case "json":
				assertAuthorKeys(input, path, [...authorKeys, "type"]);
				copyAnnotations(input, node);
				break;
			case "object":
				assertAuthorKeys(input, path, [
					...authorKeys,
					"type",
					"properties",
					"additionalProperties"
				]);
				if (!Object.hasOwn(input, "additionalProperties") || typeof input.additionalProperties !== "boolean") authorError(`${path}.additionalProperties must be explicitly true or false`);
				node.type = "object";
				copyAnnotations(input, node);
				node.additionalProperties = input.additionalProperties;
				if (Object.hasOwn(input, "properties")) tasks.push({
					kind: "property-map",
					input: input.properties,
					path: `${path}.properties`,
					destination: {
						kind: "object",
						target: node
					}
				});
				break;
			case "array":
				assertAuthorKeys(input, path, [
					...authorKeys,
					"type",
					"items"
				]);
				node.type = "array";
				copyAnnotations(input, node);
				if (Object.hasOwn(input, "items")) tasks.push({
					kind: "value",
					input: input.items,
					path: `${path}.items`,
					allowRequired: false,
					destination: {
						kind: "item",
						target: node
					}
				});
				break;
			case "string":
			case "number":
			case "integer":
			case "boolean":
			case "null":
				assertAuthorKeys(input, path, [
					...authorKeys,
					"type",
					"enum",
					"const"
				]);
				node.type = inputType;
				copyAnnotations(input, node);
				if (Object.hasOwn(input, "enum")) {
					if (!isPlainJsonArray(input.enum)) authorError(`${path}.enum must be a non-empty array of scalar values`);
					node.enum = Array.from(input.enum, (entry) => entry);
				}
				if (Object.hasOwn(input, "const")) node.const = input.const;
				break;
			default: authorError(`${path}.type must be string/number/integer/boolean/null/array/object/json, or use oneOf`);
		}
	}
}
/** Compile one implicit property map, collecting per-property requiredness. */
function compilePropertyMap(input, path) {
	const holder = {};
	runSchemaCompiler({
		kind: "property-map",
		input,
		path,
		destination: {
			kind: "root",
			holder
		}
	});
	/* v8 ignore next -- the root task assigns before scheduling any descendants. */
	return holder.value ?? authorError(`${path} did not compile`);
}
/** Compile one author node without applying any consumer root restriction. */
function compileValueSchema(input, path) {
	const holder = {};
	runSchemaCompiler({
		kind: "value",
		input,
		path,
		allowRequired: false,
		destination: {
			kind: "root",
			holder
		}
	});
	/* v8 ignore next -- the root task assigns before scheduling any descendants. */
	return holder.value ?? authorError(`${path} did not compile`);
}
/**
* Compile one author-facing value schema to the enforced raw JSON Schema
* subset. The author-only `json` node becomes an annotation-only schema.
* @param spec - schema for any JSON-value root.
* @returns The asserted raw schema projection.
*/
function valueSchemaSpecToJsonSchema(spec) {
	const schema = compileValueSchema(spec, "schema");
	assertSupportedJsonSchema(schema);
	return schema;
}
/**
* Compile the implicit open parameter object into raw JSON Schema.
* @param spec - per-property parameter definitions.
* @returns An object-rooted raw schema with no implicit-root openness override.
*/
function parameterSchemaSpecToJsonSchema(spec) {
	const compiled = compilePropertyMap(spec, "parameters");
	const schema = {
		type: "object",
		properties: compiled.properties,
		...compiled.required === void 0 ? {} : { required: compiled.required }
	};
	assertSupportedJsonSchema(schema);
	return schema;
}
/** Invalid model-generated arguments for a typed tool. */
var ToolArgsError = class extends HarnessError {
	/** Individual violations in schema-walk order. */
	violations;
	constructor(violations) {
		super(`invalid arguments: ${violations.join("; ")}`, "INVALID_ARGS");
		this.name = "ToolArgsError";
		this.violations = violations;
	}
};
/**
* Define a first-party tool with inferred arguments and strict execution
* validation. Replay-only presenters validate softly and fall back to generic
* rendering for obsolete logged arguments.
* @param options - typed definition and optional finalizer and presenters.
* @returns A registry-ready definition.
*/
function defineTool(options) {
	const userExecute = options.execute;
	const userFinalizeContent = options.finalizeContent;
	const userProjectContent = options.projectContent;
	const userRender = options.output.render;
	const userPresentationMeta = options.output.presentationMeta;
	const userPresentCall = options.presentCall;
	const userPresentResult = options.presentResult;
	const userIsConcurrencySafe = options.isConcurrencySafe;
	if (options.timeoutMs !== void 0 && (!Number.isFinite(options.timeoutMs) || options.timeoutMs <= 0)) throw new Error(`defineTool(${options.name}): timeoutMs must be a positive finite number`);
	const parameters = parameterSchemaSpecToJsonSchema(options.parameters);
	const outputSchema = valueSchemaSpecToJsonSchema(options.output.schema);
	const validate = (args) => validateJsonSchemaValue(parameters, args, "");
	const tool = {
		name: options.name,
		description: options.description,
		parameters,
		output: {
			schema: outputSchema,
			render(args, value) {
				return userRender(args, value);
			},
			...userPresentationMeta !== void 0 ? { presentationMeta(args, value) {
				return userPresentationMeta(args, value);
			} } : {}
		},
		...options.deferLoading === true ? { deferLoading: options.deferLoading } : {},
		...options.timeoutMs !== void 0 ? { timeoutMs: options.timeoutMs } : {},
		async execute(args, exec) {
			const violations = validate(args);
			if (violations.length > 0) throw new ToolArgsError(violations);
			return userExecute(args, exec);
		}
	};
	if (userProjectContent) tool.projectContent = (exec, result) => userProjectContent(exec, result);
	if (userFinalizeContent) tool.finalizeContent = (exec, result) => userFinalizeContent(exec, result);
	if (userPresentCall) tool.presentCall = (args) => {
		if (validate(args).length > 0) return void 0;
		return userPresentCall(args);
	};
	if (userPresentResult) tool.presentResult = (args, result) => {
		if (validate(args).length > 0) return void 0;
		return userPresentResult(args, result);
	};
	if (userIsConcurrencySafe) tool.isConcurrencySafe = (args) => {
		if (validate(args).length > 0) return false;
		return userIsConcurrencySafe(args);
	};
	return tool;
}
/**
* PTC mode `run_code` transport. Programs call the registry's agent-visible
* tools through nested executions scheduled under the native concurrency
* contract; each sub-dispatch is logged for reconstruction, while only the
* outer curated result enters model history.
* @module @deepseek-ai/dsh-tools/src/ptc
*/
/** The model-facing name of the PTC mode tool. */
const RUN_CODE_NAME = "run_code";
/**
* The TypeScript flavor: the fallback for a schema read with no runtime
* mounted ({@link resolveFlavor} owns which readers reach that). A real
* assembly always resolves a runtime first, so the model never sees this
* fallback outside its own language.
*/
const TYPESCRIPT_FLAVOR = {
	description: "Execute a TypeScript program against the available tools. Takes two required arguments: `code`, the BODY of an async function (erasable syntax only; top-level `await` and `return` work), and `description`, a short summary of what the program does. Call tools as `await tools.name(args)` per the declarations in the system prompt. Only what you print or return is program output — curate it. Image-bearing subtool results are attached after the run.",
	codeDescription: "The program: the body of an async TypeScript function."
};
/** Per-language `run_code` schema flavors (see {@link RunCodeFlavor}); one entry per {@link PtcSdkLanguage}. */
const RUN_CODE_FLAVORS = {
	typescript: TYPESCRIPT_FLAVOR,
	python: {
		description: "Execute a Python program against the available tools. Takes two required arguments: `code`, the BODY of an async function (top-level `await` and `return` work), and `description`, a short summary of what the program does. Call tools as `await tools.name(args)` per the declarations in the system prompt. Use `print(...)` and/or `return <value>` for program output — curate it. Image-bearing subtool results are attached after the run.",
		codeDescription: "The program: the body of an async Python function."
	}
};
/**
* The `description` parameter's model-facing description: language-independent
* (the UI label contract is the same for every runtime), shared between the
* static spec and the language-aware `parameters` getter so the two emissions
* can never drift.
*/
const RUN_CODE_DESCRIPTION_PARAM_DESCRIPTION = "Clear, concise description of what this program does in active voice, 5-10 words (shown in the UI). Examples: \"Count TODO markers across packages\"; \"Read failing test and its fixture\"; \"Rename config key in every cordis.yml\".";
const RUN_CODE_CONTROLS = {
	timeoutMs: {
		type: "number",
		description: "Positive elapsed-time budget in milliseconds, capped by the deployment maximum."
	},
	sandbox_permissions: {
		type: "string",
		enum: [...ESCALATION_TARGETS],
		description: "Wider sandbox mode for this complete program execution; requires justification and approval."
	},
	justification: {
		type: "string",
		description: "Reason this complete program needs wider access, shown to the user for approval. Use the language of the user’s current request."
	}
};
function controlParameters(runtime) {
	if (runtime === void 0) return RUN_CODE_CONTROLS;
	return {
		...runtime.timeout === void 0 ? {} : { timeoutMs: {
			...RUN_CODE_CONTROLS.timeoutMs,
			description: `Positive elapsed-time budget in milliseconds, including nested tool and approval waits. Default ${runtime.timeout.defaultMs}; capped at ${runtime.timeout.maxMs}. Zero does not disable the deadline.`
		} },
		...runtime.sandboxMode === void 0 ? {} : {
			sandbox_permissions: RUN_CODE_CONTROLS.sandbox_permissions,
			justification: RUN_CODE_CONTROLS.justification
		}
	};
}
function escalationGuidance(runtime) {
	return runtime?.sandboxMode === void 0 ? "" : " A sandbox escalation approves this complete program for one execution only. Nested tools retain their own policies and approvals. Request wider access only after evidence of a denial. Earlier effects may already have completed: inspect them before explicitly retrying. Programs are never replayed automatically.";
}
/**
* Resolve the {@link RunCodeFlavor} for the loaded runtime's language, read at
* schema-emission time so the model-visible `run_code` schema always matches
* the SDK section's language. `peekRuntime` returns `undefined` only when no
* runtime is mounted, which reaches this function through definition readers
* and `schemas()` — the doc-catalog harvest is the only shipped one, and none
* of them feeds a model, because `wireSchemas` calls `requirePtcRuntime`
* before projecting — so that path degrades to {@link TYPESCRIPT_FLAVOR}. A
* mounted runtime whose language has no flavor entry fails loud, exactly as
* `requirePtcRuntime` rejects it at assembly. Keeping this table in step with
* `SDK_RENDERERS` is the compiler's job ({@link PtcSdkLanguage}); what this
* guard owns is the runtime-supplied language neither table knows, which never
* yields a wrong-language schema for a real runtime.
*/
function resolveFlavor(peekRuntime) {
	const runtime = peekRuntime();
	if (runtime === void 0) return TYPESCRIPT_FLAVOR;
	const flavor = RUN_CODE_FLAVORS[runtime.language];
	if (!Object.hasOwn(RUN_CODE_FLAVORS, runtime.language) || flavor === void 0) {
		const known = Object.keys(RUN_CODE_FLAVORS).map((name) => JSON.stringify(name)).join(", ");
		throw new Error(`dsh-tools: no run_code schema flavor registered for runtime language ${JSON.stringify(runtime.language)} (known: ${known})`);
	}
	return flavor;
}
/**
* Thrown by `run_code` when the program run itself failed — a program
* exception, a budget expiry, an abort, or substrate death. Extends
* {@link HarnessError} (`code: 'CODE_RUN_FAILED'`); the registry's execution
* pipeline converts it into a structured `isError` result whose text carries
* the failure kind plus the captured logs, so the model can self-correct.
*/
var CodeRunFailedError = class extends HarnessError {
	constructor(message) {
		super(message, "CODE_RUN_FAILED");
		this.name = "CodeRunFailedError";
	}
};
/**
* Snapshot one binding call's argument as lossless JSON, then snapshot that
* detached value again so dispatch and logging stay independent without
* reintroducing structured-clone's platform-specific nesting limit.
*/
function jsonNormalizeArgs(value) {
	let snapshot;
	try {
		snapshot = snapshotJsonValue(value);
	} catch (error) {
		throw new Error(`tool arguments must be lossless JSON: ${error instanceof Error ? error.message : String(error)}`);
	}
	if (snapshot === void 0) throw new Error("tool arguments must be lossless JSON (call the tool with an arguments object, e.g. `{}`)");
	const logged = snapshotJsonValue(snapshot);
	/* v8 ignore next -- snapshot is already a detached lossless JSON value. */
	if (logged === void 0) throw new Error("tool arguments could not be detached for durable logging");
	return {
		dispatched: snapshot,
		logged
	};
}
/** Two-space JSON presentation, matching the existing shallow `run_code` text contract. */
const JSON_INDENT = "  ";
/**
* ECMAScript caps `JSON.stringify`'s `space` string at ten characters. The
* renderer also caps TOTAL indentation there, compacting deeper subtrees, so
* formatted output remains linear in the canonical JSON size.
*/
const MAX_JSON_INDENT_CHARS = 10;
/** Render one non-string JSON root without recursive traversal or unbounded indentation growth. */
function renderJsonValue(value) {
	const chunks = [];
	const tasks = [{
		kind: "value",
		value,
		depth: 0,
		compact: false
	}];
	for (let task = tasks.pop(); task !== void 0; task = tasks.pop()) {
		if (task.kind === "text") {
			chunks.push(task.text);
			continue;
		}
		const current = task.value;
		if (current === null || typeof current === "boolean" || typeof current === "number") {
			chunks.push(String(current));
			continue;
		}
		if (typeof current === "string") {
			chunks.push(JSON.stringify(current));
			continue;
		}
		const compact = task.compact || (task.depth + 1) * 2 > MAX_JSON_INDENT_CHARS;
		const childDepth = task.depth + 1;
		if (Array.isArray(current)) {
			chunks.push("[");
			if (current.length === 0) {
				chunks.push("]");
				continue;
			}
			tasks.push({
				kind: "text",
				text: compact ? "]" : `\n${JSON_INDENT.repeat(task.depth)}]`
			});
			for (let index = current.length - 1; index >= 0; index--) {
				const item = current[index];
				/* v8 ignore next -- canonical JsonValue arrays are dense. */
				if (item === void 0) throw new Error("cannot render a sparse JSON array");
				tasks.push({
					kind: "value",
					value: item,
					depth: childDepth,
					compact
				});
				tasks.push({
					kind: "text",
					text: compact ? index === 0 ? "" : "," : `${index === 0 ? "\n" : ",\n"}${JSON_INDENT.repeat(childDepth)}`
				});
			}
			continue;
		}
		const keys = Object.keys(current);
		chunks.push("{");
		if (keys.length === 0) {
			chunks.push("}");
			continue;
		}
		tasks.push({
			kind: "text",
			text: compact ? "}" : `\n${JSON_INDENT.repeat(task.depth)}}`
		});
		for (let index = keys.length - 1; index >= 0; index--) {
			const key = keys[index];
			/* v8 ignore next -- the loop is bounded by the captured key count. */
			if (key === void 0) throw new Error("cannot render a missing JSON object key");
			const item = current[key];
			/* v8 ignore next -- canonical JsonValue records contain no undefined properties. */
			if (item === void 0) throw new Error("cannot render an undefined JSON object property");
			tasks.push({
				kind: "value",
				value: item,
				depth: childDepth,
				compact
			});
			tasks.push({
				kind: "text",
				text: compact ? `${index === 0 ? "" : ","}${JSON.stringify(key)}:` : `${index === 0 ? "\n" : ",\n"}${JSON_INDENT.repeat(childDepth)}${JSON.stringify(key)}: `
			});
		}
	}
	return chunks.join("");
}
/** Render one present program completion value for the model-facing result text. */
function renderValue(value) {
	return typeof value === "string" ? value : renderJsonValue(value);
}
/**
* Build the `run_code` {@link ToolDefinition}: required `code` and
* `description` parameters, executed through the dispatch bridge described
* above. The
* registry reserves it as presentation infrastructure under non-native modes,
* outside the filterable global/scoped capability layers.
* @param registry - the owning registry (sub-calls go through its `execute`,
*   bindings cover its registered tools).
* @param options - the registry-private capabilities described above.
* @returns the registry-ready definition.
*/
function createRunCodeTool(registry, options) {
	const { requireRuntime, peekRuntime, maxParallel, shapeDispatchLog } = options;
	const definition = defineTool({
		name: RUN_CODE_NAME,
		description: TYPESCRIPT_FLAVOR.description,
		parameters: {
			code: {
				type: "string",
				required: true,
				description: TYPESCRIPT_FLAVOR.codeDescription
			},
			description: {
				type: "string",
				required: true,
				description: RUN_CODE_DESCRIPTION_PARAM_DESCRIPTION
			},
			...RUN_CODE_CONTROLS
		},
		output: {
			schema: {
				type: "object",
				additionalProperties: false,
				properties: {
					logs: {
						type: "array",
						required: true,
						items: { type: "string" }
					},
					result: { type: "json" },
					sandbox: {
						type: "object",
						additionalProperties: false,
						properties: {
							mode: {
								type: "string",
								required: true,
								enum: [
									"read-only",
									"workspace-write",
									"danger-full-access"
								]
							},
							denied: {
								type: "boolean",
								required: true
							},
							enforcement: {
								type: "string",
								enum: ["full", "partial"]
							}
						}
					}
				}
			},
			render: (_args, value) => {
				const rendered = value.result === void 0 ? "" : renderValue(value.result);
				const parts = [value.logs.join("\n"), rendered].filter((part) => part.length > 0);
				if (value.sandbox?.enforcement === "partial") parts.push("File sandbox enforcement is partial on this host.");
				if (value.sandbox?.denied) parts.push(`The ${value.sandbox.mode} file sandbox denied an operation.${escalationGuidance(peekRuntime())}`);
				return [{
					type: "text",
					text: parts.length > 0 ? parts.join("\n") : "(run_code completed with no output)"
				}];
			}
		},
		async execute(args, exec) {
			if (args.description.trim().length === 0) throw new Error("invalid description: expected a non-empty string");
			const runtime = requireRuntime();
			validateEscalationArgs(args.sandbox_permissions, args.justification);
			if (args.timeoutMs !== void 0 && runtime.timeout === void 0) throw new Error("timeoutMs is not available for this PTC runtime");
			if (args.timeoutMs !== void 0 && (!Number.isFinite(args.timeoutMs) || args.timeoutMs <= 0)) throw new Error("invalid timeoutMs: expected a positive finite number");
			const standingPolicy = runtime.sandboxMode === void 0 ? void 0 : options.resolveSandboxPolicy(exec);
			let policy = standingPolicy;
			if (args.sandbox_permissions !== void 0 && args.justification !== void 0) {
				if (standingPolicy === void 0) throw new Error("sandbox_permissions is not available for this PTC runtime");
				const approvedMode = await approveEscalation({
					requestedMode: args.sandbox_permissions,
					justification: args.justification,
					effectiveMode: standingPolicy.mode,
					subject: "program"
				}, {
					approver: options.peekApprover(),
					agent: exec.agent,
					callId: exec.callId,
					toolName: RUN_CODE_NAME,
					signal: exec.signal
				});
				policy = {
					...standingPolicy,
					mode: approvedMode
				};
			}
			exec.signal.throwIfAborted();
			const runController = new AbortController();
			const onOuterAbort = () => {
				runController.abort(exec.signal.reason);
			};
			exec.signal.addEventListener("abort", onOuterAbort, { once: true });
			let dispatches = 0;
			const pendingQueue = [];
			const inFlight = /* @__PURE__ */ new Set();
			/** Tracked settle-event side work (log-content listener + append), drained at run settlement. */
			const logWork = /* @__PURE__ */ new Set();
			const commitQueue = [];
			let exclusiveActive = false;
			let driving = false;
			let driverRun = Promise.resolve();
			let wake;
			const wakeup = () => {
				const release = wake;
				wake = void 0;
				release?.();
			};
			/**
			* The single ordered lane. Each pass commits the head-of-line settled
			* dispatch (ordered post-execute), then starts the next queued entry if
			* its slot is free (ordered pre-execute), and otherwise sleeps until a
			* body settles or a new submission arrives. One run reaching the
			* empty-queues/empty-pool state is quiescence.
			*/
			const drive = () => {
				if (driving) return driverRun;
				driving = true;
				driverRun = (async () => {
					try {
						for (;;) {
							const signal = new Promise((resolve) => {
								wake = resolve;
							});
							const commitHead = commitQueue[0];
							if (commitHead !== void 0 && commitHead.settled) {
								commitQueue.shift();
								await commitHead.commit();
								if (commitHead.mode === "exclusive") exclusiveActive = false;
								continue;
							}
							const head = pendingQueue[0];
							if (head !== void 0) {
								if (runController.signal.aborted) {
									pendingQueue.shift();
									head.abandon();
									continue;
								}
								const mode = head.classify();
								if (!exclusiveActive && (mode === "exclusive" ? inFlight.size === 0 : inFlight.size < maxParallel)) {
									if (mode === "exclusive") exclusiveActive = true;
									head.mode = mode;
									pendingQueue.shift();
									commitQueue.push(head);
									await head.start();
									const flight = head.flight.finally(() => {
										inFlight.delete(flight);
										wakeup();
									});
									inFlight.add(flight);
									continue;
								}
							}
							if (pendingQueue.length === 0 && commitQueue.length === 0 && inFlight.size === 0) return;
							await signal;
						}
					} finally {
						driving = false;
						wake = void 0;
					}
				})();
				return driverRun;
			};
			/** Every dispatch settled AND committed; nothing can start (the run is aborted at call time). */
			const drainDispatches = async () => {
				await drive();
				while (logWork.size > 0) await Promise.allSettled([...logWork]);
			};
			const runOver = () => runController.signal.aborted;
			const binding = (schema) => async (rawArgs) => {
				const { name } = schema;
				if (runOver()) throw new Error(`run_code run is over (${String(runController.signal.reason)}); ${name} not dispatched`);
				const normalized = jsonNormalizeArgs(rawArgs);
				const n = ++dispatches;
				const subCallId = brandString(`${String(exec.callId)}:ptc:${n}`);
				const input = {
					callId: subCallId,
					rootCallId: exec.rootCallId,
					name,
					schema,
					arguments: normalized.dispatched,
					...exec.agent ? { agent: exec.agent } : {},
					parent: exec.token,
					signal: runController.signal
				};
				const scheduler = registry[TOOL_RUNTIME_SCHEDULER];
				const outcome = await new Promise((resolve, reject) => {
					let parked;
					const settle = (result) => {
						resolve(result.isError ? {
							isError: true,
							message: result.error.message
						} : {
							isError: false,
							value: result.value
						});
						const agent = exec.agent;
						if (agent === void 0) return;
						const task = (async () => {
							const logged = await shapeDispatchLog({
								exec,
								agent,
								subCallId,
								name,
								isError: result.isError,
								content: result.content
							});
							agent.session.append("tool/ptc-dispatch", {
								rootCallId: exec.rootCallId,
								parentCallId: exec.callId,
								subCallId,
								name,
								arguments: normalized.logged,
								isError: result.isError,
								...result.error?.info === void 0 ? {} : { error: result.error.info },
								content: logged
							});
						})().finally(() => {
							logWork.delete(task);
						});
						logWork.add(task);
					};
					pendingQueue.push({
						flight: Promise.resolve(),
						settled: false,
						classify: () => registry.executionMode(input).kind,
						abandon: () => {
							reject(/* @__PURE__ */ new Error(`run_code run is over (${String(runController.signal.reason)}); ${name} tool call abandoned`));
						},
						async start() {
							exec.agent?.session.append("tool/ptc-dispatch-start", {
								rootCallId: exec.rootCallId,
								parentCallId: exec.callId,
								subCallId,
								name,
								arguments: normalized.logged
							});
							const prepared = await scheduler.prepare(input);
							if (prepared.kind === "dispatch") {
								this.flight = scheduler.dispatch(prepared.exec).then((dispatchOutcome) => {
									parked = {
										kind: dispatchOutcome.kind,
										exec: prepared.exec,
										result: dispatchOutcome.result
									};
									this.settled = true;
								});
								return;
							}
							parked = {
								kind: prepared.kind,
								exec: prepared.exec,
								result: prepared.result
							};
							this.settled = true;
						},
						async commit() {
							/* v8 ignore next -- commit() runs only after `settled` flipped, which set parked. */
							if (parked === void 0) return;
							const result = parked.kind === "post-result" ? await scheduler.finalize(parked.exec, parked.result) : scheduler.finish(parked.exec, parked.result);
							if (!result.isError && result.content.some((block) => block.type === "image")) exec.deferContext(createUserMessage({
								content: result.content,
								source: { kind: "ptc-mode" }
							}));
							for (const context of result.additionalContexts ?? []) exec.deferContext(context);
							if (result.concludesTurn) exec.concludeTurn();
							settle(result);
							while (logWork.size > maxParallel) await Promise.race(logWork);
						}
					});
					wakeup();
					drive();
				});
				if (runOver()) throw new Error(`run_code run is over (${String(runController.signal.reason)}); ${name} result discarded`);
				if (outcome.isError) throw new Error(outcome.message);
				return outcome.value;
			};
			const functions = Object.create(null);
			for (const schema of registry.schemas(exec.agent)) {
				if (schema.name === "run_code") continue;
				Object.defineProperty(functions, schema.name, {
					enumerable: true,
					value: binding(deepFreeze(schema))
				});
			}
			try {
				let result;
				try {
					result = await runtime.run(runtime.resolve({
						program: args.code,
						bindings: [{
							global: "tools",
							functions,
							errorClass: {
								name: "ToolCallError",
								memberNameProperty: "toolName"
							}
						}],
						signal: runController.signal,
						...exec.agent?.session.header.cwd !== void 0 ? { cwd: exec.agent.session.header.cwd } : {},
						...policy !== void 0 ? { sandboxPolicy: policy } : {},
						...args.timeoutMs !== void 0 ? { timeoutMs: args.timeoutMs } : {}
					}));
				} finally {
					runController.abort("run_code settled");
					await drainDispatches();
				}
				if (result.error) {
					const logsText = result.logs.length > 0 ? `\nCaptured output:\n${result.logs.join("\n")}` : "";
					const sandboxText = result.sandbox === void 0 ? "" : `\nFile sandbox: ${result.sandbox.mode}${result.sandbox.enforcement === void 0 ? "" : `; enforcement: ${result.sandbox.enforcement}`}${result.sandbox.denied ? "; operation denied" : ""}.`;
					throw new CodeRunFailedError(`code run failed (${result.error.kind}): ${result.error.message}${logsText}${sandboxText}${result.sandbox?.denied ? escalationGuidance(runtime) : ""}`);
				}
				return {
					logs: result.logs,
					...result.sandbox === void 0 ? {} : { sandbox: result.sandbox },
					...result.value !== void 0 ? { result: result.value } : {}
				};
			} finally {
				exec.signal.removeEventListener("abort", onOuterAbort);
			}
		},
		presentCall: (args) => ({
			card: "generic",
			title: args.description,
			kind: "execute",
			rawInput: args.code
		})
	});
	Object.defineProperty(definition, "description", {
		enumerable: true,
		get: () => {
			const runtime = peekRuntime();
			const instructions = runtime?.executionInstructions;
			return resolveFlavor(peekRuntime).description + (instructions ? ` ${instructions}` : "") + (runtime === void 0 ? "" : " The working directory is the Session's current directory.") + escalationGuidance(runtime);
		}
	});
	Object.defineProperty(definition, "parameters", {
		enumerable: true,
		get: () => parameterSchemaSpecToJsonSchema({
			code: {
				type: "string",
				required: true,
				description: resolveFlavor(peekRuntime).codeDescription
			},
			description: {
				type: "string",
				required: true,
				description: RUN_CODE_DESCRIPTION_PARAM_DESCRIPTION
			},
			...controlParameters(peekRuntime())
		})
	});
	return definition;
}
/**
* PTC mode codegen: the pure projection from registered tool schemas to the TypeScript SDK
* text the model programs against (the `tools:sdk` prompt section). Sibling of
* `json-schema.ts` — `schemas()` (native function calling) and this module (the generated
* `declare const tools` API) are two projections of the same store.
* @module @deepseek-ai/dsh-tools/src/ts-types
*/
/** Property names that are valid bare TS identifiers; anything else is quoted. */
const IDENTIFIER$1 = /^[A-Za-z_$][A-Za-z0-9_$]*$/;
/** Render an object key: bare when it is a valid identifier, quoted otherwise (every name stays reachable, no aliasing). */
function renderKey(name) {
	return IDENTIFIER$1.test(name) ? name : JSON.stringify(name);
}
/** One `indent`-deep line prefix (two spaces per level). */
function pad$1(indent) {
	return "  ".repeat(indent);
}
/** A one-line JSDoc block for a schema `description`, or no lines when there is none. */
function docLines$1(description, indent) {
	if (typeof description !== "string" || description.length === 0) return [];
	const collapsed = description.replace(/\s+/g, " ").trim();
	return [`${pad$1(indent)}/** ${collapsed.replaceAll("*/", String.raw`*\/`)} */`];
}
/** Render one scalar already validated by the unified schema boundary. */
function renderScalar(value) {
	return JSON.stringify(value);
}
/** Render a validated scalar `const`/`enum`, falling back to the broad type. */
function renderConstrainedScalar$1(node, type) {
	const broad = type === "integer" ? "number" : type;
	if (Object.hasOwn(node, "const")) return renderScalar(node.const);
	if (Object.hasOwn(node, "enum")) return node.enum.map(renderScalar).join(" | ");
	return broad;
}
/** Build one document from captured parts while retaining the legacy array-parenthesization test. */
function typeDocumentFrom(parts) {
	return {
		parts,
		containsUnionOrIntersection: parts.some((part) => typeof part === "string" ? part.includes("|") || part.includes("&") : part.containsUnionOrIntersection)
	};
}
/** Build a small document without an intermediate array at each call site. */
function typeDocument(...parts) {
	return typeDocumentFrom(parts);
}
/** Flatten a nested document with an explicit work stack. */
function flattenTypeDocument(document) {
	const chunks = [];
	const tasks = [document];
	for (let task = tasks.pop(); task !== void 0; task = tasks.pop()) {
		if (typeof task === "string") {
			chunks.push(task);
			continue;
		}
		for (let index = task.parts.length - 1; index >= 0; index--) {
			const part = task.parts[index];
			/* v8 ignore next -- the loop is bounded by the captured part count. */
			if (part !== void 0) tasks.push(part);
		}
	}
	return chunks.join("");
}
/** Initialize one schema-render frame with empty aggregation state. */
function schemaRenderFrame(node, indent) {
	return {
		node,
		indent,
		phase: "start",
		children: [],
		childIndex: 0,
		childDocuments: [],
		entries: []
	};
}
/** Render an already asserted schema to a composable document. */
function renderSupportedSchema(schema, indent) {
	const frames = [schemaRenderFrame(schema, indent)];
	let rootDocument;
	const finish = (document) => {
		frames.pop();
		const parent = frames.at(-1);
		if (parent === void 0) rootDocument = document;
		else parent.childDocuments.push(document);
	};
	while (frames.length > 0) {
		const frame = frames.at(-1);
		/* v8 ignore next -- the loop condition guarantees a current frame. */
		if (frame === void 0) break;
		if (frame.phase === "children") {
			if (frame.childIndex < frame.children.length) {
				const child = frame.children[frame.childIndex];
				/* v8 ignore next -- childIndex is bounded by children.length. */
				if (child === void 0) throw new Error("missing schema render child");
				frame.childIndex++;
				frames.push(schemaRenderFrame(child.node, child.indent));
				continue;
			}
			if (frame.kind === "oneOf") {
				const parts = [];
				for (let index = 0; index < frame.childDocuments.length; index++) {
					if (index > 0) parts.push(" | ");
					const child = frame.childDocuments[index];
					/* v8 ignore next -- child documents correspond one-to-one with children. */
					if (child !== void 0) parts.push(child);
				}
				finish(typeDocumentFrom(parts));
				continue;
			}
			if (frame.kind === "array") {
				const child = frame.childDocuments[0];
				/* v8 ignore next -- array frames always schedule exactly one child. */
				if (child === void 0) throw new Error("missing array item type");
				finish(child.containsUnionOrIntersection ? typeDocument("(", child, ")[]") : typeDocument(child, "[]"));
				continue;
			}
			const required = new Set(frame.node.required);
			const parts = ["{"];
			for (let index = 0; index < frame.entries.length; index++) {
				const entry = frame.entries[index];
				const child = frame.childDocuments[index];
				/* v8 ignore next -- object entries and child documents have the same length. */
				if (entry === void 0 || child === void 0) throw new Error("missing object property type");
				const [name, prop] = entry;
				for (const line of docLines$1(prop.description, frame.indent + 1)) parts.push("\n", line);
				parts.push("\n", `${pad$1(frame.indent + 1)}${renderKey(name)}${required.has(name) ? "" : "?"}: `, child, ";");
			}
			parts.push("\n", `${pad$1(frame.indent)}}`);
			const declared = typeDocumentFrom(parts);
			finish(frame.node.additionalProperties === false ? declared : typeDocument(declared, " & Record<string, JsonValue>"));
			continue;
		}
		const node = frame.node;
		if (node.oneOf !== void 0) {
			frame.kind = "oneOf";
			frame.children = Array.from(node.oneOf, (child) => ({
				node: child,
				indent: frame.indent
			}));
			frame.childIndex = 0;
			frame.childDocuments = [];
			frame.phase = "children";
			continue;
		}
		if (node.type === void 0) {
			finish(typeDocument("JsonValue"));
			continue;
		}
		switch (node.type) {
			case "string":
			case "number":
			case "integer":
			case "boolean":
			case "null":
				finish(typeDocument(renderConstrainedScalar$1(node, node.type)));
				break;
			case "array":
				if (node.items === void 0) finish(typeDocument("JsonValue[]"));
				else {
					frame.kind = "array";
					frame.children = [{
						node: node.items,
						indent: frame.indent
					}];
					frame.childIndex = 0;
					frame.childDocuments = [];
					frame.phase = "children";
				}
				break;
			case "object": {
				const open = node.additionalProperties !== false;
				const entries = Object.entries(node.properties ?? {});
				if (entries.length === 0) finish(typeDocument(open ? "Record<string, JsonValue>" : "Record<string, never>"));
				else {
					frame.kind = "object";
					frame.entries = entries;
					frame.children = entries.map(([, child]) => ({
						node: child,
						indent: frame.indent + 1
					}));
					frame.childIndex = 0;
					frame.childDocuments = [];
					frame.phase = "children";
				}
				break;
			}
			/* v8 ignore next -- assertSupportedJsonSchema narrowed this closed type union. */
			default: finish(typeDocument("unknown"));
		}
	}
	/* v8 ignore next -- every root frame produces one document. */
	return rootDocument ?? typeDocument("unknown");
}
/**
* Map one enforced JSON-Schema node to a TypeScript type literal. Supports
* every unified schema construct and returns `unknown` for malformed or
* unsupported inputs without throwing.
* @param schema - the JSON-Schema node (any shape; hostile inputs degrade).
* @param indent - the indentation level for nested object members.
* @returns the TS type text (multi-line for objects with properties).
*/
function jsonSchemaToTs(schema, indent = 0) {
	try {
		assertSupportedJsonSchema(schema);
		return flattenTypeDocument(renderSupportedSchema(schema, indent));
	} catch {
		return "unknown";
	}
}
/** The fixed model-facing usage contract rendered above the declarations (see the PTC mode Agent Note's "What the model sees"). */
const SDK_INSTRUCTIONS$1 = `## Writing code for run_code

\`run_code\` takes two required arguments: \`code\` — the body of an async TypeScript function (erasable syntax only — no \`enum\` or namespaces; type annotations are advisory, the code runs type-stripped) — and \`description\`, a short summary of what the program does. The declarations below are SDK bindings for this program. A declaration does not make its name a directly callable tool; only names supplied as separate tool schemas may be called directly.`;
const SDK_PROGRAM_INSTRUCTIONS = `Inside the program:

- Call tools as \`await tools.name(args)\` — quoted access for exotic names: \`tools["my-tool"](args)\`. Every call resolves to the tool's typed canonical JSON value. Tool arguments must be lossless JSON.
- A FAILED tool call rejects with \`ToolCallError\`, whose \`toolName\` identifies the failed tool and whose \`message\` is human-readable — \`try/catch\` it to handle and continue.
- Independent read-only calls MAY overlap under \`Promise.all\` (safe calls run concurrently; mutating calls run alone, in submission order). Sequence dependent work with \`await\`.
- Emit results with \`return\` and/or \`console.log(...)\`. Only what you print or return is program output. A successful tool result containing an image is attached after the run so you can inspect it on the next step; every other intermediate result stays out of the conversation, so extract just what you need.

Program-only SDK bindings:`;
/** Whether one string schema accepts the literal used by the bash example. */
function acceptsExampleString(schema, value) {
	return schema?.type === "string" && (schema.const === void 0 || schema.const === value) && (schema.enum === void 0 || schema.enum.includes(value));
}
/** Render the bash example only when its literal arguments satisfy the current parameter schema. */
function renderBashExample(schemas) {
	const bash = schemas.find((schema) => schema.name === "bash");
	if (bash === void 0) return "";
	const parameters = bash.parameters;
	if (parameters.type !== "object") return "";
	const required = parameters.required ?? [];
	if (required.some((name) => name !== "command" && name !== "description")) return "";
	if (!acceptsExampleString(parameters.properties?.command, "pwd")) return "";
	const needsDescription = required.includes("description");
	if (needsDescription && !acceptsExampleString(parameters.properties?.description, "Show current directory")) return "";
	return ` When no separate \`bash\` schema is supplied, invoke a declared \`bash\` binding inside \`run_code\`:\n\n\`run_code({ code: "return await tools.bash({ command: 'pwd'${needsDescription ? ", description: 'Show current directory'" : ""} })", description: "Show current directory" })\``;
}
/**
* Render the full `tools:sdk` prompt section: the fixed usage instructions
* plus one `declare const tools` interface covering every given tool.
* Deterministic — tools are emitted in lexicographic name order, so an
* unchanged tool set produces byte-identical text across assemblies. The sort
* is not a total order on byte-equal names, so two schemas sharing a name
* would render in argument order; the caller's visible-capability map is keyed
* by name, so the input never carries a duplicate.
* @param schemas - the tool schemas to declare (the caller excludes
*   `run_code` itself).
* @returns the complete section text.
*/
function renderToolsSdk(schemas) {
	const sorted = [...schemas].sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
	const argsMembers = [];
	const outputMembers = [];
	for (const schema of sorted) {
		argsMembers.push(...docLines$1(schema.description, 1));
		argsMembers.push(`${pad$1(1)}${renderKey(schema.name)}: ${jsonSchemaToTs(schema.parameters, 1)};`);
		outputMembers.push(`${pad$1(1)}${renderKey(schema.name)}: ${jsonSchemaToTs(schema.output, 1)};`);
	}
	const declaration = [
		`interface ToolArgsMap {${argsMembers.length > 0 ? `\n${argsMembers.join("\n")}\n` : ""}}`,
		`interface ToolOutputMap {${outputMembers.length > 0 ? `\n${outputMembers.join("\n")}\n` : ""}}`,
		"type ToolName = keyof ToolOutputMap",
		[
			"declare class ToolCallError extends Error {",
			"  readonly name: \"ToolCallError\";",
			"  readonly toolName: ToolName;",
			"}"
		].join("\n"),
		[
			"declare const tools: {",
			"  [K in ToolName]: (args: ToolArgsMap[K]) => Promise<ToolOutputMap[K]>;",
			"}"
		].join("\n")
	].join("\n\n");
	return `${SDK_INSTRUCTIONS$1}${renderBashExample(sorted)}\n\n${SDK_PROGRAM_INSTRUCTIONS}\n\n\`\`\`ts\ntype JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue }\n\n${declaration}\n\`\`\``;
}
/**
* PTC mode codegen — Python flavor. The pure projection from registered tool schemas to the
* Python SDK text the model programs against under `runtime.language === 'python'`. Sibling of
* {@link ./ts-types.ts | ts-types.ts}; the two files are two projections of the same registry
* store, keyed by the loaded {@link @deepseek-ai/dsh-ptc-runtime#PtcRuntime.language | PTC
* runtime's language}.
*
* Under `mode: 'ptc'` the native tool schemas are omitted from the request, so this generated
* SDK is the model's ONLY source for each tool's argument names, required fields, types,
* descriptions, and canonical output shapes; under `mode: 'both'` the native schemas ship
* alongside it and it is one of two. Object-shaped arguments and outputs therefore render as one
* named `TypedDict` per tool (and per nested object), not an opaque `dict[str, Any]`, so the
* shape survives into the program under the mode that has nothing else to carry it.
* @module @deepseek-ai/dsh-tools/src/py-types
*/
/**
* The reference grammar's `xid_start xid_continue*` — the set
* `str.isidentifier()` accepts on a CPython whose Unicode tables match the
* engine's. See {@link isBareIdentifier} for what a version skew does.
*/
const IDENTIFIER = /^[\p{XID_Start}_]\p{XID_Continue}*$/u;
/**
* Whether a name can be emitted as a bare Python identifier rather than
* routed to the subscript/`dict[str, Any]` path.
*
* Python identifiers are not ASCII: `路径` is as legal a field name as `path`,
* and rejecting it would degrade the whole enclosing object, dropping every
* field's name, requiredness, and type — information whose only source under
* `mode: 'ptc'` is this generated text.
*
* NFKC stability is a second and separate condition, because CPython
* normalizes identifiers at compile time while JSON keys are compared as
* written: `ﬁeld` would be declared and reachable as `field`, so the SDK would
* advertise a key under a spelling the harness never accepts, and two keys
* that normalize together would collapse into one declaration. Those names
* take the subscript path, which carries their exact bytes.
*
* `IDENTIFIER` matches `str.isidentifier()` (measured on Node 22.23.1 vs
* CPython 3.9.6 tables): the equivalence holds inside the two versions' shared
* tables, and the skew characters below are exactly where that pair diverges.
* The predicate as a whole is deliberately stricter than `isidentifier()`,
* which does not test NFKC stability: `'ﬁeld'.isidentifier()` is True and
* this returns false.
*
* Both conditions are evaluated against the ENGINE's Unicode tables, and the
* two sides are versioned independently — `\p{XID_Start}`/`\p{XID_Continue}`
* follow the running engine (Node 22.23.1 reports Unicode 17.0) while CPython
* follows its own (3.9.6 reports 13.0.0). The skew is not symmetric. A CPython
* older than the engine is the dangerous direction: a character added to either
* property since its tables (U+10570 Vithkuqi and U+1E290 Toto, 14.0; U+1E4D0
* Nag Mundari, 15.0; U+1C89 Cyrillic TJE, 16.0 — ages per `DerivedAge.txt`; all
* four are NFKC-stable and accepted here, and all four are `Cn` on that 3.9.6,
* which rejects them) is emitted bare and its tokenizer refuses the character,
* taking the whole SDK block down — the same parseability invariant
* {@link UNPRINTABLE}, {@link LONE_SURROGATE} and {@link MAX_LIST_NESTING}
* exist for. Both properties carry it: a character added only to `XID_Continue`
* passes the trailing `\p{XID_Continue}*` in a tail position and fails the same
* way — U+200C ZWNJ and U+200D ZWJ are that case, gaining `XID_Continue` in UCD
* 15.1 and absent from it in 13.0.0, 14.0.0 and 15.0.0, so `a\u{200C}b` is
* emitted bare here while `isidentifier()` is False on 3.9.6 and on 3.12.13
* (15.0.0). A CPython newer than the engine only routes a legal name to the
* subscript/`dict[str, Any]` path: less readable, still correct. The NFKC
* condition reduces to the same skew, since normalization stability guarantees
* an assigned character's normalization never changes afterwards.
*
* This predicate is not the only reader of engine tables. {@link camelCase}
* reads them at three further points — its split set, its head test, and its
* `toUpperCase()` case mapping — and this predicate's verdict gates none of
* them: a class name derived there reaches emitted text whenever any object
* shape in the tool's schema declares a `TypedDict`, including for a tool this
* predicate rejected. A tool named `zz-\u{1E4D0}x` with such parameters never
* reaches the skew here (the `-` rejects it outright) yet emits `class
* Zz\u{1E4D0}xArgs`, which that same 3.9.6 refuses — Nag Mundari arrived two
* releases after its tables. The case mapping is a separate table rather than
* an XID membership test, and it fails on names both conditions above accept:
* `\u{019B}` is XID_Start and NFKC-stable, so this predicate accepts it and
* `async def \u{019B}` compiles on 3.9.6, but Node uppercases it to
* `\u{A7DC}` — unassigned in that CPython, whose own `.upper()` is the identity
* here — and the declared `class \u{A7DC}Args` fails with `invalid
* non-printable character U+A7DC`. Closing the exposure therefore covers all
* four read points, not this predicate alone; it needs the target interpreter's
* version, which the backend reporting `language: 'python'` owns; the
* language-dispatch Agent Note records the deferral.
*
* The `ts-types` sibling keeps its own ASCII rule rather than sharing this
* one: ECMAScript identifiers are a different set (`$`) and are never
* normalized, so one predicate cannot be correct for both. ZWJ/ZWNJ are not
* part of that difference — both sets carry them on the engine's tables; what
* separates the two there is the CPython table version above.
* @param name - the raw schema field or tool name.
* @returns whether the name can be emitted bare.
*/
function isBareIdentifier(name) {
	return IDENTIFIER.test(name) && name.normalize("NFKC") === name;
}
/**
* Python hard keywords: reserved everywhere, so a tool or field named
* ``class`` or ``lambda`` is legal on the wire but not as an attribute
* (``tools.class`` would be a SyntaxError in the model program) and not as a
* class-syntax `TypedDict` field. Such a tool renders under subscript access
* and such an object degrades to ``dict[str, Any]`` — the model still reaches
* every tool and field without collisions.
* Soft keywords (``match``, ``case``, ``type``, ``_`` — the language
* reference's whole set) are deliberately ABSENT: each is special in exactly
* one syntactic position — a statement head (``match``, ``type``), a ``match``
* statement's clause head (``case``), or a pattern (``_``) — so ``match: str``
* as a field and ``async def match(...)`` as a method are both legal, and
* including them would needlessly degrade common search/regex tool fields to
* ``dict[str, Any]``. Underscore-leading names are handled separately, not
* here: a non-dunder ``__token`` name-mangles, a dunder present on
* ``object``/``type`` resolves before the proxy hook, and implicit
* special-method lookup bypasses the hook.
*/
const RESERVED = /* @__PURE__ */ new Set([
	"False",
	"None",
	"True",
	"and",
	"as",
	"assert",
	"async",
	"await",
	"break",
	"class",
	"continue",
	"def",
	"del",
	"elif",
	"else",
	"except",
	"finally",
	"for",
	"from",
	"global",
	"if",
	"import",
	"in",
	"is",
	"lambda",
	"nonlocal",
	"not",
	"or",
	"pass",
	"raise",
	"return",
	"try",
	"while",
	"with",
	"yield",
	"__debug__"
]);
/** `typing` symbols this module may emit, in the deterministic import order. */
const TYPING_ORDER = [
	"Any",
	"Literal",
	"NotRequired",
	"Protocol",
	"TypedDict"
];
/** `indent`-deep line prefix (four spaces per level to match PEP 8 output). */
function pad(indent) {
	return "    ".repeat(indent);
}
/**
* The `Cc` code points that survive the whitespace collapse in {@link describe}
* and have no printable form: the C0 controls, DEL, and the C1 controls. Only
* U+0009 to U+000D are absent, because ECMAScript `\s` already collapsed them —
* `\s` is TAB/VT/FF/SP/NBSP/ZWNBSP/Zs plus LF/CR/LS/PS, so no C1 code point is
* in it and the whole U+0080 to U+009F block reaches this rule intact. Those
* are not hypothetical input: they are what Windows-1252 bytes 0x80 to 0x9F
* (smart quotes, em dash) become when decoded as Latin-1.
* CPython rejects source containing a NUL outright
* (`SyntaxError: source code string cannot contain null bytes`), whether it
* sits in a docstring or in a comment, so one such byte anywhere in a schema
* description would make the whole generated SDK unparseable — under
* `mode: 'ptc'`, the model's only declaration of the tools. The rest are
* legal but invisible; escaping them with the same rule keeps the emitted text
* readable and the treatment uniform.
*
* The boundary is the category, not per-code-point addressability: `\xNN`
* addresses U+0000 to U+00FF, so one escape form covers `Cc` exactly. The
* invisible `Cf` formatting characters pass through by design — of them only
* U+00AD soft hyphen would fit `\xNN` at all, and escaping that one while
* U+200B ZWSP, U+200E/U+200F bidi marks, and U+2060 word joiner passed through
* would leave a rule that is neither category- nor addressability-shaped. The
* whole family is legal in both consumers, since only LF and CR terminate a
* Python string literal or a `#` comment. That set is the tokenizer's, not
* `str.splitlines()`': NEL (U+0085), LS (U+2028), and PS (U+2029) split a
* string at run time but do not end a physical line in source — measured on
* CPython 3.9.6 and 3.12.13, each accepted in both positions with the value
* round-tripping — so they are safe raw wherever they reach emitted text
* unescaped, which for all three is `JSON.stringify`, at two call sites:
* {@link pyScalar}'s literal path, and the subscript tool-name comment's own
* call, which a name carrying any of them always reaches, none being
* `XID_Continue`. The `description` path escapes NEL under the class above and
* folds LS and PS in {@link describe}'s `\s+` collapse, both being `\s`.
*/
const UNPRINTABLE = /[\u0000-\u0008\u000e-\u001f\u007f-\u009f]/g;
/**
* Unpaired surrogate code points, escaped by {@link describe} as `\uNNNN` —
* its own form, since `\xNN` stops at U+00FF. The `u` flag is what makes this
* the LONE ones: in Unicode mode a well-formed pair is a single astral code
* point outside D800 to DFFF, so an emoji in a description survives untouched.
*
* This is the NUL case from {@link UNPRINTABLE}, not the invisible-character
* case. Python source must be UTF-8-encodable and a lone surrogate is not, so
* `compile()` raises `UnicodeEncodeError: surrogates not allowed` for one
* anywhere in the text — measured on 3.9 for a string literal and for a `#`
* comment alike. A raw or MCP tool description reaches this: `JSON.parse` on a
* wire `"\ud800"` escape yields exactly such a code point.
*/
const LONE_SURROGATE = /[\ud800-\udfff]/gu;
/**
* The collapsed one-line `description` of a schema node (byte-stable across
* formatting churn), or `undefined` when the node carries none. Every caller
* passes an object — a validated property node, the `ToolSdkSchema` itself, or
* the `{ description }` wrapper {@link docLines} synthesizes — so only the
* description field needs guarding. A description that collapses
* to nothing (empty, or whitespace only) is `undefined` too: it documents the
* node no better than an absent one, and emitting it would leave an empty
* `"""` docstring or a bare `#   ` line in the SDK. Only ECMAScript whitespace
* folds, so a description of whitespace plus one surviving control character is
* NOT absent: it collapses to that character's visible escape.
*
* Control characters left over after the whitespace collapse are rendered as
* their `\xNN` escapes (see {@link UNPRINTABLE}) and unpaired surrogates as
* their `\uNNNN` escapes (see {@link LONE_SURROGATE}); the escape's own backslash is
* emitted literally by both consumers, since {@link docLines} doubles it into a
* Python source escape and a `#` comment carries it verbatim.
*/
function describe(schema) {
	const description = schema.description;
	if (typeof description !== "string") return void 0;
	const collapsed = description.replace(/\s+/g, " ").replace(UNPRINTABLE, (char) => `\\x${char.charCodeAt(0).toString(16).padStart(2, "0")}`).replace(LONE_SURROGATE, (char) => `\\u${char.charCodeAt(0).toString(16).padStart(4, "0")}`).trim();
	return collapsed.length === 0 ? void 0 : collapsed;
}
/**
* One-line docstring for a tool `description`, or no lines when there is none.
* Backslashes are doubled first, every quote is escaped, and a trailing
* backslash cannot survive: a description ending in `"` or an odd backslash
* would otherwise merge with (or escape) the closing triple quote and make
* the generated block — PTC mode's only SDK — syntactically invalid Python.
*/
function docLines(description, indent) {
	const collapsed = describe({ description });
	if (collapsed === void 0) return [];
	const escaped = collapsed.replaceAll("\\", "\\\\").replaceAll("\"", "\\\"");
	return [`${pad(indent)}"""${escaped}"""`];
}
/**
* CamelCase a name into a Python type identifier: non-identifier characters
* split words, `_` splits too (it is `XID_Continue`, so the split set names it
* explicitly), and a head that cannot start an identifier takes a `Tool`
* prefix. Unicode survives, so a `路径` field yields `路径`-based class names
* instead of collapsing to the bare prefix. A character that is not
* `XID_Continue` splits even when it is a letter, so a name whose NFKC folding
* would leave the identifier set is not carried through — the split set is the
* grammar's, not an ASCII approximation of it.
*
* The result is NFKC-normalized: these names are generated, never matched
* against a JSON key, so normalizing is free here and keeps what CPython
* compiles identical to what is emitted — unlike {@link isBareIdentifier},
* which must reject unstable names outright. Normalizing AFTER the prefix
* decision is what makes that hold at the seam the prefix creates: `Tool` +
* a combining-mark head composes there (`U+0301` gives `Tooĺ`, U+013A), so
* normalizing only the un-prefixed part would emit a name CPython compiles to
* a different symbol. The second call is idempotent on the un-prefixed arm.
*
* The split set, the head test, and `toUpperCase()` all read the engine's
* Unicode tables, so this function carries the same version skew
* {@link isBareIdentifier} documents, by paths independent of it: a class name
* derived here reaches emitted text whenever any object shape in the tool's
* schema declares a `TypedDict`, and the predicate's verdict on the tool name
* does not gate that. The case mapping is the one that can fail on a name the
* predicate accepted; the worked example is there.
* @param raw - the schema field or tool name to derive from.
* @returns a class-name segment safe to emit.
*/
function camelCase(raw) {
	const joined = raw.split(/[^\p{XID_Continue}]+|_+/u).filter((part) => part.length > 0).map((part) => `${part.charAt(0).toUpperCase()}${part.slice(1)}`).join("").normalize("NFKC");
	return (/^\p{XID_Start}/u.test(joined) ? joined : `Tool${joined}`).normalize("NFKC");
}
/** Class-name base cap keeping each emitted name — and total text — linear in schema depth. */
const MAX_CLASS_NAME_BASE = 120;
/**
* Deepest `list[…]` nesting emitted into one annotation before the item type
* degrades to `Any`. CPython's tokenizer rejects a logical line holding more
* than 200 simultaneously-open brackets (`MAXLEVEL`, `SyntaxError: too many
* nested parentheses`), so an array chain deeper than that would render an SDK
* block that is not valid Python at all — the same failure the docstring
* escaping in {@link docLines} exists to prevent. 180 leaves headroom for the
* few brackets an annotation can add around the chain, all of which count
* toward the same limit. Per emission site, counting brackets open at the
* chain's innermost point:
*
* - Return annotation, `async def f(self, args: X) -> chain:` — 180 `list[`
*   plus an innermost `Literal[`. The parameter list's `(` closed at the `)`
*   before the `->`, so it is NOT open here: 181.
* - TypedDict field, `field: NotRequired[chain]` — a class-body line with no
*   other open bracket, and its children start at `listDepth: 1` to reserve
*   the `NotRequired[`, so 179 `list[` plus `Literal[`: 181. Required fields
*   share that start for uniformity, spending one level of representable depth
*   on a bracket they never emit.
* - Argument annotation, `async def f(self, args: chain) -> Y:` — the `(` IS
*   still open around it: 180 `list[` plus `Literal[` plus the paren, 182, the
*   worst case. Reachable only through a raw `register()` whose `parameters`
*   is an array reached from the root through `oneOf` arms alone — the root
*   array itself, or one nested under any depth of unions, since an arm
*   inherits the enclosing depth unchanged (`A | B` opens no bracket). An
*   object ancestor takes it out of this case: its fields restart the chain at
*   the 181 site. `defineTool` compiles an object root, so the annotation is a
*   bare TypedDict class name or a one-bracket `dict[str, Any]` when that
*   object degrades — never a chain.
*
* A CPython grammar limit, not a deployment choice, so it is fixed rather than
* configurable. The sibling `ts-types` renderer needs no counterpart: nothing
* in the TypeScript grammar bounds nesting, and its SDK block is never type-
* checked. Only bracket nesting counts — a `oneOf` renders as a flat `A | B`
* chain and nested objects render as separate `class` statements, so neither
* accumulates open brackets at any depth. The invariant this cap serves is
* grammatical validity; see the `oneOf` arm in {@link renderType} for the one
* interpreter limit deliberately left uncapped.
*/
const MAX_LIST_NESTING = 180;
/**
* Cap a class-name base at {@link MAX_CLASS_NAME_BASE} (see the callers for
* why capping keeps the render linear). `slice` counts UTF-16 code units, so
* an astral character straddling the boundary would be cut in half and leave a
* lone surrogate — not an identifier character, and not even well-formed text;
* drop it rather than emit it.
*/
function capClassNameBase(base) {
	if (base.length <= MAX_CLASS_NAME_BASE) return base;
	const capped = base.slice(0, MAX_CLASS_NAME_BASE);
	return /[\uD800-\uDBFF]$/.test(capped) ? capped.slice(0, -1) : capped;
}
/**
* Reserve a unique class name from a base, suffixing `2`, `3`, … on collision.
* The base is capped at {@link MAX_CLASS_NAME_BASE} first: child class names
* derive from their parent's allocated name (`ParentChild`), so an unbounded
* schema of single-field objects would otherwise grow each name by one field
* per level and the sum of all names to Θ(depth²). Capping the base keeps each
* name — and the total emitted text — linear in depth. Collisions resume from
* the per-base counter in `state.nextClassCounter` rather than rescanning from
* `2`, so a deep chain sharing one capped base stays O(1) per allocation
* (amortized) instead of Θ(depth²) in time.
*/
function allocateClassName(base, state) {
	const capped = capClassNameBase(base);
	let name = capped;
	if (state.usedClassNames.has(name)) {
		let n = state.nextClassCounter.get(capped) ?? 2;
		while (state.usedClassNames.has(`${capped}${n}`)) n++;
		name = `${capped}${n}`;
		state.nextClassCounter.set(capped, n + 1);
	}
	state.usedClassNames.add(name);
	return name;
}
/**
* Append a child-name segment to a parent class-name base, capping the result
* at {@link MAX_CLASS_NAME_BASE}. Capping AT PROPAGATION (not only inside
* {@link allocateClassName}) keeps each level O(1): a deep `oneOf`- or
* object-chain would otherwise carry an ever-growing ConsString down the tree
* and re-materialize it (via `.length`/`.slice`) at every level — Θ(depth²).
* The bounded base plus the collision counter still yields unique names.
*
* The join is NFKC-normalized because both sides are separately normalized yet
* their concatenation need not be: a base ending in a Hangul L jamo or LV
* syllable composes with a following V or T jamo head (`가` + `ᆨ` gives `각`),
* so the emitted class name would differ from the symbol CPython compiles, and
* two byte-distinct names could fold onto one — `usedClassNames` dedupes by the
* raw bytes, so the collision counter would not see it. Normalizing costs
* O(cap + segment) per level, the same order as the `slice` it feeds. The other
* two join points need no counterpart: `Args`/`Output` start with `A`/`O` and
* {@link allocateClassName}'s suffix is digits, none of which compose backwards.
*/
function childClassName(base, segment) {
	return capClassNameBase(`${base}${segment}`.normalize("NFKC"));
}
/**
* Render one validated scalar as Python literal text (`True`/`False`,
* JSON-quoted strings, bare numbers). `null` cannot reach here: the `null`
* type renders directly as `None`, and the unified validator rejects a null
* `const`/`enum` entry on every other scalar type.
*
* A beyond-safe-range integral number takes `BigInt` digits rather than
* `String`: Python integers are arbitrary-precision, so the emitted digits ARE
* the value the model programs against, and `String` can give a different
* integer than the double holds (`2 ** 60` prints the rounded `...847000`, not
* the exact `...846976`) or no integer literal at all (`1e21` prints `1e+21`).
* `String`'s rounding is not a bug in it: `Number::toString` emits the shortest
* decimal string that re-reads to the same double, then pads to the exponent
* with zeros (1 significant digit for `1e20`, 16 for `2 ** 60`) — and when the
* shortest string is shorter than the double's exact value, those padded digits
* name an integer no double holds. Passing one back would have to cross the
* argument boundary as a JSON number — a double again — so the SDK would
* document a value no program can pass. `BigInt` needs no case split: where
* `String` is already exact (`2 ** 53`, `1e20`) the two agree byte for byte,
* and where it is not, `BigInt` is the exact one. The TS flavor needs no
* counterpart at all: its literal is re-read by a JS parser back into the same
* double.
*
* `JSON.stringify` is also what keeps this path's output parseable, and it is
* the only thing that does. It covers both classes of hazard: the two kinds of
* code point CPython refuses anywhere in source — NUL among the C0 controls,
* and the whole D800–DFFF unpaired-surrogate block, escaped under ES2019
* well-formed stringification, which the engines range guarantees — and the
* ones that break this line in particular, a bare `"` closing the literal
* early, a trailing odd backslash eating the closing quote, and a bare LF/CR
* ending it before its terminator. The `description` path carries
* {@link UNPRINTABLE} and {@link LONE_SURROGATE} because nothing quotes it,
* and folds newlines in {@link describe}.
*
* That leans on a coincidence worth naming: every escape `JSON.stringify` can
* emit (`\"`, `\\`, `\b`, `\f`, `\n`, `\r`, `\t`, `\uXXXX`) is also a Python
* escape denoting the same character, so the emitted `Literal[...]` both
* parses and decodes back to the value the schema declared. DEL, the C1
* controls (NEL among them), and LS/PS (U+2028/U+2029) do reach it raw —
* legal but invisible, byte-for-byte as in the TS flavor; escaping them is a
* both-flavors change. Those last three are legal here for the reason
* {@link UNPRINTABLE} records: they are `str.splitlines()` boundaries, not
* tokenizer line terminators. The subscript tool-name comment quotes its name
* through its own call to the same `JSON.stringify`, never through this
* function, and inherits both halves — escapes and pass-throughs alike.
*/
function pyScalar(value) {
	if (value === true) return "True";
	if (value === false) return "False";
	if (typeof value === "string") return JSON.stringify(value);
	if (typeof value === "number" && Number.isInteger(value) && !Number.isSafeInteger(value)) return BigInt(value).toString();
	return String(value);
}
/**
* Render a validated scalar `const`/`enum` as `Literal[...]`, falling back to
* the broad type. Deliberately deviates from PEP 586, which restricts `Literal`
* parameters to int/bool/str/bytes/enum/None: a non-integral number
* `const`/`enum` emits a float literal (`Literal[1.5]`) a strict checker would
* reject. An integral one does not deviate — {@link pyScalar} emits int digits,
* including for the beyond-safe-range values it widens through `BigInt`, and
* PEP 586 admits int parameters. Harmless either way — the stub is advisory
* prompt text, only required to parse — and keeping the exact value
* communicates the constraint to the model.
*/
function renderConstrainedScalar(node, broad, state) {
	if (node.const !== void 0) {
		state.typing.add("Literal");
		return `Literal[${pyScalar(node.const)}]`;
	}
	if (node.enum !== void 0) {
		state.typing.add("Literal");
		return `Literal[${node.enum.map(pyScalar).join(", ")}]`;
	}
	return broad;
}
/**
* Map one JSON-Schema node to a Python type expression, threading `state` to
* collect the `TypedDict` declarations and `typing` symbols a full render
* needs. `className` is the name to give an object node with properties (and
* the prefix for its nested objects). Handles every unified schema construct —
* `oneOf` (→ `X | Y`), `const`/`enum` (→ `Literal[...]`), `integer` (→ `int`),
* `null` (→ `None`) — and degrades an unsupported or malformed schema to `Any`
* without throwing, the same trusted-after-validation stance as the sibling
* {@link ./ts-types.ts | ts-types} renderer. {@link jsonSchemaToPy} is the
* context-free entry point; this is the collecting core.
*/
function renderType(schema, className, state) {
	const newFrame = (schema, className, listDepth) => ({
		schema,
		className,
		phase: "start",
		listDepth,
		children: [],
		childIndex: 0,
		childTypes: [],
		entries: []
	});
	try {
		assertSupportedJsonSchema(schema);
		const frames = [newFrame(schema, className, 0)];
		let result;
		const finish = (type) => {
			frames.pop();
			const parent = frames.at(-1);
			if (parent === void 0) result = type;
			else parent.childTypes.push(type);
		};
		while (frames.length > 0) {
			const frame = frames.at(-1);
			/* v8 ignore next -- the loop condition guarantees a current frame. */
			if (frame === void 0) break;
			if (frame.phase === "children") {
				if (frame.childIndex < frame.children.length) {
					const child = frame.children[frame.childIndex];
					/* v8 ignore next -- childIndex is bounded by children.length. */
					if (child === void 0) throw new Error("missing python render child");
					frame.childIndex++;
					frames.push(newFrame(child.schema, child.className, child.listDepth));
					continue;
				}
				if (frame.kind === "oneOf") {
					let union = "";
					for (const [index, childType] of frame.childTypes.entries()) union = index === 0 ? childType : `${union} | ${childType}`;
					finish(union);
					continue;
				}
				if (frame.kind === "array") {
					/* v8 ignore next -- the ?? arm needs a childless array frame, which start never builds. */
					finish(`list[${frame.childTypes[0] ?? "Any"}]`);
					continue;
				}
				const node = frame.node;
				const name = frame.allocated;
				/* v8 ignore next -- typeddict frames always set node and allocated at start. */
				if (node === void 0 || name === void 0) throw new Error("missing typeddict frame state");
				const required = new Set(node.required);
				const lines = [`class ${name}(TypedDict):`];
				for (let index = 0; index < frame.entries.length; index++) {
					const entry = frame.entries[index];
					const fieldType = frame.childTypes[index];
					/* v8 ignore next -- entries and childTypes correspond one-to-one. */
					if (entry === void 0 || fieldType === void 0) throw new Error("missing typeddict field type");
					const [field, fieldSchema] = entry;
					const description = describe(fieldSchema);
					if (description !== void 0) lines.push(`${pad(1)}# ${description}`);
					if (required.has(field)) lines.push(`${pad(1)}${field}: ${fieldType}`);
					else {
						state.typing.add("NotRequired");
						lines.push(`${pad(1)}${field}: NotRequired[${fieldType}]`);
					}
				}
				if (node.additionalProperties !== false) lines.push(`${pad(1)}# Additional keys beyond those declared are allowed.`);
				if (lines.length === 1) lines.push(`${pad(1)}pass`);
				state.classes.push(lines.join("\n"));
				finish(name);
				continue;
			}
			frame.phase = "children";
			const node = frame.schema;
			if (node.oneOf !== void 0) {
				frame.kind = "oneOf";
				frame.children = node.oneOf.map((branch, index) => ({
					schema: branch,
					className: childClassName(frame.className, `${index + 1}`),
					listDepth: frame.listDepth
				}));
				continue;
			}
			if (node.type === void 0) {
				state.typing.add("Any");
				finish("Any");
				continue;
			}
			switch (node.type) {
				case "string":
					finish(renderConstrainedScalar(node, "str", state));
					break;
				case "number":
					finish(renderConstrainedScalar(node, "float", state));
					break;
				case "integer":
					finish(renderConstrainedScalar(node, "int", state));
					break;
				case "boolean":
					finish(renderConstrainedScalar(node, "bool", state));
					break;
				case "null":
					finish("None");
					break;
				case "array":
					if (node.items === void 0) {
						state.typing.add("Any");
						finish("list[Any]");
						break;
					}
					if (frame.listDepth >= MAX_LIST_NESTING) {
						state.typing.add("Any");
						finish("Any");
						break;
					}
					frame.kind = "array";
					frame.children = [{
						schema: node.items,
						className: frame.className,
						listDepth: frame.listDepth + 1
					}];
					break;
				case "object": {
					const entries = Object.entries(node.properties ?? {});
					if (className === "" || !entries.every(([name]) => isBareIdentifier(name) && !RESERVED.has(name) && !(name.startsWith("__") && !name.endsWith("__")))) {
						state.typing.add("Any");
						finish("dict[str, Any]");
						break;
					}
					if (entries.length === 0 && node.additionalProperties !== false) {
						state.typing.add("Any");
						finish("dict[str, Any]");
						break;
					}
					frame.kind = "typeddict";
					frame.node = node;
					frame.allocated = allocateClassName(frame.className, state);
					state.typing.add("TypedDict");
					frame.entries = entries;
					/* v8 ignore next -- allocated is always set before children are built. */
					frame.children = entries.map(([field, child]) => ({
						schema: child,
						className: childClassName(frame.allocated ?? "", camelCase(field)),
						listDepth: 1
					}));
					break;
				}
				/* v8 ignore next 4 -- assertSupportedJsonSchema narrowed this closed type union. */
				default:
					state.typing.add("Any");
					finish("Any");
			}
		}
		/* v8 ignore next -- every root frame produces one expression. */
		return result ?? "Any";
	} catch {
		state.typing.add("Any");
		return "Any";
	}
}
/** The fixed model-facing usage contract rendered above the declarations. */
const SDK_INSTRUCTIONS = `## Writing code for run_code

\`run_code\` takes two required arguments: \`code\` — the body of an async Python function (top-level \`await\` and \`return\` both work) — and \`description\`, a short summary of what the program does. At run time exactly two of the names declared below are bound: \`tools\` and \`ToolCallError\`. Everything else is a STATIC STUB describing argument and return types — in particular the \`TypedDict\` classes do NOT exist at run time, so build arguments as plain \`dict\`/\`list\` JSON values: \`await tools.name({"field": 1})\`, never \`FooArgs(field=1)\`, which raises \`NameError\`. Inside the program:

- Call tools as \`await tools.name(args)\` — subscript access for exotic, reserved, or underscore-leading names: \`await tools["my-tool"](args)\`. Every call resolves to the tool's typed canonical JSON value (each method's return type below). Tool arguments must be lossless JSON.
- A FAILED tool call raises \`ToolCallError\`, whose \`toolName\` identifies the failed tool and whose message is human-readable — wrap in \`try/except\` to handle and continue.
- Independent read-only calls MAY overlap under \`asyncio.gather\` (safe calls run concurrently; mutating calls run alone, in submission order). Sequence dependent work with \`await\`.
- Emit the run's answer with \`print(...)\` and/or a top-level \`return <value>\`; the returned value must be lossless JSON. Only what you print and return is program output. A successful tool result containing an image is attached after the run so you can inspect it on the next step; every other intermediate result stays out of the conversation, so extract just what you need.

The available tools:`;
/**
* Render the full `tools:sdk` prompt section under `runtime.language ===
* 'python'`: the Python-flavored usage instructions plus one named `TypedDict`
* per tool argument or output object (and per nested object) and one awaitable
* method per visible tool on a `Tools` protocol — typed args in, the tool's
* canonical output value out — with a `tools: Tools` singleton the model calls
* into. The `typing` import line lists exactly the symbols the render used.
* Deterministic — tools are emitted in lexicographic name order, and class
* declarations precede the protocol in that same order (nested classes before
* the parent that references them), so an unchanged tool set produces
* byte-identical text across assemblies. The sort is not a total order on
* byte-equal names, so two schemas sharing a name would render in argument
* order; the caller's visible-capability map is keyed by name, so the input
* never carries a duplicate.
* @param schemas - the tool schemas plus canonical output schemas to declare
*   (the caller excludes `run_code` itself).
* @returns the complete section text.
*/
function renderToolsSdkPy(schemas) {
	const sorted = [...schemas].sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
	const state = {
		classes: [],
		usedClassNames: /* @__PURE__ */ new Set(),
		nextClassCounter: /* @__PURE__ */ new Map(),
		typing: /* @__PURE__ */ new Set(["Protocol"])
	};
	const members = [];
	let statements = 0;
	for (const schema of sorted) {
		const argType = renderType(schema.parameters, `${camelCase(schema.name)}Args`, state);
		const outputType = renderType(schema.output, `${camelCase(schema.name)}Output`, state);
		if (isBareIdentifier(schema.name) && !RESERVED.has(schema.name) && !schema.name.startsWith("_")) {
			const doc = docLines(schema.description, 2);
			members.push(doc.length > 0 ? `${pad(1)}async def ${schema.name}(self, args: ${argType}) -> ${outputType}:` : `${pad(1)}async def ${schema.name}(self, args: ${argType}) -> ${outputType}: ...`);
			members.push(...doc);
			statements += 1;
		} else {
			members.push(`${pad(1)}# tools[${JSON.stringify(schema.name)}](args: ${argType}) -> ${outputType}`);
			const description = describe(schema);
			if (description !== void 0) members.push(`${pad(1)}#   ${description}`);
		}
	}
	const body = (statements > 0 ? members : [`${pad(1)}pass`, ...members]).join("\n");
	const imports = TYPING_ORDER.filter((symbol) => state.typing.has(symbol));
	const classBlock = state.classes.length > 0 ? `${state.classes.join("\n\n")}\n\n` : "";
	return `${SDK_INSTRUCTIONS}\n\n\`\`\`python\n${`from typing import ${imports.join(", ")}\n\nclass ToolCallError(Exception):
    toolName: str\n\n${classBlock}class Tools(Protocol):\n${body}\n\ntools: Tools`}\n\`\`\``;
}
/**
* Tool registry, model presentation modes, and pre/guard/around/post/result
* execution pipeline.
* @module @deepseek-ai/dsh-tools
*/
/**
* Language → SDK-section renderer. The registry looks up the loaded
* `ctx.ptcRuntime.language` in this table when assembling the `tools:sdk`
* section under a non-native mode; a runtime whose language is not a key
* fails the assembly loudly (same idiom as `toolOrder` violations). Adding a
* new backend language is three parallel edits — a {@link PtcSdkLanguage}
* member, an entry here, and a `RUN_CODE_FLAVORS` entry in `ptc.ts` for
* its `run_code` schema strings — plus the renderer function this table points
* at. The `satisfies` clause pins this table's key set to that union, which
* the flavor table is checked against too, so any of the three left out is a
* typecheck failure. What no check reaches is the prose that names the values
* instead of deriving them: the seam's `dsh-ptc-runtime` README pair, its
* `PtcRuntime.language` JSDoc, and `docs/subsystems/ptc-runtime.md`
* with its zh pair, plus this package's own README pair and the
* {@link Config.mode} JSDoc.
*/
/**
* The model-facing statement of the `ptc` collapse. Names the consequence
* (the call fails) and the route (inside the program), because a rule the
* model can only discover by being denied is one it corrects too late.
*/
const PTC_ONLY_INSTRUCTION = `\`${RUN_CODE_NAME}\` is the only tool you can call directly — a tool call naming any other tool fails. Reach every tool the SDK declares below from inside the program.`;
const SDK_RENDERERS = {
	typescript: renderToolsSdk,
	python: renderToolsSdkPy
};
/**
* Scheduler entry point omitted from the generated named service API.
* @internal
*/
const TOOL_RUNTIME_SCHEDULER = Symbol("@deepseek-ai/dsh-tools.scheduler");
/** Canonical error code for cancellation after a tool body was invoked. */
const TOOL_ABORTED = "ABORTED";
/** Canonical error code for cancellation before a tool body was invoked. */
const TOOL_ABORTED_BEFORE_DISPATCH = "ABORTED_BEFORE_DISPATCH";
/**
* Thrown (internally) when the model requests a tool that isn't registered.
* Extends {@link HarnessError} (`code: 'UNKNOWN_TOOL'`) so an unknown-tool
* failure is as routable as a tool-thrown one — retry/sandbox/replay code can
* distinguish it from a tool body's own error.
*/
var ToolNotFoundError = class extends HarnessError {
	/**
	* @param toolName - the name the caller asked for.
	* @param reachableFrom - how the model reaches this tool instead, when the
	*   name IS visible and only the presentation denies calling it directly.
	*   Omitted for a name that is registered nowhere.
	*/
	constructor(toolName, reachableFrom) {
		super(reachableFrom === void 0 ? `unknown tool "${toolName}"` : `unknown tool "${toolName}": ${reachableFrom}`, "UNKNOWN_TOOL");
		this.name = "ToolNotFoundError";
	}
};
/** Thrown when a tool body or post-policy value violates its declared output. */
var ToolOutputError = class extends HarnessError {
	/** Schema/value violations in validation order. */
	violations;
	constructor(toolName, violations) {
		super(`tool "${toolName}" returned invalid output: ${violations.join("; ")}`, "INVALID_TOOL_OUTPUT");
		this.name = "ToolOutputError";
		this.violations = violations;
	}
};
/** Convert one projector exception into the canonical invalid-output failure. */
function projectionError(toolName, projector, error) {
	return new ToolOutputError(toolName, [`output.${projector} failed: ${errorMessage(error)}`]);
}
/** Snapshot one projector result before later durable-result materialization. */
function snapshotProjection(toolName, projector, candidate) {
	try {
		const detached = snapshotJsonValue(candidate);
		if (detached === void 0) throw new ToolOutputError(toolName, [`output.${projector} returned non-lossless JSON`]);
		return detached;
	} catch (error) {
		if (error instanceof ToolOutputError) throw error;
		throw projectionError(toolName, projector, error);
	}
}
/** Snapshot one body or policy value into the canonical invalid-output failure class. */
function snapshotToolValue(toolName, candidate) {
	try {
		const detached = snapshotJsonValue(candidate);
		if (detached === void 0) throw new ToolOutputError(toolName, ["value is not lossless JSON"]);
		return detached;
	} catch (error) {
		if (error instanceof ToolOutputError) throw error;
		throw new ToolOutputError(toolName, [`value snapshot failed: ${errorMessage(error)}`]);
	}
}
/**
* Best-effort human-readable message from an arbitrary thrown value: Error
* instances use `.message`; non-Error objects with a string `message`
* property (e.g. `throw { message: 'denied' }`) use it too; everything else
* is stringified.
*/
function errorMessage(error) {
	try {
		if (error instanceof Error) return error.message;
		if (typeof error === "object" && error !== null && "message" in error && typeof error.message === "string") return error.message;
		return String(error);
	} catch {
		return "<unprintable thrown value>";
	}
}
/** Derive one failure message from policy feedback without changing its rendered blocks. */
function failureMessageFromContent(content) {
	const text = content.map((block) => block.type === "text" ? block.text : `[${block.type} content]`).join("\n");
	return text.length > 0 ? text : "tool result blocked by post-execute policy";
}
/** Snapshot and freeze one durable tool-result projection or reject lossy data. */
function materializePresentation(candidate) {
	const detached = snapshotJsonValue(candidate);
	if (detached === void 0) throw new TypeError("tool result must be losslessly JSON-serializable");
	return deepFreeze(detached);
}
/** Structured `{ name, code }` for a thrown HarnessError, else undefined. */
function errorInfo(error) {
	try {
		return error instanceof HarnessError ? {
			name: error.name,
			code: error.code
		} : void 0;
	} catch {
		return;
	}
}
/** One scope's complete tool-registry contribution. */
var ToolLayer = class {
	tools;
	restrictions = new AnonymousEntries();
	guards = new AnonymousEntries();
	/**
	* Presentation this scope's agent declared for itself, shadowing the
	* deployment default. One cell rather than an entry table: two answers to
	* "which form does the model see" is a contradiction, not a merge.
	*/
	mode;
	constructor(scope) {
		this.tools = new NamedEntries((name) => /* @__PURE__ */ new Error(scope === void 0 ? `tool "${name}" is already registered (for a per-agent variant, register through that agent's \`agent.ctx\` instead)` : `tool "${name}" is already registered in this scope`));
	}
	/** Whether every contribution table in this aggregate layer is empty. */
	isEmpty() {
		return this.tools.isEmpty() && this.restrictions.isEmpty() && this.guards.isEmpty() && this.mode === void 0;
	}
	/** Whether every compiled restriction in this layer admits a global tool name. */
	admits(name) {
		for (const filter of this.restrictions.values()) if (filter.allow !== void 0 && !filter.allow.has(name) || filter.deny !== void 0 && filter.deny.has(name)) return false;
		return true;
	}
	/** First monotonic denial from this layer's live guard registrations. */
	guardReason(exec) {
		for (const guard of this.guards.values()) {
			const reason = guard(exec);
			if (reason !== void 0) return reason;
		}
	}
};
/** Resolve the run_code overlap cap at the owning config boundary (direct construction bypasses the Loader schema). */
function resolveMaxParallelSubCalls(value) {
	const maxParallelSubCalls = value ?? 10;
	if (!Number.isInteger(maxParallelSubCalls) || maxParallelSubCalls < 1) throw new Error("maxParallelSubCalls must be a positive integer");
	return maxParallelSubCalls;
}
/**
* Tool registry and execution pipeline. Scoped registrations shadow globals;
* one visibility resolver feeds presentation, lookup, and dispatch.
*/
var ToolRuntime = class extends Service {
	static inject = ["systemPrompt"];
	static Config = z.object({
		mode: z.union([
			"native",
			"ptc",
			"both"
		]).default("native"),
		maxParallelSubCalls: z.natural().min(1).default(10)
	});
	/** Internal staged view consumed by `dsh-agent-loop`'s parallel scheduler. */
	[TOOL_RUNTIME_SCHEDULER] = {
		prepare: (exec) => this.prepareScheduledExecution(exec),
		dispatch: (exec) => this.dispatchScheduledExecution(exec),
		finalize: (exec, result) => this.finalizeScheduledExecution(exec, result),
		finish: (exec, result) => this.finishScheduledExecution(exec, result)
	};
	/** Context deferred by a running tool body, keyed by its scheduler-owned execution. */
	deferredContexts = /* @__PURE__ */ new WeakMap();
	/** Executions whose tool body declared the current turn complete. */
	concludingExecutions = /* @__PURE__ */ new WeakSet();
	/** Original caller cancellation, kept outside the wrapper-mutable execution object. */
	cancellationStates = /* @__PURE__ */ new WeakMap();
	/** Definition-owned final content transform snapshotted before policy begins. */
	contentFinalizers = /* @__PURE__ */ new WeakMap();
	/** Execution-prepared content installed before post-execute policy. */
	contentProjectors = /* @__PURE__ */ new WeakMap();
	layers = new ScopedLayers((scope) => new ToolLayer(scope), () => {
		this.ctx.emit("tools/change");
	});
	/** Presentation for scopes that declare none; {@link presentAs} shadows it per scope. */
	defaultMode;
	maxParallelSubCalls;
	/**
	* Reserved presentation transport, kept outside the filterable registration
	* layers. Built on first need rather than at construction: which agents run
	* a PTC mode is no longer known when the service is constructed, and the
	* transport is stateless beyond its closures over `this`.
	*/
	ptcTransport;
	constructor(ctx, config = {}) {
		super(ctx, "tools");
		this.defaultMode = config.mode ?? "native";
		this.maxParallelSubCalls = resolveMaxParallelSubCalls(config.maxParallelSubCalls);
		ctx.systemPrompt.tools((context) => this.wireSchemas(context.scope));
		if (this.defaultMode !== "native") {
			ctx.systemPrompt.section(this.collapseSection());
			ctx.systemPrompt.section(this.sdkSection());
		}
	}
	/**
	* The prompt statement of the `ptc` executor collapse, registered wherever
	* {@link sdkSection} is and rendering empty outside an effective `ptc`.
	*
	* Every tool contributes its own guidance section naming its tool, none of
	* them qualify how that tool is reached, and they all render before the SDK.
	* Without this the model reads a catalog of tools it is told to use and no
	* statement that only `run_code` may be called, so it emits a native call,
	* receives `UNKNOWN_TOOL` for a tool the prompt just declared, and concludes
	* the deployment is inconsistent. Its order places the rule before that
	* guidance rather than after it.
	*
	* `both` renders empty: native calls do execute there, so the rule is false.
	* @returns the section registration.
	*/
	collapseSection() {
		return {
			name: "tools:ptc-only",
			order: this.ctx.systemPrompt.getSectionOrder("PTC_ONLY"),
			text: (context) => this.modeFor(context.scope) === "ptc" ? PTC_ONLY_INSTRUCTION : ""
		};
	}
	/**
	* The generated-SDK prompt section, registered globally by a PTC mode
	* deployment and per scope by {@link presentAs}.
	*
	* The body regenerates from the CALLING scope, and renders empty for an
	* agent presenting natively — an agent that opted out under a PTC mode
	* deployment still sees the global registration, and an empty section is
	* dropped from the rendered prompt.
	* @returns the section registration.
	*/
	sdkSection() {
		return {
			name: "tools:sdk",
			order: this.ctx.systemPrompt.getSectionOrder("TOOLS_SDK"),
			interpolate: false,
			text: (context) => {
				const mode = this.modeFor(context.scope);
				if (mode === "native") return "";
				const runtime = this.requirePtcRuntime(mode);
				const render = SDK_RENDERERS[runtime.language];
				/* v8 ignore next -- requirePtcRuntime rejects an unknown language before this runs. */
				if (render === void 0) throw new Error(`dsh-tools: no SDK renderer for ${runtime.language}`);
				return render(this.sdkSchemas(context.scope));
			}
		};
	}
	/**
	* The presentation one scope's agent sees: its own declaration, else the
	* deployment default.
	* @param scope - the calling agent, or undefined for the global view.
	* @returns the resolved presentation mode.
	*/
	modeFor(scope) {
		const layers = this.layers.chainLayers(scope);
		for (let index = layers.length - 1; index >= 0; index -= 1) {
			const mode = layers[index]?.mode;
			if (mode !== void 0) return mode;
		}
		return this.defaultMode;
	}
	/**
	* The reserved `run_code` transport, built on first need.
	*
	* It never enters the global layer: per-agent restrictions must not remove
	* it, and a scoped registration must not shadow it. The visibility resolver
	* appends it after resolving the filterable global/scoped capability layers,
	* and only for scopes whose mode actually presents it.
	* @returns the shared transport definition.
	*/
	requirePtcTransport() {
		this.ptcTransport ??= createRunCodeTool(this, {
			requireRuntime: () => this.requirePtcRuntime(this.defaultMode),
			peekApprover: () => this.ctx.get("approval"),
			resolveSandboxPolicy: (exec) => {
				const policy = this.ctx.get("sandboxPolicy");
				if (policy === void 0) throw new Error("dsh-tools: confined PTC runtime requires sandboxPolicy");
				return policy.resolve(exec.agent === void 0 ? {} : { session: exec.agent.session });
			},
			peekRuntime: () => this.ctx.get("ptcRuntime"),
			maxParallel: this.maxParallelSubCalls,
			shapeDispatchLog: (dispatch) => this.shapeDispatchLog(dispatch)
		});
		return this.ptcTransport;
	}
	/**
	* Present the calling scope's tools in `mode` instead of the deployment
	* default. Nearest scope on the chain wins, so a preset's standing
	* declaration covers every agent joined under it.
	*
	* Scoped only, and one declaration per scope: this is how an agent preset
	* composes PTC mode agents beside native ones in the same process, and a
	* process-global override would be the `mode` config field instead.
	* @param mode - the presentation the covered agents' models see.
	* @returns the exact disposer that restores the deployment default.
	*/
	presentAs(mode) {
		const ctx = this.ctx;
		if (scopeOf(ctx) === void 0) throw new Error("tools.presentAs() requires a scoped context (agent.ctx): a context-global presentation is the `mode` config field on the tools row");
		return ctx.effect(function* () {
			yield this.layers.effect(ctx, (layer) => {
				if (layer.mode !== void 0) throw new Error(`tools.presentAs("${mode}") conflicts with "${layer.mode}" already declared for this scope; one composition selects one presentation`);
				layer.mode = mode;
				return () => {
					layer.mode = void 0;
				};
			}, { label: "tools.presentAs()" });
			if (mode !== "native") {
				yield ctx.systemPrompt.section(this.collapseSection());
				yield ctx.systemPrompt.section(this.sdkSection());
			}
		}.bind(this), "tools.presentAs()");
	}
	/**
	* Build one scope's wire schemas and names for prompt-order validation.
	* Restrictions do not make known tools invalid, but a mode collapse does.
	*/
	wireSchemas(scope) {
		const view = this.view(scope);
		const mode = this.modeFor(scope);
		if (mode === "native") return {
			schemas: [...view.visible.values()].map((definition) => this.schemaOf(definition, false)),
			knownNames: [...view.knownNames]
		};
		this.requirePtcRuntime(mode);
		const schemas = [...view.visible.values()].map((definition) => this.schemaOf(definition, false));
		if (mode === "ptc") return {
			schemas: schemas.filter((schema) => schema.name === RUN_CODE_NAME),
			knownNames: [RUN_CODE_NAME]
		};
		return {
			schemas,
			knownNames: [...view.knownNames, RUN_CODE_NAME]
		};
	}
	/**
	* Resolve the PTC runtime or throw the actionable misconfiguration error.
	* Read at use time (assembly / run_code execution), NOT via static
	* `inject`: an inject entry would hold `ctx.tools` — and every tool plugin
	* behind it — hostage to a PTC runtime existing even under `mode:
	* 'native'`.
	*
	* Assembly and `run_code` execution read separately, so the language is not
	* bound to a request. Harmless while one published backend exists — both
	* reads return the same flavor — but a reload that swapped in a second
	* language between them would hand a program written against one SDK to the
	* other. Binding it is deferred until a second backend ships (the first
	* point it is testable).
	*/
	requirePtcRuntime(mode) {
		const runtime = this.ctx.get("ptcRuntime");
		if (!runtime) throw new Error(`dsh-tools: mode "${mode}" requires a PTC runtime — load a ctx.ptcRuntime implementation (e.g. @deepseek-ai/dsh-ptc-runtime-node) or set tools mode to "native"`);
		if (!Object.hasOwn(SDK_RENDERERS, runtime.language)) {
			const known = Object.keys(SDK_RENDERERS).map((name) => JSON.stringify(name)).join(", ");
			throw new Error(`dsh-tools: no SDK renderer registered for runtime language ${JSON.stringify(runtime.language)} (known: ${known})`);
		}
		return runtime;
	}
	/**
	* Register globally or in the calling agent scope. Scoped tools shadow
	* globals; duplicates within one layer and the reserved `run_code` name fail.
	* @param definition - tool schema, execution, and optional finalization/presentation callbacks.
	* @returns the exact disposer that unregisters the tool.
	*/
	register(definition) {
		const name = definition.name;
		const output = definition.output;
		if (output === void 0 || typeof output !== "object" || typeof output.render !== "function" || output.presentationMeta !== void 0 && typeof output.presentationMeta !== "function") throw new TypeError(`tool "${name}" must declare output { schema, render, presentationMeta? }`);
		assertSupportedJsonSchema(output.schema);
		const timeoutMs = definition.timeoutMs;
		if (timeoutMs !== void 0 && (!Number.isFinite(timeoutMs) || timeoutMs <= 0)) throw new TypeError(`tool "${name}" timeoutMs must be a positive finite number`);
		if (name === "run_code") throw new Error(`tool name "${RUN_CODE_NAME}" is reserved for the PTC mode presentation transport and cannot be registered or shadowed`);
		return this.layers.effect(this.ctx, (layer) => layer.tools.insert(name, definition), { label: "tools.register()" });
	}
	/**
	* Restrict global tools for the calling agent scope. Empty filters, unknown
	* names, scope-local names, and reserved transport names fail. Restrictions
	* intersect; scoped registrations remain visible.
	* @param filter - global-tool mask: `allow` (keep only) and/or `deny` (remove).
	* @returns the exact disposer that lifts this restriction.
	*/
	restrict(filter) {
		const scope = scopeOf(this.ctx);
		if (scope === void 0) throw new Error("tools.restrict() requires a scoped context (agent.ctx): a context-global restriction would mask every agent — deny the tool for the intended agent instead");
		const allow = filter.allow;
		const deny = filter.deny;
		if (allow === void 0 && deny === void 0) throw new Error("tools.restrict({}) is a no-op: pass `allow` and/or `deny` (an empty filter is almost always a materialized-empty-config bug)");
		const compiled = {
			...allow !== void 0 ? { allow: new Set(allow) } : {},
			...deny !== void 0 ? { deny: new Set(deny) } : {}
		};
		if ([...allow ?? [], ...deny ?? []].includes("run_code")) throw new Error(`tools.restrict() cannot name reserved PTC mode presentation transport "${RUN_CODE_NAME}"; restrict end-capability tools instead`);
		const known = this.view(scope).restrictableNames;
		const unknown = [...allow ?? [], ...deny ?? []].filter((name) => !known.has(name));
		if (unknown.length > 0) throw new Error(`tools.restrict() names unknown global tool${unknown.length > 1 ? "s" : ""} ${unknown.map((n) => `"${n}"`).join(", ")}; known global tools: ${[...known].sort().join(", ") || "(none)"}`);
		return this.layers.effect(this.ctx, (layer) => layer.restrictions.append(compiled), { label: "tools.restrict()" });
	}
	/**
	* Register a monotonic guard after the extensible `tools/pre-execute`
	* waterfall. A plain-context guard applies globally; one registered through
	* `agent.ctx` applies only to that agent. Any matching guard may deny by
	* returning a reason, while no guard can force-allow a call another guard
	* denied. The exact effect disposer is returned for ordered ownership and
	* HMR cleanup.
	* @param guard - synchronous check; a returned string denies the execution.
	* @returns the exact disposer that unregisters the guard.
	*/
	guard(guard) {
		return this.layers.effect(this.ctx, (layer) => layer.guards.append(guard), {
			label: "tools.guard()",
			notify: false
		});
	}
	/** First monotonic denial from the global then the scope chain's guard layers, farthest first. */
	guardReason(exec) {
		const globalReason = this.layers.global.guardReason(exec);
		if (globalReason !== void 0) return globalReason;
		if (exec.agent === void 0) return void 0;
		for (const layer of this.layers.chainLayers(exec.agent)) {
			const reason = layer.guardReason(exec);
			if (reason !== void 0) return reason;
		}
	}
	/**
	* Resolve every registry fact one scope needs in one layer traversal. The
	* visible map applies restrictions to the INHERITED surface, then the
	* scope's own registrations and the reserved presentation transport; the
	* other sets retain the pre-restriction facts needed by restriction and
	* prompt-order validation.
	*
	* A restriction filters what a scope inherits — the global layer and every
	* ancestor layer on its chain — and never what its OWN layer registers.
	* That exemption is what a per-child capability filter has to keep intact:
	* the delegation runtime registers a child's structured-output tool into the
	* child's own layer, and a filter naming the capabilities the child may use
	* must not strip the machinery it answers through.
	*
	* Reading the exempt set as "the global layer" instead of "not mine" held
	* only while every model-facing tool sat in the host composition. Once
	* presets moved them onto the agent plane they became an ANCESTOR
	* contribution, so a child's filter silently stopped constraining anything
	* it was given.
	* @param scope - the viewing scope (the agent), or undefined for the global view.
	* @returns the complete derived view for that scope.
	*/
	view(scope) {
		const layers = this.layers.chainLayers(scope);
		const own = this.layers.peek(scope);
		const inherited = new Map(this.layers.global.tools.entries());
		for (const layer of layers) {
			if (layer === own) continue;
			for (const [name, definition] of layer.tools.entries()) inherited.set(name, definition);
		}
		const visible = /* @__PURE__ */ new Map();
		const knownNames = /* @__PURE__ */ new Set();
		const restrictableNames = /* @__PURE__ */ new Set();
		for (const [name, definition] of inherited) {
			knownNames.add(name);
			restrictableNames.add(name);
			if (layers.every((layer) => layer.admits(name))) visible.set(name, definition);
		}
		if (own !== void 0) for (const [name, definition] of own.tools.entries()) {
			knownNames.add(name);
			visible.set(name, definition);
		}
		if (this.modeFor(scope) !== "native") visible.set(RUN_CODE_NAME, this.requirePtcTransport());
		return {
			visible,
			knownNames,
			restrictableNames
		};
	}
	/**
	* Look up a tool as one scope sees it (scoped
	* shadows global; a restricted-away global reads as absent). Presenters pass
	* the calling agent so the rendered card matches the definition that
	* actually executed.
	* @param name - the tool name as registered.
	* @param scope - the viewing scope (the agent); omitted = the global view.
	* @returns the definition the scope resolves, or undefined when none is visible.
	*/
	get(name, scope) {
		return this.view(scope).visible.get(name);
	}
	/**
	* Resolve the definition that MAY EXECUTE for a call, applying the mode
	* collapse at the operation boundary that owns it. The registry view
	* (`get`) is presentation-agnostic; here a MODEL-DIRECT call under `ptc`
	* may only name the reserved `run_code` transport, while a nested
	* sub-dispatch (a `parent` token set — the `run_code` SDK calling a tool
	* it bound) may call any visible tool. Denial surfaces as `UNKNOWN_TOOL`
	* through the executor, matching an absent definition.
	* @param name - the tool name as registered.
	* @param scope - the viewing scope (the agent); omitted = the global view.
	* @param nested - whether the call is a transport sub-dispatch, not a model-direct call.
	* @returns the definition that may run, or undefined when the call must be rejected.
	*/
	resolveExecution(name, scope, nested) {
		const tool = this.get(name, scope);
		if (tool === void 0) return void 0;
		if (this.collapses(name, scope, nested)) return void 0;
		return tool;
	}
	/**
	* Project visible definitions onto the allowlisted model-facing schema fields,
	* excluding execution and presentation callbacks.
	* @param scope - the viewing scope (the agent); omitted = the global view.
	* @returns one deep-cloned schema per visible tool.
	*/
	schemas(scope) {
		return [...this.view(scope).visible.values()].map((definition) => this.schemaOf(definition, true));
	}
	/** Project visible callable tools onto the generated PTC mode SDK contract. */
	sdkSchemas(scope) {
		return [...this.view(scope).visible.values()].filter((definition) => definition.name !== RUN_CODE_NAME).map((definition) => {
			const output = snapshotJsonValue(definition.output.schema);
			/* v8 ignore next -- registration already validated and retained this schema as lossless JSON. */
			if (output === void 0) throw new Error(`tool "${definition.name}" output schema must be lossless JSON before SDK projection`);
			return {
				...this.schemaOf(definition, true),
				output
			};
		});
	}
	/** Project one definition onto the model-facing schema fields. */
	schemaOf(definition, detachParameters) {
		const { name, description, parameters, deferLoading } = definition;
		const detached = detachParameters ? snapshotJsonValue(parameters) : parameters;
		if (detached === void 0) throw new Error(`tool "${name}" parameters must be lossless JSON before schema projection`);
		return {
			name,
			description,
			parameters: detached,
			...deferLoading === true ? { deferLoading } : {}
		};
	}
	/**
	* Classify a pending call through the caller's visible tool definition. Only
	* an exact `true` is parallel; unknown, hidden, undeclared, invalid, or
	* throwing classifiers are exclusive.
	* @param exec - call name, parsed arguments, and optional agent scope.
	* @returns the fail-closed scheduling mode.
	*/
	executionMode(exec) {
		const tool = this.resolveExecution(exec.name, exec.agent, exec.parent !== void 0);
		if (!tool?.isConcurrencySafe) return { kind: "exclusive" };
		try {
			return tool.isConcurrencySafe(exec.arguments) === true ? { kind: "parallel" } : { kind: "exclusive" };
		} catch {
			return { kind: "exclusive" };
		}
	}
	/**
	* Run the `tools/ptc-dispatch-log` waterfall over one settled sub-dispatch
	* and return the content the bridge should log on `tool/ptc-dispatch`.
	* Contained: when a listener throws, the method logs the original settled
	* content; that failure must not fail the dispatch or omit the settle event. Private:
	* the ONE consumer is the `run_code` bridge this registry constructs, which
	* receives it as a capability parameter (the `requireRuntime` idiom) — the
	* waterfall, not this invoker, is the public extension point.
	*/
	async shapeDispatchLog(dispatch) {
		try {
			return await this.ctx.waterfall(scopeTarget(this, dispatch.agent), "tools/ptc-dispatch-log", dispatch, () => Promise.resolve(dispatch.content));
		} catch (error) {
			this.ctx.logger.warn(`tools: ptc-dispatch-log listener failed for ${dispatch.name}: ${errorMessage(error)}; logging the original settled content`);
			return dispatch.content;
		}
	}
	/**
	* Whether the `ptc` mode collapse denies a model-direct call: only the
	* reserved `run_code` transport may be named. Nested sub-dispatches (a
	* `parent` token set) bypass the collapse. One home for the
	* security-relevant predicate, shared by {@link resolveExecution} and
	* {@link createExecution} so the two can never drift apart.
	*
	* Resolved through {@link modeFor}, NOT `defaultMode`: an agent given `ptc`
	* by an agent preset under a native deployment is the composition
	* `dsh-agent-tool-presentation` exists for, and reading the deployment default would
	* leave exactly that agent uncollapsed — announcing one surface while
	* executing another, which is the bypass this collapse closes.
	* @param name - the tool name as registered.
	* @param scope - the viewing scope whose effective presentation mode applies.
	* @param nested - whether the call is a transport sub-dispatch, not a model-direct call.
	*/
	collapses(name, scope, nested) {
		return !nested && this.modeFor(scope) === "ptc" && name !== "run_code";
	}
	/**
	* Execute through pre-policy, guards, around-dispatch, post-policy,
	* definition-owned content finalization, and final notification. Tool and
	* listener failures resolve as materialized error results; an invisible tool
	* reports `UNKNOWN_TOOL`. The returned outcome is the same lossless, frozen
	* snapshot final observers receive. Cancellation
	* arriving after entry and before final result materialization skips a
	* not-yet-started body with `ABORTED_BEFORE_DISPATCH` or replaces a
	* successful started outcome with `ABORTED`; already-started work is still
	* drained and may retain a tool-owned structured error.
	* @param exec - the typed same-process call input. The registry assigns its
	*   correlation token before policy begins.
	* @returns the materialized final result.
	*/
	async execute(exec) {
		return this.prepareExecution(exec, (prepared) => this.completeScheduledExecution(prepared));
	}
	async completeScheduledExecution(prepared) {
		switch (prepared.kind) {
			case "dispatch": {
				const dispatched = await this.dispatchScheduledExecution(prepared.exec);
				return dispatched.kind === "post-result" ? await this.finalizeScheduledExecution(prepared.exec, dispatched.result) : this.finishScheduledExecution(prepared.exec, dispatched.result);
			}
			case "post-result": return await this.finalizeScheduledExecution(prepared.exec, prepared.result);
			case "final-result": return this.finishScheduledExecution(prepared.exec, prepared.result);
			/* v8 ignore next -- closed-union exhaustiveness guard */
			default: return assertNever(prepared, "scheduled tool preparation");
		}
	}
	createExecution(exec) {
		const deferredContexts = [];
		const token = createExecutionToken();
		const callId = exec.callId;
		const rootCallId = exec.rootCallId ?? callId;
		const name = exec.name;
		const agent = exec.agent;
		const parent = exec.parent;
		const signal = exec.signal;
		const visible = this.get(name, agent);
		const collapsed = visible !== void 0 && this.collapses(name, agent, parent !== void 0);
		const concludingExecutions = this.concludingExecutions;
		const base = {
			token,
			callId,
			rootCallId,
			name,
			signal,
			...agent !== void 0 ? { agent } : {},
			...parent !== void 0 ? { parent } : {},
			...exec.schema !== void 0 ? { schema: exec.schema } : {},
			deferContext(context) {
				deferredContexts.push(context);
			},
			concludeTurn() {
				concludingExecutions.add(this);
			}
		};
		const capturedFinalizer = visible?.finalizeContent?.bind(visible);
		const capturedProjector = visible?.projectContent?.bind(visible);
		const finalizerFor = () => collapsed && !signal.aborted ? void 0 : capturedFinalizer;
		try {
			const detached = snapshotJsonValue(exec.arguments);
			if (detached === void 0) throw new TypeError("tool execution arguments must be losslessly JSON-serializable");
			const execution = {
				...base,
				arguments: deepFreeze(detached)
			};
			this.deferredContexts.set(execution, deferredContexts);
			this.contentFinalizers.set(execution, finalizerFor());
			if (!collapsed) this.contentProjectors.set(execution, capturedProjector);
			this.cancellationStates.set(execution, {
				callerSignal: signal,
				bodyInvoked: false
			});
			if (collapsed) {
				if (signal.aborted) return {
					kind: "final-result",
					exec: execution,
					result: toolAbortedBeforeDispatchResult()
				};
				return {
					kind: "final-result",
					exec: execution,
					result: toolErrorResult(new ToolNotFoundError(name, `only \`${RUN_CODE_NAME}\` is callable directly — call \`${name}\` from inside a \`${RUN_CODE_NAME}\` program instead`))
				};
			}
			return {
				kind: "ready",
				exec: execution
			};
		} catch (error) {
			const execution = {
				...base,
				arguments: void 0
			};
			this.contentFinalizers.set(execution, finalizerFor());
			return {
				kind: "final-result",
				exec: execution,
				result: toolErrorResult(error)
			};
		}
	}
	/**
	* Run the ordered pre-execute and monotonic guard stages for the scheduler.
	* @param input - the caller-supplied execution input.
	* @returns the prepared execution plus the next scheduler stage.
	* @internal
	*/
	async prepareScheduledExecution(input) {
		return this.prepareExecution(input, (prepared) => prepared);
	}
	async prepareExecution(input, next) {
		const created = this.createExecution(input);
		if (created.kind !== "ready") return next(created);
		const exec = created.exec;
		if (this.callerCancelled(exec)) return next({
			kind: "final-result",
			exec,
			result: toolAbortedBeforeDispatchResult()
		});
		try {
			const carrier = scopeTarget(this, exec.agent);
			const gate = await this.ctx.waterfall(carrier, "tools/pre-execute", exec, () => Promise.resolve({ kind: "allow" }));
			const askResolution = gate.kind === "ask" ? await this.serviceAsk(exec, gate) : {
				decision: gate,
				approvalCancelled: false
			};
			const { decision } = askResolution;
			if (this.callerCancelled(exec) && askResolution.approvalCancelled) return await next({
				kind: "post-result",
				exec,
				result: toolAbortedBeforeDispatchResult()
			});
			if (decision.kind === "cancel") return await next({
				kind: "post-result",
				exec,
				result: toolAbortedBeforeDispatchResult()
			});
			const denialReason = decision.kind === "allow" ? this.guardReason(exec) : decision.reason;
			const denialInfo = decision.kind === "deny" ? decision.info : void 0;
			if (denialReason !== void 0) return await next({
				kind: "post-result",
				exec,
				result: this.materializeFinalResult({
					content: [{
						type: "text",
						text: `Error: ${denialReason}`
					}],
					isError: true,
					error: {
						message: denialReason,
						...denialInfo === void 0 ? {} : { info: denialInfo }
					}
				})
			});
			if (this.callerCancelled(exec)) return await next({
				kind: "post-result",
				exec,
				result: toolAbortedBeforeDispatchResult()
			});
			return await next({
				kind: "dispatch",
				exec
			});
		} catch (error) {
			return next({
				kind: "final-result",
				exec,
				result: toolErrorResult(error)
			});
		}
	}
	/** Whether the original caller signal is currently aborted. */
	callerCancelled(exec) {
		const state = this.cancellationStates.get(exec);
		/* v8 ignore next -- only registry-minted executions reach the staged scheduler methods */
		if (state === void 0) throw new Error("tool registry scheduler invariant violated: missing cancellation state");
		return state.callerSignal.aborted;
	}
	/** Canonical cancellation outcome selected by whether the tool body started. */
	cancellationResult(exec, prior) {
		const state = this.cancellationStates.get(exec);
		/* v8 ignore next -- only registry-minted executions reach the staged scheduler methods */
		if (state === void 0) throw new Error("tool registry scheduler invariant violated: missing cancellation state");
		return state.bodyInvoked ? toolAbortedResult(prior) : toolAbortedBeforeDispatchResult(prior);
	}
	/**
	* Dispatch the registered body with the original caller signal fused back
	* into any around-wrapper replacement. Cancellation never abandons the body:
	* a started promise reaches quiescence before its outcome becomes `ABORTED`.
	*/
	async dispatchToolBody(exec) {
		const state = this.cancellationStates.get(exec);
		/* v8 ignore next -- only registry-minted executions reach the staged scheduler methods */
		if (state === void 0) throw new Error("tool registry scheduler invariant violated: missing cancellation state");
		const wrapperSignal = exec.signal;
		const fused = fuseToolSignals(state.callerSignal, wrapperSignal);
		const signal = fused.signal;
		if (isAborted(signal)) {
			fused.dispose();
			return toolAbortedBeforeDispatchResult();
		}
		exec.signal = signal;
		try {
			const tool = this.resolveExecution(exec.name, exec.agent, exec.parent !== void 0);
			if (!tool) throw new ToolNotFoundError(exec.name);
			state.bodyInvoked = true;
			const returned = await tool.execute(exec.arguments, exec);
			const result = this.createSuccessResult(exec, tool, returned);
			return isAborted(signal) ? toolAbortedResult(result) : result;
		} catch (error) {
			return toolErrorResult(error);
		} finally {
			fused.dispose();
			exec.signal = wrapperSignal;
		}
	}
	/**
	* Run around-dispatch and the tool body. Tool and unknown-tool failures still
	* receive post-execute; pipeline failures are already final.
	* @param exec - the prepared execution.
	* @returns whether the result still needs post-execute.
	* @internal
	*/
	async dispatchScheduledExecution(exec) {
		try {
			const mutableExec = exec;
			const carrier = scopeTarget(this, exec.agent);
			const result = await this.ctx.waterfall(carrier, "tools/execute", mutableExec, () => this.dispatchToolBody(mutableExec));
			const normalized = this.normalizeDispatchResult(exec, result);
			const deferredContexts = this.deferredContexts.get(exec);
			/* v8 ignore next -- dispatch only receives executions minted by this registry's prepare stage */
			if (deferredContexts === void 0) throw new Error("tool registry scheduler invariant violated: unprepared execution");
			const resultWithDeferredContexts = deferredContexts.length === 0 ? normalized : this.markCanonical(exec, {
				...normalized,
				additionalContexts: [...deferredContexts, ...normalized.additionalContexts ?? []]
			});
			return {
				kind: "post-result",
				result: this.callerCancelled(exec) && !resultWithDeferredContexts.isError ? this.cancellationResult(exec, resultWithDeferredContexts) : resultWithDeferredContexts
			};
		} catch (error) {
			return {
				kind: "final-result",
				result: toolErrorResult(error)
			};
		}
	}
	/**
	* Run ordered post-execute, then apply definition-owned content finalization,
	* materialize, and notify the final outcome.
	* @param exec - the prepared execution.
	* @param result - dispatch/pre result that still needs post-execute.
	* @returns the materialized final result.
	* @internal
	*/
	async finalizeScheduledExecution(exec, result) {
		try {
			const project = this.contentProjectors.get(exec);
			this.contentProjectors.delete(exec);
			const content = project?.(exec, result);
			const projected = content === void 0 ? result : this.markCanonical(exec, this.materializeFinalResult({
				...result,
				content
			}));
			const postResult = await this.postExecute(exec, projected);
			return this.finishScheduledExecution(exec, this.callerCancelled(exec) && !postResult.isError ? this.cancellationResult(exec, postResult) : postResult);
		} catch (error) {
			return this.finishScheduledExecution(exec, toolErrorResult(error));
		}
	}
	/**
	* Materialize the candidate, apply definition-owned content finalization,
	* then materialize and notify the authoritative result.
	* @param exec - the prepared execution.
	* @param result - final result.
	* @returns the materialized final result.
	* @internal
	*/
	finishScheduledExecution(exec, result) {
		let materializedResult;
		try {
			materializedResult = this.materializeFinalResult(result);
		} catch (error) {
			materializedResult = this.materializeFinalResult(toolErrorResult(error));
		}
		let finalResult;
		try {
			finalResult = this.materializeFinalResult(this.applyFinalContent(exec, materializedResult));
		} catch (error) {
			finalResult = this.materializeFinalResult(toolErrorResult(error));
		}
		this.notifyResult(exec, finalResult);
		return finalResult;
	}
	/** Apply the snapshotted tool-owned content transform without exposing other result fields. */
	applyFinalContent(exec, result) {
		const finalizeContent = this.contentFinalizers.get(exec);
		if (finalizeContent === void 0) return result;
		const content = finalizeContent(exec, result);
		return content === void 0 ? result : {
			...result,
			content
		};
	}
	/** Notify observers without exposing a mutation or error channel into the outcome. */
	notifyResult(exec, result) {
		Object.freeze(exec);
		const { name: toolName, callId } = exec;
		const reportFailure = (error) => {
			this.ctx.logger.warn(`tool "${toolName}" (${callId}): tools/result observer failed: ${errorMessage(error)}`);
		};
		const callbacks = this.ctx.events.dispatch("emit", [
			scopeTarget(this, exec.agent),
			"tools/result",
			exec,
			result
		]);
		for (const callback of callbacks) try {
			const returned = callback(exec, result);
			Promise.resolve(returned).catch(reportFailure);
		} catch (error) {
			reportFailure(error);
		}
	}
	/**
	* Resolve an `ask` decision to allow/deny through the approval seam. The
	* seam is consumed opportunistically with `ctx.get('approval')` — a
	* deployment that composes no ApprovalService keeps the historical degrade
	* to deny, and an unmount mid-session degrades the same way on the next ask.
	* An agent-less execution also degrades: without an agent there is no
	* session to audit to and no UI to route to. Otherwise the outcome maps
	* one-to-one — `allowed-once` proceeds; the three non-grants deny with
	* distinct reasons so the model can tell a human "no" from an absent
	* approval channel.
	*/
	async serviceAsk(exec, ask) {
		const approval = this.ctx.get("approval");
		if (approval === void 0) return {
			decision: {
				kind: "deny",
				reason: ask.reason ?? `tool "${exec.name}" requires approval (not yet supported)`
			},
			approvalCancelled: false
		};
		if (exec.agent === void 0) return {
			decision: {
				kind: "deny",
				reason: `tool "${exec.name}" requires approval, but the call has no agent to route it through`
			},
			approvalCancelled: false
		};
		const outcome = await approval.request({
			agent: exec.agent,
			toolName: exec.name,
			callId: exec.callId,
			...ask.reason !== void 0 ? { reason: ask.reason } : {},
			...ask.displayReason !== void 0 ? { displayReason: ask.displayReason } : {},
			signal: exec.signal
		});
		switch (outcome) {
			case "allowed-once": return {
				decision: { kind: "allow" },
				approvalCancelled: false
			};
			case "rejected": return {
				decision: {
					kind: "deny",
					reason: `the user rejected tool "${exec.name}"`
				},
				approvalCancelled: false
			};
			case "cancelled": return {
				decision: {
					kind: "deny",
					reason: `approval for tool "${exec.name}" was cancelled`
				},
				approvalCancelled: true
			};
			case "unavailable": return {
				decision: {
					kind: "deny",
					reason: `tool "${exec.name}" requires approval, but no approval channel is available`
				},
				approvalCancelled: false
			};
			default: return assertNever(outcome, "ApprovalOutcome");
		}
	}
	/**
	* Run the `tools/post-execute` waterfall over a dispatched `result` and apply
	* its {@link PostToolDecision}: `accept` keeps the call successful (replacing
	* `content` when given), `block` turns it into an `isError` whose content is
	* the corrective `feedback`. Either decision may attach `additionalContexts`,
	* which are ferried on the returned result for the loop's active-batch FIFO.
	* Context deferred by the tool body survives an accepted result but is
	* discarded when the outer call is blocked; a block exposes only context the
	* blocking decision explicitly supplied.
	* Runs inside `execute`'s outer try/catch (a throwing listener → isError).
	*/
	async postExecute(exec, result) {
		const decision = await this.ctx.waterfall(scopeTarget(this, exec.agent), "tools/post-execute", exec, result, () => Promise.resolve({ kind: "accept" }));
		const decisionContexts = decision.additionalContexts ?? [];
		if (decision.kind === "block") {
			const message = failureMessageFromContent(decision.feedback);
			return this.markCanonical(exec, {
				content: decision.feedback,
				isError: true,
				error: { message },
				...decisionContexts.length > 0 ? { additionalContexts: decisionContexts } : {}
			});
		}
		if (Object.hasOwn(decision, "content") && Object.hasOwn(decision, "value")) throw new TypeError("tools/post-execute accept decision cannot replace both value and content");
		const additionalContexts = [...result.additionalContexts ?? [], ...decisionContexts];
		if (Object.hasOwn(decision, "value")) {
			if (result.isError) throw new TypeError("tools/post-execute cannot replace the value of a failed result");
			const tool = this.resolveExecution(exec.name, exec.agent, exec.parent !== void 0);
			if (tool === void 0) throw new ToolNotFoundError(exec.name);
			const replaced = this.createSuccessResult(exec, tool, decision.value);
			return this.markCanonical(exec, {
				...replaced,
				...additionalContexts.length > 0 ? { additionalContexts } : {}
			});
		}
		return this.markCanonical(exec, {
			...result,
			...decision.content !== void 0 ? { content: decision.content } : {},
			...additionalContexts.length > 0 ? { additionalContexts } : {}
		});
	}
	/** Registry-normalized results and the exact dispatch that validated each value. */
	canonicalResults = /* @__PURE__ */ new WeakMap();
	/** Mark one registry-normalized result as canonical only for its owning dispatch. */
	markCanonical(exec, result) {
		this.canonicalResults.set(result, exec.token);
		return result;
	}
	/** Snapshot, validate, render, and optionally project one successful body value. */
	createSuccessResult(exec, tool, candidate) {
		const detached = snapshotToolValue(tool.name, candidate);
		const violations = validateJsonSchemaValue(tool.output.schema, detached, "value");
		if (violations.length > 0) throw new ToolOutputError(tool.name, violations);
		const value = deepFreeze(detached);
		let rendered;
		try {
			rendered = tool.output.render(exec.arguments, value);
		} catch (error) {
			throw projectionError(tool.name, "render", error);
		}
		const content = snapshotProjection(tool.name, "render", rendered);
		let meta;
		if (exec.parent === void 0 && tool.output.presentationMeta !== void 0) {
			let projected;
			try {
				projected = tool.output.presentationMeta(exec.arguments, value);
			} catch (error) {
				throw projectionError(tool.name, "presentationMeta", error);
			}
			meta = snapshotProjection(tool.name, "presentationMeta", projected);
		}
		const concludesTurn = this.concludingExecutions.has(exec);
		return this.markCanonical(exec, this.materializeFinalResult({
			isError: false,
			value,
			content,
			...meta !== void 0 ? { meta } : {},
			...concludesTurn ? { concludesTurn: true } : {}
		}));
	}
	/** Normalize an around-dispatch wrapper's authored result through the owning output contract. */
	normalizeDispatchResult(exec, result) {
		if (this.canonicalResults.get(result) === exec.token) return result;
		if (result.isError) return this.markCanonical(exec, {
			isError: true,
			error: result.error,
			content: result.content,
			...result.meta !== void 0 ? { meta: result.meta } : {},
			...result.additionalContexts !== void 0 ? { additionalContexts: result.additionalContexts } : {}
		});
		const tool = this.resolveExecution(exec.name, exec.agent, exec.parent !== void 0);
		if (tool === void 0) throw new ToolNotFoundError(exec.name);
		const normalized = this.createSuccessResult(exec, tool, result.value);
		return this.markCanonical(exec, {
			...normalized,
			...result.additionalContexts !== void 0 ? { additionalContexts: result.additionalContexts } : {}
		});
	}
	/** Materialize the authoritative commit outcome once, immediately before `tools/result`. */
	materializeFinalResult(result) {
		const presentation = {
			content: result.content,
			...result.meta !== void 0 ? { meta: result.meta } : {},
			...result.additionalContexts !== void 0 ? { additionalContexts: result.additionalContexts } : {}
		};
		if (result.isError) return materializePresentation({
			isError: true,
			error: result.error,
			...presentation
		});
		return deepFreeze({
			...materializePresentation({
				isError: false,
				...presentation,
				...result.concludesTurn === true ? { concludesTurn: true } : {}
			}),
			value: result.value
		});
	}
};
/** Mint a same-process correlation token whose identity is its value. */
function createExecutionToken() {
	return Symbol("dsh.tool.execution");
}
function toolErrorResult(error) {
	const info = errorInfo(error);
	const message = errorMessage(error);
	return {
		content: [{
			type: "text",
			text: `Error: ${message}`
		}],
		isError: true,
		error: {
			message,
			...info ? { info } : {}
		}
	};
}
/** Read live abort state across an await without treating it as synchronously immutable. */
function isAborted(signal) {
	return signal.aborted;
}
/**
* Fuse caller and wrapper cancellation without nesting `AbortSignal.any`.
* Keeping the relay dispatch-scoped also removes listeners when work settles.
*/
function fuseToolSignals(caller, wrapper) {
	if (caller === wrapper) return {
		signal: caller,
		dispose() {}
	};
	const controller = new AbortController();
	let listening = false;
	const dispose = () => {
		if (!listening) return;
		listening = false;
		caller.removeEventListener("abort", abortFromCaller);
		wrapper.removeEventListener("abort", abortFromWrapper);
	};
	const abortFrom = (source) => {
		const reason = source.reason;
		controller.abort(reason);
		dispose();
	};
	const abortFromCaller = () => {
		abortFrom(caller);
	};
	const abortFromWrapper = () => {
		abortFrom(wrapper);
	};
	if (wrapper.aborted) abortFromWrapper();
	else if (caller.aborted) abortFromCaller();
	else {
		listening = true;
		caller.addEventListener("abort", abortFromCaller, { once: true });
		wrapper.addEventListener("abort", abortFromWrapper, { once: true });
	}
	return {
		signal: controller.signal,
		dispose
	};
}
/** Canonical result when cancellation supersedes success after body invocation. */
function toolAbortedResult(prior) {
	const additionalContexts = prior?.additionalContexts ?? [];
	return {
		content: [{
			type: "text",
			text: "Error: tool call aborted"
		}],
		isError: true,
		error: {
			message: "tool call aborted",
			info: {
				name: "AbortError",
				code: TOOL_ABORTED
			}
		},
		...additionalContexts.length > 0 ? { additionalContexts } : {}
	};
}
/** Canonical result when cancellation prevents tool body invocation. */
function toolAbortedBeforeDispatchResult(prior) {
	const additionalContexts = prior?.additionalContexts ?? [];
	return {
		content: [{
			type: "text",
			text: "Error: tool call aborted before dispatch"
		}],
		isError: true,
		error: {
			message: "tool call aborted before dispatch",
			info: {
				name: "AbortError",
				code: TOOL_ABORTED_BEFORE_DISPATCH
			}
		},
		...additionalContexts.length > 0 ? { additionalContexts } : {}
	};
}
//#endregion
//#region lib/types/format.js
/**
* What the helper reports, turned into the short plain text the model reads.
* Kept terse: every line here is paid for on each later request of the turn.
*/
const APP_NAMES = {
	word: "Word",
	excel: "Excel",
	ppt: "PowerPoint"
};
const WPS_NAMES = {
	word: "WPS Writer",
	excel: "WPS Spreadsheets",
	ppt: "WPS Presentation"
};
/** The app as the model should name it: WPS has its own names, the tools and operations are the same. */
const appName = (on) => (on.suite === "wps" ? WPS_NAMES : APP_NAMES)[on.app];
const state = (doc) => [
	doc.saved ? "saved" : "unsaved changes",
	doc.readOnly ? "read-only" : "",
	doc.active ? "active" : "",
	doc.background ? "in the background, no window" : ""
].filter(Boolean).join(", ");
function formatStatus(apps) {
	const lines = [];
	for (const app of apps) {
		const name = appName(app);
		if (!app.installed) {
			lines.push(`${name}: not installed`);
			continue;
		}
		if (!app.running && app.documents.length === 0) {
			lines.push(`${name}: installed, not running (office_open starts it)`);
			continue;
		}
		if (app.error) {
			lines.push(`${name}: ${app.error}`);
			continue;
		}
		if (app.documents.length === 0) {
			lines.push(`${name}: running, nothing open`);
			continue;
		}
		lines.push(`${name}:`);
		for (const doc of app.documents) lines.push(`  ${doc.path ?? `${doc.name} (never saved)`} — ${state(doc)}`);
	}
	return lines.join("\n");
}
function formatOpened(doc) {
	return `${doc.how === "created" ? "Created" : doc.how === "opened" ? "Opened" : "Already open, attached to"} ${doc.path ?? doc.name} in ${appName(doc)} (${state(doc)}).${doc.suite === "wps" ? ` This is WPS Office: the same tools and operations apply${doc.app === "ppt" ? "," : ""}${doc.app === "ppt" ? " except that a formula ($...$) on a slide is a picture laid over blanks kept for it in the text, not an equation the user can edit: after you move, resize or rewrite a text shape that has formulas, write its text again (set_text) so they are placed anew, and look at the slide (office_render)" : ""}.` : ""}${doc.readOnly ? " It is read-only: edits cannot be saved to this file." : ""}`;
}
const BODY_STYLES = /* @__PURE__ */ new Set([
	"normal",
	"正文",
	"body text"
]);
function formatWord(read) {
	const notes = read.footnotes?.length ? ["Footnotes (i is the collection index, not the displayed number; edits need expect from the note text):", ...read.footnotes.map((note) => `footnote ${note.i} [display ${note.label ?? `unavailable, number style ${note.numberStyle}`}; para ${note.para}, offset ${note.offset}, page ${note.page}, section ${note.section}] ${note.text}`)] : [];
	if (read.cells) return [
		`Table ${read.table} (${read.cells.length} rows):`,
		...read.cells.map((row, index) => `${index + 1}: ${row.map((cell) => cell ?? "(merged)").join(" | ")}`),
		...notes
	].join("\n");
	return [
		`${read.paragraphs} paragraphs, ${read.tables} table(s)${read.pages === void 0 ? "" : `, ${read.pages} page(s)`}.`,
		...(read.items ?? []).map((item) => {
			if (item.size !== void 0) {
				const name = item.table === void 0 ? "table inside a cell" : `table ${item.table}`;
				return `${item.i}–${item.to} [${name}, ${item.size}]${item.open ? " holds the text below; write inside it by paragraph number:" : ""}`;
			}
			const tag = item.level !== void 0 ? `H${item.level}` : item.style !== void 0 && !BODY_STYLES.has(item.style.toLowerCase()) ? item.style : "";
			return `${item.inTable ? "  " : ""}${item.i}${tag ? ` [${tag}]` : ""} ${item.text ? item.text : "(empty)"}`;
		}),
		...read.more ? [read.more] : [],
		...notes
	].join("\n");
}
const cell = (value) => value === null || value === void 0 ? "" : String(value);
function formatExcel(read) {
	const sheets = read.sheets.map((sheet) => `${sheet.name} (${sheet.used}${sheet.charts ? `, ${sheet.charts} chart(s)` : ""})`).join(", ");
	const first = Number(/\d+/.exec(read.range)?.[0] ?? 1);
	const rows = read.values.map((row, index) => `${first + index}: ${row.map(cell).join(" | ")}`);
	const formulas = read.formulas?.length ? [`Formulas: ${read.formulas.map((item) => `${item.cell} ${item.formula}`).join("; ")}`] : [];
	const objects = read.objects?.length ? [`On this sheet: ${read.objects.join("; ")}.`] : [];
	const profile = read.profile?.length ? ["What the columns hold (the whole block, not only the rows shown):", ...read.profile.map((line) => `  ${line}`)] : [];
	return [
		`Sheets: ${sheets}.`,
		`${read.sheet}!${read.range} (values; one line per row):`,
		...rows,
		...formulas,
		...read.clipped ? [read.clipped] : [],
		...objects,
		...profile
	].join("\n");
}
function formatPpt(read) {
	const lines = [`${read.slides} slide(s), ${read.width}×${read.height} pt.`];
	for (const slide of read.items) {
		lines.push(`Slide ${slide.slide}${slide.layout ? ` (${slide.layout})` : ""}`);
		if (slide.shapes.length === 0) lines.push("  (empty)");
		for (const shape of slide.shapes) lines.push(`  "${shape.name}"${shape.id === void 0 ? "" : ` #${shape.id}`} ${shape.type} [${shape.box.join(", ")}]${shape.text === void 0 ? "" : `: ${shape.text.replace(/\n/g, " / ")}`}`);
		if (slide.plain) lines.push(`  (+ ${slide.plain} shape(s) without text: decoration; full: true lists them)`);
	}
	return lines.join("\n");
}
/** Failures found before anything was touched. */
const PRECHECKED = /* @__PURE__ */ new Set([
	"BAD_ARGS",
	"ANCHOR_MISSING",
	"ANCHOR_MISMATCH",
	"NOT_FOUND",
	"STYLE_MISSING",
	"BUSY"
]);
function formatEdit(result) {
	const done = result.done.map((line, index) => `${index + 1}. ${line}`);
	if (!result.failed) return [
		`All ${result.total} operation(s) applied (not saved yet):`,
		...done,
		"Before you finish, look at the pages you changed with office_render and fix what is off."
	].join("\n");
	const { index, op, error } = result.failed;
	return [
		`Stopped at operation ${index + 1} (${op}): ${error}`,
		done.length ? `Applied before it:\n${done.join("\n")}` : "Nothing before it was applied.",
		PRECHECKED.has(result.failed.code) ? "" : "The failed operation itself may have changed part of the document: read it before you retry.",
		result.total - index - 1 > 0 ? `The ${result.total - index - 1} operation(s) after it did not run.` : ""
	].filter(Boolean).join("\n");
}
//#endregion
//#region node_modules/.pnpm/temml@0.13.5/node_modules/temml/dist/temml.mjs
/**
* This is the ParseError class, which is the main error thrown by Temml
* functions when something has gone wrong. This is used to distinguish internal
* errors from errors in the expression that the user provided.
*
* If possible, a caller should provide a Token or ParseNode with information
* about where in the source string the problem occurred.
*/
var ParseError = class ParseError {
	constructor(message, token) {
		let error = " " + message;
		let start;
		const loc = token && token.loc;
		if (loc && loc.start <= loc.end) {
			const input = loc.lexer.input;
			start = loc.start;
			const end = loc.end;
			if (start === input.length) error += " at end of input: ";
			else error += " at position " + (start + 1) + ": \n";
			const underlined = input.slice(start, end).replace(/[^]/g, "$&̲");
			let left;
			if (start > 15) left = "…" + input.slice(start - 15, start);
			else left = input.slice(0, start);
			let right;
			if (end + 15 < input.length) right = input.slice(end, end + 15) + "…";
			else right = input.slice(end);
			error += left + underlined + right;
		}
		const self = new Error(error);
		self.name = "ParseError";
		self.__proto__ = ParseError.prototype;
		self.position = start;
		return self;
	}
};
ParseError.prototype.__proto__ = Error.prototype;
/**
* This file contains a list of utility functions which are useful in other
* files.
*/
/**
* Look up `key` on a user-supplied `options` object and return its value, or
* `defaultIfUndefined` if the setting is absent. Reads via
* `Object.prototype.hasOwnProperty.call` rather than plain property access,
* so a key that only exists because `Object.prototype` was polluted
* elsewhere is never mistaken for a setting the caller actually provided.
*/
const deflt = function(options, key, defaultIfUndefined) {
	if (!Object.prototype.hasOwnProperty.call(options, key)) return defaultIfUndefined;
	const setting = options[key];
	return setting === void 0 ? defaultIfUndefined : setting;
};
const uppercase = /([A-Z])/g;
const hyphenate = function(str) {
	return str.replace(uppercase, "-$1").toLowerCase();
};
const ESCAPE_LOOKUP = {
	"&": "&amp;",
	">": "&gt;",
	"<": "&lt;",
	"\"": "&quot;",
	"'": "&#x27;"
};
const ESCAPE_REGEX = /[&><"']/g;
/**
* Escapes text to prevent scripting attacks.
*/
function escape(text) {
	return String(text).replace(ESCAPE_REGEX, (match) => ESCAPE_LOOKUP[match]);
}
/**
* Sometimes we want to pull out the innermost element of a group. In most
* cases, this will just be the group itself, but when ordgroups and colors have
* a single element, we want to pull that out.
*/
const getBaseElem = function(group) {
	if (group.type === "ordgroup") {
		if (group.body.length === 1) return getBaseElem(group.body[0]);
		else return group;
	} else if (group.type === "color") {
		if (group.body.length === 1) return getBaseElem(group.body[0]);
		else return group;
	} else if (group.type === "font") return getBaseElem(group.body);
	else return group;
};
/**
* TeXbook algorithms often reference "character boxes", which are simply groups
* with a single character in them. To decide if something is a character box,
* we find its innermost group, and see if it is a single character.
*/
const isCharacterBox = function(group) {
	const baseElem = getBaseElem(group);
	return baseElem.type === "mathord" || baseElem.type === "textord" || baseElem.type === "atom";
};
const assert = function(value) {
	if (!value) throw new Error("Expected non-null, but got " + String(value));
	return value;
};
/**
* Return the protocol of a URL, or "_relative" if the URL does not specify a
* protocol (and thus is relative), or `null` if URL has invalid protocol
* (so should be outright rejected).
*/
const protocolFromUrl = function(url) {
	const protocol = /^[\x00-\x20]*([^\\/#?]*?)(:|&#0*58|&#x0*3a|&colon)/i.exec(url);
	if (!protocol) return "_relative";
	if (protocol[2] !== ":") return null;
	if (!/^[a-zA-Z][a-zA-Z0-9+\-.]*$/.test(protocol[1])) return null;
	return protocol[1].toLowerCase();
};
/**
* Round `n` to 4 decimal places, or to the nearest 1/10,000th em. The TeXbook
* gives an acceptable rounding error of 100sp (which would be the nearest
* 1/6551.6em with our ptPerEm = 10):
* http://www.ctex.org/documents/shredder/src/texbook.pdf#page=69
*/
const round = function(n) {
	return +n.toFixed(4);
};
const smalls = "acegıȷmnopqrsuvwxyzαγεηικμνοπρςστυχωϕ𝐚𝐜𝐞𝐠𝐦𝐧𝐨𝐩𝐪𝐫𝐬𝐮𝐯𝐰𝐱𝐲𝐳";
/**
* This is a module for storing settings passed into Temml. It correctly handles
* default settings.
*/
const FORBIDDEN_MACRO_KEYS = [
	"__proto__",
	"prototype",
	"constructor"
];
const sanitizeMacros = (macros) => {
	if (macros == null || typeof macros !== "object") return Object.create(null);
	for (const key of FORBIDDEN_MACRO_KEYS) if (Object.prototype.hasOwnProperty.call(macros, key)) delete macros[key];
	return Object.setPrototypeOf(macros, null);
};
/**
* The main Settings object
*/
var Settings = class {
	constructor(options) {
		options = options || {};
		this.displayMode = deflt(options, "displayMode", false);
		this.annotate = deflt(options, "annotate", false);
		this.leqno = deflt(options, "leqno", false);
		this.throwOnError = deflt(options, "throwOnError", false);
		this.errorColor = deflt(options, "errorColor", "#b22222");
		this.macros = sanitizeMacros(Object.prototype.hasOwnProperty.call(options, "macros") ? options.macros : void 0);
		this.wrap = deflt(options, "wrap", "none");
		this.xml = deflt(options, "xml", false);
		this.colorIsTextColor = deflt(options, "colorIsTextColor", false);
		this.strict = deflt(options, "strict", false);
		this.trust = deflt(options, "trust", false);
		const maxSize = Object.prototype.hasOwnProperty.call(options, "maxSize") ? options.maxSize : void 0;
		this.maxSize = maxSize === void 0 ? [Infinity, Infinity] : Array.isArray(maxSize) ? maxSize : [Infinity, Infinity];
		this.maxExpand = Math.max(0, deflt(options, "maxExpand", 1e3));
		this.wrapDelimiterPairs = true;
	}
	/**
	* Check whether to test potentially dangerous input, and return
	* `true` (trusted) or `false` (untrusted).  The sole argument `context`
	* should be an object with `command` field specifying the relevant LaTeX
	* command (as a string starting with `\`), and any other arguments, etc.
	* If `context` has a `url` field, a `protocol` field will automatically
	* get added by this function (changing the specified object).
	*/
	isTrusted(context) {
		if (context.url && !context.protocol) {
			const protocol = protocolFromUrl(context.url);
			if (protocol == null) return false;
			context.protocol = protocol;
		}
		const trust = typeof this.trust === "function" ? this.trust(context) : this.trust;
		return Boolean(trust);
	}
};
/**
* All registered functions.
* `functions.js` just exports this same dictionary again and makes it public.
* `Parser.js` requires this dictionary.
*/
const _functions = {};
/**
* All MathML builders. Should be only used in the `define*` and the `build*ML`
* functions.
*/
const _mathmlGroupBuilders = {};
function defineFunction({ type, names, props, handler, mathmlBuilder }) {
	const data = {
		type,
		numArgs: props.numArgs,
		argTypes: props.argTypes,
		allowedInArgument: !!props.allowedInArgument,
		allowedInText: !!props.allowedInText,
		allowedInMath: props.allowedInMath === void 0 ? true : props.allowedInMath,
		numOptionalArgs: props.numOptionalArgs || 0,
		infix: !!props.infix,
		primitive: !!props.primitive,
		handler
	};
	for (let i = 0; i < names.length; ++i) _functions[names[i]] = data;
	if (type) {
		if (mathmlBuilder) _mathmlGroupBuilders[type] = mathmlBuilder;
	}
}
/**
* Use this to register only the MathML builder for a function(e.g.
* if the function's ParseNode is generated in Parser.js rather than via a
* stand-alone handler provided to `defineFunction`).
*/
function defineFunctionBuilders({ type, mathmlBuilder }) {
	defineFunction({
		type,
		names: [],
		props: { numArgs: 0 },
		handler() {
			throw new Error("Should never be called.");
		},
		mathmlBuilder
	});
}
const normalizeArgument = function(arg) {
	return arg.type === "ordgroup" && arg.body.length === 1 ? arg.body[0] : arg;
};
const ordargument = function(arg) {
	return arg.type === "ordgroup" ? arg.body : [arg];
};
/**
* This node represents a document fragment, which contains elements, but when
* placed into the DOM doesn't have any representation itself. It only contains
* children and doesn't have any DOM node properties.
*/
var DocumentFragment = class {
	constructor(children) {
		this.children = children;
		this.classes = [];
		this.style = {};
	}
	hasClass(className) {
		return this.classes.includes(className);
	}
	/** Convert the fragment into a node. */
	toNode() {
		const frag = document.createDocumentFragment();
		for (let i = 0; i < this.children.length; i++) frag.appendChild(this.children[i].toNode());
		return frag;
	}
	/** Convert the fragment into HTML markup. */
	toMarkup() {
		let markup = "";
		for (let i = 0; i < this.children.length; i++) markup += this.children[i].toMarkup();
		return markup;
	}
	/**
	* Converts the math node into a string, similar to innerText. Applies to
	* MathDomNode's only.
	*/
	toText() {
		const toText = (child) => child.toText();
		return this.children.map(toText).join("");
	}
};
/**
* These objects store the data about the DOM nodes we create, as well as some
* extra data. They can then be transformed into real DOM nodes with the
* `toNode` function or HTML markup using `toMarkup`. They are useful for both
* storing extra properties on the nodes, as well as providing a way to easily
* work with the DOM.
*
* Similar functions for working with MathML nodes exist in mathMLTree.js.
*
*/
/**
* Create an HTML className based on a list of classes. In addition to joining
* with spaces, we also remove empty classes.
*/
const createClass = function(classes) {
	return classes.filter((cls) => cls).join(" ");
};
const initNode = function(classes, style) {
	this.classes = classes || [];
	this.attributes = {};
	this.style = style || {};
};
/**
* Convert into an HTML node
*/
const toNode = function(tagName) {
	const node = document.createElement(tagName);
	node.className = createClass(this.classes);
	for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) node.style[style] = this.style[style];
	for (const attr in this.attributes) if (Object.prototype.hasOwnProperty.call(this.attributes, attr)) node.setAttribute(attr, this.attributes[attr]);
	for (let i = 0; i < this.children.length; i++) node.appendChild(this.children[i].toNode());
	return node;
};
/**
* Convert into an HTML markup string
*/
const toMarkup = function(tagName) {
	let markup = `<${tagName}`;
	if (this.classes.length) markup += ` class="${escape(createClass(this.classes))}"`;
	let styles = "";
	for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) styles += `${hyphenate(style)}:${this.style[style]};`;
	if (styles) markup += ` style="${styles}"`;
	for (const attr in this.attributes) if (Object.prototype.hasOwnProperty.call(this.attributes, attr)) markup += ` ${attr}="${escape(this.attributes[attr])}"`;
	markup += ">";
	for (let i = 0; i < this.children.length; i++) markup += this.children[i].toMarkup();
	markup += `</${tagName}>`;
	return markup;
};
/**
* This node represents a span node, with a className, a list of children, and
* an inline style.
*
*/
var Span = class {
	constructor(classes, children, style) {
		initNode.call(this, classes, style);
		this.children = children || [];
	}
	setAttribute(attribute, value) {
		this.attributes[attribute] = value;
	}
	toNode() {
		return toNode.call(this, "span");
	}
	toMarkup() {
		return toMarkup.call(this, "span");
	}
};
let TextNode$1 = class TextNode {
	constructor(text) {
		this.text = text;
	}
	toNode() {
		return document.createTextNode(this.text);
	}
	toMarkup() {
		return escape(this.text);
	}
};
var AnchorNode = class {
	constructor(href, classes, children) {
		this.href = href;
		this.classes = classes;
		this.children = children || [];
	}
	toNode() {
		const node = document.createElement("a");
		node.setAttribute("href", this.href);
		if (this.classes.length > 0) node.className = createClass(this.classes);
		for (let i = 0; i < this.children.length; i++) node.appendChild(this.children[i].toNode());
		return node;
	}
	toMarkup() {
		let markup = `<a href='${escape(this.href)}'`;
		if (this.classes.length > 0) markup += ` class="${escape(createClass(this.classes))}"`;
		markup += ">";
		for (let i = 0; i < this.children.length; i++) markup += this.children[i].toMarkup();
		markup += "</a>";
		return markup;
	}
};
var Img = class {
	constructor(src, alt, style) {
		this.alt = alt;
		this.src = src;
		this.classes = ["mord"];
		this.style = style;
	}
	hasClass(className) {
		return this.classes.includes(className);
	}
	toNode() {
		const node = document.createElement("img");
		node.src = this.src;
		node.alt = this.alt;
		node.className = "mord";
		for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) node.style[style] = this.style[style];
		return node;
	}
	toMarkup() {
		let markup = `<img src='${this.src}' alt='${this.alt}'`;
		let styles = "";
		for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) styles += `${hyphenate(style)}:${this.style[style]};`;
		if (styles) markup += ` style="${escape(styles)}"`;
		markup += ">";
		return markup;
	}
};
/**
* These objects store data about MathML nodes.
* The `toNode` and `toMarkup` functions  create namespaced DOM nodes and
* HTML text markup respectively.
*/
function newDocumentFragment(children) {
	return new DocumentFragment(children);
}
/**
* This node represents a general purpose MathML node of any type,
* for example, `"mo"` or `"mspace"`, corresponding to `<mo>` and
* `<mspace>` tags).
*/
var MathNode = class {
	constructor(type, children, classes, style) {
		this.type = type;
		this.attributes = {};
		this.children = children || [];
		this.classes = classes || [];
		this.style = style || {};
		this.label = "";
	}
	/**
	* Sets an attribute on a MathML node. MathML depends on attributes to convey a
	* semantic content, so this is used heavily.
	*/
	setAttribute(name, value) {
		this.attributes[name] = value;
	}
	/**
	* Gets an attribute on a MathML node.
	*/
	getAttribute(name) {
		return this.attributes[name];
	}
	setLabel(value) {
		this.label = value;
	}
	/**
	* Converts the math node into a MathML-namespaced DOM element.
	*/
	toNode() {
		const node = document.createElementNS("http://www.w3.org/1998/Math/MathML", this.type);
		for (const attr in this.attributes) if (Object.prototype.hasOwnProperty.call(this.attributes, attr)) node.setAttribute(attr, this.attributes[attr]);
		if (this.classes.length > 0) node.className = createClass(this.classes);
		for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) node.style[style] = this.style[style];
		for (let i = 0; i < this.children.length; i++) node.appendChild(this.children[i].toNode());
		return node;
	}
	/**
	* Converts the math node into an HTML markup string.
	*/
	toMarkup() {
		let markup = "<" + this.type;
		for (const attr in this.attributes) if (Object.prototype.hasOwnProperty.call(this.attributes, attr)) {
			markup += " " + attr + "=\"";
			markup += escape(this.attributes[attr]);
			markup += "\"";
		}
		if (this.classes.length > 0) markup += ` class="${escape(createClass(this.classes))}"`;
		let styles = "";
		for (const style in this.style) if (Object.prototype.hasOwnProperty.call(this.style, style)) styles += `${hyphenate(style)}:${this.style[style]};`;
		if (styles) markup += ` style="${styles}"`;
		markup += ">";
		for (let i = 0; i < this.children.length; i++) markup += this.children[i].toMarkup();
		markup += "</" + this.type + ">";
		return markup;
	}
	/**
	* Converts the math node into a string, similar to innerText, but escaped.
	*/
	toText() {
		return this.children.map((child) => child.toText()).join("");
	}
};
/**
* This node represents a piece of text.
*/
var TextNode = class {
	constructor(text) {
		this.text = text;
	}
	/**
	* Converts the text node into a DOM text node.
	*/
	toNode() {
		return document.createTextNode(this.text);
	}
	/**
	* Converts the text node into escaped HTML markup
	* (representing the text itself).
	*/
	toMarkup() {
		return escape(this.toText());
	}
	/**
	* Converts the text node into a string
	* (representing the text itself).
	*/
	toText() {
		return this.text;
	}
};
const wrapWithMstyle = (expression) => {
	let node;
	if (expression.length === 1 && expression[0].type === "mrow") {
		node = expression.pop();
		node.type = "mstyle";
	} else node = new MathNode("mstyle", expression);
	return node;
};
/**
* This file provides support for building horizontal stretchy elements.
*/
const estimatedWidth = (node) => {
	let width = 0;
	if (node.body && Array.isArray(node.body)) for (const item of node.body) width += estimatedWidth(item);
	else if (node.body) width += estimatedWidth(node.body);
	else if (node.type === "supsub") {
		width += estimatedWidth(node.base);
		if (node.sub) width += .7 * estimatedWidth(node.sub);
		if (node.sup) width += .7 * estimatedWidth(node.sup);
	} else if (node.type === "mathord" || node.type === "textord") for (const ch of node.text.split("")) {
		const codePoint = ch.codePointAt(0);
		if (96 < codePoint && codePoint < 123 || 944 < codePoint && codePoint < 970) width += .56;
		else if (47 < codePoint && codePoint < 58) width += .5;
		else width += .92;
	}
	else width += 1;
	return width;
};
const stretchyCodePoint = {
	widehat: "ˆ",
	widecheck: "ˇ",
	widetilde: "~",
	wideparen: "⏜",
	utilde: "~",
	overleftarrow: "←",
	underleftarrow: "←",
	xleftarrow: "←",
	overrightarrow: "→",
	underrightarrow: "→",
	xrightarrow: "→",
	underbrace: "⏟",
	overbrace: "⏞",
	overbracket: "⎴",
	underbracket: "⎵",
	overgroup: "⏠",
	overparen: "⏜",
	undergroup: "⏡",
	underparen: "⏝",
	overleftrightarrow: "↔",
	underleftrightarrow: "↔",
	xleftrightarrow: "↔",
	Overrightarrow: "⇒",
	xRightarrow: "⇒",
	overleftharpoon: "↼",
	xleftharpoonup: "↼",
	overrightharpoon: "⇀",
	xrightharpoonup: "⇀",
	xLeftarrow: "⇐",
	xLeftrightarrow: "⇔",
	xhookleftarrow: "↩",
	xhookrightarrow: "↪",
	xmapsto: "↦",
	xrightharpoondown: "⇁",
	xleftharpoondown: "↽",
	xtwoheadleftarrow: "↞",
	xtwoheadrightarrow: "↠",
	xlongequal: "=",
	xrightleftarrows: "⇄",
	xtofrom: "⇄",
	xleftrightharpoons: "⇋",
	xrightleftharpoons: "⇌",
	yields: "→",
	yieldsLeft: "←",
	mesomerism: "↔",
	longrightharpoonup: "⇀",
	longleftharpoondown: "↽",
	eqrightharpoonup: "⇀",
	eqleftharpoondown: "↽",
	"\\cdrightarrow": "→",
	"\\cdleftarrow": "←",
	"\\cdlongequal": "=",
	yieldsLeftRight: "⇄",
	chemequilibrium: "⇌"
};
const mathMLnode = function(label) {
	const node = new MathNode("mo", [new TextNode(stretchyCodePoint[label.slice(1)])]);
	node.setAttribute("stretchy", "true");
	return node;
};
const wideHats = ["\\widehat", "\\widecheck"];
const accentNode = (group) => {
	const mo = mathMLnode(group.label);
	if (group.label.indexOf("tilde") > -1) {
		const width = estimatedWidth(group.base);
		if (1 < width && width < 1.6) mo.classes.push("tml-tilde-2");
		else if (1.6 <= width && width < 2.5) mo.classes.push("tml-tilde-3");
		else if (2.5 <= width) mo.classes.push("tml-tilde-4");
	} else if (wideHats.includes(group.label)) {
		const width = estimatedWidth(group.base);
		if (0 < width && width <= 1) mo.classes.push("tml-hat-1");
		else if (1 < width && width < 1.6) mo.classes.push("tml-hat-2");
		else if (1.6 <= width && width < 2.5) mo.classes.push("tml-hat-3");
		else if (2.5 <= width) mo.classes.push("tml-hat-4");
	}
	return mo;
};
/**
* This file holds a list of all no-argument functions and single-character
* symbols (like 'a' or ';').
*
* For each of the symbols, there are two properties they can have:
* - group (required): the ParseNode group type the symbol should have (i.e.
"textord", "mathord", etc).
* - replace: the character that this symbol or function should be
*   replaced with (i.e. "\phi" has a replace value of "\u03d5", the phi
*   character in the main font).
*
* The outermost map in the table indicates what mode the symbols should be
* accepted in (e.g. "math" or "text").
*/
const ATOMS = {
	bin: 1,
	close: 1,
	inner: 1,
	open: 1,
	punct: 1,
	rel: 1
};
const NON_ATOMS = {
	"accent-token": 1,
	mathord: 1,
	"op-token": 1,
	spacing: 1,
	textord: 1
};
const symbols = Object.create(null);
symbols.math = Object.create(null);
symbols.text = Object.create(null);
function defineSymbol(mode, group, replace, name, acceptUnicodeChar) {
	if (mode !== "math" && mode !== "text") throw new TypeError(`Invalid mode: ${mode} called in defineSymbol`);
	const entry = {
		group,
		replace
	};
	symbols[mode][name] = entry;
	/** `acceptUnicodeChar = true` is only applicable if `replace` is set. */
	if (acceptUnicodeChar && replace) symbols[mode][replace] = entry;
}
const math = "math";
const text = "text";
const accent = "accent-token";
const bin = "bin";
const close = "close";
const inner = "inner";
const mathord = "mathord";
const op = "op-token";
const open = "open";
const punct = "punct";
const rel = "rel";
const spacing = "spacing";
const textord = "textord";
defineSymbol(math, rel, "≡", "\\equiv", true);
defineSymbol(math, rel, "≺", "\\prec", true);
defineSymbol(math, rel, "≻", "\\succ", true);
defineSymbol(math, rel, "∼", "\\sim", true);
defineSymbol(math, rel, "⟂", "\\perp", true);
defineSymbol(math, rel, "⪯", "\\preceq", true);
defineSymbol(math, rel, "⪰", "\\succeq", true);
defineSymbol(math, rel, "≃", "\\simeq", true);
defineSymbol(math, rel, "≌", "\\backcong", true);
defineSymbol(math, rel, "|", "\\mid", true);
defineSymbol(math, rel, "≪", "\\ll", true);
defineSymbol(math, rel, "≫", "\\gg", true);
defineSymbol(math, rel, "≍", "\\asymp", true);
defineSymbol(math, rel, "∥", "\\parallel");
defineSymbol(math, rel, "⌣", "\\smile", true);
defineSymbol(math, rel, "⊑", "\\sqsubseteq", true);
defineSymbol(math, rel, "⊒", "\\sqsupseteq", true);
defineSymbol(math, rel, "≐", "\\doteq", true);
defineSymbol(math, rel, "⌢", "\\frown", true);
defineSymbol(math, rel, "∋", "\\ni", true);
defineSymbol(math, rel, "∌", "\\notni", true);
defineSymbol(math, rel, "∝", "\\propto", true);
defineSymbol(math, rel, "⊢", "\\vdash", true);
defineSymbol(math, rel, "⊣", "\\dashv", true);
defineSymbol(math, rel, "∋", "\\owns");
defineSymbol(math, rel, "≘", "\\arceq", true);
defineSymbol(math, rel, "≙", "\\wedgeq", true);
defineSymbol(math, rel, "≚", "\\veeeq", true);
defineSymbol(math, rel, "≛", "\\stareq", true);
defineSymbol(math, rel, "≝", "\\eqdef", true);
defineSymbol(math, rel, "≞", "\\measeq", true);
defineSymbol(math, rel, "≟", "\\questeq", true);
defineSymbol(math, rel, "≠", "\\ne", true);
defineSymbol(math, rel, "≠", "\\neq");
defineSymbol(math, rel, "⩵", "\\eqeq", true);
defineSymbol(math, rel, "⩶", "\\eqeqeq", true);
defineSymbol(math, rel, "∷", "\\dblcolon", true);
defineSymbol(math, rel, "≔", "\\coloneqq", true);
defineSymbol(math, rel, "≕", "\\eqqcolon", true);
defineSymbol(math, rel, "∹", "\\eqcolon", true);
defineSymbol(math, rel, "⩴", "\\Coloneqq", true);
defineSymbol(math, punct, ".", "\\ldotp");
defineSymbol(math, punct, "·", "\\cdotp");
defineSymbol(math, textord, "#", "\\#");
defineSymbol(text, textord, "#", "\\#");
defineSymbol(math, textord, "&", "\\&");
defineSymbol(text, textord, "&", "\\&");
defineSymbol(math, textord, "ℵ", "\\aleph", true);
defineSymbol(math, textord, "∀", "\\forall", true);
defineSymbol(math, textord, "ℏ", "\\hbar", true);
defineSymbol(math, textord, "∃", "\\exists", true);
defineSymbol(math, open, "∇", "\\nabla", true);
defineSymbol(math, textord, "♭", "\\flat", true);
defineSymbol(math, textord, "ℓ", "\\ell", true);
defineSymbol(math, textord, "♮", "\\natural", true);
defineSymbol(math, textord, "Å", "\\Angstrom", true);
defineSymbol(text, textord, "Å", "\\Angstrom", true);
defineSymbol(math, textord, "♣", "\\clubsuit", true);
defineSymbol(math, textord, "♧", "\\varclubsuit", true);
defineSymbol(math, textord, "℘", "\\wp", true);
defineSymbol(math, textord, "♯", "\\sharp", true);
defineSymbol(math, textord, "♢", "\\diamondsuit", true);
defineSymbol(math, textord, "♦", "\\vardiamondsuit", true);
defineSymbol(math, textord, "ℜ", "\\Re", true);
defineSymbol(math, textord, "♡", "\\heartsuit", true);
defineSymbol(math, textord, "♥", "\\varheartsuit", true);
defineSymbol(math, textord, "ℑ", "\\Im", true);
defineSymbol(math, textord, "♠", "\\spadesuit", true);
defineSymbol(math, textord, "♤", "\\varspadesuit", true);
defineSymbol(math, textord, "♀", "\\female", true);
defineSymbol(math, textord, "♂", "\\male", true);
defineSymbol(math, textord, "§", "\\S", true);
defineSymbol(text, textord, "§", "\\S");
defineSymbol(math, textord, "¶", "\\P", true);
defineSymbol(text, textord, "¶", "\\P");
defineSymbol(text, textord, "☺", "\\smiley", true);
defineSymbol(math, textord, "☺", "\\smiley", true);
defineSymbol(math, textord, "†", "\\dag");
defineSymbol(text, textord, "†", "\\dag");
defineSymbol(text, textord, "†", "\\textdagger");
defineSymbol(math, textord, "‡", "\\ddag");
defineSymbol(text, textord, "‡", "\\ddag");
defineSymbol(text, textord, "‡", "\\textdaggerdbl");
defineSymbol(math, close, "⎱", "\\rmoustache", true);
defineSymbol(math, open, "⎰", "\\lmoustache", true);
defineSymbol(math, close, "⟯", "\\rgroup", true);
defineSymbol(math, open, "⟮", "\\lgroup", true);
defineSymbol(math, bin, "∓", "\\mp", true);
defineSymbol(math, bin, "⊖", "\\ominus", true);
defineSymbol(math, bin, "⊎", "\\uplus", true);
defineSymbol(math, bin, "⊓", "\\sqcap", true);
defineSymbol(math, bin, "∗", "\\ast");
defineSymbol(math, bin, "⊔", "\\sqcup", true);
defineSymbol(math, bin, "◯", "\\bigcirc", true);
defineSymbol(math, bin, "∙", "\\bullet", true);
defineSymbol(math, bin, "‡", "\\ddagger");
defineSymbol(math, bin, "≀", "\\wr", true);
defineSymbol(math, bin, "⨿", "\\amalg");
defineSymbol(math, bin, "&", "\\And");
defineSymbol(math, bin, "⫽", "\\sslash", true);
defineSymbol(math, rel, "⟵", "\\longleftarrow", true);
defineSymbol(math, rel, "⇐", "\\Leftarrow", true);
defineSymbol(math, rel, "⟸", "\\Longleftarrow", true);
defineSymbol(math, rel, "⟶", "\\longrightarrow", true);
defineSymbol(math, rel, "⇒", "\\Rightarrow", true);
defineSymbol(math, rel, "⟹", "\\Longrightarrow", true);
defineSymbol(math, rel, "↔", "\\leftrightarrow", true);
defineSymbol(math, rel, "⟷", "\\longleftrightarrow", true);
defineSymbol(math, rel, "⇔", "\\Leftrightarrow", true);
defineSymbol(math, rel, "⟺", "\\Longleftrightarrow", true);
defineSymbol(math, rel, "↤", "\\mapsfrom", true);
defineSymbol(math, rel, "↦", "\\mapsto", true);
defineSymbol(math, rel, "⟼", "\\longmapsto", true);
defineSymbol(math, rel, "↗", "\\nearrow", true);
defineSymbol(math, rel, "↩", "\\hookleftarrow", true);
defineSymbol(math, rel, "↪", "\\hookrightarrow", true);
defineSymbol(math, rel, "↘", "\\searrow", true);
defineSymbol(math, rel, "↼", "\\leftharpoonup", true);
defineSymbol(math, rel, "⇀", "\\rightharpoonup", true);
defineSymbol(math, rel, "↙", "\\swarrow", true);
defineSymbol(math, rel, "↽", "\\leftharpoondown", true);
defineSymbol(math, rel, "⇁", "\\rightharpoondown", true);
defineSymbol(math, rel, "↖", "\\nwarrow", true);
defineSymbol(math, rel, "⇌", "\\rightleftharpoons", true);
defineSymbol(math, mathord, "↯", "\\lightning", true);
defineSymbol(math, mathord, "∎", "\\QED", true);
defineSymbol(math, mathord, "‰", "\\permil", true);
defineSymbol(text, textord, "‰", "\\permil");
defineSymbol(math, mathord, "☉", "\\astrosun", true);
defineSymbol(math, mathord, "☼", "\\sun", true);
defineSymbol(math, mathord, "☾", "\\leftmoon", true);
defineSymbol(math, mathord, "☽", "\\rightmoon", true);
defineSymbol(math, mathord, "⊕", "\\Earth");
defineSymbol(math, rel, "≮", "\\nless", true);
defineSymbol(math, rel, "⪇", "\\lneq", true);
defineSymbol(math, rel, "≨", "\\lneqq", true);
defineSymbol(math, rel, "≨︀", "\\lvertneqq");
defineSymbol(math, rel, "⋦", "\\lnsim", true);
defineSymbol(math, rel, "⪉", "\\lnapprox", true);
defineSymbol(math, rel, "⊀", "\\nprec", true);
defineSymbol(math, rel, "⋠", "\\npreceq", true);
defineSymbol(math, rel, "⋨", "\\precnsim", true);
defineSymbol(math, rel, "⪹", "\\precnapprox", true);
defineSymbol(math, rel, "≁", "\\nsim", true);
defineSymbol(math, rel, "∤", "\\nmid", true);
defineSymbol(math, rel, "∤", "\\nshortmid");
defineSymbol(math, rel, "⊬", "\\nvdash", true);
defineSymbol(math, rel, "⊭", "\\nvDash", true);
defineSymbol(math, rel, "⋪", "\\ntriangleleft");
defineSymbol(math, rel, "⋬", "\\ntrianglelefteq", true);
defineSymbol(math, rel, "⊄", "\\nsubset", true);
defineSymbol(math, rel, "⊅", "\\nsupset", true);
defineSymbol(math, rel, "⊊", "\\subsetneq", true);
defineSymbol(math, rel, "⊊︀", "\\varsubsetneq");
defineSymbol(math, rel, "⫋", "\\subsetneqq", true);
defineSymbol(math, rel, "⫋︀", "\\varsubsetneqq");
defineSymbol(math, rel, "≯", "\\ngtr", true);
defineSymbol(math, rel, "⪈", "\\gneq", true);
defineSymbol(math, rel, "≩", "\\gneqq", true);
defineSymbol(math, rel, "≩︀", "\\gvertneqq");
defineSymbol(math, rel, "⋧", "\\gnsim", true);
defineSymbol(math, rel, "⪊", "\\gnapprox", true);
defineSymbol(math, rel, "⊁", "\\nsucc", true);
defineSymbol(math, rel, "⋡", "\\nsucceq", true);
defineSymbol(math, rel, "⋩", "\\succnsim", true);
defineSymbol(math, rel, "⪺", "\\succnapprox", true);
defineSymbol(math, rel, "≆", "\\ncong", true);
defineSymbol(math, rel, "∦", "\\nparallel", true);
defineSymbol(math, rel, "∦", "\\nshortparallel");
defineSymbol(math, rel, "⊯", "\\nVDash", true);
defineSymbol(math, rel, "⋫", "\\ntriangleright");
defineSymbol(math, rel, "⋭", "\\ntrianglerighteq", true);
defineSymbol(math, rel, "⊋", "\\supsetneq", true);
defineSymbol(math, rel, "⊋", "\\varsupsetneq");
defineSymbol(math, rel, "⫌", "\\supsetneqq", true);
defineSymbol(math, rel, "⫌︀", "\\varsupsetneqq");
defineSymbol(math, rel, "⊮", "\\nVdash", true);
defineSymbol(math, rel, "⪵", "\\precneqq", true);
defineSymbol(math, rel, "⪶", "\\succneqq", true);
defineSymbol(math, bin, "⊴", "\\unlhd");
defineSymbol(math, bin, "⊵", "\\unrhd");
defineSymbol(math, rel, "↚", "\\nleftarrow", true);
defineSymbol(math, rel, "↛", "\\nrightarrow", true);
defineSymbol(math, rel, "⇍", "\\nLeftarrow", true);
defineSymbol(math, rel, "⇏", "\\nRightarrow", true);
defineSymbol(math, rel, "↮", "\\nleftrightarrow", true);
defineSymbol(math, rel, "⇎", "\\nLeftrightarrow", true);
defineSymbol(math, rel, "△", "\\vartriangle");
defineSymbol(math, textord, "ℏ", "\\hslash");
defineSymbol(math, textord, "▽", "\\triangledown");
defineSymbol(math, textord, "◊", "\\lozenge");
defineSymbol(math, textord, "Ⓢ", "\\circledS");
defineSymbol(math, textord, "®", "\\circledR", true);
defineSymbol(text, textord, "®", "\\circledR");
defineSymbol(text, textord, "®", "\\textregistered");
defineSymbol(math, textord, "∡", "\\measuredangle", true);
defineSymbol(math, textord, "∄", "\\nexists");
defineSymbol(math, textord, "℧", "\\mho");
defineSymbol(math, textord, "Ⅎ", "\\Finv", true);
defineSymbol(math, textord, "⅁", "\\Game", true);
defineSymbol(math, textord, "‵", "\\backprime");
defineSymbol(math, textord, "‶", "\\backdprime");
defineSymbol(math, textord, "‷", "\\backtrprime");
defineSymbol(math, textord, "▲", "\\blacktriangle");
defineSymbol(math, textord, "▼", "\\blacktriangledown");
defineSymbol(math, textord, "■", "\\blacksquare");
defineSymbol(math, textord, "⧫", "\\blacklozenge");
defineSymbol(math, textord, "★", "\\bigstar");
defineSymbol(math, textord, "∢", "\\sphericalangle", true);
defineSymbol(math, textord, "∁", "\\complement", true);
defineSymbol(math, textord, "╱", "\\diagup");
defineSymbol(math, textord, "╲", "\\diagdown");
defineSymbol(math, textord, "□", "\\square");
defineSymbol(math, textord, "□", "\\Box");
defineSymbol(math, textord, "◊", "\\Diamond");
defineSymbol(math, textord, "¥", "\\yen", true);
defineSymbol(text, textord, "¥", "\\yen", true);
defineSymbol(math, textord, "✓", "\\checkmark", true);
defineSymbol(text, textord, "✓", "\\checkmark");
defineSymbol(math, textord, "✗", "\\ballotx", true);
defineSymbol(text, textord, "✗", "\\ballotx");
defineSymbol(text, textord, "•", "\\textbullet");
defineSymbol(math, textord, "ℶ", "\\beth", true);
defineSymbol(math, textord, "ℸ", "\\daleth", true);
defineSymbol(math, textord, "ℷ", "\\gimel", true);
defineSymbol(math, textord, "ϝ", "\\digamma", true);
defineSymbol(math, textord, "ϰ", "\\varkappa");
defineSymbol(math, open, "⌜", "\\ulcorner", true);
defineSymbol(math, close, "⌝", "\\urcorner", true);
defineSymbol(math, open, "⌞", "\\llcorner", true);
defineSymbol(math, close, "⌟", "\\lrcorner", true);
defineSymbol(math, rel, "≦", "\\leqq", true);
defineSymbol(math, rel, "⩽", "\\leqslant", true);
defineSymbol(math, rel, "⪕", "\\eqslantless", true);
defineSymbol(math, rel, "≲", "\\lesssim", true);
defineSymbol(math, rel, "⪅", "\\lessapprox", true);
defineSymbol(math, rel, "≊", "\\approxeq", true);
defineSymbol(math, bin, "⋖", "\\lessdot");
defineSymbol(math, rel, "⋘", "\\lll", true);
defineSymbol(math, rel, "≶", "\\lessgtr", true);
defineSymbol(math, rel, "⋚", "\\lesseqgtr", true);
defineSymbol(math, rel, "⪋", "\\lesseqqgtr", true);
defineSymbol(math, rel, "≑", "\\doteqdot");
defineSymbol(math, rel, "≓", "\\risingdotseq", true);
defineSymbol(math, rel, "≒", "\\fallingdotseq", true);
defineSymbol(math, rel, "∽", "\\backsim", true);
defineSymbol(math, rel, "⋍", "\\backsimeq", true);
defineSymbol(math, rel, "⫅", "\\subseteqq", true);
defineSymbol(math, rel, "⋐", "\\Subset", true);
defineSymbol(math, rel, "⊏", "\\sqsubset", true);
defineSymbol(math, rel, "≼", "\\preccurlyeq", true);
defineSymbol(math, rel, "⋞", "\\curlyeqprec", true);
defineSymbol(math, rel, "≾", "\\precsim", true);
defineSymbol(math, rel, "⪷", "\\precapprox", true);
defineSymbol(math, rel, "⊲", "\\vartriangleleft");
defineSymbol(math, rel, "⊴", "\\trianglelefteq");
defineSymbol(math, rel, "⊨", "\\vDash", true);
defineSymbol(math, rel, "⊫", "\\VDash", true);
defineSymbol(math, rel, "⊪", "\\Vvdash", true);
defineSymbol(math, rel, "⌣", "\\smallsmile");
defineSymbol(math, rel, "⌢", "\\smallfrown");
defineSymbol(math, rel, "≏", "\\bumpeq", true);
defineSymbol(math, rel, "≎", "\\Bumpeq", true);
defineSymbol(math, rel, "≧", "\\geqq", true);
defineSymbol(math, rel, "⩾", "\\geqslant", true);
defineSymbol(math, rel, "⪖", "\\eqslantgtr", true);
defineSymbol(math, rel, "≳", "\\gtrsim", true);
defineSymbol(math, rel, "⪆", "\\gtrapprox", true);
defineSymbol(math, bin, "⋗", "\\gtrdot");
defineSymbol(math, rel, "⋙", "\\ggg", true);
defineSymbol(math, rel, "≷", "\\gtrless", true);
defineSymbol(math, rel, "⋛", "\\gtreqless", true);
defineSymbol(math, rel, "⪌", "\\gtreqqless", true);
defineSymbol(math, rel, "≖", "\\eqcirc", true);
defineSymbol(math, rel, "≗", "\\circeq", true);
defineSymbol(math, rel, "≜", "\\triangleq", true);
defineSymbol(math, rel, "∼", "\\thicksim");
defineSymbol(math, rel, "≈", "\\thickapprox");
defineSymbol(math, rel, "⫆", "\\supseteqq", true);
defineSymbol(math, rel, "⋑", "\\Supset", true);
defineSymbol(math, rel, "⊐", "\\sqsupset", true);
defineSymbol(math, rel, "≽", "\\succcurlyeq", true);
defineSymbol(math, rel, "⋟", "\\curlyeqsucc", true);
defineSymbol(math, rel, "≿", "\\succsim", true);
defineSymbol(math, rel, "⪸", "\\succapprox", true);
defineSymbol(math, rel, "⊳", "\\vartriangleright");
defineSymbol(math, rel, "⊵", "\\trianglerighteq");
defineSymbol(math, rel, "⊩", "\\Vdash", true);
defineSymbol(math, rel, "∣", "\\shortmid");
defineSymbol(math, rel, "∥", "\\shortparallel");
defineSymbol(math, rel, "≬", "\\between", true);
defineSymbol(math, rel, "⋔", "\\pitchfork", true);
defineSymbol(math, rel, "∝", "\\varpropto");
defineSymbol(math, rel, "◀", "\\blacktriangleleft");
defineSymbol(math, rel, "∴", "\\therefore", true);
defineSymbol(math, rel, "∍", "\\backepsilon");
defineSymbol(math, rel, "▶", "\\blacktriangleright");
defineSymbol(math, rel, "∵", "\\because", true);
defineSymbol(math, rel, "⋘", "\\llless");
defineSymbol(math, rel, "⋙", "\\gggtr");
defineSymbol(math, bin, "⊲", "\\lhd");
defineSymbol(math, bin, "⊳", "\\rhd");
defineSymbol(math, rel, "≂", "\\eqsim", true);
defineSymbol(math, rel, "≑", "\\Doteq", true);
defineSymbol(math, rel, "⥽", "\\strictif", true);
defineSymbol(math, rel, "⥼", "\\strictfi", true);
defineSymbol(math, bin, "∔", "\\dotplus", true);
defineSymbol(math, bin, "∖", "\\smallsetminus");
defineSymbol(math, bin, "⋒", "\\Cap", true);
defineSymbol(math, bin, "⋓", "\\Cup", true);
defineSymbol(math, bin, "⩞", "\\doublebarwedge", true);
defineSymbol(math, bin, "⊟", "\\boxminus", true);
defineSymbol(math, bin, "⊞", "\\boxplus", true);
defineSymbol(math, bin, "⧄", "\\boxslash", true);
defineSymbol(math, bin, "⋇", "\\divideontimes", true);
defineSymbol(math, bin, "⋉", "\\ltimes", true);
defineSymbol(math, bin, "⋊", "\\rtimes", true);
defineSymbol(math, bin, "⋋", "\\leftthreetimes", true);
defineSymbol(math, bin, "⋌", "\\rightthreetimes", true);
defineSymbol(math, bin, "⋏", "\\curlywedge", true);
defineSymbol(math, bin, "⋎", "\\curlyvee", true);
defineSymbol(math, bin, "⊝", "\\circleddash", true);
defineSymbol(math, bin, "⊛", "\\circledast", true);
defineSymbol(math, bin, "⊺", "\\intercal", true);
defineSymbol(math, bin, "⋒", "\\doublecap");
defineSymbol(math, bin, "⋓", "\\doublecup");
defineSymbol(math, bin, "⊠", "\\boxtimes", true);
defineSymbol(math, bin, "⋈", "\\bowtie", true);
defineSymbol(math, bin, "⋈", "\\Join");
defineSymbol(math, bin, "⟕", "\\leftouterjoin", true);
defineSymbol(math, bin, "⟖", "\\rightouterjoin", true);
defineSymbol(math, bin, "⟗", "\\fullouterjoin", true);
defineSymbol(math, bin, "∸", "\\dotminus", true);
defineSymbol(math, bin, "⟑", "\\wedgedot", true);
defineSymbol(math, bin, "⟇", "\\veedot", true);
defineSymbol(math, bin, "⩢", "\\doublebarvee", true);
defineSymbol(math, bin, "⩣", "\\veedoublebar", true);
defineSymbol(math, bin, "⩟", "\\wedgebar", true);
defineSymbol(math, bin, "⩠", "\\wedgedoublebar", true);
defineSymbol(math, bin, "⩔", "\\Vee", true);
defineSymbol(math, bin, "⩓", "\\Wedge", true);
defineSymbol(math, bin, "⩃", "\\barcap", true);
defineSymbol(math, bin, "⩂", "\\barcup", true);
defineSymbol(math, bin, "⩈", "\\capbarcup", true);
defineSymbol(math, bin, "⩀", "\\capdot", true);
defineSymbol(math, bin, "⩇", "\\capovercup", true);
defineSymbol(math, bin, "⩆", "\\cupovercap", true);
defineSymbol(math, bin, "⩍", "\\closedvarcap", true);
defineSymbol(math, bin, "⩌", "\\closedvarcup", true);
defineSymbol(math, bin, "⨪", "\\minusdot", true);
defineSymbol(math, bin, "⨫", "\\minusfdots", true);
defineSymbol(math, bin, "⨬", "\\minusrdots", true);
defineSymbol(math, bin, "⊻", "\\Xor", true);
defineSymbol(math, bin, "⊼", "\\Nand", true);
defineSymbol(math, bin, "⊽", "\\Nor", true);
defineSymbol(math, bin, "⊽", "\\barvee");
defineSymbol(math, bin, "⫴", "\\interleave", true);
defineSymbol(math, bin, "⧢", "\\shuffle", true);
defineSymbol(math, bin, "⫶", "\\threedotcolon", true);
defineSymbol(math, bin, "⦂", "\\typecolon", true);
defineSymbol(math, bin, "∾", "\\invlazys", true);
defineSymbol(math, bin, "⩋", "\\twocaps", true);
defineSymbol(math, bin, "⩊", "\\twocups", true);
defineSymbol(math, bin, "⩎", "\\Sqcap", true);
defineSymbol(math, bin, "⩏", "\\Sqcup", true);
defineSymbol(math, bin, "⩖", "\\veeonvee", true);
defineSymbol(math, bin, "⩕", "\\wedgeonwedge", true);
defineSymbol(math, bin, "⧗", "\\blackhourglass", true);
defineSymbol(math, bin, "⧆", "\\boxast", true);
defineSymbol(math, bin, "⧈", "\\boxbox", true);
defineSymbol(math, bin, "⧇", "\\boxcircle", true);
defineSymbol(math, bin, "⊜", "\\circledequal", true);
defineSymbol(math, bin, "⦷", "\\circledparallel", true);
defineSymbol(math, bin, "⦶", "\\circledvert", true);
defineSymbol(math, bin, "⦵", "\\circlehbar", true);
defineSymbol(math, bin, "⟡", "\\concavediamond", true);
defineSymbol(math, bin, "⟢", "\\concavediamondtickleft", true);
defineSymbol(math, bin, "⟣", "\\concavediamondtickright", true);
defineSymbol(math, bin, "⋄", "\\diamond", true);
defineSymbol(math, bin, "⧖", "\\hourglass", true);
defineSymbol(math, bin, "⟠", "\\lozengeminus", true);
defineSymbol(math, bin, "⌽", "\\obar", true);
defineSymbol(math, bin, "⦸", "\\obslash", true);
defineSymbol(math, bin, "⨸", "\\odiv", true);
defineSymbol(math, bin, "⧁", "\\ogreaterthan", true);
defineSymbol(math, bin, "⧀", "\\olessthan", true);
defineSymbol(math, bin, "⦹", "\\operp", true);
defineSymbol(math, bin, "⨷", "\\Otimes", true);
defineSymbol(math, bin, "⨶", "\\otimeshat", true);
defineSymbol(math, bin, "⋆", "\\star", true);
defineSymbol(math, bin, "△", "\\triangle", true);
defineSymbol(math, bin, "⨺", "\\triangleminus", true);
defineSymbol(math, bin, "⨹", "\\triangleplus", true);
defineSymbol(math, bin, "⨻", "\\triangletimes", true);
defineSymbol(math, bin, "⟤", "\\whitesquaretickleft", true);
defineSymbol(math, bin, "⟥", "\\whitesquaretickright", true);
defineSymbol(math, bin, "⨳", "\\smashtimes", true);
defineSymbol(math, rel, "⇢", "\\dashrightarrow", true);
defineSymbol(math, rel, "⇠", "\\dashleftarrow", true);
defineSymbol(math, rel, "⇇", "\\leftleftarrows", true);
defineSymbol(math, rel, "⇆", "\\leftrightarrows", true);
defineSymbol(math, rel, "⇚", "\\Lleftarrow", true);
defineSymbol(math, rel, "↞", "\\twoheadleftarrow", true);
defineSymbol(math, rel, "↢", "\\leftarrowtail", true);
defineSymbol(math, rel, "↫", "\\looparrowleft", true);
defineSymbol(math, rel, "⇋", "\\leftrightharpoons", true);
defineSymbol(math, rel, "↶", "\\curvearrowleft", true);
defineSymbol(math, rel, "↺", "\\circlearrowleft", true);
defineSymbol(math, rel, "↰", "\\Lsh", true);
defineSymbol(math, rel, "⇈", "\\upuparrows", true);
defineSymbol(math, rel, "↿", "\\upharpoonleft", true);
defineSymbol(math, rel, "⇃", "\\downharpoonleft", true);
defineSymbol(math, rel, "⊶", "\\origof", true);
defineSymbol(math, rel, "⊷", "\\imageof", true);
defineSymbol(math, rel, "⊸", "\\multimap", true);
defineSymbol(math, rel, "↭", "\\leftrightsquigarrow", true);
defineSymbol(math, rel, "⇉", "\\rightrightarrows", true);
defineSymbol(math, rel, "⇄", "\\rightleftarrows", true);
defineSymbol(math, rel, "↠", "\\twoheadrightarrow", true);
defineSymbol(math, rel, "↣", "\\rightarrowtail", true);
defineSymbol(math, rel, "↬", "\\looparrowright", true);
defineSymbol(math, rel, "↷", "\\curvearrowright", true);
defineSymbol(math, rel, "↻", "\\circlearrowright", true);
defineSymbol(math, rel, "↱", "\\Rsh", true);
defineSymbol(math, rel, "⇊", "\\downdownarrows", true);
defineSymbol(math, rel, "↾", "\\upharpoonright", true);
defineSymbol(math, rel, "⇂", "\\downharpoonright", true);
defineSymbol(math, rel, "⇝", "\\rightsquigarrow", true);
defineSymbol(math, rel, "⇝", "\\leadsto");
defineSymbol(math, rel, "⇛", "\\Rrightarrow", true);
defineSymbol(math, rel, "↾", "\\restriction");
defineSymbol(math, textord, "‘", "`");
defineSymbol(math, textord, "$", "\\$");
defineSymbol(text, textord, "$", "\\$");
defineSymbol(text, textord, "$", "\\textdollar");
defineSymbol(math, textord, "¢", "\\cent");
defineSymbol(text, textord, "¢", "\\cent");
defineSymbol(math, textord, "%", "\\%");
defineSymbol(text, textord, "%", "\\%");
defineSymbol(math, textord, "_", "\\_");
defineSymbol(text, textord, "_", "\\_");
defineSymbol(text, textord, "_", "\\textunderscore");
defineSymbol(text, textord, "␣", "\\textvisiblespace", true);
defineSymbol(math, textord, "∠", "\\angle", true);
defineSymbol(math, textord, "∞", "\\infty", true);
defineSymbol(math, textord, "′", "\\prime");
defineSymbol(math, textord, "″", "\\dprime");
defineSymbol(math, textord, "‴", "\\trprime");
defineSymbol(math, textord, "⁗", "\\qprime");
defineSymbol(math, textord, "△", "\\triangle");
defineSymbol(text, textord, "Α", "\\Alpha", true);
defineSymbol(text, textord, "Β", "\\Beta", true);
defineSymbol(text, textord, "Γ", "\\Gamma", true);
defineSymbol(text, textord, "Δ", "\\Delta", true);
defineSymbol(text, textord, "Ε", "\\Epsilon", true);
defineSymbol(text, textord, "Ζ", "\\Zeta", true);
defineSymbol(text, textord, "Η", "\\Eta", true);
defineSymbol(text, textord, "Θ", "\\Theta", true);
defineSymbol(text, textord, "Ι", "\\Iota", true);
defineSymbol(text, textord, "Κ", "\\Kappa", true);
defineSymbol(text, textord, "Λ", "\\Lambda", true);
defineSymbol(text, textord, "Μ", "\\Mu", true);
defineSymbol(text, textord, "Ν", "\\Nu", true);
defineSymbol(text, textord, "Ξ", "\\Xi", true);
defineSymbol(text, textord, "Ο", "\\Omicron", true);
defineSymbol(text, textord, "Π", "\\Pi", true);
defineSymbol(text, textord, "Ρ", "\\Rho", true);
defineSymbol(text, textord, "Σ", "\\Sigma", true);
defineSymbol(text, textord, "Τ", "\\Tau", true);
defineSymbol(text, textord, "Υ", "\\Upsilon", true);
defineSymbol(text, textord, "Φ", "\\Phi", true);
defineSymbol(text, textord, "Χ", "\\Chi", true);
defineSymbol(text, textord, "Ψ", "\\Psi", true);
defineSymbol(text, textord, "Ω", "\\Omega", true);
defineSymbol(math, mathord, "Α", "\\Alpha", true);
defineSymbol(math, mathord, "Β", "\\Beta", true);
defineSymbol(math, mathord, "Γ", "\\Gamma", true);
defineSymbol(math, mathord, "Δ", "\\Delta", true);
defineSymbol(math, mathord, "Ε", "\\Epsilon", true);
defineSymbol(math, mathord, "Ζ", "\\Zeta", true);
defineSymbol(math, mathord, "Η", "\\Eta", true);
defineSymbol(math, mathord, "Θ", "\\Theta", true);
defineSymbol(math, mathord, "Ι", "\\Iota", true);
defineSymbol(math, mathord, "Κ", "\\Kappa", true);
defineSymbol(math, mathord, "Λ", "\\Lambda", true);
defineSymbol(math, mathord, "Μ", "\\Mu", true);
defineSymbol(math, mathord, "Ν", "\\Nu", true);
defineSymbol(math, mathord, "Ξ", "\\Xi", true);
defineSymbol(math, mathord, "Ο", "\\Omicron", true);
defineSymbol(math, mathord, "Π", "\\Pi", true);
defineSymbol(math, mathord, "Ρ", "\\Rho", true);
defineSymbol(math, mathord, "Σ", "\\Sigma", true);
defineSymbol(math, mathord, "Τ", "\\Tau", true);
defineSymbol(math, mathord, "Υ", "\\Upsilon", true);
defineSymbol(math, mathord, "Φ", "\\Phi", true);
defineSymbol(math, mathord, "Χ", "\\Chi", true);
defineSymbol(math, mathord, "Ψ", "\\Psi", true);
defineSymbol(math, mathord, "Ω", "\\Omega", true);
defineSymbol(math, open, "¬", "\\neg", true);
defineSymbol(math, open, "¬", "\\lnot");
defineSymbol(math, textord, "⊤", "\\top");
defineSymbol(math, textord, "⊥", "\\bot");
defineSymbol(math, textord, "∅", "\\emptyset");
defineSymbol(math, textord, "⌀", "\\varnothing");
defineSymbol(math, mathord, "α", "\\alpha", true);
defineSymbol(math, mathord, "β", "\\beta", true);
defineSymbol(math, mathord, "γ", "\\gamma", true);
defineSymbol(math, mathord, "δ", "\\delta", true);
defineSymbol(math, mathord, "ϵ", "\\epsilon", true);
defineSymbol(math, mathord, "ζ", "\\zeta", true);
defineSymbol(math, mathord, "η", "\\eta", true);
defineSymbol(math, mathord, "θ", "\\theta", true);
defineSymbol(math, mathord, "ι", "\\iota", true);
defineSymbol(math, mathord, "κ", "\\kappa", true);
defineSymbol(math, mathord, "λ", "\\lambda", true);
defineSymbol(math, mathord, "μ", "\\mu", true);
defineSymbol(math, mathord, "ν", "\\nu", true);
defineSymbol(math, mathord, "ξ", "\\xi", true);
defineSymbol(math, mathord, "ο", "\\omicron", true);
defineSymbol(math, mathord, "π", "\\pi", true);
defineSymbol(math, mathord, "ρ", "\\rho", true);
defineSymbol(math, mathord, "σ", "\\sigma", true);
defineSymbol(math, mathord, "τ", "\\tau", true);
defineSymbol(math, mathord, "υ", "\\upsilon", true);
defineSymbol(math, mathord, "ϕ", "\\phi", true);
defineSymbol(math, mathord, "χ", "\\chi", true);
defineSymbol(math, mathord, "ψ", "\\psi", true);
defineSymbol(math, mathord, "ω", "\\omega", true);
defineSymbol(math, mathord, "ε", "\\varepsilon", true);
defineSymbol(math, mathord, "ϑ", "\\vartheta", true);
defineSymbol(math, mathord, "ϖ", "\\varpi", true);
defineSymbol(math, mathord, "ϱ", "\\varrho", true);
defineSymbol(math, mathord, "ς", "\\varsigma", true);
defineSymbol(math, mathord, "φ", "\\varphi", true);
defineSymbol(math, mathord, "Ϙ", "\\Coppa", true);
defineSymbol(math, mathord, "ϙ", "\\coppa", true);
defineSymbol(math, mathord, "ϙ", "\\varcoppa", true);
defineSymbol(math, mathord, "Ϟ", "\\Koppa", true);
defineSymbol(math, mathord, "ϟ", "\\koppa", true);
defineSymbol(math, mathord, "Ϡ", "\\Sampi", true);
defineSymbol(math, mathord, "ϡ", "\\sampi", true);
defineSymbol(math, mathord, "Ϛ", "\\Stigma", true);
defineSymbol(math, mathord, "ϛ", "\\stigma", true);
defineSymbol(math, mathord, "⫫", "\\Bot");
defineSymbol(math, textord, "ð", "\\eth", true);
defineSymbol(text, textord, "ð", "ð");
defineSymbol(math, textord, "Å", "\\AA");
defineSymbol(text, textord, "Å", "\\AA", true);
defineSymbol(math, textord, "Æ", "\\AE", true);
defineSymbol(text, textord, "Æ", "\\AE", true);
defineSymbol(math, textord, "Ð", "\\DH", true);
defineSymbol(text, textord, "Ð", "\\DH", true);
defineSymbol(math, textord, "Þ", "\\TH", true);
defineSymbol(text, textord, "Þ", "\\TH", true);
defineSymbol(math, textord, "ß", "\\ss", true);
defineSymbol(text, textord, "ß", "\\ss", true);
defineSymbol(math, textord, "å", "\\aa");
defineSymbol(text, textord, "å", "\\aa", true);
defineSymbol(math, textord, "æ", "\\ae", true);
defineSymbol(text, textord, "æ", "\\ae", true);
defineSymbol(math, textord, "ð", "\\dh");
defineSymbol(text, textord, "ð", "\\dh", true);
defineSymbol(math, textord, "þ", "\\th", true);
defineSymbol(text, textord, "þ", "\\th", true);
defineSymbol(math, textord, "Đ", "\\DJ", true);
defineSymbol(text, textord, "Đ", "\\DJ", true);
defineSymbol(math, textord, "đ", "\\dj", true);
defineSymbol(text, textord, "đ", "\\dj", true);
defineSymbol(math, textord, "Ł", "\\L", true);
defineSymbol(text, textord, "Ł", "\\L", true);
defineSymbol(math, textord, "Ł", "\\l", true);
defineSymbol(text, textord, "Ł", "\\l", true);
defineSymbol(math, textord, "Ŋ", "\\NG", true);
defineSymbol(text, textord, "Ŋ", "\\NG", true);
defineSymbol(math, textord, "ŋ", "\\ng", true);
defineSymbol(text, textord, "ŋ", "\\ng", true);
defineSymbol(math, textord, "Œ", "\\OE", true);
defineSymbol(text, textord, "Œ", "\\OE", true);
defineSymbol(math, textord, "œ", "\\oe", true);
defineSymbol(text, textord, "œ", "\\oe", true);
defineSymbol(math, bin, "∗", "∗", true);
defineSymbol(math, bin, "+", "+");
defineSymbol(math, bin, "∗", "*");
defineSymbol(math, bin, "⁄", "/", true);
defineSymbol(math, bin, "⁄", "⁄");
defineSymbol(math, bin, "−", "-", true);
defineSymbol(math, bin, "⋅", "\\cdot", true);
defineSymbol(math, bin, "∘", "\\circ", true);
defineSymbol(math, bin, "÷", "\\div", true);
defineSymbol(math, bin, "±", "\\pm", true);
defineSymbol(math, bin, "×", "\\times", true);
defineSymbol(math, bin, "∩", "\\cap", true);
defineSymbol(math, bin, "∪", "\\cup", true);
defineSymbol(math, bin, "∖", "\\setminus", true);
defineSymbol(math, bin, "∧", "\\land");
defineSymbol(math, bin, "∨", "\\lor");
defineSymbol(math, bin, "∧", "\\wedge", true);
defineSymbol(math, bin, "∨", "\\vee", true);
defineSymbol(math, open, "⟦", "\\llbracket", true);
defineSymbol(math, close, "⟧", "\\rrbracket", true);
defineSymbol(math, open, "⟨", "\\langle", true);
defineSymbol(math, open, "⟪", "\\lAngle", true);
defineSymbol(math, open, "⦉", "\\llangle", true);
defineSymbol(math, open, "|", "\\lvert");
defineSymbol(math, open, "‖", "\\lVert", true);
defineSymbol(math, textord, "!", "\\oc");
defineSymbol(math, textord, "?", "\\wn");
defineSymbol(math, textord, "↓", "\\shpos");
defineSymbol(math, textord, "↕", "\\shift");
defineSymbol(math, textord, "↑", "\\shneg");
defineSymbol(math, close, "?", "?");
defineSymbol(math, close, "!", "!");
defineSymbol(math, close, "‼", "‼");
defineSymbol(math, close, "⟩", "\\rangle", true);
defineSymbol(math, close, "⟫", "\\rAngle", true);
defineSymbol(math, close, "⦊", "\\rrangle", true);
defineSymbol(math, close, "|", "\\rvert");
defineSymbol(math, close, "‖", "\\rVert");
defineSymbol(math, open, "⦃", "\\lBrace", true);
defineSymbol(math, close, "⦄", "\\rBrace", true);
defineSymbol(math, rel, "=", "\\equal", true);
defineSymbol(math, rel, ":", ":");
defineSymbol(math, rel, "≈", "\\approx", true);
defineSymbol(math, rel, "≅", "\\cong", true);
defineSymbol(math, rel, "≥", "\\ge");
defineSymbol(math, rel, "≥", "\\geq", true);
defineSymbol(math, rel, "←", "\\gets");
defineSymbol(math, rel, ">", "\\gt", true);
defineSymbol(math, rel, "∈", "\\in", true);
defineSymbol(math, rel, "∉", "\\notin", true);
defineSymbol(math, rel, "", "\\@not");
defineSymbol(math, rel, "⊂", "\\subset", true);
defineSymbol(math, rel, "⊃", "\\supset", true);
defineSymbol(math, rel, "⊆", "\\subseteq", true);
defineSymbol(math, rel, "⊇", "\\supseteq", true);
defineSymbol(math, rel, "⊈", "\\nsubseteq", true);
defineSymbol(math, rel, "⊈", "\\nsubseteqq");
defineSymbol(math, rel, "⊉", "\\nsupseteq", true);
defineSymbol(math, rel, "⊉", "\\nsupseteqq");
defineSymbol(math, rel, "⊨", "\\models");
defineSymbol(math, rel, "←", "\\leftarrow", true);
defineSymbol(math, rel, "≤", "\\le");
defineSymbol(math, rel, "≤", "\\leq", true);
defineSymbol(math, rel, "<", "\\lt", true);
defineSymbol(math, rel, "→", "\\rightarrow", true);
defineSymbol(math, rel, "→", "\\to");
defineSymbol(math, rel, "≱", "\\ngeq", true);
defineSymbol(math, rel, "≱", "\\ngeqq");
defineSymbol(math, rel, "≱", "\\ngeqslant");
defineSymbol(math, rel, "≰", "\\nleq", true);
defineSymbol(math, rel, "≰", "\\nleqq");
defineSymbol(math, rel, "≰", "\\nleqslant");
defineSymbol(math, rel, "⫫", "\\Perp", true);
defineSymbol(math, spacing, "\xA0", "\\ ");
defineSymbol(math, spacing, "\xA0", "\\space");
defineSymbol(math, spacing, "\xA0", "\\nobreakspace");
defineSymbol(text, spacing, "\xA0", "\\ ");
defineSymbol(text, spacing, "\xA0", " ");
defineSymbol(text, spacing, "\xA0", "\\space");
defineSymbol(text, spacing, "\xA0", "\\nobreakspace");
defineSymbol(math, spacing, null, "\\nobreak");
defineSymbol(math, spacing, null, "\\allowbreak");
defineSymbol(math, punct, ",", ",");
defineSymbol(text, punct, ":", ":");
defineSymbol(math, punct, ";", ";");
defineSymbol(math, bin, "⊼", "\\barwedge");
defineSymbol(math, bin, "⊻", "\\veebar");
defineSymbol(math, bin, "⊙", "\\odot", true);
defineSymbol(math, bin, "⊕︎", "\\oplus");
defineSymbol(math, bin, "⊗", "\\otimes", true);
defineSymbol(math, textord, "∂", "\\partial", true);
defineSymbol(math, bin, "⊘", "\\oslash", true);
defineSymbol(math, bin, "⊚", "\\circledcirc", true);
defineSymbol(math, bin, "⊡", "\\boxdot", true);
defineSymbol(math, bin, "△", "\\bigtriangleup");
defineSymbol(math, bin, "▽", "\\bigtriangledown");
defineSymbol(math, bin, "†", "\\dagger");
defineSymbol(math, bin, "⋄", "\\diamond");
defineSymbol(math, bin, "◃", "\\triangleleft");
defineSymbol(math, bin, "▹", "\\triangleright");
defineSymbol(math, open, "{", "\\{");
defineSymbol(text, textord, "{", "\\{");
defineSymbol(text, textord, "{", "\\textbraceleft");
defineSymbol(math, close, "}", "\\}");
defineSymbol(text, textord, "}", "\\}");
defineSymbol(text, textord, "}", "\\textbraceright");
defineSymbol(math, open, "{", "\\lbrace");
defineSymbol(math, close, "}", "\\rbrace");
defineSymbol(math, open, "[", "\\lbrack", true);
defineSymbol(text, textord, "[", "\\lbrack", true);
defineSymbol(math, close, "]", "\\rbrack", true);
defineSymbol(text, textord, "]", "\\rbrack", true);
defineSymbol(math, open, "(", "\\lparen", true);
defineSymbol(math, close, ")", "\\rparen", true);
defineSymbol(math, open, "⦇", "\\llparenthesis", true);
defineSymbol(math, close, "⦈", "\\rrparenthesis", true);
defineSymbol(text, textord, "<", "\\textless", true);
defineSymbol(text, textord, ">", "\\textgreater", true);
defineSymbol(math, open, "⌊", "\\lfloor", true);
defineSymbol(math, close, "⌋", "\\rfloor", true);
defineSymbol(math, open, "⌈", "\\lceil", true);
defineSymbol(math, close, "⌉", "\\rceil", true);
defineSymbol(math, textord, "\\", "\\backslash");
defineSymbol(math, textord, "|", "|");
defineSymbol(math, textord, "|", "\\vert");
defineSymbol(text, textord, "|", "\\textbar", true);
defineSymbol(math, textord, "‖", "\\|");
defineSymbol(math, textord, "‖", "\\Vert");
defineSymbol(text, textord, "‖", "\\textbardbl");
defineSymbol(text, textord, "~", "\\textasciitilde");
defineSymbol(text, textord, "\\", "\\textbackslash");
defineSymbol(text, textord, "^", "\\textasciicircum");
defineSymbol(math, rel, "↑", "\\uparrow", true);
defineSymbol(math, rel, "⇑", "\\Uparrow", true);
defineSymbol(math, rel, "↓", "\\downarrow", true);
defineSymbol(math, rel, "⇓", "\\Downarrow", true);
defineSymbol(math, rel, "↕", "\\updownarrow", true);
defineSymbol(math, rel, "⇕", "\\Updownarrow", true);
defineSymbol(math, op, "∐", "\\coprod");
defineSymbol(math, op, "⋁", "\\bigvee");
defineSymbol(math, op, "⋀", "\\bigwedge");
defineSymbol(math, op, "⨄", "\\biguplus");
defineSymbol(math, op, "⨄", "\\bigcupplus");
defineSymbol(math, op, "⨃", "\\bigcupdot");
defineSymbol(math, op, "⨇", "\\bigdoublevee");
defineSymbol(math, op, "⨈", "\\bigdoublewedge");
defineSymbol(math, op, "⋂", "\\bigcap");
defineSymbol(math, op, "⋃", "\\bigcup");
defineSymbol(math, op, "∫", "\\int");
defineSymbol(math, op, "∫", "\\intop");
defineSymbol(math, op, "∬", "\\iint");
defineSymbol(math, op, "∭", "\\iiint");
defineSymbol(math, op, "∏", "\\prod");
defineSymbol(math, op, "∑", "\\sum");
defineSymbol(math, op, "⨂", "\\bigotimes");
defineSymbol(math, op, "⨁", "\\bigoplus");
defineSymbol(math, op, "⨀", "\\bigodot");
defineSymbol(math, op, "⨉", "\\bigtimes");
defineSymbol(math, op, "∮", "\\oint");
defineSymbol(math, op, "∯", "\\oiint");
defineSymbol(math, op, "∰", "\\oiiint");
defineSymbol(math, op, "∱", "\\intclockwise");
defineSymbol(math, op, "∲", "\\varointclockwise");
defineSymbol(math, op, "⨌", "\\iiiint");
defineSymbol(math, op, "⨍", "\\intbar");
defineSymbol(math, op, "⨎", "\\intBar");
defineSymbol(math, op, "⨏", "\\fint");
defineSymbol(math, op, "⨒", "\\rppolint");
defineSymbol(math, op, "⨓", "\\scpolint");
defineSymbol(math, op, "⨕", "\\pointint");
defineSymbol(math, op, "⨖", "\\sqint");
defineSymbol(math, op, "⨗", "\\intlarhk");
defineSymbol(math, op, "⨘", "\\intx");
defineSymbol(math, op, "⨙", "\\intcap");
defineSymbol(math, op, "⨚", "\\intcup");
defineSymbol(math, op, "⨅", "\\bigsqcap");
defineSymbol(math, op, "⨆", "\\bigsqcup");
defineSymbol(math, op, "∫", "\\smallint");
defineSymbol(text, inner, "…", "\\textellipsis");
defineSymbol(math, inner, "…", "\\mathellipsis");
defineSymbol(text, inner, "…", "\\ldots", true);
defineSymbol(math, inner, "…", "\\ldots", true);
defineSymbol(math, inner, "⋰", "\\iddots", true);
defineSymbol(math, inner, "⋯", "\\@cdots", true);
defineSymbol(math, inner, "⋱", "\\ddots", true);
defineSymbol(math, textord, "⋮", "\\varvdots");
defineSymbol(text, textord, "⋮", "\\varvdots");
defineSymbol(math, accent, "´", "\\acute");
defineSymbol(math, accent, "`", "\\grave");
defineSymbol(math, accent, "¨", "\\ddot");
defineSymbol(math, accent, "…", "\\dddot");
defineSymbol(math, accent, "….", "\\ddddot");
defineSymbol(math, accent, "~", "\\tilde");
defineSymbol(math, accent, "‾", "\\bar");
defineSymbol(math, accent, "˘", "\\breve");
defineSymbol(math, accent, "ˇ", "\\check");
defineSymbol(math, accent, "ˆ", "\\hat");
defineSymbol(math, accent, "→", "\\vec");
defineSymbol(math, accent, "˙", "\\dot");
defineSymbol(math, accent, "˚", "\\mathring");
defineSymbol(math, mathord, "ı", "\\imath", true);
defineSymbol(math, mathord, "ȷ", "\\jmath", true);
defineSymbol(math, textord, "ı", "ı");
defineSymbol(math, textord, "ȷ", "ȷ");
defineSymbol(text, textord, "ı", "\\i", true);
defineSymbol(text, textord, "ȷ", "\\j", true);
defineSymbol(text, textord, "ø", "\\o", true);
defineSymbol(math, mathord, "ø", "\\o", true);
defineSymbol(text, textord, "Ø", "\\O", true);
defineSymbol(math, mathord, "Ø", "\\O", true);
defineSymbol(text, accent, "ˊ", "\\'");
defineSymbol(text, accent, "ˋ", "\\`");
defineSymbol(text, accent, "ˆ", "\\^");
defineSymbol(text, accent, "~", "\\~");
defineSymbol(text, accent, "ˉ", "\\=");
defineSymbol(text, accent, "˘", "\\u");
defineSymbol(text, accent, "˙", "\\.");
defineSymbol(text, accent, "¸", "\\c");
defineSymbol(text, accent, "˚", "\\r");
defineSymbol(text, accent, "ˇ", "\\v");
defineSymbol(text, accent, "¨", "\\\"");
defineSymbol(text, accent, "˝", "\\H");
defineSymbol(math, accent, "ˊ", "\\'");
defineSymbol(math, accent, "ˋ", "\\`");
defineSymbol(math, accent, "ˆ", "\\^");
defineSymbol(math, accent, "~", "\\~");
defineSymbol(math, accent, "ˉ", "\\=");
defineSymbol(math, accent, "˘", "\\u");
defineSymbol(math, accent, "˙", "\\.");
defineSymbol(math, accent, "¸", "\\c");
defineSymbol(math, accent, "˚", "\\r");
defineSymbol(math, accent, "ˇ", "\\v");
defineSymbol(math, accent, "¨", "\\\"");
defineSymbol(math, accent, "˝", "\\H");
const ligatures = {
	"--": true,
	"---": true,
	"``": true,
	"''": true
};
defineSymbol(text, textord, "–", "--", true);
defineSymbol(text, textord, "–", "\\textendash");
defineSymbol(text, textord, "—", "---", true);
defineSymbol(text, textord, "—", "\\textemdash");
defineSymbol(text, textord, "‘", "`", true);
defineSymbol(text, textord, "‘", "\\textquoteleft");
defineSymbol(text, textord, "’", "'", true);
defineSymbol(text, textord, "’", "\\textquoteright");
defineSymbol(text, textord, "“", "``", true);
defineSymbol(text, textord, "“", "\\textquotedblleft");
defineSymbol(text, textord, "”", "''", true);
defineSymbol(text, textord, "”", "\\textquotedblright");
defineSymbol(math, textord, "°", "\\degree", true);
defineSymbol(text, textord, "°", "\\degree");
defineSymbol(text, textord, "°", "\\textdegree", true);
defineSymbol(math, textord, "£", "\\pounds");
defineSymbol(math, textord, "£", "\\mathsterling", true);
defineSymbol(text, textord, "£", "\\pounds");
defineSymbol(text, textord, "£", "\\textsterling", true);
defineSymbol(math, textord, "✠", "\\maltese");
defineSymbol(text, textord, "✠", "\\maltese");
defineSymbol(math, textord, "€", "\\euro", true);
defineSymbol(text, textord, "€", "\\euro", true);
defineSymbol(text, textord, "€", "\\texteuro");
defineSymbol(math, textord, "©", "\\copyright", true);
defineSymbol(text, textord, "©", "\\textcopyright");
defineSymbol(math, textord, "⌀", "\\diameter", true);
defineSymbol(text, textord, "⌀", "\\diameter");
defineSymbol(math, textord, "𝛤", "\\varGamma");
defineSymbol(math, textord, "𝛥", "\\varDelta");
defineSymbol(math, textord, "𝛩", "\\varTheta");
defineSymbol(math, textord, "𝛬", "\\varLambda");
defineSymbol(math, textord, "𝛯", "\\varXi");
defineSymbol(math, textord, "𝛱", "\\varPi");
defineSymbol(math, textord, "𝛴", "\\varSigma");
defineSymbol(math, textord, "𝛶", "\\varUpsilon");
defineSymbol(math, textord, "𝛷", "\\varPhi");
defineSymbol(math, textord, "𝛹", "\\varPsi");
defineSymbol(math, textord, "𝛺", "\\varOmega");
defineSymbol(text, textord, "𝛤", "\\varGamma");
defineSymbol(text, textord, "𝛥", "\\varDelta");
defineSymbol(text, textord, "𝛩", "\\varTheta");
defineSymbol(text, textord, "𝛬", "\\varLambda");
defineSymbol(text, textord, "𝛯", "\\varXi");
defineSymbol(text, textord, "𝛱", "\\varPi");
defineSymbol(text, textord, "𝛴", "\\varSigma");
defineSymbol(text, textord, "𝛶", "\\varUpsilon");
defineSymbol(text, textord, "𝛷", "\\varPhi");
defineSymbol(text, textord, "𝛹", "\\varPsi");
defineSymbol(text, textord, "𝛺", "\\varOmega");
const mathTextSymbols = "0123456789/@.\"";
for (let i = 0; i < 14; i++) {
	const ch = mathTextSymbols.charAt(i);
	defineSymbol(math, textord, ch, ch);
}
const textSymbols = "0123456789!@*()-=+\";:?/.,";
for (let i = 0; i < 25; i++) {
	const ch = textSymbols.charAt(i);
	defineSymbol(text, textord, ch, ch);
}
const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
for (let i = 0; i < 52; i++) {
	const ch = letters.charAt(i);
	defineSymbol(math, mathord, ch, ch);
	defineSymbol(text, textord, ch, ch);
}
const narrow = "ÇÐÞçþℂℍℕℙℚℝℤℎℏℊℋℌℐℑℒℓ℘ℛℜℬℰℱℳℭℨ";
for (let i = 0; i < 30; i++) {
	const ch = narrow.charAt(i);
	defineSymbol(math, mathord, ch, ch);
	defineSymbol(text, textord, ch, ch);
}
let wideChar = "";
for (let i = 0; i < 52; i++) {
	wideChar = String.fromCharCode(55349, 56320 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56372 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56424 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56580 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56736 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56788 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56840 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56944 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 56632 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	const ch = letters.charAt(i);
	wideChar = String.fromCharCode(55349, 56476 + i);
	defineSymbol(math, mathord, ch, wideChar);
	defineSymbol(text, textord, ch, wideChar);
}
for (let i = 0; i < 10; i++) {
	wideChar = String.fromCharCode(55349, 57294 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 57314 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 57324 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
	wideChar = String.fromCharCode(55349, 57334 + i);
	defineSymbol(math, mathord, wideChar, wideChar);
	defineSymbol(text, textord, wideChar, wideChar);
}
function setLineBreaks(expression, wrapMode, isDisplayMode) {
	const mtrs = [];
	let mrows = [];
	let block = [];
	let numTopLevelEquals = 0;
	let i = 0;
	while (i < expression.length) {
		while (expression[i] instanceof DocumentFragment) expression.splice(i, 1, ...expression[i].children);
		const node = expression[i];
		if (node.attributes && node.attributes.linebreak && node.attributes.linebreak === "newline") {
			if (block.length > 0) mrows.push(new MathNode("mrow", block));
			mrows.push(node);
			block = [];
			const mtd = new MathNode("mtd", mrows);
			mtd.style.textAlign = "left";
			mtrs.push(new MathNode("mtr", [mtd]));
			mrows = [];
			i += 1;
			continue;
		}
		block.push(node);
		if (node.type && node.type === "mo" && node.children.length === 1 && !(node.attributes.form && node.attributes.form === "prefix") && !Object.prototype.hasOwnProperty.call(node.attributes, "movablelimits")) {
			const ch = node.children[0].text;
			if (wrapMode === "=" && ch === "=") {
				numTopLevelEquals += 1;
				if (numTopLevelEquals > 1) {
					block.pop();
					const element = new MathNode("mrow", block);
					mrows.push(element);
					block = [node];
				}
			} else if (wrapMode === "tex") {
				const next = i < expression.length - 1 ? expression[i + 1] : null;
				let glueIsFreeOfNobreak = true;
				if (!(next && next.type === "mtext" && next.attributes.linebreak && next.attributes.linebreak === "nobreak")) for (let j = i + 1; j < expression.length; j++) {
					const nd = expression[j];
					if (nd.type && nd.type === "mspace" && !(nd.attributes.linebreak && nd.attributes.linebreak === "newline")) {
						block.push(nd);
						i += 1;
						if (nd.attributes && nd.attributes.linebreak && nd.attributes.linebreak === "nobreak") glueIsFreeOfNobreak = false;
					} else break;
				}
				if (glueIsFreeOfNobreak) {
					const element = new MathNode("mrow", block);
					mrows.push(element);
					block = [];
				}
			}
		}
		i += 1;
	}
	if (block.length > 0) {
		const element = new MathNode("mrow", block);
		mrows.push(element);
	}
	if (mtrs.length > 0) {
		const mtd = new MathNode("mtd", mrows);
		mtd.style.textAlign = "left";
		const mtr = new MathNode("mtr", [mtd]);
		mtrs.push(mtr);
		const mtable = new MathNode("mtable", mtrs);
		if (!isDisplayMode) {
			mtable.setAttribute("columnalign", "left");
			mtable.setAttribute("rowspacing", "0em");
		}
		return mtable;
	}
	return newDocumentFragment(mrows);
}
/**
* This file converts a parse tree into a corresponding MathML tree. The main
* entry point is the `buildMathML` function, which takes a parse tree from the
* parser.
*/
/**
* Takes a symbol and converts it into a MathML text node after performing
* optional replacement from symbols.js.
*/
const makeText = function(text, mode, style) {
	if (symbols[mode][text] && symbols[mode][text].replace && text.charCodeAt(0) !== 55349 && !(Object.prototype.hasOwnProperty.call(ligatures, text) && style && (style.fontFamily && style.fontFamily.slice(4, 6) === "tt" || style.font && style.font.slice(4, 6) === "tt"))) text = symbols[mode][text].replace;
	return new TextNode(text);
};
const copyChar = (newRow, child) => {
	if (newRow.children.length === 0 || newRow.children[newRow.children.length - 1].type !== "mtext") {
		const mtext = new MathNode("mtext", [new TextNode(child.children[0].text)]);
		newRow.children.push(mtext);
	} else newRow.children[newRow.children.length - 1].children[0].text += child.children[0].text;
};
const consolidateText = (mrow) => {
	if (mrow.type !== "mrow" && mrow.type !== "mstyle") return mrow;
	if (mrow.children.length === 0) return mrow;
	const newRow = new MathNode("mrow");
	for (let i = 0; i < mrow.children.length; i++) {
		const child = mrow.children[i];
		if (child.type === "mtext" && Object.keys(child.attributes).length === 0) copyChar(newRow, child);
		else if (child.type === "mrow") {
			let canConsolidate = true;
			for (let j = 0; j < child.children.length; j++) if (child.children[j].type !== "mtext" || Object.keys(child.attributes).length !== 0) {
				canConsolidate = false;
				break;
			}
			if (canConsolidate) for (let j = 0; j < child.children.length; j++) {
				const grandChild = child.children[j];
				copyChar(newRow, grandChild);
			}
			else newRow.children.push(child);
		} else newRow.children.push(child);
	}
	for (let i = 0; i < newRow.children.length; i++) if (newRow.children[i].type === "mtext") {
		const mtext = newRow.children[i];
		if (mtext.children[0].text.charAt(0) === " ") mtext.children[0].text = "\xA0" + mtext.children[0].text.slice(1);
		const L = mtext.children[0].text.length;
		if (L > 0 && mtext.children[0].text.charAt(L - 1) === " ") mtext.children[0].text = mtext.children[0].text.slice(0, -1) + "\xA0";
		for (const [key, value] of Object.entries(mrow.attributes)) mtext.attributes[key] = value;
	}
	if (newRow.children.length === 1 && newRow.children[0].type === "mtext") return newRow.children[0];
	else return newRow;
};
/**
* Wrap the given array of nodes in an <mrow> node if needed, i.e.,
* unless the array has length 1.  Always returns a single node.
*/
const makeRow = function(body, semisimple = false) {
	if (body.length === 1 && !(body[0] instanceof DocumentFragment)) return body[0];
	else if (!semisimple) {
		if (body[0] instanceof MathNode && body[0].type === "mo" && !body[0].attributes.fence) {
			body[0].attributes.lspace = "0em";
			body[0].attributes.rspace = "0em";
		}
		const end = body.length - 1;
		if (body[end] instanceof MathNode && body[end].type === "mo" && !body[end].attributes.fence) {
			body[end].attributes.lspace = "0em";
			body[end].attributes.rspace = "0em";
		}
	}
	return new MathNode("mrow", body);
};
/**
* Check for <mi>.</mi> which is how a dot renders in MathML,
* or <mo separator="true" lspace="0em" rspace="0em">,</mo>
* which is how a braced comma {,} renders in MathML
*/
function isNumberPunctuation(group) {
	if (!group) return false;
	if (group.type === "mi" && group.children.length === 1) {
		const child = group.children[0];
		return child instanceof TextNode && child.text === ".";
	} else if (group.type === "mtext" && group.children.length === 1) {
		const child = group.children[0];
		return child instanceof TextNode && child.text === " ";
	} else if (group.type === "mo" && group.children.length === 1 && group.getAttribute("separator") === "true" && group.getAttribute("lspace") === "0em" && group.getAttribute("rspace") === "0em") {
		const child = group.children[0];
		return child instanceof TextNode && child.text === ",";
	} else return false;
}
const isComma = (expression, i) => {
	const node = expression[i];
	const followingNode = expression[i + 1];
	return node.type === "atom" && node.text === "," && node.loc && followingNode.loc && node.loc.end === followingNode.loc.start;
};
const isRel = (item) => {
	return item.type === "atom" && item.family === "rel" || item.type === "mclass" && item.mclass === "mrel";
};
/**
* Takes a list of nodes, builds them, and returns a list of the generated
* MathML nodes.  Also do a couple chores along the way:
* (1) Suppress spacing when an author wraps an operator w/braces, as in {=}.
* (2) Suppress spacing between two adjacent relations.
*/
const buildExpression = function(expression, style, semisimple = false) {
	if (!semisimple && expression.length === 1) {
		const group = buildGroup$1(expression[0], style);
		if (group instanceof MathNode && group.type === "mo") {
			group.setAttribute("lspace", "0em");
			group.setAttribute("rspace", "0em");
		}
		return [group];
	}
	const groups = [];
	const groupArray = [];
	let lastGroup;
	for (let i = 0; i < expression.length; i++) groupArray.push(buildGroup$1(expression[i], style));
	for (let i = 0; i < groupArray.length; i++) {
		const group = groupArray[i];
		if (i < expression.length - 1 && isRel(expression[i]) && isRel(expression[i + 1])) group.setAttribute("rspace", "0em");
		if (i > 0 && isRel(expression[i]) && isRel(expression[i - 1])) group.setAttribute("lspace", "0em");
		if (group.type === "mn" && lastGroup && lastGroup.type === "mn") {
			lastGroup.children.push(...group.children);
			continue;
		} else if (isNumberPunctuation(group) && lastGroup && lastGroup.type === "mn") {
			lastGroup.children.push(...group.children);
			continue;
		} else if (lastGroup && lastGroup.type === "mn" && i < groupArray.length - 1 && groupArray[i + 1].type === "mn" && isComma(expression, i)) {
			lastGroup.children.push(...group.children);
			continue;
		} else if (group.type === "mn" && isNumberPunctuation(lastGroup)) {
			group.children = [...lastGroup.children, ...group.children];
			groups.pop();
		} else if ((group.type === "msup" || group.type === "msub") && group.children.length >= 1 && lastGroup && (lastGroup.type === "mn" || isNumberPunctuation(lastGroup))) {
			const base = group.children[0];
			if (base instanceof MathNode && base.type === "mn" && lastGroup) {
				base.children = [...lastGroup.children, ...base.children];
				groups.pop();
			}
		}
		groups.push(group);
		lastGroup = group;
	}
	return groups;
};
/**
* Equivalent to buildExpression, but wraps the elements in an <mrow>
* if there's more than one.  Returns a single node instead of an array.
*/
const buildExpressionRow = function(expression, style, semisimple = false) {
	return makeRow(buildExpression(expression, style, semisimple), semisimple);
};
/**
* Takes a group from the parser and calls the appropriate groupBuilders function
* on it to produce a MathML node.
*/
const buildGroup$1 = function(group, style) {
	if (!group) return new MathNode("mrow");
	if (_mathmlGroupBuilders[group.type]) return _mathmlGroupBuilders[group.type](group, style);
	else throw new ParseError("Got group of unknown type: '" + group.type + "'");
};
const glue$1 = (_) => {
	return new MathNode("mtd", [], [], {
		padding: "0",
		width: "50%"
	});
};
const labelContainers = [
	"mrow",
	"mtd",
	"mtable",
	"mtr"
];
const getLabel = (parent) => {
	for (const node of parent.children) if (node.type && labelContainers.includes(node.type)) {
		if (node.classes && node.classes[0] === "tml-label") return node.label;
		else {
			const label = getLabel(node);
			if (label) return label;
		}
	} else if (!node.type) {
		const label = getLabel(node);
		if (label) return label;
	}
};
const taggedExpression = (expression, tag, style, leqno) => {
	tag = buildExpressionRow(tag[0].body, style);
	tag = consolidateText(tag);
	tag.classes.push("tml-tag");
	const label = getLabel(expression);
	expression = new MathNode("mtd", [expression]);
	const rowArray = [
		glue$1(),
		expression,
		glue$1()
	];
	rowArray[leqno ? 0 : 2].children.push(tag);
	const mtr = new MathNode("mtr", rowArray, ["tml-tageqn"]);
	if (label) mtr.setAttribute("id", label);
	const table = new MathNode("mtable", [mtr]);
	table.style.width = "100%";
	table.setAttribute("displaystyle", "true");
	return table;
};
/**
* Takes a full parse tree and settings and builds a MathML representation of
* it.
*/
function buildMathML(tree, texExpression, style, settings) {
	let tag = null;
	if (tree.length === 1 && tree[0].type === "tag") {
		tag = tree[0].tag;
		tree = tree[0].body;
	}
	const expression = buildExpression(tree, style);
	if (expression.length === 1 && expression[0] instanceof AnchorNode) return expression[0];
	const wrap = settings.displayMode || settings.annotate ? "none" : settings.wrap;
	const n1 = expression.length === 0 ? null : expression[0];
	let wrapper = expression.length === 1 && tag === null && n1 instanceof MathNode ? expression[0] : setLineBreaks(expression, wrap, settings.displayMode);
	if (tag) wrapper = taggedExpression(wrapper, tag, style, settings.leqno);
	if (settings.annotate) {
		const annotation = new MathNode("annotation", [new TextNode(texExpression)]);
		annotation.setAttribute("encoding", "application/x-tex");
		wrapper = new MathNode("semantics", [wrapper, annotation]);
	}
	const math = new MathNode("math", [wrapper]);
	if (settings.xml) math.setAttribute("xmlns", "http://www.w3.org/1998/Math/MathML");
	if (settings.displayMode) {
		math.setAttribute("display", "block");
		math.style.display = "block math";
		math.classes = ["tml-display"];
	}
	return math;
}
const smallNudge = "DHKLUcegorsuvxyzΠΥΨαδηιμνοτυχϵ";
const mediumNudge = "BCEGIMNOPQRSTXZlpqtwΓΘΞΣΦΩβεζθξρςφψϑϕϱ";
const largeNudge = "AFJdfΔΛ";
const mathmlBuilder$a = (group, style) => {
	const accentNode$1 = group.isStretchy ? accentNode(group) : new MathNode("mo", [makeText(group.label, group.mode)]);
	if (!group.isStretchy) accentNode$1.setAttribute("stretchy", "false");
	if (group.label !== "\\vec") accentNode$1.style.mathDepth = "0";
	const tag = group.label === "\\c" ? "munder" : "mover";
	const needsWbkVertShift = needsWebkitVerticalShift.has(group.label);
	if (tag === "mover" && group.mode === "math" && !group.isStretchy && group.base.text && group.base.text.length === 1) {
		const text = group.base.text;
		const isVec = group.label === "\\vec";
		const vecPostfix = isVec === "\\vec" ? "-vec" : "";
		if (isVec) accentNode$1.classes.push("tml-vec");
		const wbkPostfix = isVec ? "-vec" : needsWbkVertShift ? "-acc" : "";
		if (smallNudge.indexOf(text) > -1) {
			accentNode$1.classes.push(`chr-sml${vecPostfix}`);
			accentNode$1.classes.push(`wbk-sml${wbkPostfix}`);
		} else if (mediumNudge.indexOf(text) > -1) {
			accentNode$1.classes.push(`chr-med${vecPostfix}`);
			accentNode$1.classes.push(`wbk-med${wbkPostfix}`);
		} else if (largeNudge.indexOf(text) > -1) {
			accentNode$1.classes.push(`chr-lrg${vecPostfix}`);
			accentNode$1.classes.push(`wbk-lrg${wbkPostfix}`);
		} else if (isVec) accentNode$1.classes.push(`wbk-vec`);
		else if (needsWbkVertShift) accentNode$1.classes.push(`wbk-acc`);
	} else if (needsWbkVertShift) accentNode$1.classes.push("wbk-acc");
	return new MathNode(tag, [buildGroup$1(group.base, style), accentNode$1]);
};
const nonStretchyAccents = /* @__PURE__ */ new Set([
	"\\acute",
	"\\check",
	"\\grave",
	"\\ddot",
	"\\dddot",
	"\\ddddot",
	"\\tilde",
	"\\bar",
	"\\breve",
	"\\check",
	"\\hat",
	"\\vec",
	"\\dot",
	"\\mathring"
]);
const needsWebkitVerticalShift = /* @__PURE__ */ new Set([
	"\\acute",
	"\\bar",
	"\\breve",
	"\\check",
	"\\dot",
	"\\ddot",
	"\\grave",
	"\\hat",
	"\\mathring",
	"\\`",
	"\\'",
	"\\^",
	"\\=",
	"\\u",
	"\\.",
	"\\\"",
	"\\r",
	"\\H",
	"\\v"
]);
const combiningChar = {
	"\\`": "̀",
	"\\'": "́",
	"\\^": "̂",
	"\\~": "̃",
	"\\=": "̄",
	"\\u": "̆",
	"\\.": "̇",
	"\\\"": "̈",
	"\\r": "̊",
	"\\H": "̋",
	"\\v": "̌",
	"\\c": "̧"
};
defineFunction({
	type: "accent",
	names: [
		"\\acute",
		"\\grave",
		"\\ddot",
		"\\dddot",
		"\\ddddot",
		"\\tilde",
		"\\bar",
		"\\breve",
		"\\check",
		"\\hat",
		"\\vec",
		"\\dot",
		"\\mathring",
		"\\overparen",
		"\\widecheck",
		"\\widehat",
		"\\wideparen",
		"\\widetilde",
		"\\overrightarrow",
		"\\overleftarrow",
		"\\Overrightarrow",
		"\\overleftrightarrow",
		"\\overgroup",
		"\\overleftharpoon",
		"\\overrightharpoon"
	],
	props: { numArgs: 1 },
	handler: (context, args) => {
		const base = normalizeArgument(args[0]);
		const isStretchy = !nonStretchyAccents.has(context.funcName);
		return {
			type: "accent",
			mode: context.parser.mode,
			label: context.funcName,
			isStretchy,
			base
		};
	},
	mathmlBuilder: mathmlBuilder$a
});
defineFunction({
	type: "accent",
	names: [
		"\\'",
		"\\`",
		"\\^",
		"\\~",
		"\\=",
		"\\c",
		"\\u",
		"\\.",
		"\\\"",
		"\\r",
		"\\H",
		"\\v"
	],
	props: {
		numArgs: 1,
		allowedInText: true,
		allowedInMath: true,
		argTypes: ["primitive"]
	},
	handler: (context, args) => {
		const base = normalizeArgument(args[0]);
		const mode = context.parser.mode;
		if (mode === "math" && context.parser.settings.strict) console.log(`Temml parse error: Command ${context.funcName} is invalid in math mode.`);
		if (mode === "text" && base.text && base.text.length === 1 && context.funcName in combiningChar && smalls.indexOf(base.text) > -1) return {
			type: "textord",
			mode: "text",
			text: base.text + combiningChar[context.funcName]
		};
		else if (context.funcName === "\\c" && mode === "text" && base.text && base.text.length === 1) return {
			type: "textord",
			mode: "text",
			text: base.text + "̧"
		};
		else return {
			type: "accent",
			mode,
			label: context.funcName,
			isStretchy: false,
			base
		};
	},
	mathmlBuilder: mathmlBuilder$a
});
defineFunction({
	type: "accentUnder",
	names: [
		"\\underleftarrow",
		"\\underrightarrow",
		"\\underleftrightarrow",
		"\\undergroup",
		"\\underparen",
		"\\utilde"
	],
	props: { numArgs: 1 },
	handler: ({ parser, funcName }, args) => {
		const base = args[0];
		return {
			type: "accentUnder",
			mode: parser.mode,
			label: funcName,
			base
		};
	},
	mathmlBuilder: (group, style) => {
		const accentNode$1 = accentNode(group);
		accentNode$1.style["math-depth"] = 0;
		return new MathNode("munder", [buildGroup$1(group.base, style), accentNode$1]);
	}
});
/**
* This file does conversion between units.  In particular, it provides
* calculateSize to convert other units into CSS units.
*/
const ptPerUnit = {
	pt: 800 / 803,
	pc: 9600 / 803,
	dd: 1238 / 1157 * 800 / 803,
	cc: 14856 / 1157 * 800 / 803,
	nd: 685 / 642 * 800 / 803,
	nc: 1370 / 107 * 800 / 803,
	sp: 1 / 65536 * 800 / 803,
	mm: 25.4 / 72,
	cm: 2.54 / 72,
	in: 1 / 72,
	px: 96 / 72
};
/**
* Determine whether the specified unit (either a string defining the unit
* or a "size" parse node containing a unit field) is valid.
*/
const validUnits = [
	"em",
	"ex",
	"mu",
	"pt",
	"mm",
	"cm",
	"in",
	"px",
	"bp",
	"pc",
	"dd",
	"cc",
	"nd",
	"nc",
	"sp"
];
const validUnit = function(unit) {
	if (typeof unit !== "string") unit = unit.unit;
	return validUnits.indexOf(unit) > -1;
};
const emScale = (styleLevel) => {
	return [
		1,
		.7,
		.5
	][Math.max(styleLevel - 1, 0)];
};
const calculateSize = function(sizeValue, style) {
	let number = sizeValue.number;
	if (style.maxSize[0] < 0 && number > 0) return {
		number: 0,
		unit: "em"
	};
	const unit = sizeValue.unit;
	switch (unit) {
		case "mm":
		case "cm":
		case "in":
		case "px":
			if (number * ptPerUnit[unit] > style.maxSize[1]) return {
				number: style.maxSize[1],
				unit: "pt"
			};
			return {
				number,
				unit
			};
		case "em":
		case "ex":
			if (unit === "ex") number *= .431;
			number = Math.min(number / emScale(style.level), style.maxSize[0]);
			return {
				number: round(number),
				unit: "em"
			};
		case "bp":
			if (number > style.maxSize[1]) number = style.maxSize[1];
			return {
				number,
				unit: "pt"
			};
		case "pt":
		case "pc":
		case "dd":
		case "cc":
		case "nd":
		case "nc":
		case "sp":
			number = Math.min(number * ptPerUnit[unit], style.maxSize[1]);
			return {
				number: round(number),
				unit: "pt"
			};
		case "mu":
			number = Math.min(number / 18, style.maxSize[0]);
			return {
				number: round(number),
				unit: "em"
			};
		default: throw new ParseError("Invalid unit: '" + unit + "'");
	}
};
const padding = (width) => {
	const node = new MathNode("mspace");
	node.setAttribute("width", width + "em");
	return node;
};
const paddedNode = (group, lspace = .3, rspace = 0, mustSmash = false) => {
	if (group == null && rspace === 0) return padding(lspace);
	const row = group ? [group] : [];
	if (lspace !== 0) row.unshift(padding(lspace));
	if (rspace > 0) row.push(padding(rspace));
	if (mustSmash) {
		const mpadded = new MathNode("mpadded", row);
		mpadded.setAttribute("height", "0.1px");
		return mpadded;
	} else return new MathNode("mrow", row);
};
const labelSize = (size, scriptLevel) => Number(size) / emScale(scriptLevel);
const munderoverNode = (fName, body, below, style) => {
	const arrowNode = mathMLnode(fName);
	const isEq = fName.slice(1, 3) === "eq";
	const minWidth = fName.charAt(1) === "x" ? "1.75" : fName.slice(2, 4) === "cd" ? "3.0" : isEq ? "1.0" : "2.0";
	arrowNode.setAttribute("lspace", "0");
	arrowNode.setAttribute("rspace", isEq ? "0.5em" : "0");
	const labelStyle = style.withLevel(style.level < 2 ? 2 : 3);
	const minArrowWidth = labelSize(minWidth, labelStyle.level);
	const dummyWidth = labelSize(minWidth, 3);
	const emptyLabel = paddedNode(null, minArrowWidth.toFixed(4), 0);
	const dummyNode = paddedNode(null, dummyWidth.toFixed(4), 0);
	const space = labelSize(isEq ? 0 : .3, labelStyle.level).toFixed(4);
	let upperNode;
	let lowerNode;
	const gotUpper = body && body.body && (body.body.body || body.body.length > 0);
	if (gotUpper) {
		let label = buildGroup$1(body, labelStyle);
		label = paddedNode(label, space, space, fName === "\\\\cdrightarrow" || fName === "\\\\cdleftarrow");
		upperNode = new MathNode("mover", [label, dummyNode]);
	}
	const gotLower = below && below.body && (below.body.body || below.body.length > 0);
	if (gotLower) {
		let label = buildGroup$1(below, labelStyle);
		label = paddedNode(label, space, space);
		lowerNode = new MathNode("munder", [label, dummyNode]);
	}
	let node;
	if (!gotUpper && !gotLower) node = new MathNode("mover", [arrowNode, emptyLabel]);
	else if (gotUpper && gotLower) node = new MathNode("munderover", [
		arrowNode,
		lowerNode,
		upperNode
	]);
	else if (gotUpper) node = new MathNode("mover", [arrowNode, upperNode]);
	else node = new MathNode("munder", [arrowNode, lowerNode]);
	if (minWidth === "3.0") node.style.height = "1em";
	node.setAttribute("accent", "false");
	return node;
};
defineFunction({
	type: "xArrow",
	names: [
		"\\xleftarrow",
		"\\xrightarrow",
		"\\xLeftarrow",
		"\\xRightarrow",
		"\\xleftrightarrow",
		"\\xLeftrightarrow",
		"\\xhookleftarrow",
		"\\xhookrightarrow",
		"\\xmapsto",
		"\\xrightharpoondown",
		"\\xrightharpoonup",
		"\\xleftharpoondown",
		"\\xleftharpoonup",
		"\\xlongequal",
		"\\xtwoheadrightarrow",
		"\\xtwoheadleftarrow",
		"\\xtofrom",
		"\\xleftrightharpoons",
		"\\xrightleftharpoons",
		"\\yields",
		"\\yieldsLeft",
		"\\mesomerism",
		"\\longrightharpoonup",
		"\\longleftharpoondown",
		"\\yieldsLeftRight",
		"\\chemequilibrium",
		"\\\\cdrightarrow",
		"\\\\cdleftarrow",
		"\\\\cdlongequal"
	],
	props: {
		numArgs: 1,
		numOptionalArgs: 1
	},
	handler({ parser, funcName }, args, optArgs) {
		return {
			type: "xArrow",
			mode: parser.mode,
			name: funcName,
			body: args[0],
			below: optArgs[0]
		};
	},
	mathmlBuilder(group, style) {
		const row = [munderoverNode(group.name, group.body, group.below, style)];
		row.unshift(padding(.2778));
		row.push(padding(.2778));
		return new MathNode("mrow", row);
	}
});
const arrowComponent = {
	"\\equilibriumRight": ["\\longrightharpoonup", "\\eqleftharpoondown"],
	"\\equilibriumLeft": ["\\eqrightharpoonup", "\\longleftharpoondown"]
};
defineFunction({
	type: "stackedArrow",
	names: ["\\equilibriumRight", "\\equilibriumLeft"],
	props: {
		numArgs: 1,
		numOptionalArgs: 1
	},
	handler({ parser, funcName }, args, optArgs) {
		const lowerArrowBody = args[0] ? {
			type: "hphantom",
			mode: parser.mode,
			body: args[0]
		} : null;
		const upperArrowBelow = optArgs[0] ? {
			type: "hphantom",
			mode: parser.mode,
			body: optArgs[0]
		} : null;
		return {
			type: "stackedArrow",
			mode: parser.mode,
			name: funcName,
			body: args[0],
			upperArrowBelow,
			lowerArrowBody,
			below: optArgs[0]
		};
	},
	mathmlBuilder(group, style) {
		const topLabel = arrowComponent[group.name][0];
		const botLabel = arrowComponent[group.name][1];
		const topArrow = munderoverNode(topLabel, group.body, group.upperArrowBelow, style);
		const botArrow = munderoverNode(botLabel, group.lowerArrowBody, group.below, style);
		let wrapper;
		const raiseNode = new MathNode("mpadded", [topArrow]);
		raiseNode.setAttribute("voffset", "0.3em");
		raiseNode.setAttribute("height", "+0.3em");
		raiseNode.setAttribute("depth", "-0.3em");
		if (group.name === "\\equilibriumLeft") {
			const botNode = new MathNode("mpadded", [botArrow]);
			botNode.setAttribute("width", "0.5em");
			wrapper = new MathNode("mpadded", [
				padding(.2778),
				botNode,
				raiseNode,
				padding(.2778)
			]);
		} else {
			raiseNode.setAttribute("width", group.name === "\\equilibriumRight" ? "0.5em" : "0");
			wrapper = new MathNode("mpadded", [
				padding(.2778),
				raiseNode,
				botArrow,
				padding(.2778)
			]);
		}
		wrapper.setAttribute("voffset", "-0.18em");
		wrapper.setAttribute("height", "-0.18em");
		wrapper.setAttribute("depth", "+0.18em");
		return wrapper;
	}
});
/**
* All registered environments.
* `environments.js` exports this same dictionary again and makes it public.
* `Parser.js` requires this dictionary via `environments.js`.
*/
const _environments = {};
function defineEnvironment({ type, names, props, handler, mathmlBuilder }) {
	const data = {
		type,
		numArgs: props.numArgs || 0,
		allowedInText: false,
		numOptionalArgs: 0,
		handler
	};
	for (let i = 0; i < names.length; ++i) _environments[names[i]] = data;
	if (mathmlBuilder) _mathmlGroupBuilders[type] = mathmlBuilder;
}
/**
* Asserts that the node is of the given type and returns it with stricter
* typing. Throws if the node's type does not match.
*/
function assertNodeType(node, type) {
	if (!node || node.type !== type) throw new Error(`Expected node of type ${type}, but got ` + (node ? `node of type ${node.type}` : String(node)));
	return node;
}
/**
* Returns the node more strictly typed iff it is of the given type. Otherwise,
* returns null.
*/
function assertSymbolNodeType(node) {
	const typedNode = checkSymbolNodeType(node);
	if (!typedNode) throw new Error(`Expected node of symbol group type, but got ` + (node ? `node of type ${node.type}` : String(node)));
	return typedNode;
}
/**
* Returns the node more strictly typed iff it is of the given type. Otherwise,
* returns null.
*/
function checkSymbolNodeType(node) {
	if (node && (node.type === "atom" || node.type === "delimiter" || Object.prototype.hasOwnProperty.call(NON_ATOMS, node.type))) return node;
	return null;
}
const cdArrowFunctionName = {
	">": "\\\\cdrightarrow",
	"<": "\\\\cdleftarrow",
	"=": "\\\\cdlongequal",
	A: "\\uparrow",
	V: "\\downarrow",
	"|": "\\Vert",
	".": "no arrow"
};
const newCell = () => {
	return {
		type: "styling",
		body: [],
		mode: "math",
		scriptLevel: "display"
	};
};
const isStartOfArrow = (node) => {
	return node.type === "textord" && node.text === "@";
};
const isLabelEnd = (node, endChar) => {
	return (node.type === "mathord" || node.type === "atom") && node.text === endChar;
};
function cdArrow(arrowChar, labels, parser) {
	const funcName = cdArrowFunctionName[arrowChar];
	switch (funcName) {
		case "\\\\cdrightarrow":
		case "\\\\cdleftarrow": return parser.callFunction(funcName, [labels[0]], [labels[1]]);
		case "\\uparrow":
		case "\\downarrow": {
			const leftLabel = parser.callFunction("\\\\cdleft", [labels[0]], []);
			const bareArrow = {
				type: "atom",
				text: funcName,
				mode: "math",
				family: "rel"
			};
			const arrowGroup = {
				type: "ordgroup",
				mode: "math",
				body: [
					leftLabel,
					parser.callFunction("\\Big", [bareArrow], []),
					parser.callFunction("\\\\cdright", [labels[1]], [])
				],
				semisimple: true
			};
			return parser.callFunction("\\\\cdparent", [arrowGroup], []);
		}
		case "\\\\cdlongequal": return parser.callFunction("\\\\cdlongequal", [], []);
		case "\\Vert": return parser.callFunction("\\Big", [{
			type: "textord",
			text: "\\Vert",
			mode: "math"
		}], []);
		default: return {
			type: "textord",
			text: " ",
			mode: "math"
		};
	}
}
function parseCD(parser) {
	const parsedRows = [];
	parser.gullet.beginGroup();
	parser.gullet.macros.set("\\cr", "\\\\\\relax");
	parser.gullet.beginGroup();
	while (true) {
		parsedRows.push(parser.parseExpression(false, "\\\\"));
		parser.gullet.endGroup();
		parser.gullet.beginGroup();
		const next = parser.fetch().text;
		if (next === "&" || next === "\\\\") parser.consume();
		else if (next === "\\end") {
			if (parsedRows[parsedRows.length - 1].length === 0) parsedRows.pop();
			break;
		} else throw new ParseError("Expected \\\\ or \\cr or \\end", parser.nextToken);
	}
	let row = [];
	const body = [row];
	for (let i = 0; i < parsedRows.length; i++) {
		const rowNodes = parsedRows[i];
		let cell = newCell();
		for (let j = 0; j < rowNodes.length; j++) if (!isStartOfArrow(rowNodes[j])) cell.body.push(rowNodes[j]);
		else {
			row.push(cell);
			j += 1;
			const arrowChar = assertSymbolNodeType(rowNodes[j]).text;
			const labels = new Array(2);
			labels[0] = {
				type: "ordgroup",
				mode: "math",
				body: []
			};
			labels[1] = {
				type: "ordgroup",
				mode: "math",
				body: []
			};
			if ("=|.".indexOf(arrowChar) > -1);
			else if ("<>AV".indexOf(arrowChar) > -1) for (let labelNum = 0; labelNum < 2; labelNum++) {
				let inLabel = true;
				for (let k = j + 1; k < rowNodes.length; k++) {
					if (isLabelEnd(rowNodes[k], arrowChar)) {
						inLabel = false;
						j = k;
						break;
					}
					if (isStartOfArrow(rowNodes[k])) throw new ParseError("Missing a " + arrowChar + " character to complete a CD arrow.", rowNodes[k]);
					labels[labelNum].body.push(rowNodes[k]);
				}
				if (inLabel) throw new ParseError("Missing a " + arrowChar + " character to complete a CD arrow.", rowNodes[j]);
			}
			else throw new ParseError(`Expected one of "<>AV=|." after @.`);
			const arrow = cdArrow(arrowChar, labels, parser);
			row.push(arrow);
			cell = newCell();
		}
		if (i % 2 === 0) row.push(cell);
		else row.shift();
		row = [];
		body.push(row);
	}
	body.pop();
	parser.gullet.endGroup();
	parser.gullet.endGroup();
	return {
		type: "array",
		mode: "math",
		body,
		tags: null,
		labels: new Array(body.length + 1).fill(""),
		envClasses: ["jot", "cd"],
		cols: [],
		hLinesBeforeRow: new Array(body.length + 1).fill([])
	};
}
defineFunction({
	type: "cdlabel",
	names: ["\\\\cdleft", "\\\\cdright"],
	props: { numArgs: 1 },
	handler({ parser, funcName }, args) {
		return {
			type: "cdlabel",
			mode: parser.mode,
			side: funcName.slice(4),
			label: args[0]
		};
	},
	mathmlBuilder(group, style) {
		if (group.label.body.length === 0) return new MathNode("mrow", style);
		const mrow = buildGroup$1(group.label, style);
		if (group.side === "left") mrow.classes.push("tml-shift-left");
		const mtd = new MathNode("mtd", [mrow]);
		mtd.style.padding = "0";
		const label = new MathNode("mpadded", [new MathNode("mtable", [new MathNode("mtr", [mtd])])]);
		label.setAttribute("width", "0.1px");
		label.setAttribute("displaystyle", "false");
		label.setAttribute("scriptlevel", "1");
		return label;
	}
});
defineFunction({
	type: "cdlabelparent",
	names: ["\\\\cdparent"],
	props: { numArgs: 1 },
	handler({ parser }, args) {
		return {
			type: "cdlabelparent",
			mode: parser.mode,
			fragment: args[0]
		};
	},
	mathmlBuilder(group, style) {
		return new MathNode("mrow", [buildGroup$1(group.fragment, style)]);
	}
});
const ordGroup = (body) => {
	return {
		"type": "ordgroup",
		"mode": "math",
		"body": body,
		"semisimple": true
	};
};
const phantom = (body, type) => {
	return {
		"type": type,
		"mode": "math",
		"body": ordGroup(body)
	};
};
const bordermatrixParseTree = (matrix, delimiters) => {
	const body = matrix.body;
	body[0].shift();
	const leftColumnBody = new Array(body.length - 1).fill().map(() => []);
	for (let i = 1; i < body.length; i++) {
		leftColumnBody[i - 1].push(body[i].shift());
		const phantomBody = [];
		for (let j = 0; j < body[i].length; j++) phantomBody.push(body[i][j]);
		leftColumnBody[i - 1].push(phantom(phantomBody, "vphantom"));
	}
	const topRowBody = new Array(body.length).fill().map(() => []);
	for (let j = 0; j < body[0].length; j++) topRowBody[0].push(body[0][j]);
	for (let i = 1; i < body.length; i++) for (let j = 0; j < body[0].length; j++) topRowBody[i].push(phantom(body[i][j].body, "hphantom"));
	for (let j = 0; j < body[0].length; j++) body[0][j] = phantom(body[0][j].body, "hphantom");
	const leftColumn = {
		type: "array",
		mode: "math",
		body: leftColumnBody,
		cols: [{
			type: "align",
			align: "c"
		}],
		rowGaps: new Array(leftColumnBody.length - 1).fill(null),
		hLinesBeforeRow: new Array(leftColumnBody.length + 1).fill().map(() => []),
		envClasses: [],
		scriptLevel: "text",
		arraystretch: 1,
		labels: new Array(leftColumnBody.length).fill(""),
		arraycolsep: {
			"number": .04,
			unit: "em"
		}
	};
	const topWrapper = {
		type: "styling",
		mode: "math",
		scriptLevel: "text",
		body: [{
			type: "array",
			mode: "math",
			body: topRowBody,
			cols: new Array(topRowBody.length).fill({
				type: "align",
				align: "c"
			}),
			rowGaps: new Array(topRowBody.length - 1).fill(null),
			hLinesBeforeRow: new Array(topRowBody.length + 1).fill().map(() => []),
			envClasses: [],
			scriptLevel: "text",
			arraystretch: 1,
			labels: new Array(topRowBody.length).fill(""),
			arraycolsep: null
		}]
	};
	const mover = {
		type: "supsub",
		mode: "math",
		stack: true,
		base: {
			type: "op",
			mode: "math",
			limits: true,
			alwaysHandleSupSub: true,
			parentIsSupSub: true,
			symbol: false,
			suppressBaseShift: true,
			body: [{
				type: "leftright",
				mode: "math",
				body: [matrix],
				left: delimiters ? delimiters[0] : "(",
				right: delimiters ? delimiters[1] : ")",
				rightColor: void 0
			}]
		},
		sup: topWrapper,
		sub: null
	};
	return ordGroup([leftColumn, mover]);
};
/**
* Lexing or parsing positional information for error reporting.
* This object is immutable.
*/
var SourceLocation = class SourceLocation {
	constructor(lexer, start, end) {
		this.lexer = lexer;
		this.start = start;
		this.end = end;
	}
	/**
	* Merges two `SourceLocation`s from location providers, given they are
	* provided in order of appearance.
	* - Returns the first one's location if only the first is provided.
	* - Returns a merged range of the first and the last if both are provided
	*   and their lexers match.
	* - Otherwise, returns null.
	*/
	static range(first, second) {
		if (!second) return first && first.loc;
		else if (!first || !first.loc || !second.loc || first.loc.lexer !== second.loc.lexer) return null;
		else return new SourceLocation(first.loc.lexer, first.loc.start, second.loc.end);
	}
};
/**
* Interface required to break circular dependency between Token, Lexer, and
* ParseError.
*/
/**
* The resulting token returned from `lex`.
*
* It consists of the token text plus some position information.
* The position information is essentially a range in an input string,
* but instead of referencing the bare input string, we refer to the lexer.
* That way it is possible to attach extra metadata to the input string,
* like for example a file name or similar.
*
* The position information is optional, so it is OK to construct synthetic
* tokens if appropriate. Not providing available position information may
* lead to degraded error reporting, though.
*/
var Token = class Token {
	constructor(text, loc) {
		this.text = text;
		this.loc = loc;
	}
	/**
	* Given a pair of tokens (this and endToken), compute a `Token` encompassing
	* the whole input range enclosed by these two.
	*/
	range(endToken, text) {
		return new Token(text, SourceLocation.range(this, endToken));
	}
};
const StyleLevel = {
	DISPLAY: 0,
	TEXT: 1,
	SCRIPT: 2,
	SCRIPTSCRIPT: 3
};
/**
* All registered global/built-in macros.
* `macros.js` exports this same dictionary again and makes it public.
* `Parser.js` requires this dictionary via `macros.js`.
*/
const _macros = Object.create(null);
function defineMacro(name, body) {
	_macros[name] = body;
}
/**
* Predefined macros for Temml.
* This can be used to define some commands in terms of others.
*/
const macros = _macros;
defineMacro("\\noexpand", function(context) {
	const t = context.popToken();
	if (context.isExpandable(t.text)) {
		t.noexpand = true;
		t.treatAsRelax = true;
	}
	return {
		tokens: [t],
		numArgs: 0
	};
});
defineMacro("\\expandafter", function(context) {
	const t = context.popToken();
	context.expandOnce(true);
	return {
		tokens: [t],
		numArgs: 0
	};
});
defineMacro("\\@firstoftwo", function(context) {
	return {
		tokens: context.consumeArgs(2)[0],
		numArgs: 0
	};
});
defineMacro("\\@secondoftwo", function(context) {
	return {
		tokens: context.consumeArgs(2)[1],
		numArgs: 0
	};
});
defineMacro("\\@ifnextchar", function(context) {
	const args = context.consumeArgs(3);
	context.consumeSpaces();
	const nextToken = context.future();
	if (args[0].length === 1 && args[0][0].text === nextToken.text) return {
		tokens: args[1],
		numArgs: 0
	};
	else return {
		tokens: args[2],
		numArgs: 0
	};
});
defineMacro("\\@ifstar", "\\@ifnextchar *{\\@firstoftwo{#1}}");
defineMacro("\\TextOrMath", function(context) {
	const args = context.consumeArgs(2);
	if (context.mode === "text") return {
		tokens: args[0],
		numArgs: 0
	};
	else return {
		tokens: args[1],
		numArgs: 0
	};
});
const stringFromArg = (arg) => {
	let str = "";
	for (let i = arg.length - 1; i > -1; i--) str += arg[i].text;
	return str;
};
const digitToNumber = {
	0: 0,
	1: 1,
	2: 2,
	3: 3,
	4: 4,
	5: 5,
	6: 6,
	7: 7,
	8: 8,
	9: 9,
	a: 10,
	A: 10,
	b: 11,
	B: 11,
	c: 12,
	C: 12,
	d: 13,
	D: 13,
	e: 14,
	E: 14,
	f: 15,
	F: 15
};
const nextCharNumber = (context) => {
	const numStr = context.future().text;
	if (numStr === "EOF") return [null, ""];
	return [digitToNumber[numStr.charAt(0)], numStr];
};
const appendCharNumbers = (number, numStr, base) => {
	for (let i = 1; i < numStr.length; i++) {
		const digit = digitToNumber[numStr.charAt(i)];
		number *= base;
		number += digit;
	}
	return number;
};
defineMacro("\\char", function(context) {
	let token = context.popToken();
	let base;
	let number = "";
	if (token.text === "'") {
		base = 8;
		token = context.popToken();
	} else if (token.text === "\"") {
		base = 16;
		token = context.popToken();
	} else if (token.text === "`") {
		token = context.popToken();
		if (token.text[0] === "\\") number = token.text.charCodeAt(1);
		else if (token.text === "EOF") throw new ParseError("\\char` missing argument");
		else number = token.text.charCodeAt(0);
	} else base = 10;
	if (base) {
		let numStr = token.text;
		number = digitToNumber[numStr.charAt(0)];
		if (number == null || number >= base) throw new ParseError(`Invalid base-${base} digit ${token.text}`);
		number = appendCharNumbers(number, numStr, base);
		let digit;
		[digit, numStr] = nextCharNumber(context);
		while (digit != null && digit < base) {
			number *= base;
			number += digit;
			number = appendCharNumbers(number, numStr, base);
			context.popToken();
			[digit, numStr] = nextCharNumber(context);
		}
	}
	return `\\@char{${number}}`;
});
function recreateArgStr(context) {
	const tokens = context.consumeArgs(1)[0];
	let str = "";
	let expectedLoc = tokens[tokens.length - 1].loc.start;
	for (let i = tokens.length - 1; i >= 0; i--) {
		const actualLoc = tokens[i].loc.start;
		if (actualLoc > expectedLoc) {
			str += " ";
			expectedLoc = actualLoc;
		}
		str += tokens[i].text;
		expectedLoc += tokens[i].text.length;
	}
	return str;
}
defineMacro("\\surd", "\\sqrt{\\vphantom{|}}");
defineMacro("⊕", "\\oplus");
defineMacro("\\long", "");
defineMacro("\\bgroup", "{");
defineMacro("\\egroup", "}");
defineMacro("~", "\\nobreakspace");
defineMacro("\\lq", "`");
defineMacro("\\rq", "'");
defineMacro("\\aa", "\\r a");
defineMacro("\\Bbbk", "\\Bbb{k}");
defineMacro("\\mathstrut", "\\vphantom{(}");
defineMacro("\\underbar", "\\underline{\\text{#1}}");
defineMacro("\\vdots", "{\\varvdots\\rule{0pt}{15pt}}");
defineMacro("⋮", "\\vdots");
defineMacro("\\arraystretch", "1");
defineMacro("\\arraycolsep", "6pt");
defineMacro("\\substack", "\\begin{subarray}{c}#1\\end{subarray}");
defineMacro("\\iff", "\\DOTSB\\;\\Longleftrightarrow\\;");
defineMacro("\\implies", "\\DOTSB\\;\\Longrightarrow\\;");
defineMacro("\\impliedby", "\\DOTSB\\;\\Longleftarrow\\;");
const dotsByToken = {
	",": "\\dotsc",
	"\\not": "\\dotsb",
	"+": "\\dotsb",
	"=": "\\dotsb",
	"<": "\\dotsb",
	">": "\\dotsb",
	"-": "\\dotsb",
	"*": "\\dotsb",
	":": "\\dotsb",
	"\\DOTSB": "\\dotsb",
	"\\coprod": "\\dotsb",
	"\\bigvee": "\\dotsb",
	"\\bigwedge": "\\dotsb",
	"\\biguplus": "\\dotsb",
	"\\bigcap": "\\dotsb",
	"\\bigcup": "\\dotsb",
	"\\prod": "\\dotsb",
	"\\sum": "\\dotsb",
	"\\bigotimes": "\\dotsb",
	"\\bigoplus": "\\dotsb",
	"\\bigodot": "\\dotsb",
	"\\bigsqcap": "\\dotsb",
	"\\bigsqcup": "\\dotsb",
	"\\bigtimes": "\\dotsb",
	"\\And": "\\dotsb",
	"\\longrightarrow": "\\dotsb",
	"\\Longrightarrow": "\\dotsb",
	"\\longleftarrow": "\\dotsb",
	"\\Longleftarrow": "\\dotsb",
	"\\longleftrightarrow": "\\dotsb",
	"\\Longleftrightarrow": "\\dotsb",
	"\\mapsto": "\\dotsb",
	"\\longmapsto": "\\dotsb",
	"\\hookrightarrow": "\\dotsb",
	"\\doteq": "\\dotsb",
	"\\mathbin": "\\dotsb",
	"\\mathrel": "\\dotsb",
	"\\relbar": "\\dotsb",
	"\\Relbar": "\\dotsb",
	"\\xrightarrow": "\\dotsb",
	"\\xleftarrow": "\\dotsb",
	"\\DOTSI": "\\dotsi",
	"\\int": "\\dotsi",
	"\\oint": "\\dotsi",
	"\\iint": "\\dotsi",
	"\\iiint": "\\dotsi",
	"\\iiiint": "\\dotsi",
	"\\DOTSX": "\\dotsx"
};
defineMacro("\\dots", function(context) {
	let thedots = "\\dotso";
	const next = context.expandAfterFuture().text;
	if (next in dotsByToken) thedots = dotsByToken[next];
	else if (next.slice(0, 4) === "\\not") thedots = "\\dotsb";
	else if (next in symbols.math) {
		if (["bin", "rel"].includes(symbols.math[next].group)) thedots = "\\dotsb";
	}
	return thedots;
});
const spaceAfterDots = {
	")": true,
	"]": true,
	"\\rbrack": true,
	"\\}": true,
	"\\rbrace": true,
	"\\rangle": true,
	"\\rceil": true,
	"\\rfloor": true,
	"\\rgroup": true,
	"\\rmoustache": true,
	"\\right": true,
	"\\bigr": true,
	"\\biggr": true,
	"\\Bigr": true,
	"\\Biggr": true,
	$: true,
	";": true,
	".": true,
	",": true
};
defineMacro("\\dotso", function(context) {
	if (context.future().text in spaceAfterDots) return "\\ldots\\,";
	else return "\\ldots";
});
defineMacro("\\dotsc", function(context) {
	const next = context.future().text;
	if (next in spaceAfterDots && next !== ",") return "\\ldots\\,";
	else return "\\ldots";
});
defineMacro("\\cdots", function(context) {
	if (context.future().text in spaceAfterDots) return "\\@cdots\\,";
	else return "\\@cdots";
});
defineMacro("\\dotsb", "\\cdots");
defineMacro("\\dotsm", "\\cdots");
defineMacro("\\dotsi", "\\!\\cdots");
defineMacro("\\idotsint", "\\int\\!\\cdots\\!\\int");
defineMacro("\\dotsx", "\\ldots\\,");
defineMacro("\\DOTSI", "\\relax");
defineMacro("\\DOTSB", "\\relax");
defineMacro("\\DOTSX", "\\relax");
defineMacro("\\tmspace", "\\TextOrMath{\\kern#1#3}{\\mskip#1#2}\\relax");
defineMacro("\\,", "{\\tmspace+{3mu}{.1667em}}");
defineMacro("\\thinspace", "\\,");
defineMacro("\\>", "\\mskip{4mu}");
defineMacro("\\:", "{\\tmspace+{4mu}{.2222em}}");
defineMacro("\\medspace", "\\:");
defineMacro("\\;", "{\\tmspace+{5mu}{.2777em}}");
defineMacro("\\thickspace", "\\;");
defineMacro("\\!", "{\\tmspace-{3mu}{.1667em}}");
defineMacro("\\negthinspace", "\\!");
defineMacro("\\negmedspace", "{\\tmspace-{4mu}{.2222em}}");
defineMacro("\\negthickspace", "{\\tmspace-{5mu}{.277em}}");
defineMacro("\\enspace", "\\kern.5em ");
defineMacro("\\enskip", "\\hskip.5em\\relax");
defineMacro("\\quad", "\\hskip1em\\relax");
defineMacro("\\qquad", "\\hskip2em\\relax");
defineMacro("\\AA", "\\TextOrMath{\\Angstrom}{\\mathring{A}}\\relax");
defineMacro("\\tag", "\\@ifstar\\tag@literal\\tag@paren");
defineMacro("\\tag@paren", "\\tag@literal{({#1})}");
defineMacro("\\tag@literal", (context) => {
	if (context.macros.get("\\df@tag")) throw new ParseError("Multiple \\tag");
	return "\\gdef\\df@tag{\\text{#1}}";
});
defineMacro("\\notag", "\\nonumber");
defineMacro("\\nonumber", "\\gdef\\@eqnsw{0}");
defineMacro("\\bmod", "\\mathbin{\\text{mod}}");
defineMacro("\\pod", "\\allowbreak\\mathchoice{\\mkern18mu}{\\mkern8mu}{\\mkern8mu}{\\mkern8mu}(#1)");
defineMacro("\\pmod", "\\pod{{\\rm mod}\\mkern6mu#1}");
defineMacro("\\mod", "\\allowbreak\\mathchoice{\\mkern18mu}{\\mkern12mu}{\\mkern12mu}{\\mkern12mu}{\\rm mod}\\,\\,#1");
defineMacro("\\newline", "\\\\\\relax");
defineMacro("\\TeX", "\\textrm{T}\\kern-.1667em\\raisebox{-.5ex}{E}\\kern-.125em\\textrm{X}");
defineMacro("\\LaTeX", "\\textrm{L}\\kern-.35em\\raisebox{0.2em}{\\scriptstyle A}\\kern-.15em\\TeX");
defineMacro("\\Temml", "\\textrm{T}\\kern-0.2em\\lower{0.2em}{\\textrm{E}}\\kern-0.08em{\\textrm{M}\\kern-0.08em\\raise{0.2em}\\textrm{M}\\kern-0.08em\\textrm{L}}");
defineMacro("\\hspace", "\\@ifstar\\@hspacer\\@hspace");
defineMacro("\\@hspace", "\\hskip #1\\relax");
defineMacro("\\@hspacer", "\\rule{0pt}{0pt}\\hskip #1\\relax");
defineMacro("\\colon", `\\mathpunct{\\char"3a}`);
defineMacro("\\prescript", "\\pres@cript{_{#1}^{#2}}{}{#3}");
defineMacro("\\ordinarycolon", `\\char"3a`);
defineMacro("\\vcentcolon", "\\mathrel{\\raisebox{0.035em}{\\ordinarycolon}}");
defineMacro("\\coloneq", "\\mathrel{\\raisebox{0.035em}{\\ordinarycolon}\\char\"2212}");
defineMacro("\\Coloneq", "\\mathrel{\\char\"2237\\char\"2212}");
defineMacro("\\Eqqcolon", "\\mathrel{\\char\"3d\\char\"2237}");
defineMacro("\\Eqcolon", "\\mathrel{\\char\"2212\\char\"2237}");
defineMacro("\\colonapprox", "\\mathrel{\\raisebox{0.035em}{\\ordinarycolon}\\char\"2248}");
defineMacro("\\Colonapprox", "\\mathrel{\\char\"2237\\char\"2248}");
defineMacro("\\colonsim", "\\mathrel{\\raisebox{0.035em}{\\ordinarycolon}\\char\"223c}");
defineMacro("\\Colonsim", "\\mathrel{\\raisebox{0.035em}{\\ordinarycolon}\\char\"223c}");
defineMacro("\\ratio", "\\vcentcolon");
defineMacro("\\coloncolon", "\\dblcolon");
defineMacro("\\colonequals", "\\coloneqq");
defineMacro("\\coloncolonequals", "\\Coloneqq");
defineMacro("\\equalscolon", "\\eqqcolon");
defineMacro("\\equalscoloncolon", "\\Eqqcolon");
defineMacro("\\colonminus", "\\coloneq");
defineMacro("\\coloncolonminus", "\\Coloneq");
defineMacro("\\minuscolon", "\\eqcolon");
defineMacro("\\minuscoloncolon", "\\Eqcolon");
defineMacro("\\coloncolonapprox", "\\Colonapprox");
defineMacro("\\coloncolonsim", "\\Colonsim");
defineMacro("\\notni", "\\mathrel{\\char`∌}");
defineMacro("\\limsup", "\\DOTSB\\operatorname*{lim\\,sup}");
defineMacro("\\liminf", "\\DOTSB\\operatorname*{lim\\,inf}");
defineMacro("\\injlim", "\\DOTSB\\operatorname*{inj\\,lim}");
defineMacro("\\projlim", "\\DOTSB\\operatorname*{proj\\,lim}");
defineMacro("\\varlimsup", "\\DOTSB\\operatorname*{\\overline{\\text{lim}}}");
defineMacro("\\varliminf", "\\DOTSB\\operatorname*{\\underline{\\text{lim}}}");
defineMacro("\\varinjlim", "\\DOTSB\\operatorname*{\\underrightarrow{\\text{lim}}}");
defineMacro("\\varprojlim", "\\DOTSB\\operatorname*{\\underleftarrow{\\text{lim}}}");
defineMacro("\\centerdot", "{\\medspace\\rule{0.167em}{0.189em}\\medspace}");
defineMacro("\\argmin", "\\DOTSB\\operatorname*{arg\\,min}");
defineMacro("\\argmax", "\\DOTSB\\operatorname*{arg\\,max}");
defineMacro("\\plim", "\\DOTSB\\operatorname*{plim}");
defineMacro("\\leftmodels", "\\mathop{\\reflectbox{$\\models$}}");
defineMacro("\\bra", "\\mathinner{\\langle{#1}|}");
defineMacro("\\ket", "\\mathinner{|{#1}\\rangle}");
defineMacro("\\braket", "\\mathinner{\\langle{#1}\\rangle}");
defineMacro("\\Bra", "\\left\\langle#1\\right|");
defineMacro("\\Ket", "\\left|#1\\right\\rangle");
const replaceVert = (argStr, match) => {
	const replaceStr = `}\\,\\middle${match[0] === "|" ? "\\vert" : "\\Vert"}\\,{`;
	return argStr.slice(0, match.index) + replaceStr + argStr.slice(match.index + match[0].length);
};
defineMacro("\\Braket", function(context) {
	let argStr = recreateArgStr(context);
	const regEx = /\|\||\||\\\|/g;
	let match;
	while ((match = regEx.exec(argStr)) !== null) argStr = replaceVert(argStr, match);
	return "\\left\\langle{" + argStr + "}\\right\\rangle";
});
defineMacro("\\Set", function(context) {
	let argStr = recreateArgStr(context);
	const match = /\|\||\||\\\|/.exec(argStr);
	if (match) argStr = replaceVert(argStr, match);
	return "\\left\\{\\:{" + argStr + "}\\:\\right\\}";
});
defineMacro("\\set", function(context) {
	return "\\{{" + recreateArgStr(context).replace(/\|/, "}\\mid{") + "}\\}";
});
defineMacro("\\angln", "{\\angl n}");
defineMacro("\\odv", "\\@ifstar\\odv@next\\odv@numerator");
defineMacro("\\odv@numerator", "\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}");
defineMacro("\\odv@next", "\\frac{\\mathrm{d}}{\\mathrm{d}#2}#1");
defineMacro("\\pdv", "\\@ifstar\\pdv@next\\pdv@numerator");
const pdvHelper = (args) => {
	const numerator = args[0][0].text;
	const denoms = stringFromArg(args[1]).split(",");
	const power = String(denoms.length);
	const numOp = power === "1" ? "\\partial" : `\\partial^${power}`;
	let denominator = "";
	denoms.map((e) => {
		denominator += "\\partial " + e.trim() + "\\,";
	});
	return [
		numerator,
		numOp,
		denominator.replace(/\\,$/, "")
	];
};
defineMacro("\\pdv@numerator", function(context) {
	const [numerator, numOp, denominator] = pdvHelper(context.consumeArgs(2));
	return `\\frac{${numOp} ${numerator}}{${denominator}}`;
});
defineMacro("\\pdv@next", function(context) {
	const [numerator, numOp, denominator] = pdvHelper(context.consumeArgs(2));
	return `\\frac{${numOp}}{${denominator}} ${numerator}`;
});
defineMacro("\\upalpha", "\\up@greek{\\alpha}");
defineMacro("\\upbeta", "\\up@greek{\\beta}");
defineMacro("\\upgamma", "\\up@greek{\\gamma}");
defineMacro("\\updelta", "\\up@greek{\\delta}");
defineMacro("\\upepsilon", "\\up@greek{\\epsilon}");
defineMacro("\\upzeta", "\\up@greek{\\zeta}");
defineMacro("\\upeta", "\\up@greek{\\eta}");
defineMacro("\\uptheta", "\\up@greek{\\theta}");
defineMacro("\\upiota", "\\up@greek{\\iota}");
defineMacro("\\upkappa", "\\up@greek{\\kappa}");
defineMacro("\\uplambda", "\\up@greek{\\lambda}");
defineMacro("\\upmu", "\\up@greek{\\mu}");
defineMacro("\\upnu", "\\up@greek{\\nu}");
defineMacro("\\upxi", "\\up@greek{\\xi}");
defineMacro("\\upomicron", "\\up@greek{\\omicron}");
defineMacro("\\uppi", "\\up@greek{\\pi}");
defineMacro("\\upalpha", "\\up@greek{\\alpha}");
defineMacro("\\uprho", "\\up@greek{\\rho}");
defineMacro("\\upsigma", "\\up@greek{\\sigma}");
defineMacro("\\uptau", "\\up@greek{\\tau}");
defineMacro("\\upupsilon", "\\up@greek{\\upsilon}");
defineMacro("\\upphi", "\\up@greek{\\phi}");
defineMacro("\\upchi", "\\up@greek{\\chi}");
defineMacro("\\uppsi", "\\up@greek{\\psi}");
defineMacro("\\upomega", "\\up@greek{\\omega}");
defineMacro("\\invamp", "\\mathbin{\\char\"214b}");
defineMacro("\\parr", "\\mathbin{\\char\"214b}");
defineMacro("\\upand", "\\mathbin{\\char\"214b}");
defineMacro("\\with", "\\mathbin{\\char\"26}");
defineMacro("\\multimapinv", "\\mathrel{\\char\"27dc}");
defineMacro("\\multimapboth", "\\mathrel{\\char\"29df}");
defineMacro("\\scoh", "{\\mkern5mu\\char\"2322\\mkern5mu}");
defineMacro("\\sincoh", "{\\mkern5mu\\char\"2323\\mkern5mu}");
defineMacro("\\coh", `{\\mkern5mu\\rule{}{0.7em}\\mathrlap{\\smash{\\raise2mu{\\char"2322}}}
{\\smash{\\lower4mu{\\char"2323}}}\\mkern5mu}`);
defineMacro("\\incoh", `{\\mkern5mu\\rule{}{0.7em}\\mathrlap{\\smash{\\raise2mu{\\char"2323}}}
{\\smash{\\lower4mu{\\char"2322}}}\\mkern5mu}`);
defineMacro("\\standardstate", "\\text{\\tiny\\char`⦵}");
/*************************************************************
*
*  Temml mhchem.js
*
*  This file implements a Temml version of mhchem version 3.3.0.
*  It is adapted from MathJax/extensions/TeX/mhchem.js
*  It differs from the MathJax version as follows:
*    1. The interface is changed so that it can be called from Temml, not MathJax.
*    2. \rlap and \llap are replaced with \mathrlap and \mathllap.
*    3. The reaction arrow code is simplified. All reaction arrows are rendered
*       using Temml extensible arrows instead of building non-extensible arrows.
*    4. The ~bond forms are composed entirely of \rule elements.
*    5. Two dashes in _getBond are wrapped in braces to suppress spacing. i.e., {-}
*    6. The electron dot uses \textbullet instead of \bullet.
*    7. \smash[T] has been removed. (WebKit hides anything inside \smash{…})
*
*    This code, as other Temml code, is released under the MIT license.
* 
* /*************************************************************
*
*  MathJax/extensions/TeX/mhchem.js
*
*  Implements the \ce command for handling chemical formulas
*  from the mhchem LaTeX package.
*
*  ---------------------------------------------------------------------
*
*  Copyright (c) 2011-2015 The MathJax Consortium
*  Copyright (c) 2015-2018 Martin Hensel
*
*  Licensed under the Apache License, Version 2.0 (the "License");
*  you may not use this file except in compliance with the License.
*  You may obtain a copy of the License at
*
*      http://www.apache.org/licenses/LICENSE-2.0
*
*  Unless required by applicable law or agreed to in writing, software
*  distributed under the License is distributed on an "AS IS" BASIS,
*  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
*  See the License for the specific language governing permissions and
*  limitations under the License.
*/
defineMacro("\\ce", function(context) {
	return chemParse(context.consumeArgs(1)[0], "ce");
});
defineMacro("\\pu", function(context) {
	return chemParse(context.consumeArgs(1)[0], "pu");
});
defineMacro("\\uniDash", `{\\rule{0.672em}{0.06em}}`);
defineMacro("\\triDash", `{\\rule{0.15em}{0.06em}\\kern2mu\\rule{0.15em}{0.06em}\\kern2mu\\rule{0.15em}{0.06em}}`);
defineMacro("\\tripleDash", `\\kern0.075em\\raise0.25em{\\triDash}\\kern0.075em`);
defineMacro("\\tripleDashOverLine", `\\kern0.075em\\mathrlap{\\raise0.125em{\\uniDash}}\\raise0.34em{\\triDash}\\kern0.075em`);
defineMacro("\\tripleDashOverDoubleLine", `\\kern0.075em\\mathrlap{\\mathrlap{\\raise0.48em{\\triDash}}\\raise0.27em{\\uniDash}}{\\raise0.05em{\\uniDash}}\\kern0.075em`);
defineMacro("\\tripleDashBetweenDoubleLine", `\\kern0.075em\\mathrlap{\\mathrlap{\\raise0.48em{\\uniDash}}\\raise0.27em{\\triDash}}{\\raise0.05em{\\uniDash}}\\kern0.075em`);
var chemParse = function(tokens, stateMachine) {
	var str = "";
	var expectedLoc = tokens.length && tokens[tokens.length - 1].loc.start;
	for (var i = tokens.length - 1; i >= 0; i--) {
		if (tokens[i].loc.start > expectedLoc) {
			str += " ";
			expectedLoc = tokens[i].loc.start;
		}
		str += tokens[i].text;
		expectedLoc += tokens[i].text.length;
	}
	return texify.go(mhchemParser.go(str, stateMachine));
};
/** @type {MhchemParser} */
var mhchemParser = {
	go: function(input, stateMachine) {
		if (!input) return [];
		if (stateMachine === void 0) stateMachine = "ce";
		var state = "0";
		/** @type {Buffer} */
		var buffer = {};
		buffer["parenthesisLevel"] = 0;
		input = input.replace(/\n/g, " ");
		input = input.replace(/[\u2212\u2013\u2014\u2010]/g, "-");
		input = input.replace(/[\u2026]/g, "...");
		var lastInput;
		var watchdog = 10;
		/** @type {ParserOutput[]} */
		var output = [];
		while (true) {
			if (lastInput !== input) {
				watchdog = 10;
				lastInput = input;
			} else watchdog--;
			var machine = mhchemParser.stateMachines[stateMachine];
			var t = machine.transitions[state] || machine.transitions["*"];
			iterateTransitions: for (var i = 0; i < t.length; i++) {
				var matches = mhchemParser.patterns.match_(t[i].pattern, input);
				if (matches) {
					var task = t[i].task;
					for (var iA = 0; iA < task.action_.length; iA++) {
						var o;
						if (machine.actions[task.action_[iA].type_]) o = machine.actions[task.action_[iA].type_](buffer, matches.match_, task.action_[iA].option);
						else if (mhchemParser.actions[task.action_[iA].type_]) o = mhchemParser.actions[task.action_[iA].type_](buffer, matches.match_, task.action_[iA].option);
						else throw ["MhchemBugA", "mhchem bug A. Please report. (" + task.action_[iA].type_ + ")"];
						mhchemParser.concatArray(output, o);
					}
					state = task.nextState || state;
					if (input.length > 0) {
						if (!task.revisit) input = matches.remainder;
						if (!task.toContinue) break iterateTransitions;
					} else return output;
				}
			}
			if (watchdog <= 0) throw ["MhchemBugU", "mhchem bug U. Please report."];
		}
	},
	concatArray: function(a, b) {
		if (b) {
			if (Array.isArray(b)) for (var iB = 0; iB < b.length; iB++) a.push(b[iB]);
			else a.push(b);
		}
	},
	patterns: {
		patterns: {
			"empty": /^$/,
			"else": /^./,
			"else2": /^./,
			"space": /^\s/,
			"space A": /^\s(?=[A-Z\\$])/,
			"space$": /^\s$/,
			"a-z": /^[a-z]/,
			"x": /^x/,
			"x$": /^x$/,
			"i$": /^i$/,
			"letters": /^(?:[a-zA-Z\u03B1-\u03C9\u0391-\u03A9?@]|(?:\\(?:alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|omicron|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Upsilon|Phi|Psi|Omega)(?:\s+|\{\}|(?![a-zA-Z]))))+/,
			"\\greek": /^\\(?:alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|omicron|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega|Gamma|Delta|Theta|Lambda|Xi|Pi|Sigma|Upsilon|Phi|Psi|Omega)(?:\s+|\{\}|(?![a-zA-Z]))/,
			"one lowercase latin letter $": /^(?:([a-z])(?:$|[^a-zA-Z]))$/,
			"$one lowercase latin letter$ $": /^\$(?:([a-z])(?:$|[^a-zA-Z]))\$$/,
			"one lowercase greek letter $": /^(?:\$?[\u03B1-\u03C9]\$?|\$?\\(?:alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|omicron|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega)\s*\$?)(?:\s+|\{\}|(?![a-zA-Z]))$/,
			"digits": /^[0-9]+/,
			"-9.,9": /^[+\-]?(?:[0-9]+(?:[,.][0-9]+)?|[0-9]*(?:\.[0-9]+))/,
			"-9.,9 no missing 0": /^[+\-]?[0-9]+(?:[.,][0-9]+)?/,
			"(-)(9.,9)(e)(99)": function(input) {
				var m = input.match(/^(\+\-|\+\/\-|\+|\-|\\pm\s?)?([0-9]+(?:[,.][0-9]+)?|[0-9]*(?:\.[0-9]+))?(\((?:[0-9]+(?:[,.][0-9]+)?|[0-9]*(?:\.[0-9]+))\))?(?:([eE]|\s*(\*|x|\\times|\u00D7)\s*10\^)([+\-]?[0-9]+|\{[+\-]?[0-9]+\}))?/);
				if (m && m[0]) return {
					match_: m.splice(1),
					remainder: input.substr(m[0].length)
				};
				return null;
			},
			"(-)(9)^(-9)": function(input) {
				var m = input.match(/^(\+\-|\+\/\-|\+|\-|\\pm\s?)?([0-9]+(?:[,.][0-9]+)?|[0-9]*(?:\.[0-9]+)?)\^([+\-]?[0-9]+|\{[+\-]?[0-9]+\})/);
				if (m && m[0]) return {
					match_: m.splice(1),
					remainder: input.substr(m[0].length)
				};
				return null;
			},
			"state of aggregation $": function(input) {
				var a = mhchemParser.patterns.findObserveGroups(input, "", /^\([a-z]{1,3}(?=[\),])/, ")", "");
				if (a && a.remainder.match(/^($|[\s,;\)\]\}])/)) return a;
				var m = input.match(/^(?:\((?:\\ca\s?)?\$[amothc]\$\))/);
				if (m) return {
					match_: m[0],
					remainder: input.substr(m[0].length)
				};
				return null;
			},
			"_{(state of aggregation)}$": /^_\{(\([a-z]{1,3}\))\}/,
			"{[(": /^(?:\\\{|\[|\()/,
			")]}": /^(?:\)|\]|\\\})/,
			", ": /^[,;]\s*/,
			",": /^[,;]/,
			".": /^[.]/,
			". ": /^([.\u22C5\u00B7\u2022])\s*/,
			"...": /^\.\.\.(?=$|[^.])/,
			"* ": /^([*])\s*/,
			"^{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "^{", "", "", "}");
			},
			"^($...$)": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "^", "$", "$", "");
			},
			"^a": /^\^([0-9]+|[^\\_])/,
			"^\\x{}{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "^", /^\\[a-zA-Z]+\{/, "}", "", "", "{", "}", "", true);
			},
			"^\\x{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "^", /^\\[a-zA-Z]+\{/, "}", "");
			},
			"^\\x": /^\^(\\[a-zA-Z]+)\s*/,
			"^(-1)": /^\^(-?\d+)/,
			"'": /^'/,
			"_{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "_{", "", "", "}");
			},
			"_($...$)": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "_", "$", "$", "");
			},
			"_9": /^_([+\-]?[0-9]+|[^\\])/,
			"_\\x{}{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "_", /^\\[a-zA-Z]+\{/, "}", "", "", "{", "}", "", true);
			},
			"_\\x{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "_", /^\\[a-zA-Z]+\{/, "}", "");
			},
			"_\\x": /^_(\\[a-zA-Z]+)\s*/,
			"^_": /^(?:\^(?=_)|\_(?=\^)|[\^_]$)/,
			"{}": /^\{\}/,
			"{...}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "", "{", "}", "");
			},
			"{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "{", "", "", "}");
			},
			"$...$": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "", "$", "$", "");
			},
			"${(...)}$": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "${", "", "", "}$");
			},
			"$(...)$": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "$", "", "", "$");
			},
			"=<>": /^[=<>]/,
			"#": /^[#\u2261]/,
			"+": /^\+/,
			"-$": /^-(?=[\s_},;\]/]|$|\([a-z]+\))/,
			"-9": /^-(?=[0-9])/,
			"- orbital overlap": /^-(?=(?:[spd]|sp)(?:$|[\s,;\)\]\}]))/,
			"-": /^-/,
			"pm-operator": /^(?:\\pm|\$\\pm\$|\+-|\+\/-)/,
			"operator": /^(?:\+|(?:[\-=<>]|<<|>>|\\approx|\$\\approx\$)(?=\s|$|-?[0-9]))/,
			"arrowUpDown": /^(?:v|\(v\)|\^|\(\^\))(?=$|[\s,;\)\]\}])/,
			"\\bond{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\bond{", "", "", "}");
			},
			"->": /^(?:<->|<-->|->|<-|<=>>|<<=>|<=>|[\u2192\u27F6\u21CC])/,
			"CMT": /^[CMT](?=\[)/,
			"[(...)]": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "[", "", "", "]");
			},
			"1st-level escape": /^(&|\\\\|\\hline)\s*/,
			"\\,": /^(?:\\[,\ ;:])/,
			"\\x{}{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "", /^\\[a-zA-Z]+\{/, "}", "", "", "{", "}", "", true);
			},
			"\\x{}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "", /^\\[a-zA-Z]+\{/, "}", "");
			},
			"\\ca": /^\\ca(?:\s+|(?![a-zA-Z]))/,
			"\\x": /^(?:\\[a-zA-Z]+\s*|\\[_&{}%])/,
			"orbital": /^(?:[0-9]{1,2}[spdfgh]|[0-9]{0,2}sp)(?=$|[^a-zA-Z])/,
			"others": /^[\/~|]/,
			"\\frac{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\frac{", "", "", "}", "{", "", "", "}");
			},
			"\\overset{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\overset{", "", "", "}", "{", "", "", "}");
			},
			"\\underset{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\underset{", "", "", "}", "{", "", "", "}");
			},
			"\\underbrace{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\underbrace{", "", "", "}_", "{", "", "", "}");
			},
			"\\color{(...)}0": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\color{", "", "", "}");
			},
			"\\color{(...)}{(...)}1": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\color{", "", "", "}", "{", "", "", "}");
			},
			"\\color(...){(...)}2": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\color", "\\", "", /^(?=\{)/, "{", "", "", "}");
			},
			"\\ce{(...)}": function(input) {
				return mhchemParser.patterns.findObserveGroups(input, "\\ce{", "", "", "}");
			},
			"oxidation$": /^(?:[+-][IVX]+|\\pm\s*0|\$\\pm\$\s*0)$/,
			"d-oxidation$": /^(?:[+-]?\s?[IVX]+|\\pm\s*0|\$\\pm\$\s*0)$/,
			"roman numeral": /^[IVX]+/,
			"1/2$": /^[+\-]?(?:[0-9]+|\$[a-z]\$|[a-z])\/[0-9]+(?:\$[a-z]\$|[a-z])?$/,
			"amount": function(input) {
				var match = input.match(/^(?:(?:(?:\([+\-]?[0-9]+\/[0-9]+\)|[+\-]?(?:[0-9]+|\$[a-z]\$|[a-z])\/[0-9]+|[+\-]?[0-9]+[.,][0-9]+|[+\-]?\.[0-9]+|[+\-]?[0-9]+)(?:[a-z](?=\s*[A-Z]))?)|[+\-]?[a-z](?=\s*[A-Z])|\+(?!\s))/);
				if (match) return {
					match_: match[0],
					remainder: input.substr(match[0].length)
				};
				var a = mhchemParser.patterns.findObserveGroups(input, "", "$", "$", "");
				if (a) {
					match = a.match_.match(/^\$(?:\(?[+\-]?(?:[0-9]*[a-z]?[+\-])?[0-9]*[a-z](?:[+\-][0-9]*[a-z]?)?\)?|\+|-)\$$/);
					if (match) return {
						match_: match[0],
						remainder: input.substr(match[0].length)
					};
				}
				return null;
			},
			"amount2": function(input) {
				return this["amount"](input);
			},
			"(KV letters),": /^(?:[A-Z][a-z]{0,2}|i)(?=,)/,
			"formula$": function(input) {
				if (input.match(/^\([a-z]+\)$/)) return null;
				var match = input.match(/^(?:[a-z]|(?:[0-9\ \+\-\,\.\(\)]+[a-z])+[0-9\ \+\-\,\.\(\)]*|(?:[a-z][0-9\ \+\-\,\.\(\)]+)+[a-z]?)$/);
				if (match) return {
					match_: match[0],
					remainder: input.substr(match[0].length)
				};
				return null;
			},
			"uprightEntities": /^(?:pH|pOH|pC|pK|iPr|iBu)(?=$|[^a-zA-Z])/,
			"/": /^\s*(\/)\s*/,
			"//": /^\s*(\/\/)\s*/,
			"*": /^\s*[*.]\s*/
		},
		findObserveGroups: function(input, begExcl, begIncl, endIncl, endExcl, beg2Excl, beg2Incl, end2Incl, end2Excl, combine) {
			/** @type {{(input: string, pattern: string | RegExp): string | string[] | null;}} */
			var _match = function(input, pattern) {
				if (typeof pattern === "string") {
					if (input.indexOf(pattern) !== 0) return null;
					return pattern;
				} else {
					var match = input.match(pattern);
					if (!match) return null;
					return match[0];
				}
			};
			/** @type {{(input: string, i: number, endChars: string | RegExp): {endMatchBegin: number, endMatchEnd: number} | null;}} */
			var _findObserveGroups = function(input, i, endChars) {
				var braces = 0;
				while (i < input.length) {
					var a = input.charAt(i);
					var match = _match(input.substr(i), endChars);
					if (match !== null && braces === 0) return {
						endMatchBegin: i,
						endMatchEnd: i + match.length
					};
					else if (a === "{") braces++;
					else if (a === "}") {
						if (braces === 0) throw ["ExtraCloseMissingOpen", "Extra close brace or missing open brace"];
						else braces--;
					}
					i++;
				}
				if (braces > 0) return null;
				return null;
			};
			var match = _match(input, begExcl);
			if (match === null) return null;
			input = input.substr(match.length);
			match = _match(input, begIncl);
			if (match === null) return null;
			var e = _findObserveGroups(input, match.length, endIncl || endExcl);
			if (e === null) return null;
			var match1 = input.substring(0, endIncl ? e.endMatchEnd : e.endMatchBegin);
			if (!(beg2Excl || beg2Incl)) return {
				match_: match1,
				remainder: input.substr(e.endMatchEnd)
			};
			else {
				var group2 = this.findObserveGroups(input.substr(e.endMatchEnd), beg2Excl, beg2Incl, end2Incl, end2Excl);
				if (group2 === null) return null;
				/** @type {string[]} */
				var matchRet = [match1, group2.match_];
				return {
					match_: combine ? matchRet.join("") : matchRet,
					remainder: group2.remainder
				};
			}
		},
		match_: function(m, input) {
			var pattern = mhchemParser.patterns.patterns[m];
			if (pattern === void 0) throw ["MhchemBugP", "mhchem bug P. Please report. (" + m + ")"];
			else if (typeof pattern === "function") return mhchemParser.patterns.patterns[m](input);
			else {
				var match = input.match(pattern);
				if (match) {
					var mm;
					if (match[2]) mm = [match[1], match[2]];
					else if (match[1]) mm = match[1];
					else mm = match[0];
					return {
						match_: mm,
						remainder: input.substr(match[0].length)
					};
				}
				return null;
			}
		}
	},
	actions: {
		"a=": function(buffer, m) {
			buffer.a = (buffer.a || "") + m;
		},
		"b=": function(buffer, m) {
			buffer.b = (buffer.b || "") + m;
		},
		"p=": function(buffer, m) {
			buffer.p = (buffer.p || "") + m;
		},
		"o=": function(buffer, m) {
			buffer.o = (buffer.o || "") + m;
		},
		"q=": function(buffer, m) {
			buffer.q = (buffer.q || "") + m;
		},
		"d=": function(buffer, m) {
			buffer.d = (buffer.d || "") + m;
		},
		"rm=": function(buffer, m) {
			buffer.rm = (buffer.rm || "") + m;
		},
		"text=": function(buffer, m) {
			buffer.text_ = (buffer.text_ || "") + m;
		},
		"insert": function(buffer, m, a) {
			return { type_: a };
		},
		"insert+p1": function(buffer, m, a) {
			return {
				type_: a,
				p1: m
			};
		},
		"insert+p1+p2": function(buffer, m, a) {
			return {
				type_: a,
				p1: m[0],
				p2: m[1]
			};
		},
		"copy": function(buffer, m) {
			return m;
		},
		"rm": function(buffer, m) {
			return {
				type_: "rm",
				p1: m || ""
			};
		},
		"text": function(buffer, m) {
			return mhchemParser.go(m, "text");
		},
		"{text}": function(buffer, m) {
			var ret = ["{"];
			mhchemParser.concatArray(ret, mhchemParser.go(m, "text"));
			ret.push("}");
			return ret;
		},
		"tex-math": function(buffer, m) {
			return mhchemParser.go(m, "tex-math");
		},
		"tex-math tight": function(buffer, m) {
			return mhchemParser.go(m, "tex-math tight");
		},
		"bond": function(buffer, m, k) {
			return {
				type_: "bond",
				kind_: k || m
			};
		},
		"color0-output": function(buffer, m) {
			return {
				type_: "color0",
				color: m[0]
			};
		},
		"ce": function(buffer, m) {
			return mhchemParser.go(m);
		},
		"1/2": function(buffer, m) {
			/** @type {ParserOutput[]} */
			var ret = [];
			if (m.match(/^[+\-]/)) {
				ret.push(m.substr(0, 1));
				m = m.substr(1);
			}
			var n = m.match(/^([0-9]+|\$[a-z]\$|[a-z])\/([0-9]+)(\$[a-z]\$|[a-z])?$/);
			n[1] = n[1].replace(/\$/g, "");
			ret.push({
				type_: "frac",
				p1: n[1],
				p2: n[2]
			});
			if (n[3]) {
				n[3] = n[3].replace(/\$/g, "");
				ret.push({
					type_: "tex-math",
					p1: n[3]
				});
			}
			return ret;
		},
		"9,9": function(buffer, m) {
			return mhchemParser.go(m, "9,9");
		}
	},
	createTransitions: function(o) {
		var pattern, state;
		/** @type {string[]} */
		var stateArray;
		var i;
		/** @type {Transitions} */
		var transitions = {};
		for (pattern in o) for (state in o[pattern]) {
			stateArray = state.split("|");
			o[pattern][state].stateArray = stateArray;
			for (i = 0; i < stateArray.length; i++) transitions[stateArray[i]] = [];
		}
		for (pattern in o) for (state in o[pattern]) {
			stateArray = o[pattern][state].stateArray || [];
			for (i = 0; i < stateArray.length; i++) {
				/** @type {any} */
				var p = o[pattern][state];
				if (p.action_) {
					p.action_ = [].concat(p.action_);
					for (var k = 0; k < p.action_.length; k++) if (typeof p.action_[k] === "string") p.action_[k] = { type_: p.action_[k] };
				} else p.action_ = [];
				var patternArray = pattern.split("|");
				for (var j = 0; j < patternArray.length; j++) if (stateArray[i] === "*") for (var t in transitions) transitions[t].push({
					pattern: patternArray[j],
					task: p
				});
				else transitions[stateArray[i]].push({
					pattern: patternArray[j],
					task: p
				});
			}
		}
		return transitions;
	},
	stateMachines: {}
};
mhchemParser.stateMachines = {
	"ce": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"else": { "0|1|2": {
				action_: "beginsWithBond=false",
				revisit: true,
				toContinue: true
			} },
			"oxidation$": { "0": { action_: "oxidation-output" } },
			"CMT": {
				"r": {
					action_: "rdt=",
					nextState: "rt"
				},
				"rd": {
					action_: "rqt=",
					nextState: "rdt"
				}
			},
			"arrowUpDown": { "0|1|2|as": {
				action_: [
					"sb=false",
					"output",
					"operator"
				],
				nextState: "1"
			} },
			"uprightEntities": { "0|1|2": {
				action_: ["o=", "output"],
				nextState: "1"
			} },
			"orbital": { "0|1|2|3": {
				action_: "o=",
				nextState: "o"
			} },
			"->": {
				"0|1|2|3": {
					action_: "r=",
					nextState: "r"
				},
				"a|as": {
					action_: ["output", "r="],
					nextState: "r"
				},
				"*": {
					action_: ["output", "r="],
					nextState: "r"
				}
			},
			"+": {
				"o": {
					action_: "d= kv",
					nextState: "d"
				},
				"d|D": {
					action_: "d=",
					nextState: "d"
				},
				"q": {
					action_: "d=",
					nextState: "qd"
				},
				"qd|qD": {
					action_: "d=",
					nextState: "qd"
				},
				"dq": {
					action_: ["output", "d="],
					nextState: "d"
				},
				"3": {
					action_: [
						"sb=false",
						"output",
						"operator"
					],
					nextState: "0"
				}
			},
			"amount": { "0|2": {
				action_: "a=",
				nextState: "a"
			} },
			"pm-operator": { "0|1|2|a|as": {
				action_: [
					"sb=false",
					"output",
					{
						type_: "operator",
						option: "\\pm"
					}
				],
				nextState: "0"
			} },
			"operator": { "0|1|2|a|as": {
				action_: [
					"sb=false",
					"output",
					"operator"
				],
				nextState: "0"
			} },
			"-$": {
				"o|q": {
					action_: ["charge or bond", "output"],
					nextState: "qd"
				},
				"d": {
					action_: "d=",
					nextState: "d"
				},
				"D": {
					action_: ["output", {
						type_: "bond",
						option: "-"
					}],
					nextState: "3"
				},
				"q": {
					action_: "d=",
					nextState: "qd"
				},
				"qd": {
					action_: "d=",
					nextState: "qd"
				},
				"qD|dq": {
					action_: ["output", {
						type_: "bond",
						option: "-"
					}],
					nextState: "3"
				}
			},
			"-9": { "3|o": {
				action_: ["output", {
					type_: "insert",
					option: "hyphen"
				}],
				nextState: "3"
			} },
			"- orbital overlap": {
				"o": {
					action_: ["output", {
						type_: "insert",
						option: "hyphen"
					}],
					nextState: "2"
				},
				"d": {
					action_: ["output", {
						type_: "insert",
						option: "hyphen"
					}],
					nextState: "2"
				}
			},
			"-": {
				"0|1|2": {
					action_: [
						{
							type_: "output",
							option: 1
						},
						"beginsWithBond=true",
						{
							type_: "bond",
							option: "-"
						}
					],
					nextState: "3"
				},
				"3": { action_: {
					type_: "bond",
					option: "-"
				} },
				"a": {
					action_: ["output", {
						type_: "insert",
						option: "hyphen"
					}],
					nextState: "2"
				},
				"as": {
					action_: [{
						type_: "output",
						option: 2
					}, {
						type_: "bond",
						option: "-"
					}],
					nextState: "3"
				},
				"b": { action_: "b=" },
				"o": {
					action_: {
						type_: "- after o/d",
						option: false
					},
					nextState: "2"
				},
				"q": {
					action_: {
						type_: "- after o/d",
						option: false
					},
					nextState: "2"
				},
				"d|qd|dq": {
					action_: {
						type_: "- after o/d",
						option: true
					},
					nextState: "2"
				},
				"D|qD|p": {
					action_: ["output", {
						type_: "bond",
						option: "-"
					}],
					nextState: "3"
				}
			},
			"amount2": { "1|3": {
				action_: "a=",
				nextState: "a"
			} },
			"letters": {
				"0|1|2|3|a|as|b|p|bp|o": {
					action_: "o=",
					nextState: "o"
				},
				"q|dq": {
					action_: ["output", "o="],
					nextState: "o"
				},
				"d|D|qd|qD": {
					action_: "o after d",
					nextState: "o"
				}
			},
			"digits": {
				"o": {
					action_: "q=",
					nextState: "q"
				},
				"d|D": {
					action_: "q=",
					nextState: "dq"
				},
				"q": {
					action_: ["output", "o="],
					nextState: "o"
				},
				"a": {
					action_: "o=",
					nextState: "o"
				}
			},
			"space A": { "b|p|bp": {} },
			"space": {
				"a": { nextState: "as" },
				"0": { action_: "sb=false" },
				"1|2": { action_: "sb=true" },
				"r|rt|rd|rdt|rdq": {
					action_: "output",
					nextState: "0"
				},
				"*": {
					action_: ["output", "sb=true"],
					nextState: "1"
				}
			},
			"1st-level escape": {
				"1|2": { action_: ["output", {
					type_: "insert+p1",
					option: "1st-level escape"
				}] },
				"*": {
					action_: ["output", {
						type_: "insert+p1",
						option: "1st-level escape"
					}],
					nextState: "0"
				}
			},
			"[(...)]": {
				"r|rt": {
					action_: "rd=",
					nextState: "rd"
				},
				"rd|rdt": {
					action_: "rq=",
					nextState: "rdq"
				}
			},
			"...": {
				"o|d|D|dq|qd|qD": {
					action_: ["output", {
						type_: "bond",
						option: "..."
					}],
					nextState: "3"
				},
				"*": {
					action_: [{
						type_: "output",
						option: 1
					}, {
						type_: "insert",
						option: "ellipsis"
					}],
					nextState: "1"
				}
			},
			". |* ": { "*": {
				action_: ["output", {
					type_: "insert",
					option: "addition compound"
				}],
				nextState: "1"
			} },
			"state of aggregation $": { "*": {
				action_: ["output", "state of aggregation"],
				nextState: "1"
			} },
			"{[(": {
				"a|as|o": {
					action_: [
						"o=",
						"output",
						"parenthesisLevel++"
					],
					nextState: "2"
				},
				"0|1|2|3": {
					action_: [
						"o=",
						"output",
						"parenthesisLevel++"
					],
					nextState: "2"
				},
				"*": {
					action_: [
						"output",
						"o=",
						"output",
						"parenthesisLevel++"
					],
					nextState: "2"
				}
			},
			")]}": {
				"0|1|2|3|b|p|bp|o": {
					action_: ["o=", "parenthesisLevel--"],
					nextState: "o"
				},
				"a|as|d|D|q|qd|qD|dq": {
					action_: [
						"output",
						"o=",
						"parenthesisLevel--"
					],
					nextState: "o"
				}
			},
			", ": { "*": {
				action_: ["output", "comma"],
				nextState: "0"
			} },
			"^_": { "*": {} },
			"^{(...)}|^($...$)": {
				"0|1|2|as": {
					action_: "b=",
					nextState: "b"
				},
				"p": {
					action_: "b=",
					nextState: "bp"
				},
				"3|o": {
					action_: "d= kv",
					nextState: "D"
				},
				"q": {
					action_: "d=",
					nextState: "qD"
				},
				"d|D|qd|qD|dq": {
					action_: ["output", "d="],
					nextState: "D"
				}
			},
			"^a|^\\x{}{}|^\\x{}|^\\x|'": {
				"0|1|2|as": {
					action_: "b=",
					nextState: "b"
				},
				"p": {
					action_: "b=",
					nextState: "bp"
				},
				"3|o": {
					action_: "d= kv",
					nextState: "d"
				},
				"q": {
					action_: "d=",
					nextState: "qd"
				},
				"d|qd|D|qD": { action_: "d=" },
				"dq": {
					action_: ["output", "d="],
					nextState: "d"
				}
			},
			"_{(state of aggregation)}$": { "d|D|q|qd|qD|dq": {
				action_: ["output", "q="],
				nextState: "q"
			} },
			"_{(...)}|_($...$)|_9|_\\x{}{}|_\\x{}|_\\x": {
				"0|1|2|as": {
					action_: "p=",
					nextState: "p"
				},
				"b": {
					action_: "p=",
					nextState: "bp"
				},
				"3|o": {
					action_: "q=",
					nextState: "q"
				},
				"d|D": {
					action_: "q=",
					nextState: "dq"
				},
				"q|qd|qD|dq": {
					action_: ["output", "q="],
					nextState: "q"
				}
			},
			"=<>": { "0|1|2|3|a|as|o|q|d|D|qd|qD|dq": {
				action_: [{
					type_: "output",
					option: 2
				}, "bond"],
				nextState: "3"
			} },
			"#": { "0|1|2|3|a|as|o": {
				action_: [{
					type_: "output",
					option: 2
				}, {
					type_: "bond",
					option: "#"
				}],
				nextState: "3"
			} },
			"{}": { "*": {
				action_: {
					type_: "output",
					option: 1
				},
				nextState: "1"
			} },
			"{...}": {
				"0|1|2|3|a|as|b|p|bp": {
					action_: "o=",
					nextState: "o"
				},
				"o|d|D|q|qd|qD|dq": {
					action_: ["output", "o="],
					nextState: "o"
				}
			},
			"$...$": {
				"a": { action_: "a=" },
				"0|1|2|3|as|b|p|bp|o": {
					action_: "o=",
					nextState: "o"
				},
				"as|o": { action_: "o=" },
				"q|d|D|qd|qD|dq": {
					action_: ["output", "o="],
					nextState: "o"
				}
			},
			"\\bond{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "bond"],
				nextState: "3"
			} },
			"\\frac{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 1
				}, "frac-output"],
				nextState: "3"
			} },
			"\\overset{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "overset-output"],
				nextState: "3"
			} },
			"\\underset{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "underset-output"],
				nextState: "3"
			} },
			"\\underbrace{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "underbrace-output"],
				nextState: "3"
			} },
			"\\color{(...)}{(...)}1|\\color(...){(...)}2": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "color-output"],
				nextState: "3"
			} },
			"\\color{(...)}0": { "*": { action_: [{
				type_: "output",
				option: 2
			}, "color0-output"] } },
			"\\ce{(...)}": { "*": {
				action_: [{
					type_: "output",
					option: 2
				}, "ce"],
				nextState: "3"
			} },
			"\\,": { "*": {
				action_: [{
					type_: "output",
					option: 1
				}, "copy"],
				nextState: "1"
			} },
			"\\x{}{}|\\x{}|\\x": {
				"0|1|2|3|a|as|b|p|bp|o|c0": {
					action_: ["o=", "output"],
					nextState: "3"
				},
				"*": {
					action_: [
						"output",
						"o=",
						"output"
					],
					nextState: "3"
				}
			},
			"others": { "*": {
				action_: [{
					type_: "output",
					option: 1
				}, "copy"],
				nextState: "3"
			} },
			"else2": {
				"a": {
					action_: "a to o",
					nextState: "o",
					revisit: true
				},
				"as": {
					action_: ["output", "sb=true"],
					nextState: "1",
					revisit: true
				},
				"r|rt|rd|rdt|rdq": {
					action_: ["output"],
					nextState: "0",
					revisit: true
				},
				"*": {
					action_: ["output", "copy"],
					nextState: "3"
				}
			}
		}),
		actions: {
			"o after d": function(buffer, m) {
				var ret;
				if ((buffer.d || "").match(/^[0-9]+$/)) {
					var tmp = buffer.d;
					buffer.d = void 0;
					ret = this["output"](buffer);
					buffer.b = tmp;
				} else ret = this["output"](buffer);
				mhchemParser.actions["o="](buffer, m);
				return ret;
			},
			"d= kv": function(buffer, m) {
				buffer.d = m;
				buffer.dType = "kv";
			},
			"charge or bond": function(buffer, m) {
				if (buffer["beginsWithBond"]) {
					/** @type {ParserOutput[]} */
					var ret = [];
					mhchemParser.concatArray(ret, this["output"](buffer));
					mhchemParser.concatArray(ret, mhchemParser.actions["bond"](buffer, m, "-"));
					return ret;
				} else buffer.d = m;
			},
			"- after o/d": function(buffer, m, isAfterD) {
				var c1 = mhchemParser.patterns.match_("orbital", buffer.o || "");
				var c2 = mhchemParser.patterns.match_("one lowercase greek letter $", buffer.o || "");
				var c3 = mhchemParser.patterns.match_("one lowercase latin letter $", buffer.o || "");
				var c4 = mhchemParser.patterns.match_("$one lowercase latin letter$ $", buffer.o || "");
				var hyphenFollows = m === "-" && (c1 && c1.remainder === "" || c2 || c3 || c4);
				if (hyphenFollows && !buffer.a && !buffer.b && !buffer.p && !buffer.d && !buffer.q && !c1 && c3) buffer.o = "$" + buffer.o + "$";
				/** @type {ParserOutput[]} */
				var ret = [];
				if (hyphenFollows) {
					mhchemParser.concatArray(ret, this["output"](buffer));
					ret.push({ type_: "hyphen" });
				} else {
					c1 = mhchemParser.patterns.match_("digits", buffer.d || "");
					if (isAfterD && c1 && c1.remainder === "") {
						mhchemParser.concatArray(ret, mhchemParser.actions["d="](buffer, m));
						mhchemParser.concatArray(ret, this["output"](buffer));
					} else {
						mhchemParser.concatArray(ret, this["output"](buffer));
						mhchemParser.concatArray(ret, mhchemParser.actions["bond"](buffer, m, "-"));
					}
				}
				return ret;
			},
			"a to o": function(buffer) {
				buffer.o = buffer.a;
				buffer.a = void 0;
			},
			"sb=true": function(buffer) {
				buffer.sb = true;
			},
			"sb=false": function(buffer) {
				buffer.sb = false;
			},
			"beginsWithBond=true": function(buffer) {
				buffer["beginsWithBond"] = true;
			},
			"beginsWithBond=false": function(buffer) {
				buffer["beginsWithBond"] = false;
			},
			"parenthesisLevel++": function(buffer) {
				buffer["parenthesisLevel"]++;
			},
			"parenthesisLevel--": function(buffer) {
				buffer["parenthesisLevel"]--;
			},
			"state of aggregation": function(buffer, m) {
				return {
					type_: "state of aggregation",
					p1: mhchemParser.go(m, "o")
				};
			},
			"comma": function(buffer, m) {
				var a = m.replace(/\s*$/, "");
				if (a !== m && buffer["parenthesisLevel"] === 0) return {
					type_: "comma enumeration L",
					p1: a
				};
				else return {
					type_: "comma enumeration M",
					p1: a
				};
			},
			"output": function(buffer, m, entityFollows) {
				/** @type {ParserOutput | ParserOutput[]} */
				var ret;
				if (!buffer.r) {
					ret = [];
					if (!buffer.a && !buffer.b && !buffer.p && !buffer.o && !buffer.q && !buffer.d && !entityFollows) {} else {
						if (buffer.sb) ret.push({ type_: "entitySkip" });
						if (!buffer.o && !buffer.q && !buffer.d && !buffer.b && !buffer.p && entityFollows !== 2) {
							buffer.o = buffer.a;
							buffer.a = void 0;
						} else if (!buffer.o && !buffer.q && !buffer.d && (buffer.b || buffer.p)) {
							buffer.o = buffer.a;
							buffer.d = buffer.b;
							buffer.q = buffer.p;
							buffer.a = buffer.b = buffer.p = void 0;
						} else if (buffer.o && buffer.dType === "kv" && mhchemParser.patterns.match_("d-oxidation$", buffer.d || "")) buffer.dType = "oxidation";
						else if (buffer.o && buffer.dType === "kv" && !buffer.q) buffer.dType = void 0;
						ret.push({
							type_: "chemfive",
							a: mhchemParser.go(buffer.a, "a"),
							b: mhchemParser.go(buffer.b, "bd"),
							p: mhchemParser.go(buffer.p, "pq"),
							o: mhchemParser.go(buffer.o, "o"),
							q: mhchemParser.go(buffer.q, "pq"),
							d: mhchemParser.go(buffer.d, buffer.dType === "oxidation" ? "oxidation" : "bd"),
							dType: buffer.dType
						});
					}
				} else {
					/** @type {ParserOutput[]} */
					var rd;
					if (buffer.rdt === "M") rd = mhchemParser.go(buffer.rd, "tex-math");
					else if (buffer.rdt === "T") rd = [{
						type_: "text",
						p1: buffer.rd || ""
					}];
					else rd = mhchemParser.go(buffer.rd);
					/** @type {ParserOutput[]} */
					var rq;
					if (buffer.rqt === "M") rq = mhchemParser.go(buffer.rq, "tex-math");
					else if (buffer.rqt === "T") rq = [{
						type_: "text",
						p1: buffer.rq || ""
					}];
					else rq = mhchemParser.go(buffer.rq);
					ret = {
						type_: "arrow",
						r: buffer.r,
						rd,
						rq
					};
				}
				for (var p in buffer) if (p !== "parenthesisLevel" && p !== "beginsWithBond") delete buffer[p];
				return ret;
			},
			"oxidation-output": function(buffer, m) {
				var ret = ["{"];
				mhchemParser.concatArray(ret, mhchemParser.go(m, "oxidation"));
				ret.push("}");
				return ret;
			},
			"frac-output": function(buffer, m) {
				return {
					type_: "frac-ce",
					p1: mhchemParser.go(m[0]),
					p2: mhchemParser.go(m[1])
				};
			},
			"overset-output": function(buffer, m) {
				return {
					type_: "overset",
					p1: mhchemParser.go(m[0]),
					p2: mhchemParser.go(m[1])
				};
			},
			"underset-output": function(buffer, m) {
				return {
					type_: "underset",
					p1: mhchemParser.go(m[0]),
					p2: mhchemParser.go(m[1])
				};
			},
			"underbrace-output": function(buffer, m) {
				return {
					type_: "underbrace",
					p1: mhchemParser.go(m[0]),
					p2: mhchemParser.go(m[1])
				};
			},
			"color-output": function(buffer, m) {
				return {
					type_: "color",
					color1: m[0],
					color2: mhchemParser.go(m[1])
				};
			},
			"r=": function(buffer, m) {
				buffer.r = m;
			},
			"rdt=": function(buffer, m) {
				buffer.rdt = m;
			},
			"rd=": function(buffer, m) {
				buffer.rd = m;
			},
			"rqt=": function(buffer, m) {
				buffer.rqt = m;
			},
			"rq=": function(buffer, m) {
				buffer.rq = m;
			},
			"operator": function(buffer, m, p1) {
				return {
					type_: "operator",
					kind_: p1 || m
				};
			}
		}
	},
	"a": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			"1/2$": { "0": { action_: "1/2" } },
			"else": { "0": {
				nextState: "1",
				revisit: true
			} },
			"$(...)$": { "*": {
				action_: "tex-math tight",
				nextState: "1"
			} },
			",": { "*": { action_: {
				type_: "insert",
				option: "commaDecimal"
			} } },
			"else2": { "*": { action_: "copy" } }
		}),
		actions: {}
	},
	"o": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			"1/2$": { "0": { action_: "1/2" } },
			"else": { "0": {
				nextState: "1",
				revisit: true
			} },
			"letters": { "*": { action_: "rm" } },
			"\\ca": { "*": { action_: {
				type_: "insert",
				option: "circa"
			} } },
			"\\x{}{}|\\x{}|\\x": { "*": { action_: "copy" } },
			"${(...)}$|$(...)$": { "*": { action_: "tex-math" } },
			"{(...)}": { "*": { action_: "{text}" } },
			"else2": { "*": { action_: "copy" } }
		}),
		actions: {}
	},
	"text": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"{...}": { "*": { action_: "text=" } },
			"${(...)}$|$(...)$": { "*": { action_: "tex-math" } },
			"\\greek": { "*": { action_: ["output", "rm"] } },
			"\\,|\\x{}{}|\\x{}|\\x": { "*": { action_: ["output", "copy"] } },
			"else": { "*": { action_: "text=" } }
		}),
		actions: { "output": function(buffer) {
			if (buffer.text_) {
				/** @type {ParserOutput} */
				var ret = {
					type_: "text",
					p1: buffer.text_
				};
				for (var p in buffer) delete buffer[p];
				return ret;
			}
		} }
	},
	"pq": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			"state of aggregation $": { "*": { action_: "state of aggregation" } },
			"i$": { "0": {
				nextState: "!f",
				revisit: true
			} },
			"(KV letters),": { "0": {
				action_: "rm",
				nextState: "0"
			} },
			"formula$": { "0": {
				nextState: "f",
				revisit: true
			} },
			"1/2$": { "0": { action_: "1/2" } },
			"else": { "0": {
				nextState: "!f",
				revisit: true
			} },
			"${(...)}$|$(...)$": { "*": { action_: "tex-math" } },
			"{(...)}": { "*": { action_: "text" } },
			"a-z": { "f": { action_: "tex-math" } },
			"letters": { "*": { action_: "rm" } },
			"-9.,9": { "*": { action_: "9,9" } },
			",": { "*": { action_: {
				type_: "insert+p1",
				option: "comma enumeration S"
			} } },
			"\\color{(...)}{(...)}1|\\color(...){(...)}2": { "*": { action_: "color-output" } },
			"\\color{(...)}0": { "*": { action_: "color0-output" } },
			"\\ce{(...)}": { "*": { action_: "ce" } },
			"\\,|\\x{}{}|\\x{}|\\x": { "*": { action_: "copy" } },
			"else2": { "*": { action_: "copy" } }
		}),
		actions: {
			"state of aggregation": function(buffer, m) {
				return {
					type_: "state of aggregation subscript",
					p1: mhchemParser.go(m, "o")
				};
			},
			"color-output": function(buffer, m) {
				return {
					type_: "color",
					color1: m[0],
					color2: mhchemParser.go(m[1], "pq")
				};
			}
		}
	},
	"bd": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			"x$": { "0": {
				nextState: "!f",
				revisit: true
			} },
			"formula$": { "0": {
				nextState: "f",
				revisit: true
			} },
			"else": { "0": {
				nextState: "!f",
				revisit: true
			} },
			"-9.,9 no missing 0": { "*": { action_: "9,9" } },
			".": { "*": { action_: {
				type_: "insert",
				option: "electron dot"
			} } },
			"a-z": { "f": { action_: "tex-math" } },
			"x": { "*": { action_: {
				type_: "insert",
				option: "KV x"
			} } },
			"letters": { "*": { action_: "rm" } },
			"'": { "*": { action_: {
				type_: "insert",
				option: "prime"
			} } },
			"${(...)}$|$(...)$": { "*": { action_: "tex-math" } },
			"{(...)}": { "*": { action_: "text" } },
			"\\color{(...)}{(...)}1|\\color(...){(...)}2": { "*": { action_: "color-output" } },
			"\\color{(...)}0": { "*": { action_: "color0-output" } },
			"\\ce{(...)}": { "*": { action_: "ce" } },
			"\\,|\\x{}{}|\\x{}|\\x": { "*": { action_: "copy" } },
			"else2": { "*": { action_: "copy" } }
		}),
		actions: { "color-output": function(buffer, m) {
			return {
				type_: "color",
				color1: m[0],
				color2: mhchemParser.go(m[1], "bd")
			};
		} }
	},
	"oxidation": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			"roman numeral": { "*": { action_: "roman-numeral" } },
			"${(...)}$|$(...)$": { "*": { action_: "tex-math" } },
			"else": { "*": { action_: "copy" } }
		}),
		actions: { "roman-numeral": function(buffer, m) {
			return {
				type_: "roman numeral",
				p1: m || ""
			};
		} }
	},
	"tex-math": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"\\ce{(...)}": { "*": { action_: ["output", "ce"] } },
			"{...}|\\,|\\x{}{}|\\x{}|\\x": { "*": { action_: "o=" } },
			"else": { "*": { action_: "o=" } }
		}),
		actions: { "output": function(buffer) {
			if (buffer.o) {
				/** @type {ParserOutput} */
				var ret = {
					type_: "tex-math",
					p1: buffer.o
				};
				for (var p in buffer) delete buffer[p];
				return ret;
			}
		} }
	},
	"tex-math tight": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"\\ce{(...)}": { "*": { action_: ["output", "ce"] } },
			"{...}|\\,|\\x{}{}|\\x{}|\\x": { "*": { action_: "o=" } },
			"-|+": { "*": { action_: "tight operator" } },
			"else": { "*": { action_: "o=" } }
		}),
		actions: {
			"tight operator": function(buffer, m) {
				buffer.o = (buffer.o || "") + "{" + m + "}";
			},
			"output": function(buffer) {
				if (buffer.o) {
					/** @type {ParserOutput} */
					var ret = {
						type_: "tex-math",
						p1: buffer.o
					};
					for (var p in buffer) delete buffer[p];
					return ret;
				}
			}
		}
	},
	"9,9": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": {} },
			",": { "*": { action_: "comma" } },
			"else": { "*": { action_: "copy" } }
		}),
		actions: { "comma": function() {
			return { type_: "commaDecimal" };
		} }
	},
	"pu": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"space$": { "*": { action_: ["output", "space"] } },
			"{[(|)]}": { "0|a": { action_: "copy" } },
			"(-)(9)^(-9)": { "0": {
				action_: "number^",
				nextState: "a"
			} },
			"(-)(9.,9)(e)(99)": { "0": {
				action_: "enumber",
				nextState: "a"
			} },
			"space": { "0|a": {} },
			"pm-operator": { "0|a": {
				action_: {
					type_: "operator",
					option: "\\pm"
				},
				nextState: "0"
			} },
			"operator": { "0|a": {
				action_: "copy",
				nextState: "0"
			} },
			"//": { "d": {
				action_: "o=",
				nextState: "/"
			} },
			"/": { "d": {
				action_: "o=",
				nextState: "/"
			} },
			"{...}|else": {
				"0|d": {
					action_: "d=",
					nextState: "d"
				},
				"a": {
					action_: ["space", "d="],
					nextState: "d"
				},
				"/|q": {
					action_: "q=",
					nextState: "q"
				}
			}
		}),
		actions: {
			"enumber": function(buffer, m) {
				/** @type {ParserOutput[]} */
				var ret = [];
				if (m[0] === "+-" || m[0] === "+/-") ret.push("\\pm ");
				else if (m[0]) ret.push(m[0]);
				if (m[1]) {
					mhchemParser.concatArray(ret, mhchemParser.go(m[1], "pu-9,9"));
					if (m[2]) {
						if (m[2].match(/[,.]/)) mhchemParser.concatArray(ret, mhchemParser.go(m[2], "pu-9,9"));
						else ret.push(m[2]);
					}
					m[3] = m[4] || m[3];
					if (m[3]) {
						m[3] = m[3].trim();
						if (m[3] === "e" || m[3].substr(0, 1) === "*") ret.push({ type_: "cdot" });
						else ret.push({ type_: "times" });
					}
				}
				if (m[3]) ret.push("10^{" + m[5] + "}");
				return ret;
			},
			"number^": function(buffer, m) {
				/** @type {ParserOutput[]} */
				var ret = [];
				if (m[0] === "+-" || m[0] === "+/-") ret.push("\\pm ");
				else if (m[0]) ret.push(m[0]);
				mhchemParser.concatArray(ret, mhchemParser.go(m[1], "pu-9,9"));
				ret.push("^{" + m[2] + "}");
				return ret;
			},
			"operator": function(buffer, m, p1) {
				return {
					type_: "operator",
					kind_: p1 || m
				};
			},
			"space": function() {
				return { type_: "pu-space-1" };
			},
			"output": function(buffer) {
				/** @type {ParserOutput | ParserOutput[]} */
				var ret;
				var md = mhchemParser.patterns.match_("{(...)}", buffer.d || "");
				if (md && md.remainder === "") buffer.d = md.match_;
				var mq = mhchemParser.patterns.match_("{(...)}", buffer.q || "");
				if (mq && mq.remainder === "") buffer.q = mq.match_;
				if (buffer.d) {
					buffer.d = buffer.d.replace(/\u00B0C|\^oC|\^{o}C/g, "{}^{\\circ}C");
					buffer.d = buffer.d.replace(/\u00B0F|\^oF|\^{o}F/g, "{}^{\\circ}F");
				}
				if (buffer.q) {
					buffer.q = buffer.q.replace(/\u00B0C|\^oC|\^{o}C/g, "{}^{\\circ}C");
					buffer.q = buffer.q.replace(/\u00B0F|\^oF|\^{o}F/g, "{}^{\\circ}F");
					var b5 = {
						d: mhchemParser.go(buffer.d, "pu"),
						q: mhchemParser.go(buffer.q, "pu")
					};
					if (buffer.o === "//") ret = {
						type_: "pu-frac",
						p1: b5.d,
						p2: b5.q
					};
					else {
						ret = b5.d;
						if (b5.d.length > 1 || b5.q.length > 1) ret.push({ type_: " / " });
						else ret.push({ type_: "/" });
						mhchemParser.concatArray(ret, b5.q);
					}
				} else ret = mhchemParser.go(buffer.d, "pu-2");
				for (var p in buffer) delete buffer[p];
				return ret;
			}
		}
	},
	"pu-2": {
		transitions: mhchemParser.createTransitions({
			"empty": { "*": { action_: "output" } },
			"*": { "*": {
				action_: ["output", "cdot"],
				nextState: "0"
			} },
			"\\x": { "*": { action_: "rm=" } },
			"space": { "*": {
				action_: ["output", "space"],
				nextState: "0"
			} },
			"^{(...)}|^(-1)": { "1": { action_: "^(-1)" } },
			"-9.,9": {
				"0": {
					action_: "rm=",
					nextState: "0"
				},
				"1": {
					action_: "^(-1)",
					nextState: "0"
				}
			},
			"{...}|else": { "*": {
				action_: "rm=",
				nextState: "1"
			} }
		}),
		actions: {
			"cdot": function() {
				return { type_: "tight cdot" };
			},
			"^(-1)": function(buffer, m) {
				buffer.rm += "^{" + m + "}";
			},
			"space": function() {
				return { type_: "pu-space-2" };
			},
			"output": function(buffer) {
				/** @type {ParserOutput | ParserOutput[]} */
				var ret = [];
				if (buffer.rm) {
					var mrm = mhchemParser.patterns.match_("{(...)}", buffer.rm || "");
					if (mrm && mrm.remainder === "") ret = mhchemParser.go(mrm.match_, "pu");
					else ret = {
						type_: "rm",
						p1: buffer.rm
					};
				}
				for (var p in buffer) delete buffer[p];
				return ret;
			}
		}
	},
	"pu-9,9": {
		transitions: mhchemParser.createTransitions({
			"empty": {
				"0": { action_: "output-0" },
				"o": { action_: "output-o" }
			},
			",": { "0": {
				action_: ["output-0", "comma"],
				nextState: "o"
			} },
			".": { "0": {
				action_: ["output-0", "copy"],
				nextState: "o"
			} },
			"else": { "*": { action_: "text=" } }
		}),
		actions: {
			"comma": function() {
				return { type_: "commaDecimal" };
			},
			"output-0": function(buffer) {
				/** @type {ParserOutput[]} */
				var ret = [];
				buffer.text_ = buffer.text_ || "";
				if (buffer.text_.length > 4) {
					var a = buffer.text_.length % 3;
					if (a === 0) a = 3;
					for (var i = buffer.text_.length - 3; i > 0; i -= 3) {
						ret.push(buffer.text_.substr(i, 3));
						ret.push({ type_: "1000 separator" });
					}
					ret.push(buffer.text_.substr(0, a));
					ret.reverse();
				} else ret.push(buffer.text_);
				for (var p in buffer) delete buffer[p];
				return ret;
			},
			"output-o": function(buffer) {
				/** @type {ParserOutput[]} */
				var ret = [];
				buffer.text_ = buffer.text_ || "";
				if (buffer.text_.length > 4) {
					var a = buffer.text_.length - 3;
					for (var i = 0; i < a; i += 3) {
						ret.push(buffer.text_.substr(i, 3));
						ret.push({ type_: "1000 separator" });
					}
					ret.push(buffer.text_.substr(i));
				} else ret.push(buffer.text_);
				for (var p in buffer) delete buffer[p];
				return ret;
			}
		}
	}
};
/** @type {Texify} */
var texify = {
	go: function(input, isInner) {
		if (!input) return "";
		var res = "";
		var cee = false;
		for (var i = 0; i < input.length; i++) {
			var inputi = input[i];
			if (typeof inputi === "string") res += inputi;
			else {
				res += texify._go2(inputi);
				if (inputi.type_ === "1st-level escape") cee = true;
			}
		}
		if (!isInner && !cee && res) res = "{" + res + "}";
		return res;
	},
	_goInner: function(input) {
		if (!input) return input;
		return texify.go(input, true);
	},
	_go2: function(buf) {
		/** @type {undefined | string} */
		var res;
		switch (buf.type_) {
			case "chemfive":
				res = "";
				var b5 = {
					a: texify._goInner(buf.a),
					b: texify._goInner(buf.b),
					p: texify._goInner(buf.p),
					o: texify._goInner(buf.o),
					q: texify._goInner(buf.q),
					d: texify._goInner(buf.d)
				};
				if (b5.a) {
					if (b5.a.match(/^[+\-]/)) b5.a = "{" + b5.a + "}";
					res += b5.a + "\\,";
				}
				if (b5.b || b5.p) {
					res += "{\\vphantom{X}}";
					res += "^{\\hphantom{" + (b5.b || "") + "}}_{\\hphantom{" + (b5.p || "") + "}}";
					res += "{\\vphantom{X}}";
					res += "^{\\vphantom{2}\\mathllap{" + (b5.b || "") + "}}";
					res += "_{\\vphantom{2}\\mathllap{" + (b5.p || "") + "}}";
				}
				if (b5.o) {
					if (b5.o.match(/^[+\-]/)) b5.o = "{" + b5.o + "}";
					res += b5.o;
				}
				if (buf.dType === "kv") {
					if (b5.d || b5.q) res += "{\\vphantom{X}}";
					if (b5.d) res += "^{" + b5.d + "}";
					if (b5.q) res += "_{" + b5.q + "}";
				} else if (buf.dType === "oxidation") {
					if (b5.d) {
						res += "{\\vphantom{X}}";
						res += "^{" + b5.d + "}";
					}
					if (b5.q) {
						res += "{{}}";
						res += "_{" + b5.q + "}";
					}
				} else {
					if (b5.q) {
						res += "{{}}";
						res += "_{" + b5.q + "}";
					}
					if (b5.d) {
						res += "{{}}";
						res += "^{" + b5.d + "}";
					}
				}
				break;
			case "rm":
				res = "\\mathrm{" + buf.p1 + "}";
				break;
			case "text":
				if (buf.p1.match(/[\^_]/)) {
					buf.p1 = buf.p1.replace(" ", "~").replace("-", "\\text{-}");
					res = "\\mathrm{" + buf.p1 + "}";
				} else res = "\\text{" + buf.p1 + "}";
				break;
			case "roman numeral":
				res = "\\mathrm{" + buf.p1 + "}";
				break;
			case "state of aggregation":
				res = "\\mskip2mu " + texify._goInner(buf.p1);
				break;
			case "state of aggregation subscript":
				res = "\\mskip1mu " + texify._goInner(buf.p1);
				break;
			case "bond":
				res = texify._getBond(buf.kind_);
				if (!res) throw ["MhchemErrorBond", "mhchem Error. Unknown bond type (" + buf.kind_ + ")"];
				break;
			case "frac":
				var c = "\\frac{" + buf.p1 + "}{" + buf.p2 + "}";
				res = "\\mathchoice{\\textstyle" + c + "}{" + c + "}{" + c + "}{" + c + "}";
				break;
			case "pu-frac":
				var d = "\\frac{" + texify._goInner(buf.p1) + "}{" + texify._goInner(buf.p2) + "}";
				res = "\\mathchoice{\\textstyle" + d + "}{" + d + "}{" + d + "}{" + d + "}";
				break;
			case "tex-math":
				res = buf.p1 + " ";
				break;
			case "frac-ce":
				res = "\\frac{" + texify._goInner(buf.p1) + "}{" + texify._goInner(buf.p2) + "}";
				break;
			case "overset":
				res = "\\overset{" + texify._goInner(buf.p1) + "}{" + texify._goInner(buf.p2) + "}";
				break;
			case "underset":
				res = "\\underset{" + texify._goInner(buf.p1) + "}{" + texify._goInner(buf.p2) + "}";
				break;
			case "underbrace":
				res = "\\underbrace{" + texify._goInner(buf.p1) + "}_{" + texify._goInner(buf.p2) + "}";
				break;
			case "color":
				res = "{\\color{" + buf.color1 + "}{" + texify._goInner(buf.color2) + "}}";
				break;
			case "color0":
				res = "\\color{" + buf.color + "}";
				break;
			case "arrow":
				var b6 = {
					rd: texify._goInner(buf.rd),
					rq: texify._goInner(buf.rq)
				};
				var arrow = texify._getArrow(buf.r);
				if (b6.rq) arrow += "[{\\rm " + b6.rq + "}]";
				if (b6.rd) arrow += "{\\rm " + b6.rd + "}";
				else arrow += "{}";
				res = arrow;
				break;
			case "operator":
				res = texify._getOperator(buf.kind_);
				break;
			case "1st-level escape":
				res = buf.p1 + " ";
				break;
			case "space":
				res = " ";
				break;
			case "entitySkip":
				res = "~";
				break;
			case "pu-space-1":
				res = "~";
				break;
			case "pu-space-2":
				res = "\\mkern3mu ";
				break;
			case "1000 separator":
				res = "\\mkern2mu ";
				break;
			case "commaDecimal":
				res = "{,}";
				break;
			case "comma enumeration L":
				res = "{" + buf.p1 + "}\\mkern6mu ";
				break;
			case "comma enumeration M":
				res = "{" + buf.p1 + "}\\mkern3mu ";
				break;
			case "comma enumeration S":
				res = "{" + buf.p1 + "}\\mkern1mu ";
				break;
			case "hyphen":
				res = "\\text{-}";
				break;
			case "addition compound":
				res = "\\,{\\cdot}\\,";
				break;
			case "electron dot":
				res = "\\mkern1mu \\text{\\textbullet}\\mkern1mu ";
				break;
			case "KV x":
				res = "{\\times}";
				break;
			case "prime":
				res = "\\prime ";
				break;
			case "cdot":
				res = "\\cdot ";
				break;
			case "tight cdot":
				res = "\\mkern1mu{\\cdot}\\mkern1mu ";
				break;
			case "times":
				res = "\\times ";
				break;
			case "circa":
				res = "{\\sim}";
				break;
			case "^":
				res = "uparrow";
				break;
			case "v":
				res = "downarrow";
				break;
			case "ellipsis":
				res = "\\ldots ";
				break;
			case "/":
				res = "/";
				break;
			case " / ":
				res = "\\,/\\,";
				break;
			default: throw ["MhchemBugT", "mhchem bug T. Please report."];
		}
		return res;
	},
	_getArrow: function(a) {
		switch (a) {
			case "->": return "\\yields";
			case "→": return "\\yields";
			case "⟶": return "\\yields";
			case "<-": return "\\yieldsLeft";
			case "<->": return "\\mesomerism";
			case "<-->": return "\\yieldsLeftRight";
			case "<=>": return "\\chemequilibrium";
			case "⇌": return "\\chemequilibrium";
			case "<=>>": return "\\equilibriumRight";
			case "<<=>": return "\\equilibriumLeft";
			default: throw ["MhchemBugT", "mhchem bug T. Please report."];
		}
	},
	_getBond: function(a) {
		switch (a) {
			case "-": return "{-}";
			case "1": return "{-}";
			case "=": return "{=}";
			case "2": return "{=}";
			case "#": return "{\\equiv}";
			case "3": return "{\\equiv}";
			case "~": return "{\\tripleDash}";
			case "~-": return "{\\tripleDashOverLine}";
			case "~=": return "{\\tripleDashOverDoubleLine}";
			case "~--": return "{\\tripleDashOverDoubleLine}";
			case "-~-": return "{\\tripleDashBetweenDoubleLine}";
			case "...": return "{{\\cdot}{\\cdot}{\\cdot}}";
			case "....": return "{{\\cdot}{\\cdot}{\\cdot}{\\cdot}}";
			case "->": return "{\\rightarrow}";
			case "<-": return "{\\leftarrow}";
			case "<": return "{<}";
			case ">": return "{>}";
			default: throw ["MhchemBugT", "mhchem bug T. Please report."];
		}
	},
	_getOperator: function(a) {
		switch (a) {
			case "+": return " {}+{} ";
			case "-": return " {}-{} ";
			case "=": return " {}={} ";
			case "<": return " {}<{} ";
			case ">": return " {}>{} ";
			case "<<": return " {}\\ll{} ";
			case ">>": return " {}\\gg{} ";
			case "\\pm": return " {}\\pm{} ";
			case "\\approx": return " {}\\approx{} ";
			case "$\\approx$": return " {}\\approx{} ";
			case "v": return " \\downarrow{} ";
			case "(v)": return " \\downarrow{} ";
			case "^": return " \\uparrow{} ";
			case "(^)": return " \\uparrow{} ";
			default: throw ["MhchemBugT", "mhchem bug T. Please report."];
		}
	}
};
defineMacro("\\darr", "\\downarrow");
defineMacro("\\dArr", "\\Downarrow");
defineMacro("\\Darr", "\\Downarrow");
defineMacro("\\lang", "\\langle");
defineMacro("\\rang", "\\rangle");
defineMacro("\\uarr", "\\uparrow");
defineMacro("\\uArr", "\\Uparrow");
defineMacro("\\Uarr", "\\Uparrow");
defineMacro("\\N", "\\mathbb{N}");
defineMacro("\\R", "\\mathbb{R}");
defineMacro("\\Z", "\\mathbb{Z}");
defineMacro("\\alef", "\\aleph");
defineMacro("\\alefsym", "\\aleph");
defineMacro("\\bull", "\\bullet");
defineMacro("\\clubs", "\\clubsuit");
defineMacro("\\cnums", "\\mathbb{C}");
defineMacro("\\Complex", "\\mathbb{C}");
defineMacro("\\Dagger", "\\ddagger");
defineMacro("\\diamonds", "\\diamondsuit");
defineMacro("\\empty", "\\emptyset");
defineMacro("\\exist", "\\exists");
defineMacro("\\harr", "\\leftrightarrow");
defineMacro("\\hArr", "\\Leftrightarrow");
defineMacro("\\Harr", "\\Leftrightarrow");
defineMacro("\\hearts", "\\heartsuit");
defineMacro("\\image", "\\Im");
defineMacro("\\infin", "\\infty");
defineMacro("\\isin", "\\in");
defineMacro("\\larr", "\\leftarrow");
defineMacro("\\lArr", "\\Leftarrow");
defineMacro("\\Larr", "\\Leftarrow");
defineMacro("\\lrarr", "\\leftrightarrow");
defineMacro("\\lrArr", "\\Leftrightarrow");
defineMacro("\\Lrarr", "\\Leftrightarrow");
defineMacro("\\natnums", "\\mathbb{N}");
defineMacro("\\plusmn", "\\pm");
defineMacro("\\rarr", "\\rightarrow");
defineMacro("\\rArr", "\\Rightarrow");
defineMacro("\\Rarr", "\\Rightarrow");
defineMacro("\\real", "\\Re");
defineMacro("\\reals", "\\mathbb{R}");
defineMacro("\\Reals", "\\mathbb{R}");
defineMacro("\\sdot", "\\cdot");
defineMacro("\\sect", "\\S");
defineMacro("\\spades", "\\spadesuit");
defineMacro("\\sub", "\\subset");
defineMacro("\\sube", "\\subseteq");
defineMacro("\\supe", "\\supseteq");
defineMacro("\\thetasym", "\\vartheta");
defineMacro("\\weierp", "\\wp");
/****************************************************
*
*  physics.js
*
*  Implements the Physics Package for LaTeX input.
*
*  ---------------------------------------------------------------------
*
*  The original version of this file is licensed as follows:
*  Copyright (c) 2015-2016 Kolen Cheung <https://github.com/ickc/MathJax-third-party-extensions>.
*
*  Licensed under the Apache License, Version 2.0 (the "License");
*  you may not use this file except in compliance with the License.
*  You may obtain a copy of the License at
*
*      http://www.apache.org/licenses/LICENSE-2.0
*
*  Unless required by applicable law or agreed to in writing, software
*  distributed under the License is distributed on an "AS IS" BASIS,
*  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
*  See the License for the specific language governing permissions and
*  limitations under the License.
*
*  ---------------------------------------------------------------------
*
*  This file has been revised from the original in the following ways:
*  1. The interface is changed so that it can be called from Temml, not MathJax.
*  2. \Re and \Im are not used, to avoid conflict with existing LaTeX letters.
*
*  This revision of the file is released under the MIT license.
*  https://mit-license.org/
*/
defineMacro("\\quantity", "{\\left\\{ #1 \\right\\}}");
defineMacro("\\qty", "{\\left\\{ #1 \\right\\}}");
defineMacro("\\pqty", "{\\left( #1 \\right)}");
defineMacro("\\bqty", "{\\left[ #1 \\right]}");
defineMacro("\\vqty", "{\\left\\vert #1 \\right\\vert}");
defineMacro("\\Bqty", "{\\left\\{ #1 \\right\\}}");
defineMacro("\\absolutevalue", "{\\left\\vert #1 \\right\\vert}");
defineMacro("\\abs", "{\\left\\vert #1 \\right\\vert}");
defineMacro("\\norm", "{\\left\\Vert #1 \\right\\Vert}");
defineMacro("\\evaluated", "{\\left.#1 \\right\\vert}");
defineMacro("\\eval", "{\\left.#1 \\right\\vert}");
defineMacro("\\order", "{\\mathcal{O} \\left( #1 \\right)}");
defineMacro("\\commutator", "{\\left[ #1 , #2 \\right]}");
defineMacro("\\comm", "{\\left[ #1 , #2 \\right]}");
defineMacro("\\anticommutator", "{\\left\\{ #1 , #2 \\right\\}}");
defineMacro("\\acomm", "{\\left\\{ #1 , #2 \\right\\}}");
defineMacro("\\poissonbracket", "{\\left\\{ #1 , #2 \\right\\}}");
defineMacro("\\pb", "{\\left\\{ #1 , #2 \\right\\}}");
defineMacro("\\vectorbold", "{\\boldsymbol{ #1 }}");
defineMacro("\\vb", "{\\boldsymbol{ #1 }}");
defineMacro("\\vectorarrow", "{\\vec{\\boldsymbol{ #1 }}}");
defineMacro("\\va", "{\\vec{\\boldsymbol{ #1 }}}");
defineMacro("\\vectorunit", "{{\\boldsymbol{\\hat{ #1 }}}}");
defineMacro("\\vu", "{{\\boldsymbol{\\hat{ #1 }}}}");
defineMacro("\\dotproduct", "\\mathbin{\\boldsymbol\\cdot}");
defineMacro("\\vdot", "{\\boldsymbol\\cdot}");
defineMacro("\\crossproduct", "\\mathbin{\\boldsymbol\\times}");
defineMacro("\\cross", "\\mathbin{\\boldsymbol\\times}");
defineMacro("\\cp", "\\mathbin{\\boldsymbol\\times}");
defineMacro("\\gradient", "{\\boldsymbol\\nabla}");
defineMacro("\\grad", "{\\boldsymbol\\nabla}");
defineMacro("\\divergence", "{\\grad\\vdot}");
defineMacro("\\curl", "{\\grad\\cross}");
defineMacro("\\laplacian", "\\nabla^2");
defineMacro("\\tr", "{\\operatorname{tr}}");
defineMacro("\\Tr", "{\\operatorname{Tr}}");
defineMacro("\\rank", "{\\operatorname{rank}}");
defineMacro("\\erf", "{\\operatorname{erf}}");
defineMacro("\\Res", "{\\operatorname{Res}}");
defineMacro("\\principalvalue", "{\\mathcal{P}}");
defineMacro("\\pv", "{\\mathcal{P}}");
defineMacro("\\PV", "{\\operatorname{P.V.}}");
defineMacro("\\qqtext", "{\\quad\\text{ #1 }\\quad}");
defineMacro("\\qq", "{\\quad\\text{ #1 }\\quad}");
defineMacro("\\qcomma", "{\\text{,}\\quad}");
defineMacro("\\qc", "{\\text{,}\\quad}");
defineMacro("\\qcc", "{\\quad\\text{c.c.}\\quad}");
defineMacro("\\qif", "{\\quad\\text{if}\\quad}");
defineMacro("\\qthen", "{\\quad\\text{then}\\quad}");
defineMacro("\\qelse", "{\\quad\\text{else}\\quad}");
defineMacro("\\qotherwise", "{\\quad\\text{otherwise}\\quad}");
defineMacro("\\qunless", "{\\quad\\text{unless}\\quad}");
defineMacro("\\qgiven", "{\\quad\\text{given}\\quad}");
defineMacro("\\qusing", "{\\quad\\text{using}\\quad}");
defineMacro("\\qassume", "{\\quad\\text{assume}\\quad}");
defineMacro("\\qsince", "{\\quad\\text{since}\\quad}");
defineMacro("\\qlet", "{\\quad\\text{let}\\quad}");
defineMacro("\\qfor", "{\\quad\\text{for}\\quad}");
defineMacro("\\qall", "{\\quad\\text{all}\\quad}");
defineMacro("\\qeven", "{\\quad\\text{even}\\quad}");
defineMacro("\\qodd", "{\\quad\\text{odd}\\quad}");
defineMacro("\\qinteger", "{\\quad\\text{integer}\\quad}");
defineMacro("\\qand", "{\\quad\\text{and}\\quad}");
defineMacro("\\qor", "{\\quad\\text{or}\\quad}");
defineMacro("\\qas", "{\\quad\\text{as}\\quad}");
defineMacro("\\qin", "{\\quad\\text{in}\\quad}");
defineMacro("\\differential", "{\\text{d}}");
defineMacro("\\dd", "{\\text{d}}");
defineMacro("\\derivative", "{\\frac{\\text{d}{ #1 }}{\\text{d}{ #2 }}}");
defineMacro("\\dv", "{\\frac{\\text{d}{ #1 }}{\\text{d}{ #2 }}}");
defineMacro("\\partialderivative", "{\\frac{\\partial{ #1 }}{\\partial{ #2 }}}");
defineMacro("\\variation", "{\\delta}");
defineMacro("\\var", "{\\delta}");
defineMacro("\\functionalderivative", "{\\frac{\\delta{ #1 }}{\\delta{ #2 }}}");
defineMacro("\\fdv", "{\\frac{\\delta{ #1 }}{\\delta{ #2 }}}");
defineMacro("\\innerproduct", "{\\left\\langle {#1} \\mid { #2} \\right\\rangle}");
defineMacro("\\outerproduct", "{\\left\\vert { #1 } \\right\\rangle\\left\\langle { #2} \\right\\vert}");
defineMacro("\\dyad", "{\\left\\vert { #1 } \\right\\rangle\\left\\langle { #2} \\right\\vert}");
defineMacro("\\ketbra", "{\\left\\vert { #1 } \\right\\rangle\\left\\langle { #2} \\right\\vert}");
defineMacro("\\op", "{\\left\\vert { #1 } \\right\\rangle\\left\\langle { #2} \\right\\vert}");
defineMacro("\\expectationvalue", "{\\left\\langle {#1 } \\right\\rangle}");
defineMacro("\\expval", "{\\left\\langle {#1 } \\right\\rangle}");
defineMacro("\\ev", "{\\left\\langle {#1 } \\right\\rangle}");
defineMacro("\\matrixelement", "{\\left\\langle{ #1 }\\right\\vert{ #2 }\\left\\vert{#3}\\right\\rangle}");
defineMacro("\\matrixel", "{\\left\\langle{ #1 }\\right\\vert{ #2 }\\left\\vert{#3}\\right\\rangle}");
defineMacro("\\mel", "{\\left\\langle{ #1 }\\right\\vert{ #2 }\\left\\vert{#3}\\right\\rangle}");
function getHLines(parser) {
	const hlineInfo = [];
	parser.consumeSpaces();
	let nxt = parser.fetch().text;
	if (nxt === "\\relax") {
		parser.consume();
		parser.consumeSpaces();
		nxt = parser.fetch().text;
	}
	while (nxt === "\\hline" || nxt === "\\hdashline") {
		parser.consume();
		hlineInfo.push(nxt === "\\hdashline");
		parser.consumeSpaces();
		nxt = parser.fetch().text;
	}
	return hlineInfo;
}
const validateAmsEnvironmentContext = (context) => {
	if (!context.parser.settings.displayMode) throw new ParseError(`{${context.envName}} can be used only in display mode.`);
};
const sizeRegEx$1 = /([-+]?) *(\d+(?:\.\d*)?|\.\d+) *([a-z]{2})/;
const arrayGaps = (macros) => {
	let arraystretch = macros.get("\\arraystretch");
	if (typeof arraystretch !== "string") arraystretch = stringFromArg(arraystretch.tokens);
	arraystretch = isNaN(arraystretch) ? null : Number(arraystretch);
	let arraycolsepStr = macros.get("\\arraycolsep");
	if (typeof arraycolsepStr !== "string") arraycolsepStr = stringFromArg(arraycolsepStr.tokens);
	const match = sizeRegEx$1.exec(arraycolsepStr);
	const arraycolsep = match ? {
		number: +(match[1] + match[2]),
		unit: match[3]
	} : null;
	return [arraystretch, arraycolsep];
};
const checkCellForLabels = (cell) => {
	let rowLabel = "";
	for (let i = 0; i < cell.length; i++) if (cell[i].type === "label") {
		if (rowLabel) throw new ParseError("Multiple \\labels in one row");
		rowLabel = cell[i].string;
	}
	return rowLabel;
};
function getAutoTag(name) {
	if (name.indexOf("ed") === -1) return name.indexOf("*") === -1;
}
/**
* Parse the body of the environment, with rows delimited by \\ and
* columns delimited by &, and create a nested list in row-major order
* with one group per cell.  If given an optional argument scriptLevel
* ("text", "display", etc.), then each cell is cast into that scriptLevel.
*/
function parseArray(parser, { cols, envClasses, autoTag, singleRow, emptySingleRow, maxNumCols, leqno, arraystretch, arraycolsep }, scriptLevel) {
	const endToken = envClasses && envClasses.includes("bordermatrix") ? "}" : "\\end";
	parser.gullet.beginGroup();
	if (!singleRow) parser.gullet.macros.set("\\cr", "\\\\\\relax");
	parser.gullet.beginGroup();
	let row = [];
	const body = [row];
	const rowGaps = [];
	const labels = [];
	const hLinesBeforeRow = [];
	const tags = autoTag != null ? [] : void 0;
	function beginRow() {
		if (autoTag) parser.gullet.macros.set("\\@eqnsw", "1", true);
	}
	function endRow() {
		if (tags) {
			if (parser.gullet.macros.get("\\df@tag")) {
				tags.push(parser.subparse([new Token("\\df@tag")]));
				parser.gullet.macros.set("\\df@tag", void 0, true);
			} else tags.push(Boolean(autoTag) && parser.gullet.macros.get("\\@eqnsw") === "1");
		}
	}
	beginRow();
	hLinesBeforeRow.push(getHLines(parser));
	while (true) {
		let cell = parser.parseExpression(false, singleRow ? "\\end" : "\\\\");
		parser.gullet.endGroup();
		parser.gullet.beginGroup();
		cell = {
			type: "ordgroup",
			mode: parser.mode,
			body: cell,
			semisimple: true
		};
		row.push(cell);
		const next = parser.fetch().text;
		if (next === "&") {
			if (maxNumCols && row.length === maxNumCols) {
				if (envClasses.includes("array")) {
					if (parser.settings.strict) throw new ParseError("Too few columns specified in the {array} column argument.", parser.nextToken);
				} else if (maxNumCols === 2) throw new ParseError("The split environment accepts no more than two columns", parser.nextToken);
				else throw new ParseError("The equation environment accepts only one column", parser.nextToken);
			}
			parser.consume();
		} else if (next === endToken) {
			endRow();
			if (row.length === 1 && cell.body.length === 0 && (body.length > 1 || !emptySingleRow)) body.pop();
			labels.push(checkCellForLabels(cell.body));
			if (hLinesBeforeRow.length < body.length + 1) hLinesBeforeRow.push([]);
			break;
		} else if (next === "\\\\") {
			parser.consume();
			let size;
			if (parser.gullet.future().text !== " ") size = parser.parseSizeGroup(true);
			rowGaps.push(size ? size.value : null);
			endRow();
			labels.push(checkCellForLabels(cell.body));
			hLinesBeforeRow.push(getHLines(parser));
			row = [];
			body.push(row);
			beginRow();
		} else throw new ParseError("Expected & or \\\\ or \\cr or " + endToken, parser.nextToken);
	}
	parser.gullet.endGroup();
	parser.gullet.endGroup();
	return {
		type: "array",
		mode: parser.mode,
		body,
		cols,
		rowGaps,
		hLinesBeforeRow,
		envClasses,
		autoTag,
		scriptLevel,
		tags,
		labels,
		leqno,
		arraystretch,
		arraycolsep
	};
}
function dCellStyle(envName) {
	return envName.slice(0, 1) === "d" ? "display" : "text";
}
const alignMap = {
	c: "center ",
	l: "left ",
	r: "right "
};
const glue = (group) => {
	const glueNode = new MathNode("mtd", []);
	glueNode.style = {
		padding: "0",
		width: "50%"
	};
	if (group.envClasses.includes("multline")) glueNode.style.width = "7.5%";
	return glueNode;
};
const mathmlBuilder$9 = function(group, style) {
	const tbl = [];
	const numRows = group.body.length;
	const hlines = group.hLinesBeforeRow;
	const tagIsPresent = group.tags && group.tags.some((tag) => tag);
	for (let i = 0; i < numRows; i++) {
		const rw = group.body[i];
		const row = [];
		const cellLevel = group.scriptLevel === "text" ? StyleLevel.TEXT : group.scriptLevel === "script" ? StyleLevel.SCRIPT : StyleLevel.DISPLAY;
		for (let j = 0; j < rw.length; j++) {
			const mtd = new MathNode("mtd", [buildGroup$1(rw[j], style.withLevel(cellLevel))]);
			if (group.envClasses.includes("multline")) {
				const align = i === 0 ? "left" : i === numRows - 1 ? "right" : "center";
				if (align !== "center") mtd.classes.push("tml-" + align);
			}
			row.push(mtd);
		}
		const numColumns = group.body[0].length;
		for (let k = 0; k < numColumns - rw.length; k++) row.push(new MathNode("mtd", [], [], style));
		if (tagIsPresent) {
			const tag = group.tags[i];
			let tagElement;
			if (tag === true) tagElement = new MathNode("mtext", [new Span(["tml-eqn"])]);
			else if (tag === false) tagElement = new MathNode("mtext", [], []);
			else {
				tagElement = buildExpressionRow(tag[0].body, style.withLevel(cellLevel), true);
				tagElement = consolidateText(tagElement);
				tagElement.classes = ["tml-tag"];
			}
			if (tagElement) {
				row.unshift(glue(group));
				row.push(glue(group));
				if (group.leqno) row[0].children.push(tagElement);
				else row[row.length - 1].children.push(tagElement);
			}
		}
		const mtr = new MathNode("mtr", row, []);
		const label = group.labels.shift();
		if (label && group.tags && group.tags[i]) {
			mtr.setAttribute("id", label);
			if (Array.isArray(group.tags[i])) mtr.classes.push("tml-tageqn");
		}
		if (i === 0 && hlines[0].length > 0) {
			if (hlines[0].length === 2) mtr.children.forEach((cell) => {
				cell.style.borderTop = "0.15em double";
			});
			else mtr.children.forEach((cell) => {
				cell.style.borderTop = hlines[0][0] ? "0.06em dashed" : "0.06em solid";
			});
		}
		if (hlines[i + 1].length > 0) {
			if (hlines[i + 1].length === 2) mtr.children.forEach((cell) => {
				cell.style.borderBottom = "0.15em double";
			});
			else mtr.children.forEach((cell) => {
				cell.style.borderBottom = hlines[i + 1][0] ? "0.06em dashed" : "0.06em solid";
			});
		}
		let mustSquashRow = true;
		for (let j = 0; j < mtr.children.length; j++) {
			const child = mtr.children[j].children[0];
			if (!(child && child.type === "mpadded" && child.attributes.height === "0px")) {
				mustSquashRow = false;
				break;
			}
		}
		if (mustSquashRow) {
			mtr.classes.push("ff-squash");
			for (let j = 0; j < mtr.children.length; j++) {
				mtr.children[j].style.paddingTop = "0";
				mtr.children[j].style.paddingBottom = "0";
			}
		}
		tbl.push(mtr);
	}
	if (group.arraystretch && group.arraystretch !== 1) {
		const pad = String(1.4 * group.arraystretch - .8) + "ex";
		for (let i = 0; i < tbl.length; i++) for (let j = 0; j < tbl[i].children.length; j++) {
			tbl[i].children[j].style.paddingTop = pad;
			tbl[i].children[j].style.paddingBottom = pad;
		}
	}
	let sidePadding;
	let sidePadUnit;
	if (group.envClasses.length > 0) {
		sidePadding = group.envClasses.includes("abut") ? "0" : group.envClasses.includes("cases") ? "0" : group.envClasses.includes("small") ? "0.1389" : group.envClasses.includes("cd") ? "0.25" : "0.4";
		sidePadUnit = "em";
	}
	if (group.arraycolsep) {
		const arraySidePad = calculateSize(group.arraycolsep, style);
		sidePadding = arraySidePad.number.toFixed(4);
		sidePadUnit = arraySidePad.unit;
	}
	if (sidePadding) {
		const numCols = tbl.length === 0 ? 0 : tbl[0].children.length;
		const sidePad = (j, hand) => {
			if (j === 0 && hand === 0) return "0";
			if (j === numCols - 1 && hand === 1) return "0";
			if (group.envClasses[0] !== "align") return sidePadding;
			if (hand === 1) return "0";
			if (tagIsPresent) return j % 2 ? "1" : "0";
			else return j % 2 ? "0" : "1";
		};
		for (let i = 0; i < tbl.length; i++) for (let j = 0; j < tbl[i].children.length; j++) {
			tbl[i].children[j].style.paddingLeft = `${sidePad(j, 0)}${sidePadUnit}`;
			tbl[i].children[j].style.paddingRight = `${sidePad(j, 1)}${sidePadUnit}`;
		}
	}
	if (group.envClasses.length === 0) for (let i = 0; i < tbl.length; i++) {
		tbl[i].children[0].style.paddingLeft = "0em";
		if (tbl[i].children.length === tbl[0].children.length) tbl[i].children[tbl[i].children.length - 1].style.paddingRight = "0em";
	}
	if (group.envClasses.length > 0) {
		const align = group.envClasses.includes("align") || group.envClasses.includes("alignat");
		for (let i = 0; i < tbl.length; i++) {
			const row = tbl[i];
			if (align) {
				for (let j = 0; j < row.children.length; j++) row.children[j].classes = ["tml-" + (j % 2 ? "left" : "right")];
				if (tagIsPresent) {
					const k = group.leqno ? 0 : row.children.length - 1;
					row.children[k].classes = [];
				}
			}
			if (row.children.length > 1 && group.envClasses.includes("cases")) row.children[1].style.paddingLeft = "1em";
			if (group.envClasses.includes("cases") || group.envClasses.includes("subarray")) for (const cell of row.children) cell.classes.push("tml-left");
		}
	}
	let table = new MathNode("mtable", tbl);
	if (group.envClasses.length > 0) {
		if (group.envClasses.includes("jot")) table.classes.push("tml-jot");
		else if (group.envClasses.includes("small")) table.classes.push("tml-small");
	}
	if (group.scriptLevel === "display") table.setAttribute("displaystyle", "true");
	if (group.autoTag || group.envClasses.includes("multline")) table.style.width = "100%";
	if (group.cols && group.cols.length > 0) {
		const cols = group.cols;
		let prevTypeWasAlign = false;
		let iStart = 0;
		let iEnd = cols.length;
		while (cols[iStart].type === "separator") iStart += 1;
		while (cols[iEnd - 1].type === "separator") iEnd -= 1;
		if (cols[0].type === "separator") {
			const sep = cols[1].type === "separator" ? "0.15em double" : cols[0].separator === "|" ? "0.06em solid " : "0.06em dashed ";
			for (const row of table.children) row.children[0].style.borderLeft = sep;
		}
		let iCol = tagIsPresent ? 0 : -1;
		for (let i = iStart; i < iEnd; i++) if (cols[i].type === "align") {
			const colAlign = alignMap[cols[i].align];
			iCol += 1;
			for (const row of table.children) if (colAlign.trim() !== "center" && iCol < row.children.length) row.children[iCol].classes = ["tml-" + colAlign.trim()];
			prevTypeWasAlign = true;
		} else if (cols[i].type === "separator") {
			if (prevTypeWasAlign) {
				const sep = cols[i + 1].type === "separator" ? "0.15em double" : cols[i].separator === "|" ? "0.06em solid" : "0.06em dashed";
				for (const row of table.children) if (iCol < row.children.length) row.children[iCol].style.borderRight = sep;
			}
			prevTypeWasAlign = false;
		}
		if (cols[cols.length - 1].type === "separator") {
			const sep = cols[cols.length - 2].type === "separator" ? "0.15em double" : cols[cols.length - 1].separator === "|" ? "0.06em solid" : "0.06em dashed";
			for (const row of table.children) {
				row.children[row.children.length - 1].style.borderRight = sep;
				row.children[row.children.length - 1].style.paddingRight = "0.4em";
			}
		}
	}
	if (group.envClasses.includes("small")) {
		table = new MathNode("mstyle", [table]);
		table.setAttribute("scriptlevel", "1");
	}
	return table;
};
const alignedHandler = function(context, args) {
	if (context.envName.indexOf("ed") === -1) validateAmsEnvironmentContext(context);
	const isSplit = context.envName === "split";
	const cols = [];
	const res = parseArray(context.parser, {
		cols,
		emptySingleRow: true,
		autoTag: isSplit ? void 0 : getAutoTag(context.envName),
		envClasses: ["abut", "jot"],
		maxNumCols: context.envName === "split" ? 2 : void 0,
		leqno: context.parser.settings.leqno
	}, "display");
	let numMaths;
	let numCols = 0;
	const isAlignedAt = context.envName.indexOf("at") > -1;
	if (args[0] && isAlignedAt) {
		let arg0 = "";
		for (let i = 0; i < args[0].body.length; i++) {
			const textord = assertNodeType(args[0].body[i], "textord");
			arg0 += textord.text;
		}
		if (isNaN(arg0)) throw new ParseError("The alignat enviroment requires a numeric first argument.");
		numMaths = Number(arg0);
		numCols = numMaths * 2;
	}
	res.body.forEach(function(row) {
		if (isAlignedAt) {
			const curMaths = row.length / 2;
			if (numMaths < curMaths) throw new ParseError(`Too many math in a row: expected ${numMaths}, but got ${curMaths}`, row[0]);
		} else if (numCols < row.length) numCols = row.length;
	});
	for (let i = 0; i < numCols; ++i) {
		let align = "r";
		if (i % 2 === 1) align = "l";
		cols[i] = {
			type: "align",
			align
		};
	}
	if (context.envName === "split");
	else if (isAlignedAt) res.envClasses.push("alignat");
	else res.envClasses[0] = "align";
	return res;
};
defineEnvironment({
	type: "array",
	names: ["array", "darray"],
	props: { numArgs: 1 },
	handler(context, args) {
		const cols = (checkSymbolNodeType(args[0]) ? [args[0]] : assertNodeType(args[0], "ordgroup").body).map(function(nde) {
			const ca = assertSymbolNodeType(nde).text;
			if ("lcr".indexOf(ca) !== -1) return {
				type: "align",
				align: ca
			};
			else if (ca === "|") return {
				type: "separator",
				separator: "|"
			};
			else if (ca === ":") return {
				type: "separator",
				separator: ":"
			};
			throw new ParseError("Unknown column alignment: " + ca, nde);
		});
		const [arraystretch, arraycolsep] = arrayGaps(context.parser.gullet.macros);
		const res = {
			cols,
			envClasses: ["array"],
			maxNumCols: cols.length,
			arraystretch,
			arraycolsep
		};
		return parseArray(context.parser, res, dCellStyle(context.envName));
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: [
		"matrix",
		"pmatrix",
		"bmatrix",
		"Bmatrix",
		"vmatrix",
		"Vmatrix",
		"matrix*",
		"pmatrix*",
		"bmatrix*",
		"Bmatrix*",
		"vmatrix*",
		"Vmatrix*"
	],
	props: { numArgs: 0 },
	handler(context) {
		const delimiters = {
			matrix: null,
			pmatrix: ["(", ")"],
			bmatrix: ["[", "]"],
			Bmatrix: ["\\{", "\\}"],
			vmatrix: ["|", "|"],
			Vmatrix: ["\\Vert", "\\Vert"]
		}[context.envName.replace("*", "")];
		let colAlign = "c";
		const payload = {
			envClasses: [],
			cols: []
		};
		if (context.envName.charAt(context.envName.length - 1) === "*") {
			const parser = context.parser;
			parser.consumeSpaces();
			if (parser.fetch().text === "[") {
				parser.consume();
				parser.consumeSpaces();
				colAlign = parser.fetch().text;
				if ("lcr".indexOf(colAlign) === -1) throw new ParseError("Expected l or c or r", parser.nextToken);
				parser.consume();
				parser.consumeSpaces();
				parser.expect("]");
				parser.consume();
				payload.cols = [];
			}
		}
		const res = parseArray(context.parser, payload, "text");
		res.cols = res.body.length > 0 ? new Array(res.body[0].length).fill({
			type: "align",
			align: colAlign
		}) : [];
		const [arraystretch, arraycolsep] = arrayGaps(context.parser.gullet.macros);
		res.arraystretch = arraystretch;
		if (arraycolsep && !(arraycolsep === 6 && arraycolsep === "pt")) res.arraycolsep = arraycolsep;
		return delimiters ? {
			type: "leftright",
			mode: context.mode,
			body: [res],
			left: delimiters[0],
			right: delimiters[1],
			rightColor: void 0
		} : res;
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["bordermatrix"],
	props: { numArgs: 0 },
	handler(context) {
		const res = parseArray(context.parser, {
			cols: [],
			envClasses: ["bordermatrix"]
		}, "text");
		res.cols = res.body.length > 0 ? new Array(res.body[0].length).fill({
			type: "align",
			align: "c"
		}) : [];
		res.envClasses = [];
		res.arraystretch = 1;
		if (context.envName === "matrix") return res;
		return bordermatrixParseTree(res, context.delimiters);
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["smallmatrix"],
	props: { numArgs: 0 },
	handler(context) {
		return parseArray(context.parser, { envClasses: ["small"] }, "script");
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["subarray"],
	props: { numArgs: 1 },
	handler(context, args) {
		const cols = (checkSymbolNodeType(args[0]) ? [args[0]] : assertNodeType(args[0], "ordgroup").body).map(function(nde) {
			const ca = assertSymbolNodeType(nde).text;
			if ("lc".indexOf(ca) !== -1) return {
				type: "align",
				align: ca
			};
			throw new ParseError("Unknown column alignment: " + ca, nde);
		});
		if (cols.length > 1) throw new ParseError("{subarray} can contain only one column");
		let res = {
			cols,
			envClasses: ["small"]
		};
		res = parseArray(context.parser, res, "script");
		if (res.body.length > 0 && res.body[0].length > 1) throw new ParseError("{subarray} can contain only one column");
		return res;
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: [
		"cases",
		"dcases",
		"rcases",
		"drcases"
	],
	props: { numArgs: 0 },
	handler(context) {
		const res = parseArray(context.parser, {
			cols: [],
			envClasses: ["cases"]
		}, dCellStyle(context.envName));
		return {
			type: "leftright",
			mode: context.mode,
			body: [res],
			left: context.envName.indexOf("r") > -1 ? "." : "\\{",
			right: context.envName.indexOf("r") > -1 ? "\\}" : ".",
			rightColor: void 0
		};
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: [
		"align",
		"align*",
		"aligned",
		"split"
	],
	props: { numArgs: 0 },
	handler: alignedHandler,
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: [
		"alignat",
		"alignat*",
		"alignedat"
	],
	props: { numArgs: 1 },
	handler: alignedHandler,
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: [
		"gathered",
		"gather",
		"gather*"
	],
	props: { numArgs: 0 },
	handler(context) {
		if (context.envName !== "gathered") validateAmsEnvironmentContext(context);
		const res = {
			cols: [],
			envClasses: ["abut", "jot"],
			autoTag: getAutoTag(context.envName),
			emptySingleRow: true,
			leqno: context.parser.settings.leqno
		};
		return parseArray(context.parser, res, "display");
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["equation", "equation*"],
	props: { numArgs: 0 },
	handler(context) {
		validateAmsEnvironmentContext(context);
		const res = {
			autoTag: getAutoTag(context.envName),
			emptySingleRow: true,
			singleRow: true,
			maxNumCols: 1,
			envClasses: ["align"],
			leqno: context.parser.settings.leqno
		};
		return parseArray(context.parser, res, "display");
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["multline", "multline*"],
	props: { numArgs: 0 },
	handler(context) {
		validateAmsEnvironmentContext(context);
		const res = {
			autoTag: context.envName === "multline",
			maxNumCols: 1,
			envClasses: ["jot", "multline"],
			leqno: context.parser.settings.leqno
		};
		return parseArray(context.parser, res, "display");
	},
	mathmlBuilder: mathmlBuilder$9
});
defineEnvironment({
	type: "array",
	names: ["CD"],
	props: { numArgs: 0 },
	handler(context) {
		validateAmsEnvironmentContext(context);
		return parseCD(context.parser);
	},
	mathmlBuilder: mathmlBuilder$9
});
defineFunction({
	type: "text",
	names: ["\\hline", "\\hdashline"],
	props: {
		numArgs: 0,
		allowedInText: true,
		allowedInMath: true
	},
	handler(context, args) {
		throw new ParseError(`${context.funcName} valid only within array environment`);
	}
});
const environments = _environments;
defineFunction({
	type: "bordermatrix",
	names: ["\\bordermatrix", "\\matrix"],
	props: {
		numArgs: 0,
		numOptionalArgs: 1
	},
	handler: ({ parser, funcName }, args, optArgs) => {
		let delimiters = ["(", ")"];
		if (funcName === "\\bordermatrix" && optArgs[0] && optArgs[0].body) {
			const body = optArgs[0].body;
			if (body.length === 1 && body[0].type === "delimiter") delimiters = [body[0].left, body[0].right];
		}
		parser.consumeSpaces();
		parser.consume();
		const env = environments["bordermatrix"];
		const context = {
			mode: parser.mode,
			envName: funcName.slice(1),
			delimiters,
			parser
		};
		const result = env.handler(context);
		parser.expect("}", true);
		return result;
	}
});
defineFunction({
	type: "cancelto",
	names: ["\\cancelto"],
	props: { numArgs: 2 },
	handler({ parser }, args) {
		const to = args[0];
		const body = args[1];
		return {
			type: "cancelto",
			mode: parser.mode,
			body,
			to,
			isCharacterBox: isCharacterBox(body)
		};
	},
	mathmlBuilder(group, style) {
		const fromNode = new MathNode("mrow", [buildGroup$1(group.body, style)], ["ff-narrow"]);
		const phantom = new MathNode("mphantom", [buildGroup$1(group.body, style)]);
		const arrow = new MathNode("mrow", [phantom], ["tml-cancelto"]);
		arrow.style.color = style.color;
		if (group.isCharacterBox && smalls.indexOf(group.body.body[0].text) > -1) {
			arrow.style.left = "0.1em";
			arrow.style.width = "90%";
		}
		const node = new MathNode("mrow", [fromNode, arrow], ["menclose"]);
		if (!group.isCharacterBox || /[f∫∑]/.test(group.body.body[0].text)) phantom.style.paddingRight = "0.2em";
		else {
			phantom.style.padding = "0.5ex 0.1em 0 0";
			const strut = new MathNode("mspace", []);
			strut.setAttribute("height", "0.85em");
			fromNode.children.push(strut);
		}
		let dummyNode;
		if (group.isCharacterBox) {
			dummyNode = new MathNode("mspace", []);
			dummyNode.setAttribute("height", "1em");
		} else {
			const zeroWidthNode = new MathNode("mpadded", [buildGroup$1(group.body, style)]);
			zeroWidthNode.setAttribute("width", "0.1px");
			dummyNode = new MathNode("mphantom", [zeroWidthNode]);
		}
		const toNode = buildGroup$1(group.to, style);
		toNode.style.color = style.color;
		const zeroWidthToNode = new MathNode("mpadded", [toNode]);
		if (!group.isCharacterBox || /[f∫∑]/.test(group.body.body[0].text)) {
			const w = new MathNode("mspace", []);
			w.setAttribute("width", "0.2em");
			zeroWidthToNode.children.unshift(w);
		}
		zeroWidthToNode.setAttribute("width", "0.1px");
		const mover = new MathNode("mover", [dummyNode, zeroWidthToNode]);
		const nudgeLeft = new MathNode("mrow", [], ["ff-nudge-left"]);
		return newDocumentFragment([makeRow([node, mover]), nudgeLeft]);
	}
});
defineFunction({
	type: "textord",
	names: ["\\@char"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler({ parser, token }, args) {
		const group = assertNodeType(args[0], "ordgroup").body;
		let number = "";
		for (let i = 0; i < group.length; i++) {
			const node = assertNodeType(group[i], "textord");
			number += node.text;
		}
		const code = parseInt(number);
		if (isNaN(code)) throw new ParseError(`\\@char has non-numeric argument ${number}`, token);
		return {
			type: "textord",
			mode: parser.mode,
			text: String.fromCodePoint(code)
		};
	}
});
const htmlRegEx = /^(#[a-f0-9]{3}|#?[a-f0-9]{6})$/i;
const htmlOrNameRegEx = /^(#[a-f0-9]{3}|#?[a-f0-9]{6}|[a-z]+)$/i;
const RGBregEx = /^ *\d{1,3} *(?:, *\d{1,3} *){2}$/;
const rgbRegEx = /^ *[10](?:\.\d*)? *(?:, *[10](?:\.\d*)? *){2}$/;
const xcolorHtmlRegEx = /^[a-f0-9]{6}$/i;
const toHex = (num) => {
	let str = num.toString(16);
	if (str.length === 1) str = "0" + str;
	return str;
};
const xcolors = JSON.parse(`{
  "Apricot": "#ffb484",
  "Aquamarine": "#08b4bc",
  "Bittersweet": "#c84c14",
  "blue": "#0000FF",
  "Blue": "#303494",
  "BlueGreen": "#08b4bc",
  "BlueViolet": "#503c94",
  "BrickRed": "#b8341c",
  "brown": "#BF8040",
  "Brown": "#802404",
  "BurntOrange": "#f8941c",
  "CadetBlue": "#78749c",
  "CarnationPink": "#f884b4",
  "Cerulean": "#08a4e4",
  "CornflowerBlue": "#40ace4",
  "cyan": "#00FFFF",
  "Cyan": "#08acec",
  "Dandelion": "#ffbc44",
  "darkgray": "#404040",
  "DarkOrchid": "#a8548c",
  "Emerald": "#08ac9c",
  "ForestGreen": "#089c54",
  "Fuchsia": "#90348c",
  "Goldenrod": "#ffdc44",
  "gray": "#808080",
  "Gray": "#98949c",
  "green": "#00FF00",
  "Green": "#08a44c",
  "GreenYellow": "#e0e474",
  "JungleGreen": "#08ac9c",
  "Lavender": "#f89cc4",
  "lightgray": "#c0c0c0",
  "lime": "#BFFF00",
  "LimeGreen": "#90c43c",
  "magenta": "#FF00FF",
  "Magenta": "#f0048c",
  "Mahogany": "#b0341c",
  "Maroon": "#b03434",
  "Melon": "#f89c7c",
  "MidnightBlue": "#086494",
  "Mulberry": "#b03c94",
  "NavyBlue": "#086cbc",
  "olive": "#7F7F00",
  "OliveGreen": "#407c34",
  "orange": "#FF8000",
  "Orange": "#f8843c",
  "OrangeRed": "#f0145c",
  "Orchid": "#b074ac",
  "Peach": "#f8945c",
  "Periwinkle": "#8074bc",
  "PineGreen": "#088c74",
  "pink": "#ff7f7f",
  "Plum": "#98248c",
  "ProcessBlue": "#08b4ec",
  "purple": "#BF0040",
  "Purple": "#a0449c",
  "RawSienna": "#983c04",
  "red": "#ff0000",
  "Red": "#f01c24",
  "RedOrange": "#f86434",
  "RedViolet": "#a0246c",
  "Rhodamine": "#f0549c",
  "Royallue": "#0874bc",
  "RoyalPurple": "#683c9c",
  "RubineRed": "#f0047c",
  "Salmon": "#f8948c",
  "SeaGreen": "#30bc9c",
  "Sepia": "#701404",
  "SkyBlue": "#48c4dc",
  "SpringGreen": "#c8dc64",
  "Tan": "#e09c74",
  "teal": "#007F7F",
  "TealBlue": "#08acb4",
  "Thistle": "#d884b4",
  "Turquoise": "#08b4cc",
  "violet": "#800080",
  "Violet": "#60449c",
  "VioletRed": "#f054a4",
  "WildStrawberry": "#f0246c",
  "yellow": "#FFFF00",
  "Yellow": "#fff404",
  "YellowGreen": "#98cc6c",
  "YellowOrange": "#ffa41c"
}`);
const colorFromSpec = (model, spec) => {
	let color = "";
	if (model === "HTML") {
		if (!htmlRegEx.test(spec)) throw new ParseError("Invalid HTML input.");
		color = spec;
	} else if (model === "RGB") {
		if (!RGBregEx.test(spec)) throw new ParseError("Invalid RGB input.");
		spec.split(",").map((e) => {
			color += toHex(Number(e.trim()));
		});
	} else {
		if (!rgbRegEx.test(spec)) throw new ParseError("Invalid rbg input.");
		spec.split(",").map((e) => {
			const num = Number(e.trim());
			if (num > 1) throw new ParseError("Color rgb input must be < 1.");
			color += toHex(Number((num * 255).toFixed(0)));
		});
	}
	if (color.charAt(0) !== "#") color = "#" + color;
	return color;
};
const validateColor = (color, macros, token) => {
	const macroName = `\\\\color@${color}`;
	if (!htmlOrNameRegEx.exec(color)) throw new ParseError("Invalid color: '" + color + "'", token);
	if (xcolorHtmlRegEx.test(color)) return "#" + color;
	else if (color.charAt(0) === "#") return color;
	else if (macros.has(macroName)) color = macros.get(macroName).tokens[0].text;
	else if (xcolors[color]) color = xcolors[color];
	return color;
};
const mathmlBuilder$8 = (group, style) => {
	let expr = buildExpression(group.body, style.withColor(group.color));
	if (expr.length === 0) expr.push(new MathNode("mrow"));
	expr = expr.map((e) => {
		e.style.color = group.color;
		return e;
	});
	return newDocumentFragment(expr);
};
defineFunction({
	type: "color",
	names: ["\\textcolor"],
	props: {
		numArgs: 2,
		numOptionalArgs: 1,
		allowedInText: true,
		argTypes: [
			"raw",
			"raw",
			"original"
		]
	},
	handler({ parser, token }, args, optArgs) {
		const model = optArgs[0] && assertNodeType(optArgs[0], "raw").string;
		let color = "";
		if (model) {
			const spec = assertNodeType(args[0], "raw").string;
			color = colorFromSpec(model, spec);
		} else color = validateColor(assertNodeType(args[0], "raw").string, parser.gullet.macros, token);
		const body = args[1];
		return {
			type: "color",
			mode: parser.mode,
			color,
			isTextColor: true,
			body: ordargument(body)
		};
	},
	mathmlBuilder: mathmlBuilder$8
});
defineFunction({
	type: "color",
	names: ["\\color"],
	props: {
		numArgs: 1,
		numOptionalArgs: 1,
		allowedInText: true,
		argTypes: ["raw", "raw"]
	},
	handler({ parser, breakOnTokenText, token }, args, optArgs) {
		const model = optArgs[0] && assertNodeType(optArgs[0], "raw").string;
		let color = "";
		if (model) {
			const spec = assertNodeType(args[0], "raw").string;
			color = colorFromSpec(model, spec);
		} else color = validateColor(assertNodeType(args[0], "raw").string, parser.gullet.macros, token);
		const body = parser.parseExpression(true, breakOnTokenText, true);
		return {
			type: "color",
			mode: parser.mode,
			color,
			isTextColor: false,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$8
});
defineFunction({
	type: "color",
	names: ["\\definecolor"],
	props: {
		numArgs: 3,
		allowedInText: true,
		argTypes: [
			"raw",
			"raw",
			"raw"
		]
	},
	handler({ parser, funcName, token }, args) {
		const name = assertNodeType(args[0], "raw").string;
		if (!/^[A-Za-z]+$/.test(name)) throw new ParseError("Color name must be latin letters.", token);
		const model = assertNodeType(args[1], "raw").string;
		if (![
			"HTML",
			"RGB",
			"rgb"
		].includes(model)) throw new ParseError("Color model must be HTML, RGB, or rgb.", token);
		const spec = assertNodeType(args[2], "raw").string;
		const color = colorFromSpec(model, spec);
		parser.gullet.macros.set(`\\\\color@${name}`, {
			tokens: [{ text: color }],
			numArgs: 0
		});
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
defineFunction({
	type: "cr",
	names: ["\\\\"],
	props: {
		numArgs: 0,
		numOptionalArgs: 0,
		allowedInText: true
	},
	handler({ parser }, args, optArgs) {
		const size = parser.gullet.future().text === "[" ? parser.parseSizeGroup(true) : null;
		const newLine = !parser.settings.displayMode;
		return {
			type: "cr",
			mode: parser.mode,
			newLine,
			size: size && assertNodeType(size, "size").value
		};
	},
	mathmlBuilder(group, style) {
		const node = new MathNode("mo");
		if (group.newLine) {
			node.setAttribute("linebreak", "newline");
			if (group.size) {
				const size = calculateSize(group.size, style);
				node.setAttribute("height", size.number + size.unit);
			}
		}
		return node;
	}
});
const globalMap = {
	"\\global": "\\global",
	"\\long": "\\\\globallong",
	"\\\\globallong": "\\\\globallong",
	"\\def": "\\gdef",
	"\\gdef": "\\gdef",
	"\\edef": "\\xdef",
	"\\xdef": "\\xdef",
	"\\let": "\\\\globallet",
	"\\futurelet": "\\\\globalfuture"
};
const checkControlSequence = (tok) => {
	const name = tok.text;
	if (/^(?:[\\{}$&#^_]|EOF)$/.test(name)) throw new ParseError("Expected a control sequence", tok);
	return name;
};
const getRHS = (parser) => {
	let tok = parser.gullet.popToken();
	if (tok.text === "=") {
		tok = parser.gullet.popToken();
		if (tok.text === " ") tok = parser.gullet.popToken();
	}
	return tok;
};
const letCommand = (parser, name, tok, global) => {
	let macro = parser.gullet.macros.get(tok.text);
	if (macro == null) {
		tok.noexpand = true;
		macro = {
			tokens: [tok],
			numArgs: 0,
			unexpandable: !parser.gullet.isExpandable(tok.text)
		};
	}
	parser.gullet.macros.set(name, macro, global);
};
defineFunction({
	type: "internal",
	names: [
		"\\global",
		"\\long",
		"\\\\globallong"
	],
	props: {
		numArgs: 0,
		allowedInText: true
	},
	handler({ parser, funcName }) {
		parser.consumeSpaces();
		const token = parser.fetch();
		if (globalMap[token.text]) {
			if (funcName === "\\global" || funcName === "\\\\globallong") token.text = globalMap[token.text];
			return assertNodeType(parser.parseFunction(), "internal");
		}
		throw new ParseError(`Invalid token after macro prefix`, token);
	}
});
defineFunction({
	type: "internal",
	names: [
		"\\def",
		"\\gdef",
		"\\edef",
		"\\xdef"
	],
	props: {
		numArgs: 0,
		allowedInText: true,
		primitive: true
	},
	handler({ parser, funcName }) {
		let tok = parser.gullet.popToken();
		const name = tok.text;
		if (/^(?:[\\{}$&#^_]|EOF)$/.test(name)) throw new ParseError("Expected a control sequence", tok);
		let numArgs = 0;
		let insert;
		const delimiters = [[]];
		while (parser.gullet.future().text !== "{") {
			tok = parser.gullet.popToken();
			if (tok.text === "#") {
				if (parser.gullet.future().text === "{") {
					insert = parser.gullet.future();
					delimiters[numArgs].push("{");
					break;
				}
				tok = parser.gullet.popToken();
				if (!/^[1-9]$/.test(tok.text)) throw new ParseError(`Invalid argument number "${tok.text}"`);
				if (parseInt(tok.text) !== numArgs + 1) throw new ParseError(`Argument number "${tok.text}" out of order`);
				numArgs++;
				delimiters.push([]);
			} else if (tok.text === "EOF") throw new ParseError("Expected a macro definition");
			else delimiters[numArgs].push(tok.text);
		}
		let { tokens } = parser.gullet.consumeArg();
		if (insert) tokens.unshift(insert);
		if (funcName === "\\edef" || funcName === "\\xdef") {
			tokens = parser.gullet.expandTokens(tokens);
			if (tokens.length > parser.gullet.settings.maxExpand) throw new ParseError("Too many expansions in an " + funcName);
			tokens.reverse();
		}
		parser.gullet.macros.set(name, {
			tokens,
			numArgs,
			delimiters
		}, funcName === globalMap[funcName]);
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
defineFunction({
	type: "internal",
	names: ["\\let", "\\\\globallet"],
	props: {
		numArgs: 0,
		allowedInText: true,
		primitive: true
	},
	handler({ parser, funcName }) {
		const name = checkControlSequence(parser.gullet.popToken());
		parser.gullet.consumeSpaces();
		const tok = getRHS(parser);
		letCommand(parser, name, tok, funcName === "\\\\globallet");
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
defineFunction({
	type: "internal",
	names: ["\\futurelet", "\\\\globalfuture"],
	props: {
		numArgs: 0,
		allowedInText: true,
		primitive: true
	},
	handler({ parser, funcName }) {
		const name = checkControlSequence(parser.gullet.popToken());
		const middle = parser.gullet.popToken();
		const tok = parser.gullet.popToken();
		letCommand(parser, name, tok, funcName === "\\\\globalfuture");
		parser.gullet.pushToken(tok);
		parser.gullet.pushToken(middle);
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
defineFunction({
	type: "internal",
	names: [
		"\\newcommand",
		"\\renewcommand",
		"\\providecommand"
	],
	props: {
		numArgs: 0,
		allowedInText: true,
		primitive: true
	},
	handler({ parser, funcName }) {
		let name = "";
		const tok = parser.gullet.popToken();
		if (tok.text === "{") {
			name = checkControlSequence(parser.gullet.popToken());
			parser.gullet.popToken();
		} else name = checkControlSequence(tok);
		const exists = parser.gullet.isDefined(name);
		if (exists && funcName === "\\newcommand") throw new ParseError(`\\newcommand{${name}} attempting to redefine ${name}; use \\renewcommand`);
		if (!exists && funcName === "\\renewcommand") throw new ParseError(`\\renewcommand{${name}} when command ${name} does not yet exist; use \\newcommand`);
		let numArgs = 0;
		if (parser.gullet.future().text === "[") {
			let tok = parser.gullet.popToken();
			tok = parser.gullet.popToken();
			if (!/^[0-9]$/.test(tok.text)) throw new ParseError(`Invalid number of arguments: "${tok.text}"`);
			numArgs = parseInt(tok.text);
			tok = parser.gullet.popToken();
			if (tok.text !== "]") throw new ParseError(`Invalid argument "${tok.text}"`);
		}
		const { tokens } = parser.gullet.consumeArg();
		if (!(funcName === "\\providecommand" && parser.gullet.macros.has(name))) parser.gullet.macros.set(name, {
			tokens,
			numArgs
		});
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
const delimiterSizes = {
	"\\bigl": {
		mclass: "mopen",
		size: 1
	},
	"\\Bigl": {
		mclass: "mopen",
		size: 2
	},
	"\\biggl": {
		mclass: "mopen",
		size: 3
	},
	"\\Biggl": {
		mclass: "mopen",
		size: 4
	},
	"\\bigr": {
		mclass: "mclose",
		size: 1
	},
	"\\Bigr": {
		mclass: "mclose",
		size: 2
	},
	"\\biggr": {
		mclass: "mclose",
		size: 3
	},
	"\\Biggr": {
		mclass: "mclose",
		size: 4
	},
	"\\bigm": {
		mclass: "mrel",
		size: 1
	},
	"\\Bigm": {
		mclass: "mrel",
		size: 2
	},
	"\\biggm": {
		mclass: "mrel",
		size: 3
	},
	"\\Biggm": {
		mclass: "mrel",
		size: 4
	},
	"\\big": {
		mclass: "mord",
		size: 1
	},
	"\\Big": {
		mclass: "mord",
		size: 2
	},
	"\\bigg": {
		mclass: "mord",
		size: 3
	},
	"\\Bigg": {
		mclass: "mord",
		size: 4
	}
};
const leftToRight = {
	"(": ")",
	"\\lparen": "\\rparen",
	"[": "]",
	"\\lbrack": "\\rbrack",
	"\\{": "\\}",
	"\\lbrace": "\\rbrace",
	"⦇": "⦈",
	"\\llparenthesis": "\\rrparenthesis",
	"\\lfloor": "\\rfloor",
	"⌊": "⌋",
	"\\lceil": "\\rceil",
	"⌈": "⌉",
	"\\langle": "\\rangle",
	"⟨": "⟩",
	"\\lAngle": "\\rAngle",
	"⟪": "⟫",
	"\\llangle": "\\rrangle",
	"⦉": "⦊",
	"\\lvert": "\\rvert",
	"\\lVert": "\\rVert",
	"\\lgroup": "\\rgroup",
	"⟮": "⟯",
	"\\lmoustache": "\\rmoustache",
	"⎰": "⎱",
	"\\llbracket": "\\rrbracket",
	"⟦": "⟧",
	"\\lBrace": "\\rBrace",
	"⦃": "⦄"
};
const leftDelimiterNames = new Set(Object.keys(leftToRight));
new Set(Object.values(leftToRight));
const delimiters = /* @__PURE__ */ new Set([
	"(",
	"\\lparen",
	")",
	"\\rparen",
	"[",
	"\\lbrack",
	"]",
	"\\rbrack",
	"\\{",
	"\\lbrace",
	"\\}",
	"\\rbrace",
	"⦇",
	"\\llparenthesis",
	"⦈",
	"\\rrparenthesis",
	"\\lfloor",
	"\\rfloor",
	"⌊",
	"⌋",
	"\\lceil",
	"\\rceil",
	"⌈",
	"⌉",
	"<",
	">",
	"\\langle",
	"⟨",
	"\\rangle",
	"⟩",
	"\\lAngle",
	"⟪",
	"\\rAngle",
	"⟫",
	"\\llangle",
	"⦉",
	"\\rrangle",
	"⦊",
	"\\lt",
	"\\gt",
	"\\lvert",
	"\\rvert",
	"\\lVert",
	"\\rVert",
	"\\lgroup",
	"\\rgroup",
	"⟮",
	"⟯",
	"\\lmoustache",
	"\\rmoustache",
	"⎰",
	"⎱",
	"\\llbracket",
	"\\rrbracket",
	"⟦",
	"⟧",
	"\\lBrace",
	"\\rBrace",
	"⦃",
	"⦄",
	"/",
	"\\backslash",
	"|",
	"\\vert",
	"\\|",
	"\\Vert",
	"‖",
	"\\uparrow",
	"\\Uparrow",
	"\\downarrow",
	"\\Downarrow",
	"\\updownarrow",
	"\\Updownarrow",
	"."
]);
const dels = /* @__PURE__ */ new Set([
	"}",
	"\\left",
	"\\middle",
	"\\right"
]);
const isDelimiter = (str) => str.length > 0 && (delimiters.has(str) || delimiterSizes[str] || dels.has(str));
const sizeToMaxHeight = [
	0,
	1.2,
	1.8,
	2.4,
	3
];
function checkDelimiter(delim, context) {
	if (delim.type === "ordgroup" && delim.body.length === 1) delim = delim.body[0];
	const symDelim = checkSymbolNodeType(delim);
	if (symDelim && delimiters.has(symDelim.text)) {
		if (symDelim.text === "<" || symDelim.text === "\\lt") symDelim.text = "⟨";
		if (symDelim.text === ">" || symDelim.text === "\\gt") symDelim.text = "⟩";
		return symDelim;
	} else if (symDelim) throw new ParseError(`Invalid delimiter '${symDelim.text}' after '${context.funcName}'`, delim);
	else throw new ParseError(`Invalid delimiter type '${delim.type}'`, delim);
}
const needExplicitStretch = /* @__PURE__ */ new Set([
	"/",
	"\\",
	"\\backslash",
	"∖",
	"\\vert",
	"|"
]);
const makeFenceMo = (delim, mode, form, isStretchy) => {
	const node = new MathNode("mo", [makeText(delim === "." ? "" : delim, mode)]);
	node.setAttribute("fence", "true");
	node.setAttribute("form", form);
	node.setAttribute("stretchy", isStretchy ? "true" : "false");
	return node;
};
defineFunction({
	type: "delimsizing",
	names: [
		"\\bigl",
		"\\Bigl",
		"\\biggl",
		"\\Biggl",
		"\\bigr",
		"\\Bigr",
		"\\biggr",
		"\\Biggr",
		"\\bigm",
		"\\Bigm",
		"\\biggm",
		"\\Biggm",
		"\\big",
		"\\Big",
		"\\bigg",
		"\\Bigg"
	],
	props: {
		numArgs: 1,
		argTypes: ["primitive"]
	},
	handler: (context, args) => {
		const delim = checkDelimiter(args[0], context);
		const delimNode = {
			type: "delimsizing",
			mode: context.parser.mode,
			size: delimiterSizes[context.funcName].size,
			mclass: delimiterSizes[context.funcName].mclass,
			delim: delim.text
		};
		const nextToken = context.parser.fetch().text;
		if (nextToken !== "^" && nextToken !== "_") return delimNode;
		else return {
			type: "ordgroup",
			mode: "math",
			body: [delimNode, {
				type: "ordgroup",
				mode: "math",
				body: []
			}]
		};
	},
	mathmlBuilder: (group) => {
		const children = [];
		const delim = group.delim === "." ? "" : group.delim;
		children.push(makeText(delim, group.mode));
		const node = new MathNode("mo", children);
		if (group.mclass === "mopen" || group.mclass === "mclose") node.setAttribute("fence", "true");
		else node.setAttribute("fence", "false");
		if (needExplicitStretch.has(delim) || delim.indexOf("arrow") > -1) node.setAttribute("stretchy", "true");
		node.setAttribute("symmetric", "true");
		node.setAttribute("minsize", sizeToMaxHeight[group.size] + "em");
		node.setAttribute("maxsize", sizeToMaxHeight[group.size] + "em");
		return node;
	}
});
function assertParsed(group) {
	if (!group.body) throw new Error("Bug: The delim ParseNode wasn't fully parsed.");
}
defineFunction({
	type: "leftright-right",
	names: ["\\right"],
	props: {
		numArgs: 1,
		argTypes: ["primitive"]
	},
	handler: (context, args) => {
		return {
			type: "leftright-right",
			mode: context.parser.mode,
			delim: checkDelimiter(args[0], context).text
		};
	}
});
defineFunction({
	type: "leftright",
	names: ["\\left"],
	props: {
		numArgs: 1,
		argTypes: ["primitive"]
	},
	handler: (context, args) => {
		const delim = checkDelimiter(args[0], context);
		const parser = context.parser;
		++parser.leftrightDepth;
		let body = parser.parseExpression(false, "\\right", true);
		let nextToken = parser.fetch();
		while (nextToken.text === "\\middle") {
			parser.consume();
			const middle = parser.fetch().text;
			if (!symbols.math[middle]) throw new ParseError(`Invalid delimiter '${middle}' after '\\middle'`);
			checkDelimiter({
				type: "atom",
				mode: "math",
				text: middle
			}, { funcName: "\\middle" });
			body.push({
				type: "middle",
				mode: "math",
				delim: middle
			});
			parser.consume();
			body = body.concat(parser.parseExpression(false, "\\right", true));
			nextToken = parser.fetch();
		}
		--parser.leftrightDepth;
		parser.expect("\\right", false);
		const right = assertNodeType(parser.parseFunction(), "leftright-right");
		return {
			type: "leftright",
			mode: parser.mode,
			body,
			left: delim.text,
			right: right.delim,
			isStretchy: true
		};
	},
	mathmlBuilder: (group, style) => {
		assertParsed(group);
		const inner = buildExpression(group.body, style);
		const leftNode = makeFenceMo(group.left, group.mode, "prefix", true);
		inner.unshift(leftNode);
		const rightNode = makeFenceMo(group.right, group.mode, "postfix", true);
		if (group.body.length > 0) {
			const lastElement = group.body[group.body.length - 1];
			if (lastElement.type === "color" && !lastElement.isTextColor) rightNode.setAttribute("mathcolor", lastElement.color);
		}
		inner.push(rightNode);
		return makeRow(inner);
	}
});
defineFunction({
	type: "delimiter",
	names: Array.from(leftDelimiterNames),
	props: {
		numArgs: 0,
		allowedInText: true,
		allowedInMath: true,
		allowedInArgument: true
	},
	handler: ({ parser, funcName, token }) => {
		if (parser.mode === "text") return {
			type: "textord",
			mode: "text",
			text: funcName,
			loc: token.loc
		};
		else if (!parser.settings.wrapDelimiterPairs) return {
			type: "atom",
			mode: "math",
			family: "open",
			loc: token.loc,
			text: funcName
		};
		const rightDelim = leftToRight[funcName];
		const body = parser.parseExpression(false, rightDelim, false);
		if (parser.fetch().text !== rightDelim) throw new ParseError("Unmatched delimiter");
		parser.consume();
		return {
			type: "delimiter",
			mode: parser.mode,
			body,
			left: funcName,
			right: rightDelim
		};
	},
	mathmlBuilder: (group, style) => {
		assertParsed(group);
		const inner = buildExpression(group.body, style);
		const leftNode = makeFenceMo(group.left, group.mode, "prefix", false);
		inner.unshift(leftNode);
		const rightNode = makeFenceMo(group.right, group.mode, "postfix", false);
		if (group.body.length > 0) {
			const lastElement = group.body[group.body.length - 1];
			if (lastElement.type === "color" && !lastElement.isTextColor) rightNode.setAttribute("mathcolor", lastElement.color);
		}
		inner.push(rightNode);
		return makeRow(inner);
	}
});
defineFunction({
	type: "middle",
	names: ["\\middle"],
	props: {
		numArgs: 1,
		argTypes: ["primitive"]
	},
	handler: (context, args) => {
		const delim = checkDelimiter(args[0], context);
		if (!context.parser.leftrightDepth) throw new ParseError("\\middle without preceding \\left", delim);
		return {
			type: "middle",
			mode: context.parser.mode,
			delim: delim.text
		};
	},
	mathmlBuilder: (group) => {
		const textNode = makeText(group.delim, group.mode);
		const middleNode = new MathNode("mo", [textNode]);
		middleNode.setAttribute("stretchy", "true");
		middleNode.setAttribute("form", "infix");
		if (textNode.text !== "/") {
			middleNode.setAttribute("lspace", "0.05em");
			middleNode.setAttribute("rspace", "0.05em");
		}
		return middleNode;
	}
});
const boxTags = [
	"\\boxed",
	"\\fcolorbox",
	"\\colorbox"
];
const mathmlBuilder$7 = (group, style) => {
	const node = new MathNode(boxTags.includes(group.label) ? "mrow" : "menclose", [buildGroup$1(group.body, style)]);
	switch (group.label) {
		case "\\overline":
			node.setAttribute("notation", "top");
			node.classes.push("tml-overline");
			break;
		case "\\underline":
			node.setAttribute("notation", "bottom");
			node.classes.push("tml-underline");
			break;
		case "\\cancel":
			node.setAttribute("notation", "updiagonalstrike");
			node.children.push(new MathNode("mrow", [], ["tml-cancel", "upstrike"]));
			break;
		case "\\bcancel":
			node.setAttribute("notation", "downdiagonalstrike");
			node.children.push(new MathNode("mrow", [], ["tml-cancel", "downstrike"]));
			break;
		case "\\sout":
			node.setAttribute("notation", "horizontalstrike");
			node.children.push(new MathNode("mrow", [], ["tml-cancel", "sout"]));
			break;
		case "\\xcancel":
			node.setAttribute("notation", "updiagonalstrike downdiagonalstrike");
			node.children.push(new MathNode("mrow", [], ["tml-cancel", "tml-xcancel"]));
			break;
		case "\\longdiv":
			node.setAttribute("notation", "longdiv");
			node.classes.push("longdiv-top");
			node.children.push(new MathNode("mrow", [], ["longdiv-arc"]));
			break;
		case "\\phase":
			node.setAttribute("notation", "phasorangle");
			node.classes.push("phasor-bottom");
			node.children.push(new MathNode("mrow", [], ["phasor-angle"]));
			break;
		case "\\textcircled":
			node.setAttribute("notation", "circle");
			node.classes.push("circle-pad");
			node.children.push(new MathNode("mrow", [], ["textcircle"]));
			break;
		case "\\angl":
			node.setAttribute("notation", "actuarial");
			node.classes.push("actuarial");
			break;
		case "\\boxed":
			node.style.padding = "3pt";
			node.style.border = "1px solid";
			node.setAttribute("scriptlevel", "0");
			node.setAttribute("displaystyle", "true");
			break;
		case "\\fbox":
			node.setAttribute("notation", "box");
			node.classes.push("tml-fbox");
			break;
		case "\\fcolorbox":
		case "\\colorbox":
			node.style.padding = "0.3em";
			if (group.label === "\\fcolorbox") node.style.border = "0.0667em solid " + String(group.borderColor);
	}
	if (group.backgroundColor) node.setAttribute("mathbackground", group.backgroundColor);
	return node;
};
defineFunction({
	type: "enclose",
	names: ["\\colorbox"],
	props: {
		numArgs: 2,
		numOptionalArgs: 1,
		allowedInText: true,
		argTypes: [
			"raw",
			"raw",
			"text"
		]
	},
	handler({ parser, funcName }, args, optArgs) {
		const model = optArgs[0] && assertNodeType(optArgs[0], "raw").string;
		let color = "";
		if (model) {
			const spec = assertNodeType(args[0], "raw").string;
			color = colorFromSpec(model, spec);
		} else color = validateColor(assertNodeType(args[0], "raw").string, parser.gullet.macros);
		const body = args[1];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			backgroundColor: color,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "enclose",
	names: ["\\fcolorbox"],
	props: {
		numArgs: 3,
		numOptionalArgs: 1,
		allowedInText: true,
		argTypes: [
			"raw",
			"raw",
			"raw",
			"text"
		]
	},
	handler({ parser, funcName }, args, optArgs) {
		const model = optArgs[0] && assertNodeType(optArgs[0], "raw").string;
		let borderColor = "";
		let backgroundColor;
		if (model) {
			const borderSpec = assertNodeType(args[0], "raw").string;
			const backgroundSpec = assertNodeType(args[0], "raw").string;
			borderColor = colorFromSpec(model, borderSpec);
			backgroundColor = colorFromSpec(model, backgroundSpec);
		} else {
			borderColor = validateColor(assertNodeType(args[0], "raw").string, parser.gullet.macros);
			backgroundColor = validateColor(assertNodeType(args[1], "raw").string, parser.gullet.macros);
		}
		const body = args[2];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			backgroundColor,
			borderColor,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "enclose",
	names: ["\\fbox"],
	props: {
		numArgs: 1,
		argTypes: ["hbox"],
		allowedInText: true
	},
	handler({ parser }, args) {
		return {
			type: "enclose",
			mode: parser.mode,
			label: "\\fbox",
			body: args[0]
		};
	}
});
defineFunction({
	type: "enclose",
	names: [
		"\\angl",
		"\\cancel",
		"\\bcancel",
		"\\xcancel",
		"\\overline",
		"\\boxed",
		"\\longdiv",
		"\\phase"
	],
	props: { numArgs: 1 },
	handler({ parser, funcName }, args) {
		const body = args[0];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "enclose",
	names: ["\\sout"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler({ parser, funcName }, args) {
		const body = args[0];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "enclose",
	names: ["\\underline"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler({ parser, funcName }, args) {
		const body = args[0];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "enclose",
	names: ["\\textcircled"],
	props: {
		numArgs: 1,
		argTypes: ["text"],
		allowedInArgument: true,
		allowedInText: true
	},
	handler({ parser, funcName }, args) {
		const body = args[0];
		return {
			type: "enclose",
			mode: parser.mode,
			label: funcName,
			body
		};
	},
	mathmlBuilder: mathmlBuilder$7
});
defineFunction({
	type: "environment",
	names: ["\\begin", "\\end"],
	props: {
		numArgs: 1,
		argTypes: ["text"]
	},
	handler({ parser, funcName }, args) {
		const nameGroup = args[0];
		if (nameGroup.type !== "ordgroup") throw new ParseError("Invalid environment name", nameGroup);
		let envName = "";
		for (let i = 0; i < nameGroup.body.length; ++i) envName += assertNodeType(nameGroup.body[i], "textord").text;
		if (funcName === "\\begin") {
			if (!Object.prototype.hasOwnProperty.call(environments, envName)) throw new ParseError("No such environment: " + envName, nameGroup);
			const env = environments[envName];
			const { args, optArgs } = parser.parseArguments("\\begin{" + envName + "}", env);
			const context = {
				mode: parser.mode,
				envName,
				parser
			};
			const result = env.handler(context, args, optArgs);
			parser.expect("\\end", false);
			const endNameToken = parser.nextToken;
			const end = assertNodeType(parser.parseFunction(), "environment");
			if (end.name !== envName) throw new ParseError(`Mismatch: \\begin{${envName}} matched by \\end{${end.name}}`, endNameToken);
			return result;
		}
		return {
			type: "environment",
			mode: parser.mode,
			name: envName,
			nameGroup
		};
	}
});
defineFunction({
	type: "envTag",
	names: ["\\env@tag"],
	props: {
		numArgs: 1,
		argTypes: ["math"]
	},
	handler({ parser }, args) {
		return {
			type: "envTag",
			mode: parser.mode,
			body: args[0]
		};
	},
	mathmlBuilder(group, style) {
		return new MathNode("mrow");
	}
});
defineFunction({
	type: "noTag",
	names: ["\\env@notag"],
	props: { numArgs: 0 },
	handler({ parser }) {
		return {
			type: "noTag",
			mode: parser.mode
		};
	},
	mathmlBuilder(group, style) {
		return new MathNode("mrow");
	}
});
const script = Object.freeze({
	B: 8426,
	E: 8427,
	F: 8427,
	H: 8387,
	I: 8391,
	L: 8390,
	M: 8422,
	R: 8393,
	e: 8394,
	g: 8355,
	o: 8389
});
const frak = Object.freeze({
	C: 8426,
	H: 8388,
	I: 8392,
	R: 8394,
	Z: 8398
});
const bbb = Object.freeze({
	C: 8383,
	H: 8389,
	N: 8391,
	P: 8393,
	Q: 8393,
	R: 8395,
	Z: 8394
});
const bold = Object.freeze({
	"ϵ": 119527,
	"ϑ": 119564,
	"ϰ": 119534,
	"φ": 119577,
	"ϱ": 119535,
	"ϖ": 119563
});
const boldItalic = Object.freeze({
	"ϵ": 119643,
	"ϑ": 119680,
	"ϰ": 119650,
	"φ": 119693,
	"ϱ": 119651,
	"ϖ": 119679
});
const boldsf = Object.freeze({
	"ϵ": 119701,
	"ϑ": 119738,
	"ϰ": 119708,
	"φ": 119751,
	"ϱ": 119709,
	"ϖ": 119737
});
const bisf = Object.freeze({
	"ϵ": 119759,
	"ϑ": 119796,
	"ϰ": 119766,
	"φ": 119809,
	"ϱ": 119767,
	"ϖ": 119795
});
const offset = Object.freeze({
	upperCaseLatin: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return 119743;
		},
		"italic": (ch) => {
			return 119795;
		},
		"bold-italic": (ch) => {
			return 119847;
		},
		"script": (ch) => {
			return script[ch] || 119899;
		},
		"script-bold": (ch) => {
			return 119951;
		},
		"fraktur": (ch) => {
			return frak[ch] || 120003;
		},
		"fraktur-bold": (ch) => {
			return 120107;
		},
		"double-struck": (ch) => {
			return bbb[ch] || 120055;
		},
		"sans-serif": (ch) => {
			return 120159;
		},
		"sans-serif-bold": (ch) => {
			return 120211;
		},
		"sans-serif-italic": (ch) => {
			return 120263;
		},
		"sans-serif-bold-italic": (ch) => {
			return 120380;
		},
		"monospace": (ch) => {
			return 120367;
		}
	},
	lowerCaseLatin: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return 119737;
		},
		"italic": (ch) => {
			return ch === "h" ? 8358 : 119789;
		},
		"bold-italic": (ch) => {
			return 119841;
		},
		"script": (ch) => {
			return script[ch] || 119893;
		},
		"script-bold": (ch) => {
			return 119945;
		},
		"fraktur": (ch) => {
			return 119997;
		},
		"fraktur-bold": (ch) => {
			return 120101;
		},
		"double-struck": (ch) => {
			return 120049;
		},
		"sans-serif": (ch) => {
			return 120153;
		},
		"sans-serif-bold": (ch) => {
			return 120205;
		},
		"sans-serif-italic": (ch) => {
			return 120257;
		},
		"sans-serif-bold-italic": (ch) => {
			return 120309;
		},
		"monospace": (ch) => {
			return 120361;
		}
	},
	upperCaseGreek: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return 119575;
		},
		"italic": (ch) => {
			return 119633;
		},
		"bold-italic": (ch) => {
			return 119575;
		},
		"script": (ch) => {
			return 0;
		},
		"script-bold": (ch) => {
			return 0;
		},
		"fraktur": (ch) => {
			return 0;
		},
		"fraktur-bold": (ch) => {
			return 0;
		},
		"double-struck": (ch) => {
			return 0;
		},
		"sans-serif": (ch) => {
			return 119749;
		},
		"sans-serif-bold": (ch) => {
			return 119749;
		},
		"sans-serif-italic": (ch) => {
			return 0;
		},
		"sans-serif-bold-italic": (ch) => {
			return 119807;
		},
		"monospace": (ch) => {
			return 0;
		}
	},
	lowerCaseGreek: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return 119569;
		},
		"italic": (ch) => {
			return 119627;
		},
		"bold-italic": (ch) => {
			return ch === "ϕ" ? 119678 : 119685;
		},
		"script": (ch) => {
			return 0;
		},
		"script-bold": (ch) => {
			return 0;
		},
		"fraktur": (ch) => {
			return 0;
		},
		"fraktur-bold": (ch) => {
			return 0;
		},
		"double-struck": (ch) => {
			return 0;
		},
		"sans-serif": (ch) => {
			return 119743;
		},
		"sans-serif-bold": (ch) => {
			return 119743;
		},
		"sans-serif-italic": (ch) => {
			return 0;
		},
		"sans-serif-bold-italic": (ch) => {
			return 119801;
		},
		"monospace": (ch) => {
			return 0;
		}
	},
	varGreek: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return bold[ch] || -51;
		},
		"italic": (ch) => {
			return 0;
		},
		"bold-italic": (ch) => {
			return boldItalic[ch] || 58;
		},
		"script": (ch) => {
			return 0;
		},
		"script-bold": (ch) => {
			return 0;
		},
		"fraktur": (ch) => {
			return 0;
		},
		"fraktur-bold": (ch) => {
			return 0;
		},
		"double-struck": (ch) => {
			return 0;
		},
		"sans-serif": (ch) => {
			return boldsf[ch] || 116;
		},
		"sans-serif-bold": (ch) => {
			return boldsf[ch] || 116;
		},
		"sans-serif-italic": (ch) => {
			return 0;
		},
		"sans-serif-bold-italic": (ch) => {
			return bisf[ch] || 174;
		},
		"monospace": (ch) => {
			return 0;
		}
	},
	numeral: {
		"normal": (ch) => {
			return 0;
		},
		"bold": (ch) => {
			return 120734;
		},
		"italic": (ch) => {
			return 0;
		},
		"bold-italic": (ch) => {
			return 0;
		},
		"script": (ch) => {
			return 0;
		},
		"script-bold": (ch) => {
			return 0;
		},
		"fraktur": (ch) => {
			return 0;
		},
		"fraktur-bold": (ch) => {
			return 0;
		},
		"double-struck": (ch) => {
			return 120744;
		},
		"sans-serif": (ch) => {
			return 120754;
		},
		"sans-serif-bold": (ch) => {
			return 120764;
		},
		"sans-serif-italic": (ch) => {
			return 0;
		},
		"sans-serif-bold-italic": (ch) => {
			return 0;
		},
		"monospace": (ch) => {
			return 120774;
		}
	}
});
const variantChar = (ch, variant) => {
	const codePoint = ch.codePointAt(0);
	const block = 64 < codePoint && codePoint < 91 ? "upperCaseLatin" : 96 < codePoint && codePoint < 123 ? "lowerCaseLatin" : 912 < codePoint && codePoint < 938 ? "upperCaseGreek" : 944 < codePoint && codePoint < 970 || ch === "ϕ" ? "lowerCaseGreek" : 120545 < codePoint && codePoint < 120572 || bold[ch] ? "varGreek" : 47 < codePoint && codePoint < 58 ? "numeral" : "other";
	return block === "other" ? ch : String.fromCodePoint(codePoint + offset[block][variant](ch));
};
const smallCaps = Object.freeze({
	a: "ᴀ",
	b: "ʙ",
	c: "ᴄ",
	d: "ᴅ",
	e: "ᴇ",
	f: "ꜰ",
	g: "ɢ",
	h: "ʜ",
	i: "ɪ",
	j: "ᴊ",
	k: "ᴋ",
	l: "ʟ",
	m: "ᴍ",
	n: "ɴ",
	o: "ᴏ",
	p: "ᴘ",
	q: "ǫ",
	r: "ʀ",
	s: "s",
	t: "ᴛ",
	u: "ᴜ",
	v: "ᴠ",
	w: "ᴡ",
	x: "x",
	y: "ʏ",
	z: "ᴢ"
});
const varNameFonts = ["mathrm", "mathit"];
const isLongVariableName = (group, font) => {
	if (!varNameFonts.includes(font) || !group.body || group.body.type !== "ordgroup" || group.body.body.length === 1) return false;
	if (group.body.body[0].type !== "mathord") return false;
	for (let i = 1; i < group.body.body.length; i++) {
		const parseNodeType = group.body.body[i].type;
		if (!(parseNodeType === "mathord" || parseNodeType === "textord" && !isNaN(group.body.body[i].text))) return false;
	}
	return true;
};
const mathmlBuilder$6 = (group, style) => {
	const font = group.font;
	const newStyle = style.withFont(font);
	const mathGroup = buildGroup$1(group.body, newStyle);
	if (mathGroup.children.length === 0) return mathGroup;
	if (font === "boldsymbol" && [
		"mo",
		"mpadded",
		"mrow"
	].includes(mathGroup.type)) {
		mathGroup.style.fontWeight = "bold";
		return mathGroup;
	}
	if (isLongVariableName(group, font)) {
		const mi = mathGroup.children[0].children[0].children ? mathGroup.children[0].children[0] : mathGroup.children[0];
		delete mi.attributes.mathvariant;
		for (let i = 1; i < mathGroup.children.length; i++) mi.children[0].text += mathGroup.children[i].children[0].children ? mathGroup.children[i].children[0].children[0].text : mathGroup.children[i].children[0].text;
		if (font === "mathit") {
			mi.children[0].text = mi.children[0].text.split("").map((c) => variantChar(c, "italic")).join("");
			return mi;
		}
		const mpadded = new MathNode("mpadded", [mi]);
		mpadded.setAttribute("lspace", "0");
		return mpadded;
	}
	let canConsolidate = mathGroup.children[0].type === "mo";
	for (let i = 1; i < mathGroup.children.length; i++) {
		if (mathGroup.children[i].type === "mo" && font === "boldsymbol") mathGroup.children[i].style.fontWeight = "bold";
		if (mathGroup.children[i].type !== "mi") canConsolidate = false;
		if ((mathGroup.children[i].attributes && mathGroup.children[i].attributes.mathvariant || "") !== "normal") canConsolidate = false;
	}
	if (!canConsolidate) return mathGroup;
	const mi = mathGroup.children[0];
	for (let i = 1; i < mathGroup.children.length; i++) mi.children.push(mathGroup.children[i].children[0]);
	if (mi.attributes.mathvariant && mi.attributes.mathvariant === "normal") return new MathNode("mrow", [new MathNode("mtext", new TextNode("​")), mi]);
	return mi;
};
const fontAliases = {
	"\\Bbb": "\\mathbb",
	"\\bold": "\\mathbf",
	"\\frak": "\\mathfrak",
	"\\bm": "\\boldsymbol"
};
defineFunction({
	type: "font",
	names: [
		"\\mathrm",
		"\\mathit",
		"\\mathbf",
		"\\mathnormal",
		"\\up@greek",
		"\\boldsymbol",
		"\\mathbb",
		"\\mathcal",
		"\\mathfrak",
		"\\mathscr",
		"\\mathsf",
		"\\mathsfit",
		"\\mathtt",
		"\\Bbb",
		"\\bm",
		"\\bold",
		"\\frak"
	],
	props: {
		numArgs: 1,
		allowedInArgument: true
	},
	handler: ({ parser, funcName }, args) => {
		const body = normalizeArgument(args[0]);
		let func = funcName;
		if (func in fontAliases) func = fontAliases[func];
		return {
			type: "font",
			mode: parser.mode,
			font: func.slice(1),
			body
		};
	},
	mathmlBuilder: mathmlBuilder$6
});
defineFunction({
	type: "font",
	names: [
		"\\rm",
		"\\sf",
		"\\tt",
		"\\bf",
		"\\it",
		"\\cal"
	],
	props: {
		numArgs: 0,
		allowedInText: true
	},
	handler: ({ parser, funcName, breakOnTokenText }, args) => {
		const { mode } = parser;
		const body = parser.parseExpression(true, breakOnTokenText, true);
		return {
			type: "font",
			mode,
			font: `math${funcName.slice(1)}`,
			body: {
				type: "ordgroup",
				mode: parser.mode,
				body
			}
		};
	},
	mathmlBuilder: mathmlBuilder$6
});
const stylArray = [
	"display",
	"text",
	"script",
	"scriptscript"
];
const scriptLevel = {
	auto: -1,
	display: 0,
	text: 0,
	script: 1,
	scriptscript: 2
};
const adjustStyle = (functionSize, originalStyle) => {
	let style = originalStyle;
	if (functionSize === "display") {
		const newSize = style.level >= StyleLevel.SCRIPT ? StyleLevel.TEXT : StyleLevel.DISPLAY;
		style = style.withLevel(newSize);
	} else if (functionSize === "text" && style.level === StyleLevel.DISPLAY) style = style.withLevel(StyleLevel.TEXT);
	else if (functionSize === "auto") style = style.incrementLevel();
	else if (functionSize === "script") style = style.withLevel(StyleLevel.SCRIPT);
	else if (functionSize === "scriptscript") style = style.withLevel(StyleLevel.SCRIPTSCRIPT);
	return style;
};
const mathmlBuilder$5 = (group, style) => {
	style = adjustStyle(group.scriptLevel, style);
	const numer = buildGroup$1(group.numer, style);
	const denom = buildGroup$1(group.denom, style);
	if (style.level === 3) {
		numer.style.mathDepth = "2";
		numer.setAttribute("scriptlevel", "2");
		denom.style.mathDepth = "2";
		denom.setAttribute("scriptlevel", "2");
	}
	let node = new MathNode("mfrac", [numer, denom]);
	if (!group.hasBarLine) node.setAttribute("linethickness", "0px");
	else if (group.barSize) {
		const ruleWidth = calculateSize(group.barSize, style);
		node.setAttribute("linethickness", ruleWidth.number + ruleWidth.unit);
	}
	if (group.leftDelim != null || group.rightDelim != null) {
		const withDelims = [];
		if (group.leftDelim != null) {
			const leftOp = new MathNode("mo", [new TextNode(group.leftDelim.replace("\\", ""))]);
			leftOp.setAttribute("fence", "true");
			withDelims.push(leftOp);
		}
		withDelims.push(node);
		if (group.rightDelim != null) {
			const rightOp = new MathNode("mo", [new TextNode(group.rightDelim.replace("\\", ""))]);
			rightOp.setAttribute("fence", "true");
			withDelims.push(rightOp);
		}
		node = makeRow(withDelims);
	}
	if (group.scriptLevel !== "auto") {
		node = new MathNode("mstyle", [node]);
		node.setAttribute("displaystyle", String(group.scriptLevel === "display"));
		node.setAttribute("scriptlevel", scriptLevel[group.scriptLevel]);
	}
	return node;
};
defineFunction({
	type: "genfrac",
	names: [
		"\\cfrac",
		"\\dfrac",
		"\\frac",
		"\\tfrac",
		"\\dbinom",
		"\\binom",
		"\\tbinom",
		"\\\\atopfrac",
		"\\\\bracefrac",
		"\\\\brackfrac"
	],
	props: {
		numArgs: 2,
		allowedInArgument: true
	},
	handler: ({ parser, funcName }, args) => {
		const numer = args[0];
		const denom = args[1];
		let hasBarLine = false;
		let leftDelim = null;
		let rightDelim = null;
		let scriptLevel = "auto";
		switch (funcName) {
			case "\\cfrac":
			case "\\dfrac":
			case "\\frac":
			case "\\tfrac":
				hasBarLine = true;
				break;
			case "\\\\atopfrac":
				hasBarLine = false;
				break;
			case "\\dbinom":
			case "\\binom":
			case "\\tbinom":
				leftDelim = "(";
				rightDelim = ")";
				break;
			case "\\\\bracefrac":
				leftDelim = "\\{";
				rightDelim = "\\}";
				break;
			case "\\\\brackfrac":
				leftDelim = "[";
				rightDelim = "]";
				break;
			default: throw new Error("Unrecognized genfrac command");
		}
		if (funcName === "\\cfrac" || funcName.startsWith("\\d")) scriptLevel = "display";
		else if (funcName.startsWith("\\t")) scriptLevel = "text";
		return {
			type: "genfrac",
			mode: parser.mode,
			continued: false,
			numer,
			denom,
			hasBarLine,
			leftDelim,
			rightDelim,
			scriptLevel,
			barSize: null
		};
	},
	mathmlBuilder: mathmlBuilder$5
});
defineFunction({
	type: "infix",
	names: [
		"\\over",
		"\\choose",
		"\\atop",
		"\\brace",
		"\\brack"
	],
	props: {
		numArgs: 0,
		infix: true
	},
	handler({ parser, funcName, token }) {
		let replaceWith;
		switch (funcName) {
			case "\\over":
				replaceWith = "\\frac";
				break;
			case "\\choose":
				replaceWith = "\\binom";
				break;
			case "\\atop":
				replaceWith = "\\\\atopfrac";
				break;
			case "\\brace":
				replaceWith = "\\\\bracefrac";
				break;
			case "\\brack":
				replaceWith = "\\\\brackfrac";
				break;
			default: throw new Error("Unrecognized infix genfrac command");
		}
		return {
			type: "infix",
			mode: parser.mode,
			replaceWith,
			token
		};
	}
});
const delimFromValue = function(delimString) {
	let delim = null;
	if (delimString.length > 0) {
		delim = delimString;
		delim = delim === "." ? null : delim;
	}
	return delim;
};
defineFunction({
	type: "genfrac",
	names: ["\\genfrac"],
	props: {
		numArgs: 6,
		allowedInArgument: true,
		argTypes: [
			"math",
			"math",
			"size",
			"text",
			"math",
			"math"
		]
	},
	handler({ parser }, args) {
		const numer = args[4];
		const denom = args[5];
		const leftNode = normalizeArgument(args[0]);
		const leftDelim = leftNode.type === "atom" && leftNode.family === "open" ? delimFromValue(leftNode.text) : null;
		const rightNode = normalizeArgument(args[1]);
		const rightDelim = rightNode.type === "atom" && rightNode.family === "close" ? delimFromValue(rightNode.text) : null;
		const barNode = assertNodeType(args[2], "size");
		let hasBarLine;
		let barSize = null;
		if (barNode.isBlank) hasBarLine = true;
		else {
			barSize = barNode.value;
			hasBarLine = barSize.number > 0;
		}
		let scriptLevel = "auto";
		let styl = args[3];
		if (styl.type === "ordgroup") {
			if (styl.body.length > 0) {
				const textOrd = assertNodeType(styl.body[0], "textord");
				scriptLevel = stylArray[Number(textOrd.text)];
			}
		} else {
			styl = assertNodeType(styl, "textord");
			scriptLevel = stylArray[Number(styl.text)];
		}
		return {
			type: "genfrac",
			mode: parser.mode,
			numer,
			denom,
			continued: false,
			hasBarLine,
			barSize,
			leftDelim,
			rightDelim,
			scriptLevel
		};
	},
	mathmlBuilder: mathmlBuilder$5
});
defineFunction({
	type: "infix",
	names: ["\\above"],
	props: {
		numArgs: 1,
		argTypes: ["size"],
		infix: true
	},
	handler({ parser, funcName, token }, args) {
		return {
			type: "infix",
			mode: parser.mode,
			replaceWith: "\\\\abovefrac",
			barSize: assertNodeType(args[0], "size").value,
			token
		};
	}
});
defineFunction({
	type: "genfrac",
	names: ["\\\\abovefrac"],
	props: {
		numArgs: 3,
		argTypes: [
			"math",
			"size",
			"math"
		]
	},
	handler: ({ parser, funcName }, args) => {
		const numer = args[0];
		const barSize = assert(assertNodeType(args[1], "infix").barSize);
		const denom = args[2];
		const hasBarLine = barSize.number > 0;
		return {
			type: "genfrac",
			mode: parser.mode,
			numer,
			denom,
			continued: false,
			hasBarLine,
			barSize,
			leftDelim: null,
			rightDelim: null,
			scriptLevel: "auto"
		};
	},
	mathmlBuilder: mathmlBuilder$5
});
defineFunction({
	type: "hbox",
	names: ["\\hbox"],
	props: {
		numArgs: 1,
		argTypes: ["hbox"],
		allowedInArgument: true,
		allowedInText: false
	},
	handler({ parser }, args) {
		return {
			type: "hbox",
			mode: parser.mode,
			body: ordargument(args[0])
		};
	},
	mathmlBuilder(group, style) {
		const newStyle = style.withLevel(StyleLevel.TEXT);
		const mrow = buildExpressionRow(group.body, newStyle);
		return consolidateText(mrow);
	}
});
const mathmlBuilder$4 = (group, style) => {
	const accentNode = mathMLnode(group.label);
	accentNode.style["math-depth"] = 0;
	return new MathNode(group.isOver ? "mover" : "munder", [buildGroup$1(group.base, style), accentNode]);
};
defineFunction({
	type: "horizBracket",
	names: [
		"\\overbrace",
		"\\underbrace",
		"\\overbracket",
		"\\underbracket"
	],
	props: { numArgs: 1 },
	handler({ parser, funcName }, args) {
		return {
			type: "horizBracket",
			mode: parser.mode,
			label: funcName,
			isOver: /^\\over/.test(funcName),
			base: args[0]
		};
	},
	mathmlBuilder: mathmlBuilder$4
});
defineFunction({
	type: "html",
	names: [
		"\\class",
		"\\id",
		"\\style",
		"\\data"
	],
	props: {
		numArgs: 2,
		argTypes: ["raw", "original"],
		allowedInText: true
	},
	handler: ({ parser, funcName, token }, args) => {
		const value = assertNodeType(args[0], "raw").string;
		const body = args[1];
		if (parser.settings.strict) throw new ParseError(`Function "${funcName}" is disabled in strict mode`, token);
		let trustContext;
		const attributes = {};
		switch (funcName) {
			case "\\class":
				attributes.class = value;
				trustContext = {
					command: "\\class",
					class: value
				};
				break;
			case "\\id":
				attributes.id = value;
				trustContext = {
					command: "\\id",
					id: value
				};
				break;
			case "\\style":
				attributes.style = value;
				trustContext = {
					command: "\\style",
					style: value
				};
				break;
			case "\\data": {
				const data = value.split(",");
				for (let i = 0; i < data.length; i++) {
					const keyVal = data[i].split("=");
					if (keyVal.length !== 2) throw new ParseError("Error parsing key-value for \\data");
					attributes["data-" + keyVal[0].trim()] = keyVal[1].trim();
				}
				trustContext = {
					command: "\\data",
					attributes
				};
				break;
			}
			default: throw new Error("Unrecognized html command");
		}
		if (!parser.settings.isTrusted(trustContext)) throw new ParseError(`Function "${funcName}" is not trusted`, token);
		return {
			type: "html",
			mode: parser.mode,
			attributes,
			body: ordargument(body)
		};
	},
	mathmlBuilder: (group, style) => {
		const element = buildExpressionRow(group.body, style);
		const classes = [];
		if (group.attributes.class) classes.push(...group.attributes.class.trim().split(/\s+/));
		element.classes = classes;
		for (const attr in group.attributes) if (attr !== "class" && Object.prototype.hasOwnProperty.call(group.attributes, attr)) element.setAttribute(attr, group.attributes[attr]);
		return element;
	}
});
const sizeData = function(str) {
	if (/^[-+]? *(\d+(\.\d*)?|\.\d+)$/.test(str)) return {
		number: +str,
		unit: "bp"
	};
	else {
		const match = /([-+]?) *(\d+(?:\.\d*)?|\.\d+) *([a-z]{2})/.exec(str);
		if (!match) throw new ParseError("Invalid size: '" + str + "' in \\includegraphics");
		const data = {
			number: +(match[1] + match[2]),
			unit: match[3]
		};
		if (!validUnit(data)) throw new ParseError("Invalid unit: '" + data.unit + "' in \\includegraphics.");
		return data;
	}
};
defineFunction({
	type: "includegraphics",
	names: ["\\includegraphics"],
	props: {
		numArgs: 1,
		numOptionalArgs: 1,
		argTypes: ["raw", "url"],
		allowedInText: false
	},
	handler: ({ parser, token }, args, optArgs) => {
		let width = {
			number: 0,
			unit: "em"
		};
		let height = {
			number: .9,
			unit: "em"
		};
		let totalheight = {
			number: 0,
			unit: "em"
		};
		let alt = "";
		if (optArgs[0]) {
			const attributes = assertNodeType(optArgs[0], "raw").string.split(",");
			for (let i = 0; i < attributes.length; i++) {
				const keyVal = attributes[i].split("=");
				if (keyVal.length === 2) {
					const str = keyVal[1].trim();
					switch (keyVal[0].trim()) {
						case "alt":
							alt = str;
							break;
						case "width":
							width = sizeData(str);
							break;
						case "height":
							height = sizeData(str);
							break;
						case "totalheight":
							totalheight = sizeData(str);
							break;
						default: throw new ParseError("Invalid key: '" + keyVal[0] + "' in \\includegraphics.");
					}
				}
			}
		}
		const src = assertNodeType(args[0], "url").url;
		if (alt === "") {
			alt = src;
			alt = alt.replace(/^.*[\\/]/, "");
			alt = alt.substring(0, alt.lastIndexOf("."));
		}
		if (!parser.settings.isTrusted({
			command: "\\includegraphics",
			url: src
		})) throw new ParseError(`Function "\\includegraphics" is not trusted`, token);
		return {
			type: "includegraphics",
			mode: parser.mode,
			alt,
			width,
			height,
			totalheight,
			src
		};
	},
	mathmlBuilder: (group, style) => {
		const height = calculateSize(group.height, style);
		const depth = {
			number: 0,
			unit: "em"
		};
		if (group.totalheight.number > 0) {
			if (group.totalheight.unit === height.unit && group.totalheight.number > height.number) {
				depth.number = group.totalheight.number - height.number;
				depth.unit = height.unit;
			}
		}
		let width = 0;
		if (group.width.number > 0) width = calculateSize(group.width, style);
		const graphicStyle = { height: height.number + depth.number + "em" };
		if (width.number > 0) graphicStyle.width = width.number + width.unit;
		if (depth.number > 0) graphicStyle.verticalAlign = -depth.number + depth.unit;
		const node = new Img(group.src, group.alt, graphicStyle);
		node.height = height;
		node.depth = depth;
		return new MathNode("mtext", [node]);
	}
});
defineFunction({
	type: "kern",
	names: [
		"\\kern",
		"\\mkern",
		"\\hskip",
		"\\mskip"
	],
	props: {
		numArgs: 1,
		argTypes: ["size"],
		primitive: true,
		allowedInText: true
	},
	handler({ parser, funcName, token }, args) {
		const size = assertNodeType(args[0], "size");
		if (parser.settings.strict) {
			const mathFunction = funcName[1] === "m";
			const muUnit = size.value.unit === "mu";
			if (mathFunction) {
				if (!muUnit) throw new ParseError(`LaTeX's ${funcName} supports only mu units, not ${size.value.unit} units`, token);
				if (parser.mode !== "math") throw new ParseError(`LaTeX's ${funcName} works only in math mode`, token);
			} else if (muUnit) throw new ParseError(`LaTeX's ${funcName} doesn't support mu units`, token);
		}
		return {
			type: "kern",
			mode: parser.mode,
			dimension: size.value
		};
	},
	mathmlBuilder(group, style) {
		const dimension = calculateSize(group.dimension, style);
		const ch = dimension.number > 0 && dimension.unit === "em" ? spaceCharacter(dimension.number) : "";
		if (group.mode === "text" && ch.length > 0) return new MathNode("mtext", [new TextNode(ch)]);
		else if (dimension.number >= 0) {
			const node = new MathNode("mspace");
			node.setAttribute("width", dimension.number + dimension.unit);
			return node;
		} else {
			const node = new MathNode("mrow");
			node.style.marginLeft = dimension.number + dimension.unit;
			return node;
		}
	}
});
const spaceCharacter = function(width) {
	if (width >= .05555 && width <= .05556) return " ";
	else if (width >= .1666 && width <= .1667) return " ";
	else if (width >= .2222 && width <= .2223) return " ";
	else if (width >= .2777 && width <= .2778) return "  ";
	else return "";
};
const invalidIdRegEx = /[^A-Za-z_0-9-]/g;
defineFunction({
	type: "label",
	names: ["\\label"],
	props: {
		numArgs: 1,
		argTypes: ["raw"]
	},
	handler({ parser }, args) {
		return {
			type: "label",
			mode: parser.mode,
			string: args[0].string.replace(invalidIdRegEx, "")
		};
	},
	mathmlBuilder(group, style) {
		const node = new MathNode("mrow", [], ["tml-label"]);
		if (group.string.length > 0) node.setLabel(group.string);
		return node;
	}
});
const textModeLap = [
	"\\clap",
	"\\llap",
	"\\rlap"
];
defineFunction({
	type: "lap",
	names: [
		"\\mathllap",
		"\\mathrlap",
		"\\mathclap",
		"\\clap",
		"\\llap",
		"\\rlap"
	],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler: ({ parser, funcName, token }, args) => {
		if (textModeLap.includes(funcName)) {
			if (parser.settings.strict && parser.mode !== "text") throw new ParseError(`{${funcName}} can be used only in text mode.
 Try \\math${funcName.slice(1)}`, token);
			funcName = funcName.slice(1);
		} else funcName = funcName.slice(5);
		const body = args[0];
		return {
			type: "lap",
			mode: parser.mode,
			alignment: funcName,
			body
		};
	},
	mathmlBuilder: (group, style) => {
		let strut;
		if (group.alignment === "llap") {
			strut = new MathNode("mpadded", [new MathNode("mphantom", buildExpression(ordargument(group.body), style))]);
			strut.setAttribute("width", "0.1px");
		}
		const inner = buildGroup$1(group.body, style);
		let node;
		if (group.alignment === "llap") {
			inner.style.position = "absolute";
			inner.style.right = "0";
			inner.style.bottom = `0`;
			node = new MathNode("mpadded", [strut, inner]);
		} else node = new MathNode("mpadded", [inner]);
		if (group.alignment === "rlap") {
			if (group.body.body.length > 0 && group.body.body[0].type === "genfrac") node.setAttribute("lspace", "0.16667em");
		} else {
			const offset = group.alignment === "llap" ? "-1" : "-0.5";
			node.setAttribute("lspace", offset + "width");
			if (group.alignment === "llap") node.style.position = "relative";
			else {
				node.style.display = "flex";
				node.style.justifyContent = "center";
			}
		}
		node.setAttribute("width", "0.1px");
		return node;
	}
});
defineFunction({
	type: "ordgroup",
	names: ["\\(", "$"],
	props: {
		numArgs: 0,
		allowedInText: true,
		allowedInMath: false
	},
	handler({ funcName, parser }, args) {
		const outerMode = parser.mode;
		parser.switchMode("math");
		const close = funcName === "\\(" ? "\\)" : "$";
		const body = parser.parseExpression(false, close);
		parser.expect(close);
		parser.switchMode(outerMode);
		return {
			type: "ordgroup",
			mode: parser.mode,
			body
		};
	}
});
defineFunction({
	type: "text",
	names: ["\\)", "\\]"],
	props: {
		numArgs: 0,
		allowedInText: true,
		allowedInMath: false
	},
	handler(context, token) {
		throw new ParseError(`Mismatched ${context.funcName}`, token);
	}
});
const chooseStyle = (group, style) => {
	switch (style.level) {
		case StyleLevel.DISPLAY: return group.display;
		case StyleLevel.TEXT: return group.text;
		case StyleLevel.SCRIPT: return group.script;
		case StyleLevel.SCRIPTSCRIPT: return group.scriptscript;
		default: return group.text;
	}
};
defineFunction({
	type: "mathchoice",
	names: ["\\mathchoice"],
	props: {
		numArgs: 4,
		primitive: true
	},
	handler: ({ parser }, args) => {
		return {
			type: "mathchoice",
			mode: parser.mode,
			display: ordargument(args[0]),
			text: ordargument(args[1]),
			script: ordargument(args[2]),
			scriptscript: ordargument(args[3])
		};
	},
	mathmlBuilder: (group, style) => {
		const body = chooseStyle(group, style);
		return buildExpressionRow(body, style);
	}
});
const textAtomTypes = [
	"text",
	"textord",
	"mathord",
	"atom"
];
function mathmlBuilder$3(group, style) {
	let node;
	const inner = buildExpression(group.body, style);
	if (group.mclass === "minner") node = new MathNode("mpadded", inner);
	else if (group.mclass === "mord") {
		if (group.isCharacterBox || inner[0].type === "mathord") {
			node = inner[0];
			node.type = "mi";
			if (node.children.length === 1 && node.children[0].text && node.children[0].text === "∇") node.setAttribute("mathvariant", "normal");
		} else node = new MathNode("mi", inner);
	} else {
		node = new MathNode("mrow", inner);
		if (group.mustPromote) {
			node = inner[0];
			node.type = "mo";
			if (group.isCharacterBox && group.body[0].text && /[A-Za-z]/.test(group.body[0].text)) node.setAttribute("mathvariant", "italic");
		} else node = new MathNode("mrow", inner);
		const doSpacing = style.level < 2;
		if (node.type === "mrow") {
			if (doSpacing) {
				if (group.mclass === "mbin") {
					node.children.unshift(padding(.2222));
					node.children.push(padding(.2222));
				} else if (group.mclass === "mrel") {
					node.children.unshift(padding(.2778));
					node.children.push(padding(.2778));
				} else if (group.mclass === "mpunct") node.children.push(padding(.1667));
				else if (group.mclass === "minner") {
					node.children.unshift(padding(.0556));
					node.children.push(padding(.0556));
				}
			}
		} else if (group.mclass === "mbin") {
			node.attributes.lspace = doSpacing ? "0.2222em" : "0";
			node.attributes.rspace = doSpacing ? "0.2222em" : "0";
		} else if (group.mclass === "mrel") {
			node.attributes.lspace = doSpacing ? "0.2778em" : "0";
			node.attributes.rspace = doSpacing ? "0.2778em" : "0";
		} else if (group.mclass === "mpunct") {
			node.attributes.lspace = "0em";
			node.attributes.rspace = doSpacing ? "0.1667em" : "0";
		} else if (group.mclass === "mopen" || group.mclass === "mclose") {
			node.attributes.lspace = "0em";
			node.attributes.rspace = "0em";
		} else if (group.mclass === "minner" && doSpacing) {
			node.attributes.lspace = "0.0556em";
			node.attributes.width = "+0.1111em";
		}
		if (!(group.mclass === "mopen" || group.mclass === "mclose")) {
			delete node.attributes.stretchy;
			delete node.attributes.form;
		}
	}
	return node;
}
defineFunction({
	type: "mclass",
	names: [
		"\\mathord",
		"\\mathbin",
		"\\mathrel",
		"\\mathopen",
		"\\mathclose",
		"\\mathpunct",
		"\\mathinner"
	],
	props: {
		numArgs: 1,
		primitive: true
	},
	handler({ parser, funcName }, args) {
		const body = args[0];
		const isCharacterBox$1 = isCharacterBox(body);
		let mustPromote = true;
		const mord = {
			type: "mathord",
			text: "",
			mode: parser.mode
		};
		const arr = body.body ? body.body : [body];
		for (const arg of arr) if (textAtomTypes.includes(arg.type)) {
			if (symbols[parser.mode][arg.text]) mord.text += symbols[parser.mode][arg.text].replace;
			else if (arg.text) mord.text += arg.text;
			else if (arg.body) arg.body.map((e) => {
				mord.text += e.text;
			});
		} else {
			mustPromote = false;
			break;
		}
		if (mustPromote && funcName === "\\mathord" && mord.type === "mathord" && mord.text.length > 1) return mord;
		else return {
			type: "mclass",
			mode: parser.mode,
			mclass: "m" + funcName.slice(5),
			body: ordargument(mustPromote ? mord : body),
			isCharacterBox: isCharacterBox$1,
			mustPromote
		};
	},
	mathmlBuilder: mathmlBuilder$3
});
const binrelClass = (arg) => {
	const atom = arg.type === "ordgroup" && arg.body.length && arg.body.length === 1 ? arg.body[0] : arg;
	if (atom.type === "atom") {
		const family = arg.body.length > 0 && arg.body[0].text && symbols.math[arg.body[0].text] ? symbols.math[arg.body[0].text].group : atom.family;
		if (family === "bin" || family === "rel") return "m" + family;
		else return "mord";
	} else return "mord";
};
defineFunction({
	type: "mclass",
	names: ["\\@binrel"],
	props: { numArgs: 2 },
	handler({ parser }, args) {
		return {
			type: "mclass",
			mode: parser.mode,
			mclass: binrelClass(args[0]),
			body: ordargument(args[1]),
			isCharacterBox: isCharacterBox(args[1])
		};
	}
});
defineFunction({
	type: "mclass",
	names: [
		"\\stackrel",
		"\\overset",
		"\\underset"
	],
	props: { numArgs: 2 },
	handler({ parser, funcName }, args) {
		const baseArg = args[1];
		const shiftedArg = args[0];
		let mclass;
		if (funcName !== "\\stackrel") mclass = binrelClass(baseArg);
		else mclass = "mrel";
		const baseOp = {
			type: mclass === "mrel" || mclass === "mbin" ? "op" : "ordgroup",
			mode: baseArg.mode,
			limits: true,
			alwaysHandleSupSub: true,
			parentIsSupSub: false,
			symbol: false,
			suppressBaseShift: funcName !== "\\stackrel",
			body: ordargument(baseArg)
		};
		return {
			type: "supsub",
			mode: shiftedArg.mode,
			stack: true,
			base: baseOp,
			sup: funcName === "\\underset" ? null : shiftedArg,
			sub: funcName === "\\underset" ? shiftedArg : null
		};
	},
	mathmlBuilder: mathmlBuilder$3
});
const buildGroup = (el, style, noneNode) => {
	if (!el) return noneNode;
	const node = buildGroup$1(el, style);
	if (node.type === "mrow" && node.children.length === 0) return noneNode;
	return node;
};
defineFunction({
	type: "multiscript",
	names: ["\\sideset", "\\pres@cript"],
	props: { numArgs: 3 },
	handler({ parser, funcName, token }, args) {
		if (args[2].body.length === 0) throw new ParseError(funcName + `cannot parse an empty base.`);
		const base = args[2].body[0];
		if (parser.settings.strict && funcName === "\\sideset" && !base.symbol) throw new ParseError(`The base of \\sideset must be a big operator. Try \\prescript.`);
		if (args[0].body.length > 0 && args[0].body[0].type !== "supsub" || args[1].body.length > 0 && args[1].body[0].type !== "supsub") throw new ParseError("\\sideset can parse only subscripts and superscripts in its first two arguments", token);
		const prescripts = args[0].body.length > 0 ? args[0].body[0] : null;
		const postscripts = args[1].body.length > 0 ? args[1].body[0] : null;
		if (!prescripts && !postscripts) return base;
		else if (!prescripts) return {
			type: "styling",
			mode: parser.mode,
			scriptLevel: "text",
			body: [{
				type: "supsub",
				mode: parser.mode,
				base,
				sup: postscripts.sup,
				sub: postscripts.sub
			}]
		};
		else return {
			type: "multiscript",
			mode: parser.mode,
			isSideset: funcName === "\\sideset",
			prescripts,
			postscripts,
			base
		};
	},
	mathmlBuilder(group, style) {
		const base = buildGroup$1(group.base, style);
		const prescriptsNode = new MathNode("mprescripts");
		const noneNode = new MathNode("none");
		let children = [];
		const preSub = buildGroup(group.prescripts.sub, style, noneNode);
		const preSup = buildGroup(group.prescripts.sup, style, noneNode);
		if (group.isSideset) {
			preSub.setAttribute("style", "text-align: left;");
			preSup.setAttribute("style", "text-align: left;");
		}
		if (group.postscripts) children = [
			base,
			buildGroup(group.postscripts.sub, style, noneNode),
			buildGroup(group.postscripts.sup, style, noneNode),
			prescriptsNode,
			preSub,
			preSup
		];
		else children = [
			base,
			prescriptsNode,
			preSub,
			preSup
		];
		return new MathNode("mmultiscripts", children);
	}
});
defineFunction({
	type: "not",
	names: ["\\not"],
	props: {
		numArgs: 1,
		primitive: true,
		allowedInText: false
	},
	handler({ parser }, args) {
		const isCharacterBox$1 = isCharacterBox(args[0]);
		let body;
		if (isCharacterBox$1) {
			body = ordargument(args[0]);
			if (body[0].text.charAt(0) === "\\") body[0].text = symbols.math[body[0].text].replace;
			body[0].text = body[0].text.slice(0, 1) + "̸" + body[0].text.slice(1);
		} else body = [
			{
				type: "textord",
				mode: "math",
				text: "̸"
			},
			{
				type: "kern",
				mode: "math",
				dimension: {
					number: -.6,
					unit: "em"
				}
			},
			args[0]
		];
		return {
			type: "not",
			mode: parser.mode,
			body,
			isCharacterBox: isCharacterBox$1
		};
	},
	mathmlBuilder(group, style) {
		if (group.isCharacterBox) return buildExpression(group.body, style, true)[0];
		else return buildExpressionRow(group.body, style);
	}
});
const ordAtomTypes = [
	"textord",
	"mathord",
	"atom"
];
const noSuccessor = ["\\smallint"];
const ordTypes = [
	"textord",
	"mathord",
	"ordgroup",
	"close",
	"leftright",
	"font"
];
const setSpacing = (node) => {
	node.attributes.lspace = "0.1667em";
	node.attributes.rspace = "0.1667em";
};
const mathmlBuilder$2 = (group, style) => {
	let node;
	if (group.symbol) {
		node = new MathNode("mo", [makeText(group.name, group.mode)]);
		if (noSuccessor.includes(group.name)) node.setAttribute("largeop", "false");
		else node.setAttribute("movablelimits", "false");
		if (group.fromMathOp) setSpacing(node);
	} else if (group.body) {
		node = new MathNode("mo", buildExpression(group.body, style));
		if (group.fromMathOp) setSpacing(node);
	} else {
		node = new MathNode("mi", [new TextNode(group.name.slice(1))]);
		if (!group.parentIsSupSub) {
			const operator = new MathNode("mo", [makeText("⁡", "text")]);
			const row = [node, operator];
			if (group.needsLeadingSpace) {
				const lead = new MathNode("mspace");
				lead.setAttribute("width", "0.1667em");
				row.unshift(lead);
			}
			if (!group.isFollowedByDelimiter) {
				const trail = new MathNode("mspace");
				trail.setAttribute("width", "0.1667em");
				row.push(trail);
			}
			node = new MathNode("mrow", row);
		}
	}
	return node;
};
const singleCharBigOps = {
	"∏": "\\prod",
	"∐": "\\coprod",
	"∑": "\\sum",
	"⋀": "\\bigwedge",
	"⋁": "\\bigvee",
	"⋂": "\\bigcap",
	"⋃": "\\bigcup",
	"⨀": "\\bigodot",
	"⨁": "\\bigoplus",
	"⨂": "\\bigotimes",
	"⨄": "\\biguplus",
	"⨅": "\\bigsqcap",
	"⨆": "\\bigsqcup",
	"⨃": "\\bigcupdot",
	"⨇": "\\bigdoublevee",
	"⨈": "\\bigdoublewedge",
	"⨉": "\\bigtimes"
};
defineFunction({
	type: "op",
	names: [
		"\\coprod",
		"\\bigvee",
		"\\bigwedge",
		"\\biguplus",
		"\\bigcupplus",
		"\\bigcupdot",
		"\\bigcap",
		"\\bigcup",
		"\\bigdoublevee",
		"\\bigdoublewedge",
		"\\intop",
		"\\prod",
		"\\sum",
		"\\bigotimes",
		"\\bigoplus",
		"\\bigodot",
		"\\bigsqcap",
		"\\bigsqcup",
		"\\bigtimes",
		"\\smallint",
		"∏",
		"∐",
		"∑",
		"⋀",
		"⋁",
		"⋂",
		"⋃",
		"⨀",
		"⨁",
		"⨂",
		"⨃",
		"⨄",
		"⨅",
		"⨆",
		"⨇",
		"⨈",
		"⨉"
	],
	props: { numArgs: 0 },
	handler: ({ parser, funcName }, args) => {
		let fName = funcName;
		if (fName.length === 1) fName = singleCharBigOps[fName];
		return {
			type: "op",
			mode: parser.mode,
			limits: true,
			parentIsSupSub: false,
			symbol: true,
			stack: false,
			name: fName
		};
	},
	mathmlBuilder: mathmlBuilder$2
});
defineFunction({
	type: "op",
	names: ["\\mathop"],
	props: {
		numArgs: 1,
		primitive: true
	},
	handler: ({ parser }, args) => {
		const body = args[0];
		const arr = body.body ? body.body : [body];
		const isSymbol = arr.length === 1 && ordAtomTypes.includes(arr[0].type);
		return {
			type: "op",
			mode: parser.mode,
			limits: true,
			parentIsSupSub: false,
			symbol: isSymbol,
			fromMathOp: true,
			stack: false,
			name: isSymbol ? arr[0].text : null,
			body: isSymbol ? null : ordargument(body)
		};
	},
	mathmlBuilder: mathmlBuilder$2
});
const singleCharIntegrals = {
	"∫": "\\int",
	"∬": "\\iint",
	"∭": "\\iiint",
	"∮": "\\oint",
	"∯": "\\oiint",
	"∰": "\\oiiint",
	"∱": "\\intclockwise",
	"∲": "\\varointclockwise",
	"⨌": "\\iiiint",
	"⨍": "\\intbar",
	"⨎": "\\intBar",
	"⨏": "\\fint",
	"⨒": "\\rppolint",
	"⨓": "\\scpolint",
	"⨕": "\\pointint",
	"⨖": "\\sqint",
	"⨗": "\\intlarhk",
	"⨘": "\\intx",
	"⨙": "\\intcap",
	"⨚": "\\intcup"
};
defineFunction({
	type: "op",
	names: [
		"\\arcsin",
		"\\arccos",
		"\\arctan",
		"\\arctg",
		"\\arcctg",
		"\\arg",
		"\\ch",
		"\\cos",
		"\\cosec",
		"\\cosh",
		"\\cot",
		"\\cotg",
		"\\coth",
		"\\csc",
		"\\ctg",
		"\\cth",
		"\\deg",
		"\\dim",
		"\\exp",
		"\\hom",
		"\\ker",
		"\\lg",
		"\\ln",
		"\\log",
		"\\sec",
		"\\sin",
		"\\sinh",
		"\\sh",
		"\\sgn",
		"\\tan",
		"\\tanh",
		"\\tg",
		"\\th"
	],
	props: { numArgs: 0 },
	handler({ parser, funcName }) {
		const prevAtomType = parser.prevAtomType;
		const next = parser.gullet.future().text;
		return {
			type: "op",
			mode: parser.mode,
			limits: false,
			parentIsSupSub: false,
			symbol: false,
			stack: false,
			isFollowedByDelimiter: isDelimiter(next),
			needsLeadingSpace: prevAtomType.length > 0 && ordTypes.includes(prevAtomType),
			name: funcName
		};
	},
	mathmlBuilder: mathmlBuilder$2
});
defineFunction({
	type: "op",
	names: [
		"\\det",
		"\\gcd",
		"\\inf",
		"\\lim",
		"\\max",
		"\\min",
		"\\Pr",
		"\\sup"
	],
	props: { numArgs: 0 },
	handler({ parser, funcName }) {
		const prevAtomType = parser.prevAtomType;
		const next = parser.gullet.future().text;
		return {
			type: "op",
			mode: parser.mode,
			limits: true,
			parentIsSupSub: false,
			symbol: false,
			stack: false,
			isFollowedByDelimiter: isDelimiter(next),
			needsLeadingSpace: prevAtomType.length > 0 && ordTypes.includes(prevAtomType),
			name: funcName
		};
	},
	mathmlBuilder: mathmlBuilder$2
});
defineFunction({
	type: "op",
	names: [
		"\\int",
		"\\iint",
		"\\iiint",
		"\\iiiint",
		"\\oint",
		"\\oiint",
		"\\oiiint",
		"\\intclockwise",
		"\\varointclockwise",
		"\\intbar",
		"\\intBar",
		"\\fint",
		"\\rppolint",
		"\\scpolint",
		"\\pointint",
		"\\sqint",
		"\\intlarhk",
		"\\intx",
		"\\intcap",
		"\\intcup",
		"∫",
		"∬",
		"∭",
		"∮",
		"∯",
		"∰",
		"∱",
		"∲",
		"⨌",
		"⨍",
		"⨎",
		"⨏",
		"⨒",
		"⨓",
		"⨕",
		"⨖",
		"⨗",
		"⨘",
		"⨙",
		"⨚"
	],
	props: {
		numArgs: 0,
		allowedInArgument: true
	},
	handler({ parser, funcName }) {
		let fName = funcName;
		if (fName.length === 1) fName = singleCharIntegrals[fName];
		return {
			type: "op",
			mode: parser.mode,
			limits: false,
			parentIsSupSub: false,
			symbol: true,
			stack: false,
			name: fName
		};
	},
	mathmlBuilder: mathmlBuilder$2
});
const mathmlBuilder$1 = (group, style) => {
	let expression = buildExpression(group.body, style.withFont("mathrm"));
	let isAllString = true;
	for (let i = 0; i < expression.length; i++) {
		let node = expression[i];
		if (node instanceof MathNode) {
			if ((node.type === "mrow" || node.type === "mpadded") && node.children.length === 1 && node.children[0] instanceof MathNode) node = node.children[0];
			else if (node.type === "mrow" && node.children.length === 2 && node.children[0] instanceof MathNode && node.children[1] instanceof MathNode && node.children[1].type === "mspace" && !node.children[1].attributes.width && node.children[1].children.length === 0) node = node.children[0];
			switch (node.type) {
				case "mi":
				case "mn":
				case "ms":
				case "mtext": break;
				case "mspace":
					if (node.attributes.width) {
						const width = node.attributes.width.replace("em", "");
						const ch = spaceCharacter(Number(width));
						if (ch === "") isAllString = false;
						else expression[i] = new MathNode("mtext", [new TextNode(ch)]);
					}
					break;
				case "mo": {
					const child = node.children[0];
					if (node.children.length === 1 && child instanceof TextNode) child.text = child.text.replace(/\u2212/, "-").replace(/\u2217/, "*");
					else isAllString = false;
					break;
				}
				default: isAllString = false;
			}
		} else isAllString = false;
	}
	if (isAllString) expression = [new TextNode(expression.map((node) => node.toText()).join(""))];
	else if (expression.length === 1 && ["mover", "munder"].includes(expression[0].type) && (expression[0].children[0].type === "mi" || expression[0].children[0].type === "mtext")) {
		expression[0].children[0].type = "mi";
		if (group.parentIsSupSub) return new MathNode("mrow", expression);
		else {
			const operator = new MathNode("mo", [makeText("⁡", "text")]);
			return newDocumentFragment([expression[0], operator]);
		}
	}
	let wrapper;
	if (isAllString) {
		wrapper = new MathNode("mi", expression);
		if (expression[0].text.length === 1) wrapper.setAttribute("mathvariant", "normal");
	} else wrapper = new MathNode("mrow", expression);
	if (!group.parentIsSupSub) {
		const operator = new MathNode("mo", [makeText("⁡", "text")]);
		const fragment = [wrapper, operator];
		if (group.needsLeadingSpace) {
			const space = new MathNode("mspace");
			space.setAttribute("width", "0.1667em");
			fragment.unshift(space);
		}
		if (!group.isFollowedByDelimiter) {
			const trail = new MathNode("mspace");
			trail.setAttribute("width", "0.1667em");
			fragment.push(trail);
		}
		return newDocumentFragment(fragment);
	}
	return wrapper;
};
defineFunction({
	type: "operatorname",
	names: ["\\operatorname@", "\\operatornamewithlimits"],
	props: {
		numArgs: 1,
		allowedInArgument: true
	},
	handler: ({ parser, funcName }, args) => {
		const body = args[0];
		const prevAtomType = parser.prevAtomType;
		const next = parser.gullet.future().text;
		return {
			type: "operatorname",
			mode: parser.mode,
			body: ordargument(body),
			alwaysHandleSupSub: funcName === "\\operatornamewithlimits",
			limits: false,
			parentIsSupSub: false,
			isFollowedByDelimiter: isDelimiter(next),
			needsLeadingSpace: prevAtomType.length > 0 && ordTypes.includes(prevAtomType)
		};
	},
	mathmlBuilder: mathmlBuilder$1
});
defineMacro("\\operatorname", "\\@ifstar\\operatornamewithlimits\\operatorname@");
defineFunctionBuilders({
	type: "ordgroup",
	mathmlBuilder(group, style) {
		return buildExpressionRow(group.body, style, group.semisimple);
	}
});
defineFunction({
	type: "phantom",
	names: ["\\phantom"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler: ({ parser }, args) => {
		const body = args[0];
		return {
			type: "phantom",
			mode: parser.mode,
			body: ordargument(body)
		};
	},
	mathmlBuilder: (group, style) => {
		return new MathNode("mphantom", buildExpression(group.body, style));
	}
});
defineFunction({
	type: "hphantom",
	names: ["\\hphantom"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler: ({ parser }, args) => {
		const body = args[0];
		return {
			type: "hphantom",
			mode: parser.mode,
			body
		};
	},
	mathmlBuilder: (group, style) => {
		const node = new MathNode("mpadded", [new MathNode("mphantom", buildExpression(ordargument(group.body), style))]);
		node.setAttribute("height", "0px");
		node.setAttribute("depth", "0px");
		return node;
	}
});
defineFunction({
	type: "vphantom",
	names: ["\\vphantom"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler: ({ parser }, args) => {
		const body = args[0];
		return {
			type: "vphantom",
			mode: parser.mode,
			body
		};
	},
	mathmlBuilder: (group, style) => {
		const node = new MathNode("mpadded", [new MathNode("mphantom", buildExpression(ordargument(group.body), style))]);
		node.setAttribute("width", "0.1px");
		return node;
	}
});
defineFunction({
	type: "pmb",
	names: ["\\pmb"],
	props: {
		numArgs: 1,
		allowedInText: true
	},
	handler({ parser }, args) {
		return {
			type: "pmb",
			mode: parser.mode,
			body: ordargument(args[0])
		};
	},
	mathmlBuilder(group, style) {
		const inner = buildExpression(group.body, style);
		const node = wrapWithMstyle(inner);
		node.setAttribute("style", "font-weight:bold");
		return node;
	}
});
const mathmlBuilder = (group, style) => {
	const newStyle = style.withLevel(StyleLevel.TEXT);
	const node = new MathNode("mpadded", [buildGroup$1(group.body, newStyle)]);
	const dy = calculateSize(group.dy, style);
	node.setAttribute("voffset", dy.number + dy.unit);
	if (dy.number > 0) node.style.padding = dy.number + dy.unit + " 0 0 0";
	else node.style.padding = "0 0 " + Math.abs(dy.number) + dy.unit + " 0";
	return node;
};
defineFunction({
	type: "raise",
	names: ["\\raise", "\\lower"],
	props: {
		numArgs: 2,
		argTypes: ["size", "primitive"],
		primitive: true
	},
	handler({ parser, funcName }, args) {
		const amount = assertNodeType(args[0], "size").value;
		if (funcName === "\\lower") amount.number *= -1;
		const body = args[1];
		return {
			type: "raise",
			mode: parser.mode,
			dy: amount,
			body
		};
	},
	mathmlBuilder
});
defineFunction({
	type: "raise",
	names: ["\\raisebox"],
	props: {
		numArgs: 2,
		argTypes: ["size", "hbox"],
		allowedInText: true
	},
	handler({ parser, funcName }, args) {
		const amount = assertNodeType(args[0], "size").value;
		const body = args[1];
		return {
			type: "raise",
			mode: parser.mode,
			dy: amount,
			body
		};
	},
	mathmlBuilder
});
defineFunction({
	type: "ref",
	names: ["\\ref", "\\eqref"],
	props: {
		numArgs: 1,
		argTypes: ["raw"]
	},
	handler({ parser, funcName }, args) {
		return {
			type: "ref",
			mode: parser.mode,
			funcName,
			string: args[0].string.replace(invalidIdRegEx, "")
		};
	},
	mathmlBuilder(group, style) {
		const classes = group.funcName === "\\ref" ? ["tml-ref"] : ["tml-ref", "tml-eqref"];
		return new AnchorNode("#" + group.string, classes, null);
	}
});
defineFunction({
	type: "reflect",
	names: ["\\reflectbox"],
	props: {
		numArgs: 1,
		argTypes: ["hbox"],
		allowedInText: true
	},
	handler({ parser }, args) {
		return {
			type: "reflect",
			mode: parser.mode,
			body: args[0]
		};
	},
	mathmlBuilder(group, style) {
		const node = buildGroup$1(group.body, style);
		node.style.transform = "scaleX(-1)";
		return node;
	}
});
defineFunction({
	type: "internal",
	names: ["\\relax"],
	props: {
		numArgs: 0,
		allowedInText: true,
		allowedInArgument: true
	},
	handler({ parser }) {
		return {
			type: "internal",
			mode: parser.mode
		};
	}
});
defineFunction({
	type: "rule",
	names: ["\\rule"],
	props: {
		numArgs: 2,
		numOptionalArgs: 1,
		allowedInText: true,
		allowedInMath: true,
		argTypes: [
			"size",
			"size",
			"size"
		]
	},
	handler({ parser }, args, optArgs) {
		const shift = optArgs[0];
		const width = assertNodeType(args[0], "size");
		const height = assertNodeType(args[1], "size");
		return {
			type: "rule",
			mode: parser.mode,
			shift: shift && assertNodeType(shift, "size").value,
			width: width.value,
			height: height.value
		};
	},
	mathmlBuilder(group, style) {
		const width = calculateSize(group.width, style);
		const height = calculateSize(group.height, style);
		const shift = group.shift ? calculateSize(group.shift, style) : {
			number: 0,
			unit: "em"
		};
		const color = style.color && style.getColor() || "black";
		const rule = new MathNode("mspace");
		if (width.number > 0 && height.number > 0) rule.setAttribute("mathbackground", color);
		rule.setAttribute("width", width.number + width.unit);
		rule.setAttribute("height", height.number + height.unit);
		if (shift.number === 0) return rule;
		const wrapper = new MathNode("mpadded", [rule]);
		if (shift.number >= 0) wrapper.setAttribute("height", "+" + shift.number + shift.unit);
		else {
			wrapper.setAttribute("height", shift.number + shift.unit);
			wrapper.setAttribute("depth", "+" + -shift.number + shift.unit);
		}
		wrapper.setAttribute("voffset", shift.number + shift.unit);
		return wrapper;
	}
});
const numRegEx = /^[0-9]$/;
const unicodeNumSubs = {
	"0": "₀",
	"1": "₁",
	"2": "₂",
	"3": "₃",
	"4": "₄",
	"5": "₅",
	"6": "₆",
	"7": "₇",
	"8": "₈",
	"9": "₉"
};
const unicodeNumSups = {
	"0": "⁰",
	"1": "¹",
	"2": "²",
	"3": "³",
	"4": "⁴",
	"5": "⁵",
	"6": "⁶",
	"7": "⁷",
	"8": "⁸",
	"9": "⁹"
};
defineFunction({
	type: "sfrac",
	names: ["\\sfrac"],
	props: {
		numArgs: 2,
		allowedInText: true,
		allowedInMath: true
	},
	handler({ parser }, args) {
		let numerator = "";
		for (const node of args[0].body) {
			if (node.type !== "textord" || !numRegEx.test(node.text)) throw new ParseError("Numerator must be an integer.", node);
			numerator += node.text;
		}
		let denominator = "";
		for (const node of args[1].body) {
			if (node.type !== "textord" || !numRegEx.test(node.text)) throw new ParseError("Denominator must be an integer.", node);
			denominator += node.text;
		}
		return {
			type: "sfrac",
			mode: parser.mode,
			numerator,
			denominator
		};
	},
	mathmlBuilder(group, style) {
		const numerator = group.numerator.split("").map((c) => unicodeNumSups[c]).join("");
		const denominator = group.denominator.split("").map((c) => unicodeNumSubs[c]).join("");
		return new MathNode("mn", [new TextNode(numerator + "⁄" + denominator, group.mode, style)], ["special-fraction"]);
	}
});
const sizeMap = {
	"\\tiny": .5,
	"\\sixptsize": .6,
	"\\Tiny": .6,
	"\\scriptsize": .7,
	"\\footnotesize": .8,
	"\\small": .9,
	"\\normalsize": 1,
	"\\large": 1.2,
	"\\Large": 1.44,
	"\\LARGE": 1.728,
	"\\huge": 2.074,
	"\\Huge": 2.488
};
defineFunction({
	type: "sizing",
	names: [
		"\\tiny",
		"\\sixptsize",
		"\\Tiny",
		"\\scriptsize",
		"\\footnotesize",
		"\\small",
		"\\normalsize",
		"\\large",
		"\\Large",
		"\\LARGE",
		"\\huge",
		"\\Huge"
	],
	props: {
		numArgs: 0,
		allowedInText: true
	},
	handler: ({ breakOnTokenText, funcName, parser }, args) => {
		if (parser.settings.strict && parser.mode === "math") console.log(`Temml strict-mode warning: Command ${funcName} is invalid in math mode.`);
		const body = parser.parseExpression(false, breakOnTokenText, true);
		return {
			type: "sizing",
			mode: parser.mode,
			funcName,
			body
		};
	},
	mathmlBuilder: (group, style) => {
		const newStyle = style.withFontSize(sizeMap[group.funcName]);
		const inner = buildExpression(group.body, newStyle);
		const node = wrapWithMstyle(inner);
		const factor = (sizeMap[group.funcName] / style.fontSize).toFixed(4);
		node.setAttribute("mathsize", factor + "em");
		return node;
	}
});
defineFunction({
	type: "smash",
	names: ["\\smash"],
	props: {
		numArgs: 1,
		numOptionalArgs: 1,
		allowedInText: true
	},
	handler: ({ parser }, args, optArgs) => {
		let smashHeight = false;
		let smashDepth = false;
		const tbArg = optArgs[0] && assertNodeType(optArgs[0], "ordgroup");
		if (tbArg) {
			let letter = "";
			for (let i = 0; i < tbArg.body.length; ++i) {
				letter = tbArg.body[i].text;
				if (letter === "t") smashHeight = true;
				else if (letter === "b") smashDepth = true;
				else {
					smashHeight = false;
					smashDepth = false;
					break;
				}
			}
		} else {
			smashHeight = true;
			smashDepth = true;
		}
		const body = args[0];
		return {
			type: "smash",
			mode: parser.mode,
			body,
			smashHeight,
			smashDepth
		};
	},
	mathmlBuilder: (group, style) => {
		const node = new MathNode("mpadded", [buildGroup$1(group.body, style)]);
		if (group.smashHeight) node.setAttribute("height", "0px");
		if (group.smashDepth) node.setAttribute("depth", "0px");
		return node;
	}
});
const xHeights = [
	"a",
	"c",
	"e",
	"ı",
	"m",
	"n",
	"o",
	"r",
	"s",
	"u",
	"v",
	"w",
	"x",
	"z",
	"α",
	"ε",
	"ι",
	"κ",
	"ν",
	"ο",
	"π",
	"σ",
	"τ",
	"υ",
	"ω",
	"\\alpha",
	"\\epsilon",
	"\\iota",
	"\\kappa",
	"\\nu",
	"\\omega",
	"\\pi",
	"\\tau",
	"\\omega"
];
defineFunction({
	type: "sqrt",
	names: ["\\sqrt"],
	props: {
		numArgs: 1,
		numOptionalArgs: 1
	},
	handler({ parser }, args, optArgs) {
		const index = optArgs[0];
		const body = args[0];
		if (body.body && body.body.length === 1 && body.body[0].text && xHeights.includes(body.body[0].text)) body.body.push({
			"type": "rule",
			"mode": "math",
			"shift": null,
			"width": {
				"number": 0,
				"unit": "pt"
			},
			"height": {
				"number": .5,
				"unit": "em"
			}
		});
		return {
			type: "sqrt",
			mode: parser.mode,
			body,
			index
		};
	},
	mathmlBuilder(group, style) {
		const { body, index } = group;
		return index ? new MathNode("mroot", [buildGroup$1(body, style), buildGroup$1(index, style.incrementLevel())]) : new MathNode("msqrt", [buildGroup$1(body, style)]);
	}
});
const styleMap = {
	display: 0,
	text: 1,
	script: 2,
	scriptscript: 3
};
const styleAttributes = {
	display: ["0", "true"],
	text: ["0", "false"],
	script: ["1", "false"],
	scriptscript: ["2", "false"]
};
defineFunction({
	type: "styling",
	names: [
		"\\displaystyle",
		"\\textstyle",
		"\\scriptstyle",
		"\\scriptscriptstyle"
	],
	props: {
		numArgs: 0,
		allowedInText: true,
		primitive: true
	},
	handler({ breakOnTokenText, funcName, parser }, args) {
		const body = parser.parseExpression(true, breakOnTokenText, true);
		const scriptLevel = funcName.slice(1, funcName.length - 5);
		return {
			type: "styling",
			mode: parser.mode,
			scriptLevel,
			body
		};
	},
	mathmlBuilder(group, style) {
		const newStyle = style.withLevel(styleMap[group.scriptLevel]);
		const inner = buildExpression(group.body, newStyle);
		const node = wrapWithMstyle(inner);
		const attr = styleAttributes[group.scriptLevel];
		node.setAttribute("scriptlevel", attr[0]);
		node.setAttribute("displaystyle", attr[1]);
		return node;
	}
});
/**
* Sometimes, groups perform special rules when they have superscripts or
* subscripts attached to them. This function lets the `supsub` group know that
* Sometimes, groups perform special rules when they have superscripts or
* its inner element should handle the superscripts and subscripts instead of
* handling them itself.
*/
const symbolRegEx = /^m(over|under|underover)$/;
const smallPad = "DHKLUcegorsuvxyzΠΥΨαδηιμνοτυχϵ";
const mediumPad = "BCEFGIMNOPQRSTXZlpqtwΓΘΞΣΦΩβεζθξρςφψϑϕϱ";
const largePad = "AJdfΔΛ";
defineFunctionBuilders({
	type: "supsub",
	mathmlBuilder(group, style) {
		let isBracket = false;
		let isOver;
		let isSup;
		let appendApplyFunction = false;
		let appendSpace = false;
		let needsLeadingSpace = false;
		if (group.base && group.base.type === "horizBracket") {
			isSup = !!group.sup;
			if (isSup === group.base.isOver) {
				isBracket = true;
				isOver = group.base.isOver;
			}
		}
		if (group.base && !group.stack && (group.base.type === "op" || group.base.type === "operatorname")) {
			group.base.parentIsSupSub = true;
			appendApplyFunction = !group.base.symbol;
			appendSpace = appendApplyFunction && !group.isFollowedByDelimiter;
			needsLeadingSpace = group.base.needsLeadingSpace;
		}
		const children = group.stack && group.base.body.length === 1 ? [buildGroup$1(group.base.body[0], style)] : [buildGroup$1(group.base, style)];
		const childStyle = style.inSubOrSup();
		if (group.sub) {
			const sub = buildGroup$1(group.sub, childStyle);
			if (style.level === 3) sub.setAttribute("scriptlevel", "2");
			children.push(sub);
		}
		if (group.sup) {
			const sup = buildGroup$1(group.sup, childStyle);
			if (style.level === 3) sup.setAttribute("scriptlevel", "2");
			if (group.base && group.base.text && group.base.text.length === 1) {
				const text = group.base.text;
				if (smallPad.indexOf(text) > -1) sup.classes.push("tml-sml-pad");
				else if (mediumPad.indexOf(text) > -1) sup.classes.push("tml-med-pad");
				else if (largePad.indexOf(text) > -1) sup.classes.push("tml-lrg-pad");
			}
			children.push(sup);
		}
		let nodeType;
		if (isBracket) nodeType = isOver ? "mover" : "munder";
		else if (!group.sub) {
			const base = group.base;
			if (base && base.type === "op" && base.limits && (style.level === StyleLevel.DISPLAY || base.alwaysHandleSupSub)) nodeType = "mover";
			else if (base && base.type === "operatorname" && base.alwaysHandleSupSub && (base.limits || style.level === StyleLevel.DISPLAY)) nodeType = "mover";
			else nodeType = "msup";
		} else if (!group.sup) {
			const base = group.base;
			if (group.stack) nodeType = "munder";
			else if (base && base.type === "op" && base.limits && (style.level === StyleLevel.DISPLAY || base.alwaysHandleSupSub)) nodeType = "munder";
			else if (base && base.type === "operatorname" && base.alwaysHandleSupSub && (base.limits || style.level === StyleLevel.DISPLAY)) nodeType = "munder";
			else nodeType = "msub";
		} else {
			const base = group.base;
			if (base && (base.type === "op" && base.limits || base.type === "multiscript") && (style.level === StyleLevel.DISPLAY || base.alwaysHandleSupSub)) nodeType = "munderover";
			else if (base && base.type === "operatorname" && base.alwaysHandleSupSub && (style.level === StyleLevel.DISPLAY || base.limits)) nodeType = "munderover";
			else nodeType = "msubsup";
		}
		let node = new MathNode(nodeType, children);
		if (appendApplyFunction) {
			const operator = new MathNode("mo", [makeText("⁡", "text")]);
			if (needsLeadingSpace) {
				const space = new MathNode("mspace");
				space.setAttribute("width", "0.1667em");
				node = newDocumentFragment([
					space,
					node,
					operator
				]);
			} else node = newDocumentFragment([node, operator]);
			if (appendSpace) {
				const space = new MathNode("mspace");
				space.setAttribute("width", "0.1667em");
				node.children.push(space);
			}
		} else if (symbolRegEx.test(nodeType)) node = new MathNode("mrow", [node]);
		return node;
	}
});
const short = [
	"\\shortmid",
	"\\nshortmid",
	"\\shortparallel",
	"\\nshortparallel",
	"\\smallsetminus"
];
const arrows = [
	"\\Rsh",
	"\\Lsh",
	"\\restriction"
];
const isArrow = (str) => {
	if (str.length === 1) {
		const codePoint = str.codePointAt(0);
		return 8591 < codePoint && codePoint < 8704;
	}
	return str.indexOf("arrow") > -1 || str.indexOf("harpoon") > -1 || arrows.includes(str);
};
defineFunctionBuilders({
	type: "atom",
	mathmlBuilder(group, style) {
		const node = new MathNode("mo", [makeText(group.text, group.mode)]);
		if (group.family === "punct") node.setAttribute("separator", "true");
		else if (group.family === "open" || group.family === "close") {
			if (group.family === "open") {
				node.setAttribute("form", "prefix");
				node.setAttribute("stretchy", "false");
			} else if (group.family === "close") {
				node.setAttribute("form", "postfix");
				node.setAttribute("stretchy", "false");
			}
		} else if (group.text === "\\mid") {
			node.setAttribute("lspace", "0.22em");
			node.setAttribute("rspace", "0.22em");
			node.setAttribute("stretchy", "false");
		} else if (group.family === "rel" && isArrow(group.text)) node.setAttribute("stretchy", "false");
		else if (short.includes(group.text)) node.setAttribute("mathsize", "70%");
		else if (group.text === ":") {
			node.attributes.lspace = "0.2222em";
			node.attributes.rspace = "0.2222em";
		} else if (group.needsSpacing) {
			if (group.family === "bin") return new MathNode("mrow", [
				padding(.222),
				node,
				padding(.222)
			]);
			else return new MathNode("mrow", [
				padding(.2778),
				node,
				padding(.2778)
			]);
		}
		return node;
	}
});
/**
* Maps TeX font commands to "mathvariant" attribute in buildMathML.js
*/
const fontMap = {
	mathbf: "bold",
	mathrm: "normal",
	textit: "italic",
	mathit: "italic",
	mathnormal: "italic",
	mathbb: "double-struck",
	mathcal: "script",
	mathfrak: "fraktur",
	mathscr: "script",
	mathsf: "sans-serif",
	mathtt: "monospace"
};
/**
* Returns the math variant as a string or null if none is required.
*/
const getVariant = function(group, style) {
	if (style.fontFamily === "texttt") return "monospace";
	else if (style.fontFamily === "textsc") return "normal";
	else if (style.fontFamily === "textsf") {
		if (style.fontShape === "textit" && style.fontWeight === "textbf") return "sans-serif-bold-italic";
		else if (style.fontShape === "textit") return "sans-serif-italic";
		else if (style.fontWeight === "textbf") return "sans-serif-bold";
		else return "sans-serif";
	} else if (style.fontShape === "textit" && style.fontWeight === "textbf") return "bold-italic";
	else if (style.fontShape === "textit") return "italic";
	else if (style.fontWeight === "textbf") return "bold";
	const font = style.font;
	if (!font || font === "mathnormal") return null;
	const mode = group.mode;
	switch (font) {
		case "mathit": return "italic";
		case "mathrm": {
			const codePoint = group.text.codePointAt(0);
			return 939 < codePoint && codePoint < 975 ? "italic" : "normal";
		}
		case "greekItalic": return "italic";
		case "up@greek": return "normal";
		case "boldsymbol":
		case "mathboldsymbol": return "bold-italic";
		case "mathbf": return "bold";
		case "mathbb": return "double-struck";
		case "mathfrak": return "fraktur";
		case "mathscr":
		case "mathcal": return "script";
		case "mathsf": return "sans-serif";
		case "mathsfit": return "sans-serif-italic";
		case "mathtt": return "monospace";
	}
	let text = group.text;
	if (symbols[mode][text] && symbols[mode][text].replace) text = symbols[mode][text].replace;
	return Object.prototype.hasOwnProperty.call(fontMap, font) ? fontMap[font] : null;
};
const numberRegEx = /^\d(?:[\d,.]*\d)?$/;
const latinRegEx = /[A-Ba-z]/;
const primes = /* @__PURE__ */ new Set([
	"\\prime",
	"\\dprime",
	"\\trprime",
	"\\qprime",
	"\\backprime",
	"\\backdprime",
	"\\backtrprime"
]);
const italicNumber = (text, variant, tag) => {
	const wrapper = new MathNode("mstyle", [new MathNode(tag, [text])]);
	wrapper.style["font-style"] = "italic";
	wrapper.style["font-family"] = "Cambria, 'Times New Roman', serif";
	if (variant === "bold-italic") wrapper.style["font-weight"] = "bold";
	return wrapper;
};
defineFunctionBuilders({
	type: "mathord",
	mathmlBuilder(group, style) {
		const text = makeText(group.text, group.mode, style);
		const codePoint = text.text.codePointAt(0);
		const defaultVariant = 912 < codePoint && codePoint < 938 ? "normal" : "italic";
		const variant = getVariant(group, style) || defaultVariant;
		if (variant === "script") {
			text.text = variantChar(text.text, variant);
			return new MathNode("mi", [text], [style.font]);
		} else if (variant !== "italic") text.text = variantChar(text.text, variant);
		let node = new MathNode("mi", [text]);
		if (variant === "normal") {
			node.setAttribute("mathvariant", "normal");
			if (text.text.length === 1) {
				const mspace = new MathNode("mspace", []);
				node = new MathNode("mrow", [node, mspace]);
			}
		}
		return node;
	}
});
defineFunctionBuilders({
	type: "textord",
	mathmlBuilder(group, style) {
		let ch = group.text;
		const codePoint = ch.codePointAt(0);
		if (style.fontFamily === "textsc") {
			if (96 < codePoint && codePoint < 123) ch = smallCaps[ch];
		}
		const text = makeText(ch, group.mode, style);
		const variant = getVariant(group, style) || "normal";
		let node;
		if (numberRegEx.test(group.text)) {
			const tag = group.mode === "text" ? "mtext" : "mn";
			if (variant === "italic" || variant === "bold-italic") return italicNumber(text, variant, tag);
			else {
				if (variant !== "normal") text.text = text.text.split("").map((c) => variantChar(c, variant)).join("");
				node = new MathNode(tag, [text]);
			}
		} else if (group.mode === "text") {
			if (variant !== "normal") text.text = variantChar(text.text, variant);
			node = new MathNode("mtext", [text]);
		} else if (primes.has(group.text)) {
			node = new MathNode("mo", [text]);
			node.classes.push("tml-prime");
		} else {
			const origText = text.text;
			if (variant !== "italic") text.text = variantChar(text.text, variant);
			node = new MathNode("mi", [text]);
			if (text.text === origText && latinRegEx.test(origText)) node.setAttribute("mathvariant", "italic");
		}
		return node;
	}
});
const cssSpace = {
	"\\nobreak": "nobreak",
	"\\allowbreak": "allowbreak"
};
const regularSpace = {
	" ": {},
	"\\ ": {},
	"~": { className: "nobreak" },
	"\\space": {},
	"\\nobreakspace": { className: "nobreak" }
};
defineFunctionBuilders({
	type: "spacing",
	mathmlBuilder(group, style) {
		let node;
		if (Object.prototype.hasOwnProperty.call(regularSpace, group.text)) node = new MathNode("mtext", [new TextNode("\xA0")]);
		else if (Object.prototype.hasOwnProperty.call(cssSpace, group.text)) {
			node = new MathNode("mo");
			if (group.text === "\\nobreak") node.setAttribute("linebreak", "nobreak");
		} else throw new ParseError(`Unknown type of space "${group.text}"`);
		return node;
	}
});
defineFunctionBuilders({ type: "tag" });
const textFontFamilies = {
	"\\text": void 0,
	"\\textrm": "textrm",
	"\\textsf": "textsf",
	"\\texttt": "texttt",
	"\\textnormal": "textrm",
	"\\textsc": "textsc"
};
const textFontWeights = {
	"\\textbf": "textbf",
	"\\textmd": "textmd"
};
const textFontShapes = {
	"\\textit": "textit",
	"\\textup": "textup"
};
const styleWithFont = (group, style) => {
	const font = group.font;
	if (!font) return style;
	else if (textFontFamilies[font]) return style.withTextFontFamily(textFontFamilies[font]);
	else if (textFontWeights[font]) return style.withTextFontWeight(textFontWeights[font]);
	else if (font === "\\emph") return style.fontShape === "textit" ? style.withTextFontShape("textup") : style.withTextFontShape("textit");
	return style.withTextFontShape(textFontShapes[font]);
};
defineFunction({
	type: "text",
	names: [
		"\\text",
		"\\textrm",
		"\\textsf",
		"\\texttt",
		"\\textnormal",
		"\\textsc",
		"\\textbf",
		"\\textmd",
		"\\textit",
		"\\textup",
		"\\emph"
	],
	props: {
		numArgs: 1,
		argTypes: ["text"],
		allowedInArgument: true,
		allowedInText: true
	},
	handler({ parser, funcName }, args) {
		const body = args[0];
		return {
			type: "text",
			mode: parser.mode,
			body: ordargument(body),
			font: funcName
		};
	},
	mathmlBuilder(group, style) {
		const newStyle = styleWithFont(group, style);
		const mrow = buildExpressionRow(group.body, newStyle);
		return consolidateText(mrow);
	}
});
defineFunction({
	type: "vcenter",
	names: ["\\vcenter"],
	props: {
		numArgs: 1,
		argTypes: ["original"],
		allowedInText: false
	},
	handler({ parser }, args) {
		return {
			type: "vcenter",
			mode: parser.mode,
			body: args[0]
		};
	},
	mathmlBuilder(group, style) {
		const mtd = new MathNode("mtd", [buildGroup$1(group.body, style)]);
		mtd.style.padding = "0";
		return new MathNode("mtable", [new MathNode("mtr", [mtd])]);
	}
});
defineFunction({
	type: "verb",
	names: ["\\verb"],
	props: {
		numArgs: 0,
		allowedInText: true
	},
	handler(context, args, optArgs) {
		throw new ParseError("\\verb ended by end of line instead of matching delimiter");
	},
	mathmlBuilder(group, style) {
		const node = new MathNode("mtext", [new TextNode(makeVerb(group))]);
		node.setAttribute("mathvariant", "monospace");
		return node;
	}
});
/**
* Converts verb group into body string.
*
* \verb* replaces each space with an open box \u2423
* \verb replaces each space with a no-break space \xA0
*/
const makeVerb = (group) => group.body.replace(/ /g, group.star ? "␣" : "\xA0");
/** Include this to ensure that all functions are defined. */
const functions = _functions;
/**
* The Lexer class handles tokenizing the input in various ways. Since our
* parser expects us to be able to backtrack, the lexer allows lexing from any
* given starting point.
*
* Its main exposed function is the `lex` function, which takes a position to
* lex from and a type of token to lex. It defers to the appropriate `_innerLex`
* function.
*
* The various `_innerLex` functions perform the actual lexing of different
* kinds.
*/
const spaceRegexString = "[ \r\n	]";
const controlWordRegexString = "\\\\[a-zA-Z@]+";
const controlSymbolRegexString = "\\\\[^\ud800-\udfff]";
const controlWordWhitespaceRegexString = `(${controlWordRegexString})${spaceRegexString}*`;
const controlSpaceRegexString = "\\\\(\n|[ \r	]+\n?)[ \r	]*";
const combiningDiacriticalMarkString = "[̀-ͯ]";
const combiningDiacriticalMarksEndRegex = new RegExp(`${combiningDiacriticalMarkString}+$`);
const tokenRegexString = `(${spaceRegexString}+)|${controlSpaceRegexString}|([!-\\[\\]-‧‪-퟿豈-￿]${combiningDiacriticalMarkString}*` + "|[\ud800-\udbff][\udc00-\udfff]" + `${combiningDiacriticalMarkString}*|\\\\verb\\*([^]).*?\\4|\\\\verb([^*a-zA-Z]).*?\\5|${controlWordWhitespaceRegexString}|${controlSymbolRegexString})`;
/** Main Lexer class */
var Lexer = class {
	constructor(input, settings) {
		this.input = input;
		this.settings = settings;
		this.tokenRegex = new RegExp(tokenRegexString, "g");
		this.catcodes = {
			"%": 14,
			"~": 13
		};
	}
	setCatcode(char, code) {
		this.catcodes[char] = code;
	}
	/**
	* This function lexes a single token.
	*/
	lex() {
		const input = this.input;
		const pos = this.tokenRegex.lastIndex;
		if (pos === input.length) return new Token("EOF", new SourceLocation(this, pos, pos));
		const match = this.tokenRegex.exec(input);
		if (match === null || match.index !== pos) throw new ParseError(`Unexpected character: '${input[pos]}'`, new Token(input[pos], new SourceLocation(this, pos, pos + 1)));
		const text = match[6] || match[3] || (match[2] ? "\\ " : " ");
		if (this.catcodes[text] === 14) {
			const nlIndex = input.indexOf("\n", this.tokenRegex.lastIndex);
			if (nlIndex === -1) {
				this.tokenRegex.lastIndex = input.length;
				if (this.settings.strict) throw new ParseError("% comment has no terminating newline; LaTeX would fail because of commenting the end of math mode");
			} else this.tokenRegex.lastIndex = nlIndex + 1;
			return this.lex();
		}
		return new Token(text, new SourceLocation(this, pos, this.tokenRegex.lastIndex));
	}
};
/**
* A `Namespace` refers to a space of nameable things like macros or lengths,
* which can be `set` either globally or local to a nested group, using an
* undo stack similar to how TeX implements this functionality.
* Performance-wise, `get` and local `set` take constant time, while global
* `set` takes time proportional to the depth of group nesting.
*/
var Namespace = class {
	/**
	* Both arguments are optional.  The first argument is an object of
	* built-in mappings which never change.  The second argument is an object
	* of initial (global-level) mappings, which will constantly change
	* according to any global/top-level `set`s done.
	*/
	constructor(builtins = Object.create(null), globalMacros = Object.create(null)) {
		this.current = globalMacros;
		this.builtins = builtins;
		this.undefStack = [];
	}
	/**
	* Start a new nested group, affecting future local `set`s.
	*/
	beginGroup() {
		this.undefStack.push(Object.create(null));
	}
	/**
	* End current nested group, restoring values before the group began.
	*/
	endGroup() {
		if (this.undefStack.length === 0) throw new ParseError("Unbalanced namespace destruction: attempt to pop global namespace; please report this as a bug");
		const undefs = this.undefStack.pop();
		for (const undef in undefs) if (Object.prototype.hasOwnProperty.call(undefs, undef)) {
			if (undefs[undef] === void 0) delete this.current[undef];
			else this.current[undef] = undefs[undef];
		}
	}
	/**
	* Detect whether `name` has a definition.  Equivalent to
	* `get(name) != null`.
	*/
	has(name) {
		return Object.prototype.hasOwnProperty.call(this.current, name) || Object.prototype.hasOwnProperty.call(this.builtins, name);
	}
	/**
	* Get the current value of a name, or `undefined` if there is no value.
	*
	* Note: Do not use `if (namespace.get(...))` to detect whether a macro
	* is defined, as the definition may be the empty string which evaluates
	* to `false` in JavaScript.  Use `if (namespace.get(...) != null)` or
	* `if (namespace.has(...))`.
	*/
	get(name) {
		if (Object.prototype.hasOwnProperty.call(this.current, name)) return this.current[name];
		else return this.builtins[name];
	}
	/**
	* Set the current value of a name, and optionally set it globally too.
	* Local set() sets the current value and (when appropriate) adds an undo
	* operation to the undo stack.  Global set() may change the undo
	* operation at every level, so takes time linear in their number.
	*/
	set(name, value, global = false) {
		if (global) {
			for (let i = 0; i < this.undefStack.length; i++) delete this.undefStack[i][name];
			if (this.undefStack.length > 0) this.undefStack[this.undefStack.length - 1][name] = value;
		} else {
			const top = this.undefStack[this.undefStack.length - 1];
			if (top && !Object.prototype.hasOwnProperty.call(top, name)) top[name] = this.current[name];
		}
		this.current[name] = value;
	}
};
/**
* This file contains the “gullet” where macros are expanded
* until only non-macro tokens remain.
*/
const implicitCommands = {
	"^": true,
	_: true,
	"\\limits": true,
	"\\nolimits": true
};
var MacroExpander = class {
	constructor(input, settings, mode) {
		this.settings = settings;
		this.expansionCount = 0;
		this.feed(input);
		this.macros = new Namespace(macros, settings.macros);
		this.mode = mode;
		this.stack = [];
	}
	/**
	* Feed a new input string to the same MacroExpander
	* (with existing macros etc.).
	*/
	feed(input) {
		this.lexer = new Lexer(input, this.settings);
	}
	/**
	* Switches between "text" and "math" modes.
	*/
	switchMode(newMode) {
		this.mode = newMode;
	}
	/**
	* Start a new group nesting within all namespaces.
	*/
	beginGroup() {
		this.macros.beginGroup();
	}
	/**
	* End current group nesting within all namespaces.
	*/
	endGroup() {
		this.macros.endGroup();
	}
	/**
	* Returns the topmost token on the stack, without expanding it.
	* Similar in behavior to TeX's `\futurelet`.
	*/
	future() {
		if (this.stack.length === 0) this.pushToken(this.lexer.lex());
		return this.stack[this.stack.length - 1];
	}
	/**
	* Remove and return the next unexpanded token.
	*/
	popToken() {
		this.future();
		return this.stack.pop();
	}
	/**
	* Add a given token to the token stack.  In particular, this get be used
	* to put back a token returned from one of the other methods.
	*/
	pushToken(token) {
		this.stack.push(token);
	}
	/**
	* Append an array of tokens to the token stack.
	*/
	pushTokens(tokens) {
		this.stack.push(...tokens);
	}
	/**
	* Find an macro argument without expanding tokens and append the array of
	* tokens to the token stack. Uses Token as a container for the result.
	*/
	scanArgument(isOptional) {
		let start;
		let end;
		let tokens;
		if (isOptional) {
			this.consumeSpaces();
			if (this.future().text !== "[") return null;
			start = this.popToken();
			({tokens, end} = this.consumeArg(["]"]));
		} else ({tokens, start, end} = this.consumeArg());
		this.pushToken(new Token("EOF", end.loc));
		this.pushTokens(tokens);
		return new Token("", SourceLocation.range(start, end));
	}
	/**
	* Consume all following space tokens, without expansion.
	*/
	consumeSpaces() {
		for (;;) if (this.future().text === " ") this.stack.pop();
		else break;
	}
	/**
	* Consume an argument from the token stream, and return the resulting array
	* of tokens and start/end token.
	*/
	consumeArg(delims) {
		const tokens = [];
		const isDelimited = delims && delims.length > 0;
		if (!isDelimited) this.consumeSpaces();
		const start = this.future();
		let tok;
		let depth = 0;
		let match = 0;
		do {
			tok = this.popToken();
			tokens.push(tok);
			if (tok.text === "{") ++depth;
			else if (tok.text === "}") {
				--depth;
				if (depth === -1) throw new ParseError("Extra }", tok);
			} else if (tok.text === "EOF") throw new ParseError("Unexpected end of input in a macro argument, expected '" + (delims && isDelimited ? delims[match] : "}") + "'", tok);
			if (delims && isDelimited) {
				if ((depth === 0 || depth === 1 && delims[match] === "{") && tok.text === delims[match]) {
					++match;
					if (match === delims.length) {
						tokens.splice(-match, match);
						break;
					}
				} else match = 0;
			}
		} while (depth !== 0 || isDelimited);
		if (start.text === "{" && tokens[tokens.length - 1].text === "}") {
			tokens.pop();
			tokens.shift();
		}
		tokens.reverse();
		return {
			tokens,
			start,
			end: tok
		};
	}
	/**
	* Consume the specified number of (delimited) arguments from the token
	* stream and return the resulting array of arguments.
	*/
	consumeArgs(numArgs, delimiters) {
		if (delimiters) {
			if (delimiters.length !== numArgs + 1) throw new ParseError("The length of delimiters doesn't match the number of args!");
			const delims = delimiters[0];
			for (let i = 0; i < delims.length; i++) {
				const tok = this.popToken();
				if (delims[i] !== tok.text) throw new ParseError("Use of the macro doesn't match its definition", tok);
			}
		}
		const args = [];
		for (let i = 0; i < numArgs; i++) args.push(this.consumeArg(delimiters && delimiters[i + 1]).tokens);
		return args;
	}
	/**
	* Expand the next token only once if possible.
	*
	* If the token is expanded, the resulting tokens will be pushed onto
	* the stack in reverse order, and the number of such tokens will be
	* returned.  This number might be zero or positive.
	*
	* If not, the return value is `false`, and the next token remains at the
	* top of the stack.
	*
	* In either case, the next token will be on the top of the stack,
	* or the stack will be empty (in case of empty expansion
	* and no other tokens).
	*
	* Used to implement `expandAfterFuture` and `expandNextToken`.
	*
	* If expandableOnly, only expandable tokens are expanded and
	* an undefined control sequence results in an error.
	*/
	expandOnce(expandableOnly) {
		const topToken = this.popToken();
		const name = topToken.text;
		const expansion = !topToken.noexpand ? this._getExpansion(name) : null;
		if (expansion == null || expandableOnly && expansion.unexpandable) {
			if (expandableOnly && expansion == null && name[0] === "\\" && !this.isDefined(name)) throw new ParseError("Undefined control sequence: " + name);
			this.pushToken(topToken);
			return false;
		}
		this.expansionCount++;
		if (this.expansionCount > this.settings.maxExpand) throw new ParseError("Too many expansions: infinite loop or need to increase maxExpand setting");
		let tokens = expansion.tokens;
		const args = this.consumeArgs(expansion.numArgs, expansion.delimiters);
		if (expansion.numArgs) {
			tokens = tokens.slice();
			for (let i = tokens.length - 1; i >= 0; --i) {
				let tok = tokens[i];
				if (tok.text === "#") {
					if (i === 0) throw new ParseError("Incomplete placeholder at end of macro body", tok);
					tok = tokens[--i];
					if (tok.text === "#") tokens.splice(i + 1, 1);
					else if (/^[1-9]$/.test(tok.text)) tokens.splice(i, 2, ...args[+tok.text - 1]);
					else throw new ParseError("Not a valid argument number", tok);
				}
			}
		}
		this.pushTokens(tokens);
		return tokens.length;
	}
	/**
	* Expand the next token only once (if possible), and return the resulting
	* top token on the stack (without removing anything from the stack).
	* Similar in behavior to TeX's `\expandafter\futurelet`.
	* Equivalent to expandOnce() followed by future().
	*/
	expandAfterFuture() {
		this.expandOnce();
		return this.future();
	}
	/**
	* Recursively expand first token, then return first non-expandable token.
	*/
	expandNextToken() {
		for (;;) if (this.expandOnce() === false) {
			const token = this.stack.pop();
			if (token.treatAsRelax) token.text = "\\relax";
			return token;
		}
		throw new Error();
	}
	/**
	* Fully expand the given macro name and return the resulting list of
	* tokens, or return `undefined` if no such macro is defined.
	*/
	expandMacro(name) {
		return this.macros.has(name) ? this.expandTokens([new Token(name)]) : void 0;
	}
	/**
	* Fully expand the given token stream and return the resulting list of
	* tokens.  Note that the input tokens are in reverse order, but the
	* output tokens are in forward order.
	*/
	expandTokens(tokens) {
		const output = [];
		const oldStackLength = this.stack.length;
		this.pushTokens(tokens);
		while (this.stack.length > oldStackLength) if (this.expandOnce(true) === false) {
			const token = this.stack.pop();
			if (token.treatAsRelax) {
				token.noexpand = false;
				token.treatAsRelax = false;
			}
			output.push(token);
		}
		return output;
	}
	/**
	* Fully expand the given macro name and return the result as a string,
	* or return `undefined` if no such macro is defined.
	*/
	expandMacroAsText(name) {
		const tokens = this.expandMacro(name);
		if (tokens) return tokens.map((token) => token.text).join("");
		else return tokens;
	}
	/**
	* Returns the expanded macro as a reversed array of tokens and a macro
	* argument count.  Or returns `null` if no such macro.
	*/
	_getExpansion(name) {
		const definition = this.macros.get(name);
		if (definition == null) return definition;
		if (name.length === 1) {
			const catcode = this.lexer.catcodes[name];
			if (catcode != null && catcode !== 13) return;
		}
		const expansion = typeof definition === "function" ? definition(this) : definition;
		if (typeof expansion === "string") {
			let numArgs = 0;
			if (expansion.indexOf("#") !== -1) {
				const stripped = expansion.replace(/##/g, "");
				while (stripped.indexOf("#" + (numArgs + 1)) !== -1) ++numArgs;
			}
			const bodyLexer = new Lexer(expansion, this.settings);
			const tokens = [];
			let tok = bodyLexer.lex();
			while (tok.text !== "EOF") {
				tokens.push(tok);
				tok = bodyLexer.lex();
			}
			tokens.reverse();
			return {
				tokens,
				numArgs
			};
		}
		return expansion;
	}
	/**
	* Determine whether a command is currently "defined" (has some
	* functionality), meaning that it's a macro (in the current group),
	* a function, a symbol, or one of the special commands listed in
	* `implicitCommands`.
	*/
	isDefined(name) {
		return this.macros.has(name) || Object.prototype.hasOwnProperty.call(functions, name) || Object.prototype.hasOwnProperty.call(symbols.math, name) || Object.prototype.hasOwnProperty.call(symbols.text, name) || Object.prototype.hasOwnProperty.call(implicitCommands, name);
	}
	/**
	* Determine whether a command is expandable.
	*/
	isExpandable(name) {
		const macro = this.macros.get(name);
		return macro != null ? typeof macro === "string" || typeof macro === "function" || !macro.unexpandable : Object.prototype.hasOwnProperty.call(functions, name) && !functions[name].primitive;
	}
};
const unicodeSubRegEx = /^[₊₋₌₍₎₀₁₂₃₄₅₆₇₈₉ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓᵦᵧᵨᵩᵪ]/;
const uSubsAndSups = Object.freeze({
	"₊": "+",
	"₋": "-",
	"₌": "=",
	"₍": "(",
	"₎": ")",
	"₀": "0",
	"₁": "1",
	"₂": "2",
	"₃": "3",
	"₄": "4",
	"₅": "5",
	"₆": "6",
	"₇": "7",
	"₈": "8",
	"₉": "9",
	"ₐ": "a",
	"ₑ": "e",
	"ₕ": "h",
	"ᵢ": "i",
	"ⱼ": "j",
	"ₖ": "k",
	"ₗ": "l",
	"ₘ": "m",
	"ₙ": "n",
	"ₒ": "o",
	"ₚ": "p",
	"ᵣ": "r",
	"ₛ": "s",
	"ₜ": "t",
	"ᵤ": "u",
	"ᵥ": "v",
	"ₓ": "x",
	"ᵦ": "β",
	"ᵧ": "γ",
	"ᵨ": "ρ",
	"ᵩ": "ϕ",
	"ᵪ": "χ",
	"⁺": "+",
	"⁻": "-",
	"⁼": "=",
	"⁽": "(",
	"⁾": ")",
	"⁰": "0",
	"¹": "1",
	"²": "2",
	"³": "3",
	"⁴": "4",
	"⁵": "5",
	"⁶": "6",
	"⁷": "7",
	"⁸": "8",
	"⁹": "9",
	"ᴬ": "A",
	"ᴮ": "B",
	"ᴰ": "D",
	"ᴱ": "E",
	"ᴳ": "G",
	"ᴴ": "H",
	"ᴵ": "I",
	"ᴶ": "J",
	"ᴷ": "K",
	"ᴸ": "L",
	"ᴹ": "M",
	"ᴺ": "N",
	"ᴼ": "O",
	"ᴾ": "P",
	"ᴿ": "R",
	"ᵀ": "T",
	"ᵁ": "U",
	"ⱽ": "V",
	"ᵂ": "W",
	"ᵃ": "a",
	"ᵇ": "b",
	"ᶜ": "c",
	"ᵈ": "d",
	"ᵉ": "e",
	"ᶠ": "f",
	"ᵍ": "g",
	"ʰ": "h",
	"ⁱ": "i",
	"ʲ": "j",
	"ᵏ": "k",
	"ˡ": "l",
	"ᵐ": "m",
	"ⁿ": "n",
	"ᵒ": "o",
	"ᵖ": "p",
	"ʳ": "r",
	"ˢ": "s",
	"ᵗ": "t",
	"ᵘ": "u",
	"ᵛ": "v",
	"ʷ": "w",
	"ˣ": "x",
	"ʸ": "y",
	"ᶻ": "z",
	"ᵝ": "β",
	"ᵞ": "γ",
	"ᵟ": "δ",
	"ᵠ": "ϕ",
	"ᵡ": "χ",
	"ᶿ": "θ"
});
const asciiFromScript = Object.freeze({
	"𝒜": "A",
	"ℬ": "B",
	"𝒞": "C",
	"𝒟": "D",
	"ℰ": "E",
	"ℱ": "F",
	"𝒢": "G",
	"ℋ": "H",
	"ℐ": "I",
	"𝒥": "J",
	"𝒦": "K",
	"ℒ": "L",
	"ℳ": "M",
	"𝒩": "N",
	"𝒪": "O",
	"𝒫": "P",
	"𝒬": "Q",
	"ℛ": "R",
	"𝒮": "S",
	"𝒯": "T",
	"𝒰": "U",
	"𝒱": "V",
	"𝒲": "W",
	"𝒳": "X",
	"𝒴": "Y",
	"𝒵": "Z"
});
var unicodeAccents = {
	"́": {
		text: "\\'",
		math: "\\acute"
	},
	"̀": {
		text: "\\`",
		math: "\\grave"
	},
	"̈": {
		text: "\\\"",
		math: "\\ddot"
	},
	"̃": {
		text: "\\~",
		math: "\\tilde"
	},
	"̄": {
		text: "\\=",
		math: "\\bar"
	},
	"̆": {
		text: "\\u",
		math: "\\breve"
	},
	"̌": {
		text: "\\v",
		math: "\\check"
	},
	"̂": {
		text: "\\^",
		math: "\\hat"
	},
	"̇": {
		text: "\\.",
		math: "\\dot"
	},
	"̊": {
		text: "\\r",
		math: "\\mathring"
	},
	"̋": { text: "\\H" },
	"̧": { text: "\\c" }
};
var unicodeSymbols = {
	"á": "á",
	"à": "à",
	"ä": "ä",
	"ǟ": "ǟ",
	"ã": "ã",
	"ā": "ā",
	"ă": "ă",
	"ắ": "ắ",
	"ằ": "ằ",
	"ẵ": "ẵ",
	"ǎ": "ǎ",
	"â": "â",
	"ấ": "ấ",
	"ầ": "ầ",
	"ẫ": "ẫ",
	"ȧ": "ȧ",
	"ǡ": "ǡ",
	"å": "å",
	"ǻ": "ǻ",
	"ḃ": "ḃ",
	"ć": "ć",
	"č": "č",
	"ĉ": "ĉ",
	"ċ": "ċ",
	"ď": "ď",
	"ḋ": "ḋ",
	"é": "é",
	"è": "è",
	"ë": "ë",
	"ẽ": "ẽ",
	"ē": "ē",
	"ḗ": "ḗ",
	"ḕ": "ḕ",
	"ĕ": "ĕ",
	"ě": "ě",
	"ê": "ê",
	"ế": "ế",
	"ề": "ề",
	"ễ": "ễ",
	"ė": "ė",
	"ḟ": "ḟ",
	"ǵ": "ǵ",
	"ḡ": "ḡ",
	"ğ": "ğ",
	"ǧ": "ǧ",
	"ĝ": "ĝ",
	"ġ": "ġ",
	"ḧ": "ḧ",
	"ȟ": "ȟ",
	"ĥ": "ĥ",
	"ḣ": "ḣ",
	"í": "í",
	"ì": "ì",
	"ï": "ï",
	"ḯ": "ḯ",
	"ĩ": "ĩ",
	"ī": "ī",
	"ĭ": "ĭ",
	"ǐ": "ǐ",
	"î": "î",
	"ǰ": "ǰ",
	"ĵ": "ĵ",
	"ḱ": "ḱ",
	"ǩ": "ǩ",
	"ĺ": "ĺ",
	"ľ": "ľ",
	"ḿ": "ḿ",
	"ṁ": "ṁ",
	"ń": "ń",
	"ǹ": "ǹ",
	"ñ": "ñ",
	"ň": "ň",
	"ṅ": "ṅ",
	"ó": "ó",
	"ò": "ò",
	"ö": "ö",
	"ȫ": "ȫ",
	"õ": "õ",
	"ṍ": "ṍ",
	"ṏ": "ṏ",
	"ȭ": "ȭ",
	"ō": "ō",
	"ṓ": "ṓ",
	"ṑ": "ṑ",
	"ŏ": "ŏ",
	"ǒ": "ǒ",
	"ô": "ô",
	"ố": "ố",
	"ồ": "ồ",
	"ỗ": "ỗ",
	"ȯ": "ȯ",
	"ȱ": "ȱ",
	"ő": "ő",
	"ṕ": "ṕ",
	"ṗ": "ṗ",
	"ŕ": "ŕ",
	"ř": "ř",
	"ṙ": "ṙ",
	"ś": "ś",
	"ṥ": "ṥ",
	"š": "š",
	"ṧ": "ṧ",
	"ŝ": "ŝ",
	"ṡ": "ṡ",
	"ẗ": "ẗ",
	"ť": "ť",
	"ṫ": "ṫ",
	"ú": "ú",
	"ù": "ù",
	"ü": "ü",
	"ǘ": "ǘ",
	"ǜ": "ǜ",
	"ǖ": "ǖ",
	"ǚ": "ǚ",
	"ũ": "ũ",
	"ṹ": "ṹ",
	"ū": "ū",
	"ṻ": "ṻ",
	"ŭ": "ŭ",
	"ǔ": "ǔ",
	"û": "û",
	"ů": "ů",
	"ű": "ű",
	"ṽ": "ṽ",
	"ẃ": "ẃ",
	"ẁ": "ẁ",
	"ẅ": "ẅ",
	"ŵ": "ŵ",
	"ẇ": "ẇ",
	"ẘ": "ẘ",
	"ẍ": "ẍ",
	"ẋ": "ẋ",
	"ý": "ý",
	"ỳ": "ỳ",
	"ÿ": "ÿ",
	"ỹ": "ỹ",
	"ȳ": "ȳ",
	"ŷ": "ŷ",
	"ẏ": "ẏ",
	"ẙ": "ẙ",
	"ź": "ź",
	"ž": "ž",
	"ẑ": "ẑ",
	"ż": "ż",
	"Á": "Á",
	"À": "À",
	"Ä": "Ä",
	"Ǟ": "Ǟ",
	"Ã": "Ã",
	"Ā": "Ā",
	"Ă": "Ă",
	"Ắ": "Ắ",
	"Ằ": "Ằ",
	"Ẵ": "Ẵ",
	"Ǎ": "Ǎ",
	"Â": "Â",
	"Ấ": "Ấ",
	"Ầ": "Ầ",
	"Ẫ": "Ẫ",
	"Ȧ": "Ȧ",
	"Ǡ": "Ǡ",
	"Å": "Å",
	"Ǻ": "Ǻ",
	"Ḃ": "Ḃ",
	"Ć": "Ć",
	"Č": "Č",
	"Ĉ": "Ĉ",
	"Ċ": "Ċ",
	"Ď": "Ď",
	"Ḋ": "Ḋ",
	"É": "É",
	"È": "È",
	"Ë": "Ë",
	"Ẽ": "Ẽ",
	"Ē": "Ē",
	"Ḗ": "Ḗ",
	"Ḕ": "Ḕ",
	"Ĕ": "Ĕ",
	"Ě": "Ě",
	"Ê": "Ê",
	"Ế": "Ế",
	"Ề": "Ề",
	"Ễ": "Ễ",
	"Ė": "Ė",
	"Ḟ": "Ḟ",
	"Ǵ": "Ǵ",
	"Ḡ": "Ḡ",
	"Ğ": "Ğ",
	"Ǧ": "Ǧ",
	"Ĝ": "Ĝ",
	"Ġ": "Ġ",
	"Ḧ": "Ḧ",
	"Ȟ": "Ȟ",
	"Ĥ": "Ĥ",
	"Ḣ": "Ḣ",
	"Í": "Í",
	"Ì": "Ì",
	"Ï": "Ï",
	"Ḯ": "Ḯ",
	"Ĩ": "Ĩ",
	"Ī": "Ī",
	"Ĭ": "Ĭ",
	"Ǐ": "Ǐ",
	"Î": "Î",
	"İ": "İ",
	"Ĵ": "Ĵ",
	"Ḱ": "Ḱ",
	"Ǩ": "Ǩ",
	"Ĺ": "Ĺ",
	"Ľ": "Ľ",
	"Ḿ": "Ḿ",
	"Ṁ": "Ṁ",
	"Ń": "Ń",
	"Ǹ": "Ǹ",
	"Ñ": "Ñ",
	"Ň": "Ň",
	"Ṅ": "Ṅ",
	"Ó": "Ó",
	"Ò": "Ò",
	"Ö": "Ö",
	"Ȫ": "Ȫ",
	"Õ": "Õ",
	"Ṍ": "Ṍ",
	"Ṏ": "Ṏ",
	"Ȭ": "Ȭ",
	"Ō": "Ō",
	"Ṓ": "Ṓ",
	"Ṑ": "Ṑ",
	"Ŏ": "Ŏ",
	"Ǒ": "Ǒ",
	"Ô": "Ô",
	"Ố": "Ố",
	"Ồ": "Ồ",
	"Ỗ": "Ỗ",
	"Ȯ": "Ȯ",
	"Ȱ": "Ȱ",
	"Ő": "Ő",
	"Ṕ": "Ṕ",
	"Ṗ": "Ṗ",
	"Ŕ": "Ŕ",
	"Ř": "Ř",
	"Ṙ": "Ṙ",
	"Ś": "Ś",
	"Ṥ": "Ṥ",
	"Š": "Š",
	"Ṧ": "Ṧ",
	"Ŝ": "Ŝ",
	"Ṡ": "Ṡ",
	"Ť": "Ť",
	"Ṫ": "Ṫ",
	"Ú": "Ú",
	"Ù": "Ù",
	"Ü": "Ü",
	"Ǘ": "Ǘ",
	"Ǜ": "Ǜ",
	"Ǖ": "Ǖ",
	"Ǚ": "Ǚ",
	"Ũ": "Ũ",
	"Ṹ": "Ṹ",
	"Ū": "Ū",
	"Ṻ": "Ṻ",
	"Ŭ": "Ŭ",
	"Ǔ": "Ǔ",
	"Û": "Û",
	"Ů": "Ů",
	"Ű": "Ű",
	"Ṽ": "Ṽ",
	"Ẃ": "Ẃ",
	"Ẁ": "Ẁ",
	"Ẅ": "Ẅ",
	"Ŵ": "Ŵ",
	"Ẇ": "Ẇ",
	"Ẍ": "Ẍ",
	"Ẋ": "Ẋ",
	"Ý": "Ý",
	"Ỳ": "Ỳ",
	"Ÿ": "Ÿ",
	"Ỹ": "Ỹ",
	"Ȳ": "Ȳ",
	"Ŷ": "Ŷ",
	"Ẏ": "Ẏ",
	"Ź": "Ź",
	"Ž": "Ž",
	"Ẑ": "Ẑ",
	"Ż": "Ż",
	"ά": "ά",
	"ὰ": "ὰ",
	"ᾱ": "ᾱ",
	"ᾰ": "ᾰ",
	"έ": "έ",
	"ὲ": "ὲ",
	"ή": "ή",
	"ὴ": "ὴ",
	"ί": "ί",
	"ὶ": "ὶ",
	"ϊ": "ϊ",
	"ΐ": "ΐ",
	"ῒ": "ῒ",
	"ῑ": "ῑ",
	"ῐ": "ῐ",
	"ό": "ό",
	"ὸ": "ὸ",
	"ύ": "ύ",
	"ὺ": "ὺ",
	"ϋ": "ϋ",
	"ΰ": "ΰ",
	"ῢ": "ῢ",
	"ῡ": "ῡ",
	"ῠ": "ῠ",
	"ώ": "ώ",
	"ὼ": "ὼ",
	"Ύ": "Ύ",
	"Ὺ": "Ὺ",
	"Ϋ": "Ϋ",
	"Ῡ": "Ῡ",
	"Ῠ": "Ῠ",
	"Ώ": "Ώ",
	"Ὼ": "Ὼ"
};
const binLeftCancellers = [
	"bin",
	"op",
	"open",
	"punct",
	"rel"
];
const sizeRegEx = /([-+]?) *(\d+(?:\.\d*)?|\.\d+) *([a-z]{2})/;
const textRegEx = /^ *\\text/;
/**
* This file contains the parser used to parse out a TeX expression from the
* input. Since TeX isn't context-free, standard parsers don't work particularly
* well.
*
* The strategy of this parser is as such:
*
* The main functions (the `.parse...` ones) take a position in the current
* parse string to parse tokens from. The lexer (found in Lexer.js, stored at
* this.gullet.lexer) also supports pulling out tokens at arbitrary places. When
* individual tokens are needed at a position, the lexer is called to pull out a
* token, which is then used.
*
* The parser has a property called "mode" indicating the mode that
* the parser is currently in. Currently it has to be one of "math" or
* "text", which denotes whether the current environment is a math-y
* one or a text-y one (e.g. inside \text). Currently, this serves to
* limit the functions which can be used in text mode.
*
* The main functions then return an object which contains the useful data that
* was parsed at its given point, and a new position at the end of the parsed
* data. The main functions can call each other and continue the parsing by
* using the returned position as a new starting point.
*
* There are also extra `.handle...` functions, which pull out some reused
* functionality into self-contained functions.
*
* The functions return ParseNodes.
*/
var Parser = class Parser {
	constructor(input, settings, isPreamble = false) {
		this.mode = "math";
		this.gullet = new MacroExpander(input, settings, this.mode);
		this.settings = settings;
		this.isPreamble = isPreamble;
		this.leftrightDepth = 0;
		this.prevAtomType = "";
	}
	/**
	* Checks a result to make sure it has the right type, and throws an
	* appropriate error otherwise.
	*/
	expect(text, consume = true) {
		if (this.fetch().text !== text) throw new ParseError(`Expected '${text}', got '${this.fetch().text}'`, this.fetch());
		if (consume) this.consume();
	}
	/**
	* Discards the current lookahead token, considering it consumed.
	*/
	consume() {
		this.nextToken = null;
	}
	/**
	* Return the current lookahead token, or if there isn't one (at the
	* beginning, or if the previous lookahead token was consume()d),
	* fetch the next token as the new lookahead token and return it.
	*/
	fetch() {
		if (this.nextToken == null) this.nextToken = this.gullet.expandNextToken();
		return this.nextToken;
	}
	/**
	* Switches between "text" and "math" modes.
	*/
	switchMode(newMode) {
		this.mode = newMode;
		this.gullet.switchMode(newMode);
	}
	/**
	* Main parsing function, which parses an entire input.
	*/
	parse() {
		this.gullet.beginGroup();
		if (this.settings.colorIsTextColor) this.gullet.macros.set("\\color", "\\textcolor");
		const parse = this.parseExpression(false);
		this.expect("EOF");
		if (this.isPreamble) {
			const macros = Object.create(null);
			Object.entries(this.gullet.macros.current).forEach(([key, value]) => {
				macros[key] = value;
			});
			this.gullet.endGroup();
			return macros;
		}
		const tag = this.gullet.macros.get("\\df@tag");
		this.gullet.endGroup();
		if (tag) this.gullet.macros.current["\\df@tag"] = tag;
		return parse;
	}
	static get endOfExpression() {
		return [
			"}",
			"\\endgroup",
			"\\end",
			"\\right",
			"\\endtoggle",
			"&"
		];
	}
	/**
	* Fully parse a separate sequence of tokens as a separate job.
	* Tokens should be specified in reverse order, as in a MacroDefinition.
	*/
	subparse(tokens) {
		const oldToken = this.nextToken;
		this.consume();
		this.gullet.pushToken(new Token("}"));
		this.gullet.pushTokens(tokens);
		const parse = this.parseExpression(false);
		this.expect("}");
		this.nextToken = oldToken;
		return parse;
	}
	/**
	* Parses an "expression", which is a list of atoms.
	*
	* `breakOnInfix`: Should the parsing stop when we hit infix nodes? This
	*                 happens when functions have higher precedence than infix
	*                 nodes in implicit parses.
	*
	* `breakOnTokenText`: The text of the token that the expression should end
	*                     with, or `null` if something else should end the
	*                     expression.
	*
	* `breakOnMiddle`: \color, \over, and old styling functions work on an implicit group.
	*                  These groups end just before the usual tokens, but they also
	*                  end just before `\middle`.
	*/
	parseExpression(breakOnInfix, breakOnTokenText, breakOnMiddle) {
		const body = [];
		this.prevAtomType = "";
		while (true) {
			if (this.mode === "math") this.consumeSpaces();
			const lex = this.fetch();
			if (Parser.endOfExpression.indexOf(lex.text) !== -1) break;
			if (breakOnTokenText && lex.text === breakOnTokenText) break;
			if (breakOnMiddle && lex.text === "\\middle") break;
			if (breakOnInfix && functions[lex.text] && functions[lex.text].infix) break;
			const atom = this.parseAtom(breakOnTokenText);
			if (!atom) break;
			else if (atom.type === "internal") continue;
			body.push(atom);
			this.prevAtomType = atom.type === "atom" ? atom.family : atom.type;
		}
		if (this.mode === "text") this.formLigatures(body);
		return this.handleInfixNodes(body);
	}
	/**
	* Rewrites infix operators such as \over with corresponding commands such
	* as \frac.
	*
	* There can only be one infix operator per group.  If there's more than one
	* then the expression is ambiguous.  This can be resolved by adding {}.
	*/
	handleInfixNodes(body) {
		let overIndex = -1;
		let funcName;
		for (let i = 0; i < body.length; i++) if (body[i].type === "infix") {
			if (overIndex !== -1) throw new ParseError("only one infix operator per group", body[i].token);
			overIndex = i;
			funcName = body[i].replaceWith;
		}
		if (overIndex !== -1 && funcName) {
			let numerNode;
			let denomNode;
			const numerBody = body.slice(0, overIndex);
			const denomBody = body.slice(overIndex + 1);
			if (numerBody.length === 1 && numerBody[0].type === "ordgroup") numerNode = numerBody[0];
			else numerNode = {
				type: "ordgroup",
				mode: this.mode,
				body: numerBody
			};
			if (denomBody.length === 1 && denomBody[0].type === "ordgroup") denomNode = denomBody[0];
			else denomNode = {
				type: "ordgroup",
				mode: this.mode,
				body: denomBody
			};
			let node;
			if (funcName === "\\\\abovefrac") node = this.callFunction(funcName, [
				numerNode,
				body[overIndex],
				denomNode
			], []);
			else node = this.callFunction(funcName, [numerNode, denomNode], []);
			return [node];
		} else return body;
	}
	/**
	* Handle a subscript or superscript with nice errors.
	*/
	handleSupSubscript(name) {
		const symbolToken = this.fetch();
		const symbol = symbolToken.text;
		this.consume();
		this.consumeSpaces();
		let group;
		do
			group = this.parseGroup(name);
		while (group.type && group.type === "internal");
		if (!group) throw new ParseError("Expected group after '" + symbol + "'", symbolToken);
		return group;
	}
	/**
	* Converts the textual input of an unsupported command into a text node
	* contained within a color node whose color is determined by errorColor
	*/
	formatUnsupportedCmd(text) {
		const textordArray = [];
		for (let i = 0; i < text.length; i++) textordArray.push({
			type: "textord",
			mode: "text",
			text: text[i]
		});
		const textNode = {
			type: "text",
			mode: this.mode,
			body: textordArray
		};
		return {
			type: "color",
			mode: this.mode,
			color: this.settings.errorColor,
			body: [textNode]
		};
	}
	/**
	* Parses a group with optional super/subscripts.
	*/
	parseAtom(breakOnTokenText) {
		const base = this.parseGroup("atom", breakOnTokenText);
		if (base && base.type === "internal") return base;
		if (this.mode === "text") return base;
		let superscript;
		let subscript;
		while (true) {
			this.consumeSpaces();
			const lex = this.fetch();
			if (lex.text === "\\limits" || lex.text === "\\nolimits") {
				if (base && base.type === "op") {
					base.limits = lex.text === "\\limits";
					base.alwaysHandleSupSub = true;
				} else if (base && base.type === "operatorname") {
					if (base.alwaysHandleSupSub) base.limits = lex.text === "\\limits";
				} else throw new ParseError("Limit controls must follow a math operator", lex);
				this.consume();
			} else if (lex.text === "^") {
				if (superscript) throw new ParseError("Double superscript", lex);
				superscript = this.handleSupSubscript("superscript");
			} else if (lex.text === "_") {
				if (subscript) throw new ParseError("Double subscript", lex);
				subscript = this.handleSupSubscript("subscript");
			} else if (lex.text === "'") {
				if (superscript) throw new ParseError("Double superscript", lex);
				const prime = {
					type: "textord",
					mode: this.mode,
					text: "\\prime"
				};
				const primes = [prime];
				this.consume();
				while (this.fetch().text === "'") {
					primes.push(prime);
					this.consume();
				}
				if (this.fetch().text === "^") primes.push(this.handleSupSubscript("superscript"));
				superscript = {
					type: "ordgroup",
					mode: this.mode,
					body: primes
				};
			} else if (uSubsAndSups[lex.text]) {
				const isSub = unicodeSubRegEx.test(lex.text);
				const subsupTokens = [];
				subsupTokens.push(new Token(uSubsAndSups[lex.text]));
				this.consume();
				while (true) {
					const token = this.fetch().text;
					if (!uSubsAndSups[token]) break;
					if (unicodeSubRegEx.test(token) !== isSub) break;
					subsupTokens.unshift(new Token(uSubsAndSups[token]));
					this.consume();
				}
				const body = this.subparse(subsupTokens);
				if (isSub) subscript = {
					type: "ordgroup",
					mode: "math",
					body
				};
				else superscript = {
					type: "ordgroup",
					mode: "math",
					body
				};
			} else break;
		}
		if (superscript || subscript) {
			if (base && base.type === "multiscript" && !base.postscripts) {
				base.postscripts = {
					sup: superscript,
					sub: subscript
				};
				return base;
			} else {
				const isFollowedByDelimiter = !base || base.type !== "op" && base.type !== "operatorname" ? void 0 : isDelimiter(this.nextToken.text);
				return {
					type: "supsub",
					mode: this.mode,
					base,
					sup: superscript,
					sub: subscript,
					isFollowedByDelimiter
				};
			}
		} else return base;
	}
	/**
	* Parses an entire function, including its base and all of its arguments.
	*/
	parseFunction(breakOnTokenText, name) {
		const token = this.fetch();
		const func = token.text;
		const funcData = functions[func];
		if (!funcData) return null;
		this.consume();
		if (name && name !== "atom" && !funcData.allowedInArgument) throw new ParseError("Got function '" + func + "' with no arguments" + (name ? " as " + name : ""), token);
		else if (this.mode === "text" && !funcData.allowedInText) throw new ParseError("Can't use function '" + func + "' in text mode", token);
		else if (this.mode === "math" && funcData.allowedInMath === false) throw new ParseError("Can't use function '" + func + "' in math mode", token);
		const prevAtomType = this.prevAtomType;
		const { args, optArgs } = this.parseArguments(func, funcData);
		this.prevAtomType = prevAtomType;
		return this.callFunction(func, args, optArgs, token, breakOnTokenText);
	}
	/**
	* Call a function handler with a suitable context and arguments.
	*/
	callFunction(name, args, optArgs, token, breakOnTokenText) {
		const context = {
			funcName: name,
			parser: this,
			token,
			breakOnTokenText
		};
		const func = functions[name];
		if (func && func.handler) return func.handler(context, args, optArgs);
		else throw new ParseError(`No function handler for ${name}`);
	}
	/**
	* Parses the arguments of a function or environment
	*/
	parseArguments(func, funcData) {
		const totalArgs = funcData.numArgs + funcData.numOptionalArgs;
		if (totalArgs === 0) return {
			args: [],
			optArgs: []
		};
		const args = [];
		const optArgs = [];
		for (let i = 0; i < totalArgs; i++) {
			let argType = funcData.argTypes && funcData.argTypes[i];
			const isOptional = i < funcData.numOptionalArgs;
			if (funcData.primitive && argType == null || funcData.type === "sqrt" && i === 1 && optArgs[0] == null) argType = "primitive";
			const arg = this.parseGroupOfType(`argument to '${func}'`, argType, isOptional);
			if (isOptional) optArgs.push(arg);
			else if (arg != null) args.push(arg);
			else throw new ParseError("Null argument, please report this as a bug");
		}
		return {
			args,
			optArgs
		};
	}
	/**
	* Parses a group when the mode is changing.
	*/
	parseGroupOfType(name, type, optional) {
		switch (type) {
			case "size": return this.parseSizeGroup(optional);
			case "url": return this.parseUrlGroup(optional);
			case "math":
			case "text": return this.parseArgumentGroup(optional, type);
			case "hbox": {
				const group = this.parseArgumentGroup(optional, "text");
				return group != null ? {
					type: "styling",
					mode: group.mode,
					body: [group],
					scriptLevel: "text"
				} : null;
			}
			case "raw": {
				const token = this.parseStringGroup("raw", optional);
				return token != null ? {
					type: "raw",
					mode: "text",
					string: token.text
				} : null;
			}
			case "primitive": {
				if (optional) throw new ParseError("A primitive argument cannot be optional");
				const group = this.parseGroup(name);
				if (group == null) throw new ParseError("Expected group as " + name, this.fetch());
				return group;
			}
			case "original":
			case null:
			case void 0: return this.parseArgumentGroup(optional);
			default: throw new ParseError("Unknown group type as " + name, this.fetch());
		}
	}
	/**
	* Discard any space tokens, fetching the next non-space token.
	*/
	consumeSpaces() {
		while (true) {
			const ch = this.fetch().text;
			if (ch === " " || ch === "\xA0" || ch === "︎") this.consume();
			else break;
		}
	}
	/**
	* Parses a group, essentially returning the string formed by the
	* brace-enclosed tokens plus some position information.
	*/
	parseStringGroup(modeName, optional) {
		const argToken = this.gullet.scanArgument(optional);
		if (argToken == null) return null;
		let str = "";
		let nextToken;
		while ((nextToken = this.fetch()).text !== "EOF") {
			str += nextToken.text;
			this.consume();
		}
		this.consume();
		argToken.text = str;
		return argToken;
	}
	/**
	* Parses a regex-delimited group: the largest sequence of tokens
	* whose concatenated strings match `regex`. Returns the string
	* formed by the tokens plus some position information.
	*/
	parseRegexGroup(regex, modeName) {
		const firstToken = this.fetch();
		let lastToken = firstToken;
		let str = "";
		let nextToken;
		while ((nextToken = this.fetch()).text !== "EOF" && regex.test(str + nextToken.text)) {
			lastToken = nextToken;
			str += lastToken.text;
			this.consume();
		}
		if (str === "") throw new ParseError("Invalid " + modeName + ": '" + firstToken.text + "'", firstToken);
		return firstToken.range(lastToken, str);
	}
	/**
	* Parses a size specification, consisting of magnitude and unit.
	*/
	parseSizeGroup(optional) {
		let res;
		let isBlank = false;
		this.gullet.consumeSpaces();
		if (!optional && this.gullet.future().text !== "{") res = this.parseRegexGroup(/^[-+]? *(?:$|\d+|\d+\.\d*|\.\d*) *[a-z]{0,2} *$/, "size");
		else res = this.parseStringGroup("size", optional);
		if (!res) return null;
		if (!optional && res.text.length === 0) {
			res.text = "0pt";
			isBlank = true;
		}
		const match = sizeRegEx.exec(res.text);
		if (!match) throw new ParseError("Invalid size: '" + res.text + "'", res);
		const data = {
			number: +(match[1] + match[2]),
			unit: match[3]
		};
		if (!validUnit(data)) throw new ParseError("Invalid unit: '" + data.unit + "'", res);
		return {
			type: "size",
			mode: this.mode,
			value: data,
			isBlank
		};
	}
	/**
	* Parses an URL, checking escaped letters and allowed protocols,
	* and setting the catcode of % as an active character (as in \hyperref).
	*/
	parseUrlGroup(optional) {
		this.gullet.lexer.setCatcode("%", 13);
		this.gullet.lexer.setCatcode("~", 12);
		const res = this.parseStringGroup("url", optional);
		this.gullet.lexer.setCatcode("%", 14);
		this.gullet.lexer.setCatcode("~", 13);
		if (res == null) return null;
		let url = res.text.replace(/\\([#$%&~_^{}])/g, "$1");
		url = res.text.replace(/{\u2044}/g, "/");
		return {
			type: "url",
			mode: this.mode,
			url
		};
	}
	/**
	* Parses an argument with the mode specified.
	*/
	parseArgumentGroup(optional, mode) {
		const argToken = this.gullet.scanArgument(optional);
		if (argToken == null) return null;
		const outerMode = this.mode;
		if (mode) this.switchMode(mode);
		this.gullet.beginGroup();
		const expression = this.parseExpression(false, "EOF");
		this.expect("EOF");
		this.gullet.endGroup();
		const result = {
			type: "ordgroup",
			mode: this.mode,
			loc: argToken.loc,
			body: expression
		};
		if (mode) this.switchMode(outerMode);
		return result;
	}
	/**
	* Parses an ordinary group, which is either a single nucleus (like "x")
	* or an expression in braces (like "{x+y}") or an implicit group, a group
	* that starts at the current position, and ends right before a higher explicit
	* group ends, or at EOF.
	*/
	parseGroup(name, breakOnTokenText) {
		const firstToken = this.fetch();
		const text = firstToken.text;
		if (name === "argument to '\\left'") return this.parseSymbol();
		let result;
		if (text === "{" || text === "\\begingroup" || text === "\\toggle") {
			this.consume();
			const groupEnd = text === "{" ? "}" : text === "\\begingroup" ? "\\endgroup" : "\\endtoggle";
			this.gullet.beginGroup();
			const expression = this.parseExpression(false, groupEnd);
			const lastToken = this.fetch();
			this.expect(groupEnd);
			this.gullet.endGroup();
			result = {
				type: lastToken.text === "\\endtoggle" ? "toggle" : "ordgroup",
				mode: this.mode,
				loc: SourceLocation.range(firstToken, lastToken),
				body: expression,
				semisimple: text === "\\begingroup" || void 0
			};
		} else {
			result = this.parseFunction(breakOnTokenText, name) || this.parseSymbol();
			if (result == null && text[0] === "\\" && !Object.prototype.hasOwnProperty.call(implicitCommands, text)) {
				if (this.settings.throwOnError) throw new ParseError("Unsupported function name: " + text, firstToken);
				result = this.formatUnsupportedCmd(text);
				this.consume();
			}
		}
		return result;
	}
	/**
	* Form ligature-like combinations of characters for text mode.
	* This includes inputs like "--", "---", "``" and "''".
	* The result will simply replace multiple textord nodes with a single
	* character in each value by a single textord node having multiple
	* characters in its value.  The representation is still ASCII source.
	* The group will be modified in place.
	*/
	formLigatures(group) {
		let n = group.length - 1;
		for (let i = 0; i < n; ++i) {
			const a = group[i];
			const v = a.text;
			if (v === "-" && group[i + 1].text === "-") {
				if (i + 1 < n && group[i + 2].text === "-") {
					group.splice(i, 3, {
						type: "textord",
						mode: "text",
						loc: SourceLocation.range(a, group[i + 2]),
						text: "---"
					});
					n -= 2;
				} else {
					group.splice(i, 2, {
						type: "textord",
						mode: "text",
						loc: SourceLocation.range(a, group[i + 1]),
						text: "--"
					});
					n -= 1;
				}
			}
			if ((v === "'" || v === "`") && group[i + 1].text === v) {
				group.splice(i, 2, {
					type: "textord",
					mode: "text",
					loc: SourceLocation.range(a, group[i + 1]),
					text: v + v
				});
				n -= 1;
			}
		}
	}
	/**
	* Parse a single symbol out of the string. Here, we handle single character
	* symbols and special functions like \verb.
	*/
	parseSymbol() {
		const nucleus = this.fetch();
		let text = nucleus.text;
		if (/^\\verb[^a-zA-Z]/.test(text)) {
			this.consume();
			let arg = text.slice(5);
			const star = arg.charAt(0) === "*";
			if (star) arg = arg.slice(1);
			if (arg.length < 2 || arg.charAt(0) !== arg.slice(-1)) throw new ParseError(`\\verb assertion failed --
                    please report what input caused this bug`);
			arg = arg.slice(1, -1);
			return {
				type: "verb",
				mode: "text",
				body: arg,
				star
			};
		}
		if (Object.prototype.hasOwnProperty.call(unicodeSymbols, text[0]) && this.mode === "math" && !symbols[this.mode][text[0]]) {
			if (this.settings.strict && this.mode === "math") throw new ParseError(`Accented Unicode text character "${text[0]}" used in math mode`, nucleus);
			text = unicodeSymbols[text[0]] + text.slice(1);
		}
		const match = this.mode === "math" ? combiningDiacriticalMarksEndRegex.exec(text) : null;
		if (match) {
			text = text.substring(0, match.index);
			if (text === "i") text = "ı";
			else if (text === "j") text = "ȷ";
		}
		let symbol;
		if (symbols[this.mode][text]) {
			let group = symbols[this.mode][text].group;
			if (group === "bin" && (binLeftCancellers.includes(this.prevAtomType) || this.prevAtomType === "")) group = "open";
			const loc = SourceLocation.range(nucleus);
			let s;
			if (Object.prototype.hasOwnProperty.call(ATOMS, group)) {
				const family = group;
				s = {
					type: "atom",
					mode: this.mode,
					family,
					loc,
					text
				};
				if ((family === "rel" || family === "bin") && this.prevAtomType === "text") {
					if (textRegEx.test(loc.lexer.input.slice(loc.end))) s.needsSpacing = true;
				}
			} else {
				if (asciiFromScript[text]) {
					this.consume();
					const nextCode = this.fetch().text.charCodeAt(0);
					const font = nextCode === 65025 ? "mathscr" : "mathcal";
					if (nextCode === 65024 || nextCode === 65025) this.consume();
					return {
						type: "font",
						mode: "math",
						font,
						body: {
							type: "mathord",
							mode: "math",
							loc,
							text: asciiFromScript[text]
						}
					};
				}
				s = {
					type: group,
					mode: this.mode,
					loc,
					text
				};
			}
			symbol = s;
		} else if (text.charCodeAt(0) >= 128 || combiningDiacriticalMarksEndRegex.exec(text)) {
			if (this.settings.strict && this.mode === "math") throw new ParseError(`Unicode text character "${text[0]}" used in math mode`, nucleus);
			symbol = {
				type: "textord",
				mode: "text",
				loc: SourceLocation.range(nucleus),
				text
			};
		} else return null;
		this.consume();
		if (match) for (let i = 0; i < match[0].length; i++) {
			const accent = match[0][i];
			if (!unicodeAccents[accent]) throw new ParseError(`Unknown accent ' ${accent}'`, nucleus);
			const command = unicodeAccents[accent][this.mode] || unicodeAccents[accent].text;
			if (!command) throw new ParseError(`Accent ${accent} unsupported in ${this.mode} mode`, nucleus);
			symbol = {
				type: "accent",
				mode: this.mode,
				loc: SourceLocation.range(nucleus),
				label: command,
				isStretchy: false,
				base: symbol
			};
		}
		return symbol;
	}
};
/**
* Parses an expression using a Parser, then returns the parsed result.
*/
const parseTree = function(toParse, settings) {
	if (!(typeof toParse === "string" || toParse instanceof String)) throw new TypeError("Temml can only parse string typed expression");
	let tree;
	let parser;
	try {
		parser = new Parser(toParse, settings);
		delete parser.gullet.macros.current["\\df@tag"];
		tree = parser.parse();
	} catch (error) {
		if (error.toString() === "ParseError:  Unmatched delimiter") {
			settings.wrapDelimiterPairs = false;
			parser = new Parser(toParse, settings);
			delete parser.gullet.macros.current["\\df@tag"];
			tree = parser.parse();
		} else throw error;
	}
	if (!(tree.length > 0 && tree[0].type && tree[0].type === "array" && tree[0].addEqnNum)) {
		if (parser.gullet.macros.get("\\df@tag")) {
			if (!settings.displayMode) throw new ParseError("\\tag works only in display mode");
			parser.gullet.feed("\\df@tag");
			tree = [{
				type: "tag",
				mode: "text",
				body: tree,
				tag: parser.parse()
			}];
		}
	}
	return tree;
};
/**
* This file contains information about the style that the mathmlBuilder carries
* around with it. Data is held in an `Style` object, and when
* recursing, a new `Style` object can be created with the `.with*` functions.
*/
const subOrSupLevel = [
	2,
	2,
	3,
	3
];
/**
* This is the main Style class. It contains the current style.level, color, and font.
*
* Style objects should not be modified. To create a new Style with
* different properties, call a `.with*` method.
*/
var Style = class Style {
	constructor(data) {
		this.level = data.level;
		this.color = data.color;
		this.font = data.font || "";
		this.fontFamily = data.fontFamily || "";
		this.fontSize = data.fontSize || 1;
		this.fontWeight = data.fontWeight || "";
		this.fontShape = data.fontShape || "";
		this.maxSize = data.maxSize;
	}
	/**
	* Returns a new style object with the same properties as "this".  Properties
	* from "extension" will be copied to the new style object.
	*/
	extend(extension) {
		const data = {
			level: this.level,
			color: this.color,
			font: this.font,
			fontFamily: this.fontFamily,
			fontSize: this.fontSize,
			fontWeight: this.fontWeight,
			fontShape: this.fontShape,
			maxSize: this.maxSize
		};
		for (const key in extension) if (Object.prototype.hasOwnProperty.call(extension, key)) data[key] = extension[key];
		return new Style(data);
	}
	withLevel(n) {
		return this.extend({ level: n });
	}
	incrementLevel() {
		return this.extend({ level: Math.min(this.level + 1, 3) });
	}
	inSubOrSup() {
		return this.extend({ level: subOrSupLevel[this.level] });
	}
	/**
	* Create a new style object with the given color.
	*/
	withColor(color) {
		return this.extend({ color });
	}
	/**
	* Creates a new style object with the given math font or old text font.
	* @type {[type]}
	*/
	withFont(font) {
		return this.extend({ font });
	}
	/**
	* Create a new style objects with the given fontFamily.
	*/
	withTextFontFamily(fontFamily) {
		return this.extend({
			fontFamily,
			font: ""
		});
	}
	/**
	* Creates a new style object with the given font size
	*/
	withFontSize(num) {
		return this.extend({ fontSize: num });
	}
	/**
	* Creates a new style object with the given font weight
	*/
	withTextFontWeight(fontWeight) {
		return this.extend({
			fontWeight,
			font: ""
		});
	}
	/**
	* Creates a new style object with the given font weight
	*/
	withTextFontShape(fontShape) {
		return this.extend({
			fontShape,
			font: ""
		});
	}
	/**
	* Gets the CSS color of the current style object
	*/
	getColor() {
		return this.color;
	}
};
const version = "0.13.5";
function postProcess(block) {
	const labelMap = {};
	let i = 0;
	const amsEqns = document.getElementsByClassName("tml-eqn");
	for (let parent of amsEqns) {
		i += 1;
		parent.setAttribute("id", "tml-eqn-" + String(i));
		while (true) {
			if (parent.tagName === "mtable") break;
			if (parent.getElementsByClassName("tml-label").length > 0) {
				const id = parent.attributes.id.value;
				labelMap[id] = String(i);
				break;
			} else parent = parent.parentElement;
		}
	}
	const taggedEqns = document.getElementsByClassName("tml-tageqn");
	for (const parent of taggedEqns) if (parent.getElementsByClassName("tml-label").length > 0) {
		const tags = parent.getElementsByClassName("tml-tag");
		if (tags.length > 0) {
			const id = parent.attributes.id.value;
			labelMap[id] = tags[0].textContent;
		}
	}
	[...block.getElementsByClassName("tml-ref")].forEach((ref) => {
		const attr = ref.getAttribute("href");
		let str = labelMap[attr.slice(1)];
		if (ref.className.indexOf("tml-eqref") === -1) {
			str = str.replace(/^\(/, "");
			str = str.replace(/\)$/, "");
		} else {
			if (str.charAt(0) !== "(") str = "(" + str;
			if (str.slice(-1) !== ")") str = str + ")";
		}
		const mtext = document.createElementNS("http://www.w3.org/1998/Math/MathML", "mtext");
		mtext.appendChild(document.createTextNode(str));
		const math = document.createElementNS("http://www.w3.org/1998/Math/MathML", "math");
		math.appendChild(mtext);
		ref.textContent = "";
		ref.appendChild(math);
	});
}
const findEndOfMath = function(delimiter, text, startIndex) {
	let index = startIndex;
	let braceLevel = 0;
	const delimLength = delimiter.length;
	while (index < text.length) {
		const character = text[index];
		if (braceLevel <= 0 && text.slice(index, index + delimLength) === delimiter) return index;
		else if (character === "\\") index++;
		else if (character === "{") braceLevel++;
		else if (character === "}") braceLevel--;
		index++;
	}
	return -1;
};
const escapeRegex = function(string) {
	return string.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
};
const amsRegex = /^\\(?:begin|(?:eq)?ref){/;
const splitAtDelimiters = function(text, delimiters) {
	let index;
	const data = [];
	const regexLeft = new RegExp("(" + delimiters.map((x) => escapeRegex(x.left)).join("|") + ")");
	while (true) {
		index = text.search(regexLeft);
		if (index === -1) break;
		if (index > 0) {
			data.push({
				type: "text",
				data: text.slice(0, index)
			});
			text = text.slice(index);
		}
		const i = delimiters.findIndex((delim) => text.startsWith(delim.left));
		index = findEndOfMath(delimiters[i].right, text, delimiters[i].left.length);
		if (index === -1) break;
		const rawData = text.slice(0, index + delimiters[i].right.length);
		const math = amsRegex.test(rawData) ? rawData : text.slice(delimiters[i].left.length, index);
		data.push({
			type: "math",
			data: math,
			rawData,
			display: delimiters[i].display
		});
		text = text.slice(index + delimiters[i].right.length);
	}
	if (text !== "") data.push({
		type: "text",
		data: text
	});
	return data;
};
const defaultDelimiters = [
	{
		left: "$$",
		right: "$$",
		display: true
	},
	{
		left: "\\(",
		right: "\\)",
		display: false
	},
	{
		left: "\\begin{equation}",
		right: "\\end{equation}",
		display: true
	},
	{
		left: "\\begin{equation*}",
		right: "\\end{equation*}",
		display: true
	},
	{
		left: "\\begin{align}",
		right: "\\end{align}",
		display: true
	},
	{
		left: "\\begin{align*}",
		right: "\\end{align*}",
		display: true
	},
	{
		left: "\\begin{alignat}",
		right: "\\end{alignat}",
		display: true
	},
	{
		left: "\\begin{alignat*}",
		right: "\\end{alignat*}",
		display: true
	},
	{
		left: "\\begin{gather}",
		right: "\\end{gather}",
		display: true
	},
	{
		left: "\\begin{gather*}",
		right: "\\end{gather*}",
		display: true
	},
	{
		left: "\\begin{CD}",
		right: "\\end{CD}",
		display: true
	},
	{
		left: "\\ref{",
		right: "}",
		display: false
	},
	{
		left: "\\eqref{",
		right: "}",
		display: false
	},
	{
		left: "\\[",
		right: "\\]",
		display: true
	}
];
const firstDraftDelimiters = {
	"$": [
		{
			left: "$$",
			right: "$$",
			display: true
		},
		{
			left: "$`",
			right: "`$",
			display: false
		},
		{
			left: "$",
			right: "$",
			display: false
		}
	],
	"(": [{
		left: "\\[",
		right: "\\]",
		display: true
	}, {
		left: "\\(",
		right: "\\)",
		display: false
	}]
};
const amsDelimiters = [
	{
		left: "\\begin{equation}",
		right: "\\end{equation}",
		display: true
	},
	{
		left: "\\begin{equation*}",
		right: "\\end{equation*}",
		display: true
	},
	{
		left: "\\begin{align}",
		right: "\\end{align}",
		display: true
	},
	{
		left: "\\begin{align*}",
		right: "\\end{align*}",
		display: true
	},
	{
		left: "\\begin{alignat}",
		right: "\\end{alignat}",
		display: true
	},
	{
		left: "\\begin{alignat*}",
		right: "\\end{alignat*}",
		display: true
	},
	{
		left: "\\begin{gather}",
		right: "\\end{gather}",
		display: true
	},
	{
		left: "\\begin{gather*}",
		right: "\\end{gather*}",
		display: true
	},
	{
		left: "\\begin{CD}",
		right: "\\end{CD}",
		display: true
	},
	{
		left: "\\ref{",
		right: "}",
		display: false
	},
	{
		left: "\\eqref{",
		right: "}",
		display: false
	}
];
const delimitersFromKey = (key) => {
	if (key === "$" || key === "(") return firstDraftDelimiters[key];
	else if (key === "$+" || key === "(+") return firstDraftDelimiters[key.slice(0, 1)].concat(amsDelimiters);
	else if (key === "ams") return amsDelimiters;
	else if (key === "all") return firstDraftDelimiters["("].concat(firstDraftDelimiters["$"]).concat(amsDelimiters);
	else return defaultDelimiters;
};
const renderMathInText = function(text, optionsCopy) {
	const data = splitAtDelimiters(text, optionsCopy.delimiters);
	if (data.length === 1 && data[0].type === "text") return null;
	const fragment = document.createDocumentFragment();
	for (let i = 0; i < data.length; i++) if (data[i].type === "text") fragment.appendChild(document.createTextNode(data[i].data));
	else {
		const span = document.createElement("span");
		let math = data[i].data;
		optionsCopy.displayMode = data[i].display;
		try {
			if (optionsCopy.preProcess) math = optionsCopy.preProcess(math);
			temml.render(math, span, optionsCopy);
		} catch (e) {
			if (!(e instanceof ParseError)) throw e;
			optionsCopy.errorCallback("Temml auto-render: Failed to parse `" + data[i].data + "` with ", e);
			fragment.appendChild(document.createTextNode(data[i].rawData));
			continue;
		}
		fragment.appendChild(span);
	}
	return fragment;
};
const renderElem = function(elem, optionsCopy) {
	for (let i = 0; i < elem.childNodes.length; i++) {
		const childNode = elem.childNodes[i];
		if (childNode.nodeType === 3) {
			const frag = renderMathInText(childNode.textContent, optionsCopy);
			if (frag) {
				i += frag.childNodes.length - 1;
				elem.replaceChild(frag, childNode);
			}
		} else if (childNode.nodeType === 1) {
			const className = " " + childNode.className + " ";
			if (optionsCopy.ignoredTags.indexOf(childNode.nodeName.toLowerCase()) === -1 && optionsCopy.ignoredClasses.every((x) => className.indexOf(" " + x + " ") === -1)) renderElem(childNode, optionsCopy);
		}
	}
};
const renderMathInElement = function(elem, options) {
	if (!elem) throw new Error("No element provided to render");
	const optionsCopy = {};
	for (const option in options) if (Object.prototype.hasOwnProperty.call(options, option)) optionsCopy[option] = options[option];
	if (optionsCopy.fences) optionsCopy.delimiters = delimitersFromKey(optionsCopy.fences);
	else optionsCopy.delimiters = optionsCopy.delimiters || defaultDelimiters;
	optionsCopy.ignoredTags = optionsCopy.ignoredTags || [
		"script",
		"noscript",
		"style",
		"textarea",
		"pre",
		"code",
		"option"
	];
	optionsCopy.ignoredClasses = optionsCopy.ignoredClasses || [];
	optionsCopy.errorCallback = optionsCopy.errorCallback || console.error;
	optionsCopy.macros = optionsCopy.macros || Object.create(null);
	renderElem(elem, optionsCopy);
	postProcess(elem);
};
/**
* This is the main entry point for Temml. Here, we expose functions for
* rendering expressions either to DOM nodes or to markup strings.
*
* We also expose the ParseError class to check if errors thrown from Temml are
* errors in the expression, or errors in javascript handling.
*/
/**
* @type {import('./temml').render}
* Parse and build an expression, and place that expression in the DOM node
* given.
*/
let render$1 = function(expression, baseNode, options = {}) {
	baseNode.textContent = "";
	const alreadyInMathElement = baseNode.tagName.toLowerCase() === "math";
	if (alreadyInMathElement) options.wrap = "none";
	const math = renderToMathMLTree(expression, options);
	if (alreadyInMathElement) {
		baseNode.textContent = "";
		math.children.forEach((e) => {
			baseNode.appendChild(e.toNode());
		});
	} else if (math.children.length > 1) {
		baseNode.textContent = "";
		math.children.forEach((e) => {
			baseNode.appendChild(e.toNode());
		});
	} else baseNode.appendChild(math.toNode());
};
if (typeof document !== "undefined") {
	if (document.compatMode !== "CSS1Compat") {
		typeof console !== "undefined" && console.warn("Warning: Temml doesn't work in quirks mode. Make sure your website has a suitable doctype.");
		render$1 = function() {
			throw new ParseError("Temml doesn't work in quirks mode.");
		};
	}
}
/**
* @type {import('./temml').renderToString}
* Parse and build an expression, and return the markup for that.
*/
const renderToString = function(expression, options) {
	return renderToMathMLTree(expression, options).toMarkup();
};
/**
* @type {import('./temml').generateParseTree}
* Parse an expression and return the parse tree.
*/
const generateParseTree = function(expression, options) {
	const settings = new Settings(options);
	return parseTree(expression, settings);
};
/**
* @type {import('./temml').definePreamble}
* Take an expression which contains a preamble.
* Parse it and return the macros.
*/
const definePreamble = function(expression, options) {
	const settings = new Settings(options);
	settings.macros = Object.create(null);
	if (!(typeof expression === "string" || expression instanceof String)) throw new TypeError("Temml can only parse string typed expression");
	const parser = new Parser(expression, settings, true);
	delete parser.gullet.macros.current["\\df@tag"];
	return parser.parse();
};
/**
* If the given error is a Temml ParseError,
* renders the invalid LaTeX as a span with hover title giving the Temml
* error message.  Otherwise, simply throws the error.
*/
const renderError = function(error, expression, options) {
	if (options.throwOnError || !(error instanceof ParseError)) throw error;
	const node = new Span(["temml-error"], [new TextNode$1(expression + "\n\n" + error.toString())]);
	node.style.color = options.errorColor;
	node.style.whiteSpace = "pre-line";
	return node;
};
/**
* @type {import('./temml').renderToMathMLTree}
* Generates and returns the Temml build tree. This is used for advanced
* use cases (like rendering to custom output).
*/
const renderToMathMLTree = function(expression, options) {
	const settings = new Settings(options);
	try {
		return buildMathML(parseTree(expression, settings), expression, new Style({
			level: settings.displayMode ? StyleLevel.DISPLAY : StyleLevel.TEXT,
			maxSize: settings.maxSize
		}), settings);
	} catch (error) {
		return renderError(error, expression, settings);
	}
};
//#endregion
//#region lib/types/omml.js
/**
* LaTeX formulas to OMML, the equation format of .docx files.
*
* Microsoft Word builds its equations itself, from the linear text the helper writes; WPS Office does not. For WPS
* the formulas are converted here — temml reads the LaTeX into a MathML tree, and that tree is written out as OMML —
* and the helper inserts the result as it is.
*/
const render = {
	/**
	* Current Temml version
	*/
	version,
	/**
	* Renders the given LaTeX into MathML, and adds
	* it as a child to the specified DOM node.
	*/
	render: render$1,
	/**
	* Renders the given LaTeX into MathML string,
	* for sending to the client.
	*/
	renderToString,
	/**
	* Finds all the math delimiters in a given element of a running HTML document
	* and converts the contents of each instance into a <math> element.
	*/
	renderMathInElement,
	/**
	* Post-process an entire HTML block.
	* Writes AMS auto-numbers and implements \ref{}.
	* Typcally called once, after a loop has rendered many individual spans.
	*/
	postProcess,
	/**
	* Temml error, usually during parsing.
	*/
	ParseError,
	/**
	* Creates a set of macros with document-wide scope.
	*/
	definePreamble,
	/**
	* Parses the given LaTeX into Temml's internal parse tree structure,
	* without rendering to HTML or MathML.
	*
	* NOTE: This method is not currently recommended for public use.
	* The internal tree representation is unstable and is very likely
	* to change. Use at your own risk.
	*/
	__parse: generateParseTree,
	/**
	* Renders the given LaTeX into a MathML internal DOM tree
	* representation, without flattening that representation to a string.
	*
	* NOTE: This method is not currently recommended for public use.
	* The internal tree representation is unstable and is very likely
	* to change. Use at your own risk.
	*/
	__renderToMathMLTree: renderToMathMLTree,
	/**
	* adds a new symbol to builtin symbols table
	*/
	__defineSymbol: defineSymbol,
	/**
	* adds a new macro to builtin macro list
	*/
	__defineMacro: defineMacro
}.__renderToMathMLTree;
const NARY = /* @__PURE__ */ new Set([
	"∑",
	"∏",
	"∐",
	"∫",
	"∬",
	"∭",
	"∮",
	"∯",
	"∰",
	"⋃",
	"⋂",
	"⋁",
	"⋀",
	"⨁",
	"⨂",
	"⨀"
]);
/** What MathML writes over a letter, and the combining mark OMML wants for it. */
const ACCENTS = {
	"ˆ": "̂",
	"^": "̂",
	"˜": "̃",
	"~": "̃",
	"→": "⃗",
	"˙": "̇",
	"¨": "̈",
	"ˇ": "̌",
	"˘": "̆",
	"´": "́",
	"`": "̀",
	"⃗": "⃗",
	"˚": "̊",
	"←": "⃖",
	"↔": "⃡"
};
const BARS = /* @__PURE__ */ new Set([
	"‾",
	"¯",
	"―",
	"_",
	"̲"
]);
const esc = (text) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const FONT = "<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr>";
function run(text, style = "") {
	if (text.length === 0) return "";
	return `<m:r>${style === "normal" ? "<m:rPr><m:nor/></m:rPr>" : style === "plain" ? "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" : ""}${FONT}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}
const textOf = (node) => node === void 0 ? "" : node.text !== void 0 && node.type === void 0 ? node.text : (node.children ?? []).map(textOf).join("");
const isToken = (node, type) => node !== void 0 && node.type === type;
/** The children of a node as one sequence: rows and the like that only group are opened up. */
function flat(nodes) {
	const out = [];
	for (const node of nodes) if (node.type === void 0 && node.children !== void 0) out.push(...flat(node.children));
	else if ((node.type === "mrow" || node.type === "mstyle" || node.type === "mpadded" || node.type === "semantics") && !isFenced(node)) out.push(...flat(node.children ?? []));
	else out.push(node);
	return out;
}
function isFenced(node) {
	const children = node.children ?? [];
	if (node.type !== "mrow" || children.length < 2) return false;
	const first = children[0], last = children[children.length - 1];
	return isToken(first, "mo") && first.attributes?.fence === "true" && isToken(last, "mo") && last.attributes?.fence === "true";
}
const naryOf = (node) => {
	if (![
		"msubsup",
		"munderover",
		"msub",
		"msup",
		"munder",
		"mover"
	].includes(node.type ?? "")) return void 0;
	const base = node.children?.[0];
	const sign = isToken(base, "mo") ? textOf(base) : "";
	return NARY.has(sign) ? sign : void 0;
};
function space(node) {
	const width = /^(-?[\d.]+)em$/.exec(node.attributes?.width ?? "");
	const em = width ? Number(width[1]) : 0;
	if (em >= 1.5) return run("  ");
	if (em >= .8) return run(" ");
	if (em >= .2) return run(" ");
	if (em >= .1) return run(" ");
	return "";
}
function table(node) {
	const rows = (node.children ?? []).filter((row) => row.type === "mtr" || row.type === "mlabeledtr");
	const columns = Math.max(1, ...rows.map((row) => (row.children ?? []).length));
	const side = [];
	for (let c = 0; c < columns; c++) {
		const cell = rows.map((row) => row.children?.[c]).find(Boolean);
		side.push(cell?.classes?.includes("tml-left") ? "left" : cell?.classes?.includes("tml-right") ? "right" : "center");
	}
	return `<m:m><m:mPr><m:mcs>${side.map((jc) => `<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="${jc}"/></m:mcPr></m:mc>`).join("")}</m:mcs></m:mPr>${rows.map((row) => {
		const cells = (row.children ?? []).map((cell) => `<m:e>${seq(cell.children ?? [])}</m:e>`);
		while (cells.length < columns) cells.push("<m:e/>");
		return `<m:mr>${cells.join("")}</m:mr>`;
	}).join("")}</m:m>`;
}
const arg = (node) => node === void 0 ? "" : seq([node]);
function one(node) {
	const kids = node.children ?? [];
	switch (node.type) {
		case "mi": {
			const text = textOf(node);
			return run(text, node.attributes?.mathvariant === "normal" || [...text].length > 1 ? "plain" : "");
		}
		case "mn": return run(textOf(node));
		case "mo": {
			const text = textOf(node);
			return text === "⁡" || text === "⁢" || text === "⁣" ? "" : run(text);
		}
		case "mtext": return run(textOf(node), "normal");
		case "mspace": return space(node);
		case "mfrac": return `<m:f>${node.attributes?.linethickness === "0px" || node.attributes?.linethickness === "0" ? "<m:fPr><m:type m:val=\"noBar\"/></m:fPr>" : ""}<m:num>${arg(kids[0])}</m:num><m:den>${arg(kids[1])}</m:den></m:f>`;
		case "msqrt": return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${seq(kids)}</m:e></m:rad>`;
		case "mroot": return `<m:rad><m:deg>${arg(kids[1])}</m:deg><m:e>${arg(kids[0])}</m:e></m:rad>`;
		case "msup": return `<m:sSup><m:e>${arg(kids[0])}</m:e><m:sup>${arg(kids[1])}</m:sup></m:sSup>`;
		case "msub": return `<m:sSub><m:e>${arg(kids[0])}</m:e><m:sub>${arg(kids[1])}</m:sub></m:sSub>`;
		case "msubsup": return `<m:sSubSup><m:e>${arg(kids[0])}</m:e><m:sub>${arg(kids[1])}</m:sub><m:sup>${arg(kids[2])}</m:sup></m:sSubSup>`;
		case "mmultiscripts": return `<m:sPre><m:sub>${arg(kids[3])}</m:sub><m:sup>${arg(kids[4])}</m:sup><m:e>${arg(kids[0])}</m:e></m:sPre>`;
		case "mover": {
			const mark = textOf(kids[1]);
			if (isToken(kids[1], "mo") && BARS.has(mark)) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${arg(kids[0])}</m:e></m:bar>`;
			if (isToken(kids[1], "mo") && ACCENTS[mark] !== void 0) return `<m:acc><m:accPr><m:chr m:val="${ACCENTS[mark]}"/></m:accPr><m:e>${arg(kids[0])}</m:e></m:acc>`;
			if (isToken(kids[1], "mo") && mark === "⏞") return `<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr><m:e>${arg(kids[0])}</m:e></m:groupChr>`;
			return `<m:limUpp><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limUpp>`;
		}
		case "munder": {
			const mark = textOf(kids[1]);
			if (isToken(kids[1], "mo") && BARS.has(mark)) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${arg(kids[0])}</m:e></m:bar>`;
			if (isToken(kids[1], "mo") && mark === "⏟") return `<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/><m:vertJc m:val="top"/></m:groupChrPr><m:e>${arg(kids[0])}</m:e></m:groupChr>`;
			return `<m:limLow><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limLow>`;
		}
		case "munderover": return `<m:limUpp><m:e><m:limLow><m:e>${arg(kids[0])}</m:e><m:lim>${arg(kids[1])}</m:lim></m:limLow></m:e><m:lim>${arg(kids[2])}</m:lim></m:limUpp>`;
		case "menclose": {
			const notation = node.attributes?.notation ?? "";
			if (notation.includes("top")) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${seq(kids)}</m:e></m:bar>`;
			if (notation.includes("bottom")) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${seq(kids)}</m:e></m:bar>`;
			if (notation.includes("box")) return `<m:borderBox><m:e>${seq(kids)}</m:e></m:borderBox>`;
			return seq(kids);
		}
		case "mtable": return table(node);
		case "mphantom": return "";
		case "mrow":
			if (isFenced(node)) {
				const open = textOf(kids[0]), close = textOf(kids[kids.length - 1]);
				return `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/></m:dPr><m:e>${seq(kids.slice(1, -1))}</m:e></m:d>`;
			}
			return seq(kids);
		default: return seq(kids);
	}
}
/** A run of sibling nodes. A sum, product or integral takes the node after it as what it applies to. */
function seq(nodes) {
	const list = flat(nodes);
	let out = "";
	for (let i = 0; i < list.length; i++) {
		const node = list[i];
		const sign = naryOf(node);
		if (sign === void 0) {
			out += one(node);
			continue;
		}
		const kids = node.children ?? [];
		const under = node.type === "munderover" || node.type === "munder" || node.type === "mover";
		const sub = node.type === "msup" || node.type === "mover" ? void 0 : kids[1];
		const sup = node.type === "msup" || node.type === "mover" ? kids[1] : node.type === "msubsup" || node.type === "munderover" ? kids[2] : void 0;
		let operand = "";
		while (i + 1 < list.length && list[i + 1].type === "mspace") i++;
		if (i + 1 < list.length && !(isToken(list[i + 1], "mo") && !isFenced(list[i + 1]))) operand = one(list[++i]);
		out += `<m:nary><m:naryPr><m:chr m:val="${sign}"/><m:limLoc m:val="${under ? "undOvr" : "subSup"}"/>${sub === void 0 ? "<m:subHide m:val=\"1\"/>" : ""}${sup === void 0 ? "<m:supHide m:val=\"1\"/>" : ""}</m:naryPr><m:sub>${arg(sub)}</m:sub><m:sup>${arg(sup)}</m:sup><m:e>${operand}</m:e></m:nary>`;
	}
	return out;
}
/** One formula as an `<m:oMath>` element, or undefined when the LaTeX cannot be read. */
function latexToOmml(tex, display) {
	try {
		const body = seq(render(tex, {
			displayMode: display,
			annotate: false,
			throwOnError: true,
			strict: false
		}).children ?? []);
		return body.length === 0 ? void 0 : `<m:oMath>${body}</m:oMath>`;
	} catch {
		return;
	}
}
const DOLLARS = /\$\$([^$\r\n]+?)\$\$|\$([^$\r\n]+?)\$(?!\$)/g;
const TAG = /\\tag\s*\{([^}]*)\}/;
/**
* Every formula written between dollar signs anywhere in a batch of operations, converted: keyed "d:" + source for
* one on a line of its own and "i:" + source for one inside a line, the source as the helper will read it.
*/
function formulasOf(value, found = {}) {
	if (typeof value === "string") {
		if (!value.includes("$")) return found;
		for (const match of value.matchAll(DOLLARS)) {
			const display = match[1] !== void 0;
			const source = (match[1] ?? match[2] ?? "").trim().replace(TAG, "").trim();
			const key = (display ? "d:" : "i:") + source;
			if (source.length === 0 || found[key] !== void 0) continue;
			const omml = latexToOmml(source, display);
			if (omml !== void 0) found[key] = omml;
		}
	} else if (Array.isArray(value)) for (const item of value) formulasOf(item, found);
	else if (value !== null && typeof value === "object") for (const [name, item] of Object.entries(value)) if (name === "latex" && typeof item === "string") {
		const source = item.trim().replace(/^\$+|\$+$/g, "").trim();
		const omml = source.length === 0 || found[`d:${source}`] !== void 0 ? void 0 : latexToOmml(source, true);
		if (omml !== void 0) found[`d:${source}`] = omml;
	} else formulasOf(item, found);
	return found;
}
//#endregion
//#region lib/types/tools.js
/**
* Model-facing tools. Six of them, each doing a whole step: the schemas are
* sent with every request, and a batch of edits in one call is one model round
* instead of one per change.
*/
const { extname, isAbsolute } = win32;
const IMAGE = {
	type: "object",
	additionalProperties: false,
	properties: {
		attachmentId: {
			type: "string",
			required: true
		},
		mediaType: {
			type: "string",
			required: true
		},
		bytes: {
			type: "integer",
			required: true
		},
		width: {
			type: "integer",
			required: true
		},
		height: {
			type: "integer",
			required: true
		},
		name: { type: "string" },
		originalDimensions: {
			type: "object",
			additionalProperties: false,
			properties: {
				width: {
					type: "integer",
					required: true
				},
				height: {
					type: "integer",
					required: true
				}
			}
		}
	}
};
const output = {
	schema: {
		type: "object",
		additionalProperties: false,
		properties: {
			text: {
				type: "string",
				required: true
			},
			image: IMAGE,
			more: {
				type: "array",
				items: IMAGE
			}
		}
	},
	render: (_args, value) => [{
		type: "text",
		text: value.text
	}, ...[value.image, ...value.more ?? []].filter((image) => image !== void 0).map((image) => ({
		type: "image",
		attachment: image
	}))]
};
function toJson(ref) {
	return {
		attachmentId: String(ref.attachmentId),
		mediaType: ref.mediaType,
		bytes: ref.bytes,
		width: ref.width,
		height: ref.height,
		...ref.name === void 0 ? {} : { name: ref.name },
		...ref.originalDimensions === void 0 ? {} : { originalDimensions: {
			width: ref.originalDimensions.width,
			height: ref.originalDimensions.height
		} }
	};
}
const card = (title) => ({
	card: "generic",
	title,
	kind: "other"
});
const EXTENSIONS = {
	".docx": "word",
	".doc": "word",
	".docm": "word",
	".rtf": "word",
	".dotx": "word",
	".xlsx": "excel",
	".xls": "excel",
	".xlsm": "excel",
	".xlsb": "excel",
	".csv": "excel",
	".pptx": "ppt",
	".ppt": "ppt",
	".pptm": "ppt"
};
/** Which app a call is about: said outright, or told from the file name. */
function appOf(args) {
	if (args.app !== void 0) {
		if (args.app !== "word" && args.app !== "excel" && args.app !== "ppt") throw new Error("app must be word, excel or ppt.");
		return args.app;
	}
	const kind = EXTENSIONS[extname(args.path ?? args.doc ?? "").toLowerCase()];
	if (!kind) throw new Error("Say which app: pass app (word, excel or ppt), or name the document with its extension.");
	return kind;
}
const fileName = (doc) => (doc ?? "").split(/[\\/]/).pop() ?? "";
const DOC = {
	type: "string",
	description: "The open document: its full path or its window name (e.g. report.docx). Default: the document you last opened or worked on."
};
const APP = {
	type: "string",
	enum: [
		"word",
		"excel",
		"ppt"
	],
	description: "Only to pick the active document of another app, or when doc has no extension."
};
const SILENT = `Silent mode is on (the user chose it in the settings): the document is open in the background, with no window, and the user is not watching. Work as usual — office_edit, office_render to look at the result, office_save — only unseen. When you stop, what was opened in the background is saved and closed by itself; say in your answer where the file is.
- Excel data work may use Python here when that is quicker (pandas / openpyxl for cleaning, joining and aggregating large tables): run it on the file on disk while the workbook is NOT open in Excel (before office_open, or after office_close), or compute in Python and put the results in with write_range. Charts, number formats, conditional formats and layout still go through office_edit, which lets Excel itself draw them. Word and PowerPoint documents are always built with office_edit: file libraries lose their layout.
- Look at the result with office_render before you finish, exactly as otherwise: nobody else sees it until it is done.
- If something can only be done on the window itself with mouse / keyboard tools, call office_open {path, show: true} first: it brings the document into view. That is slow; use it only when office_edit has no operation for the job.`;
const TAIL = `Colours are "#RRGGBB"; align is left | center | right | justify; file paths are absolute.`;
const EDIT_CORE = `Edit a document that is open in Word, Excel or PowerPoint, live: each change appears in the window, and the user can keep working in it. Runs ops in order and stops at the first that fails. Does not save (use office_save). Word edits are one undo step for the user; Excel and PowerPoint edits cannot be undone with Ctrl+Z.
Each op is {op: name, ...fields}. The operations of an app and their fields are listed in the result of your first office_open / office_read for that app, and again by office_help {app} at any time: read that list before you write ops, and do not guess names or fields.
Formulas are LaTeX between dollar signs in any text you write ($...$ in a sentence, $$...$$ on a line of its own) and become native equations. ${TAIL}`;
/** The operations of each app. Not sent with every request: handed over when the app is first used, and by office_help. */
const GUIDES = {
	word: `Word ops. para = paragraph number from office_read; pass expect = its first words so a number made stale by other edits is caught. where = after (default with para) | before | start | end (default without para: the very end of the document).
In a form or template whose sections are table cells, write INSIDE the cell: give para = a paragraph of that cell. Without para the text or table lands after everything, outside the form.
When you add to a document that already has headings of the kind you are writing (（一）.., 1.1 ..), give yours the same look: the same style, bold and size, and the same indent — firstLineIndent: 0 when theirs start at the margin, since a body style with a first-line indent would push yours in. Compare them on the page with office_render.
Formulas: write LaTeX between dollar signs in any text (paragraphs, table cells, captions, replace_text): $...$ inside a sentence, $$...$$ as a paragraph of its own. Every symbol, subscript, power, norm or matrix in running text goes between dollar signs too ($\\kappa(A)$, $\\|x\\|_2$, $10^{-8}$, \\begin{pmatrix}..\\end{pmatrix}, \\begin{cases}..\\end{cases}); do not type them with Unicode characters such as ‖x‖₂ or κ₂. Number a display formula when the document numbers its formulas by ending it with \\tag{1}: $$T=2\\pi\\sqrt{l/g} \\tag{1}$$ puts (1) flush right. They become native Word equations; Word itself reads the LaTeX, so ordinary LaTeX works as you would write it (\\mathfrak, \\mathcal, \\bigoplus, \\cong, \\operatorname, \\not\\equiv, matrices, cases, aligned ...) — there is no need to try symbols out first. Never write a formula as plain text such as T = 2π√(l/g). A unit after a number inside a formula is upright: $$D=(42.0\\pm0.5)\\ \\mathrm{mm}$$, $\\mathrm{m/s}$, $\\mathrm{m\\cdot s^{-1}}$.
- insert_paragraphs {items:[{text, style?}], para?, expect?, where?} — style: Normal, Title, Heading 1..6, List Bullet, List Number, Quote, or a style name the document has. Without style a paragraph is body text in the font, size and spacing of the text it is placed next to. A newline in text starts another paragraph. Items also take the format_text fields.
- set_text {para, expect, text} — rewrite one paragraph, keeping its style. expect is required ("" for an empty paragraph).
- Native footnotes: insert_footnote {para, expect, text, find?, occurrence?, where?: before|after, ...format fields}. Without find inserts at paragraph end; with find defaults to after its first occurrence. Native numbering and Footnote Text/Reference styles follow the document. Text may contain **bold** and $LaTeX$.
- edit_footnote {footnote, expect, text?, ...format fields}; delete_footnote {footnote, expect}. footnote is the collection index returned by office_read, NOT the displayed number. expect is the first words of the note text, required even for formatting. Reread after inserts/deletes. format_text also accepts {footnote, expect, find?, ...format fields} instead of para/to.
- footnote_settings {numbering: continuous|section|page, section?} — changes numbering for the document, or one section. Insertions preserve existing settings. Read all notes with office_read {footnotes:true, full:true}; by default only notes anchored in the paragraphs read are listed. The body shows [footnote:i].
- Inline footnotes in insert_paragraphs/set_text/replace_text: [fn:full citation] or \\footnote{full citation}, multiple per paragraph. Balanced braces/brackets and escaped delimiters are supported, nested/empty/unclosed notes are rejected. Escape the opener with a backslash to keep literal syntax. set_text refuses paragraphs with existing footnotes; use a smaller replace_text that does not cross their reference marks. delete_range explicitly removes notes anchored in its deleted text.
- replace_text {find, replace, all?} — find is literal text within one paragraph; replace may hold formulas and citations.
- format_text {para?, to?, find?, style?, font?, size?, bold?, italic?, latinFont?, underline?, superscript?, subscript?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing?} — para..to, or the first match of find (e.g. {find:"关键词：", bold:true}). In any text you write, **摘要：** makes that part bold. lineSpacing: a number is a multiple (1.5 = 1.5 倍行距), "20pt" an exact height (固定值 20 磅), "at least 20pt" a minimum (最小值). Paragraphs that hold a formula or a picture automatically get an exact height as a minimum instead, so nothing tall is cut off — do not fix that by hand. firstLineIndent: 2 = two characters (首行缩进 2 字符, same as indentChars), 0 = none, "0.74cm" / "24pt" a length.
- delete_range {para, expect, to?} — expect is required; para..to in one operation removes a run of paragraphs (leftover hints, surplus blank lines). A table goes with them when the run covers all of it: to remove a table, give its first and last paragraph as office_read lists them (i–to); a run that covers only part of a table empties those cells and the table stays. Paragraph numbers shift after every insert or delete; each result tells you the new numbers. A table counts one paragraph for each cell and one more at the end of each row (3 rows of 4 cells: 15).
- insert_table {data:[[cell,..],..], caption?, borders?: "grid"|"three-line"|"none", para?, where?, header?} — borders "three-line" is the 三线表 of papers (heavy rule above and below, light rule under the header, no other lines). A table goes inside a cell of another table only with nested: true. Several tables or pictures in ONE call: each insert shifts every paragraph number after it (a table adds a paragraph per cell), so give them from the END of the document towards the start, or read again between them. set_cell {table, row, col, text}
- format_table {table | para, borders?, size?, font?, latinFont?, align?, header?, rowHeight?, lineSpacing?, autofit?: "window"|"content"|"fixed", keepTogether?: true} — change a table that is already there; para (any paragraph of it) names a table that sits inside a cell of another and so has no number. keepTogether: when a table breaks across two pages and looks bad that way (two or three rows cut off from the rest), this puts it whole on one page — unless that would leave more than about a third of the page before it empty, in which case it tells you so and you change the text above instead ("force" moves it regardless). A long table may simply run over the page.
- Captions: give caption (the title only, no "表 1") on insert_table / insert_image. It is set the standard way: numbered automatically ("表 1", "图 1"), centred, above a table and below a figure. Do not write captions as ordinary paragraphs.
- insert_diagram {nodes: [{id, text, row, col, shape?}], edges: [{from, to, label?}], caption?, para?, where?} — a flow chart or block diagram drawn with Word's own shapes (the user can edit every box and line), in line with the text like a picture. Place the nodes on a grid: row 1 is the top, col 1 the left; a top-down flow is one column with the branches in the columns beside it. Shapes: rect (a step; the default), round, diamond (a decision), terminator (start / end), io (input / output), document, database, ellipse. Edges are arrows that find their own way with square corners: downwards bottom to top, sideways right to left, back up around the right; say fromSide / toSide (top | bottom | left | right) to choose, label for "是" / "否" on a branch, dashed: true, arrow: false. style: "outline" (the default: white boxes with thin black lines in the document's own type — right for papers, theses and formal documents), "modern" (flat tinted boxes, start / end in the accent colour, decisions in a warm tone, thin grey arrows, a sans type — only when the user asks for colour or the document itself is colourful), "classic" (pale blue with dark outlines); accent: "#RRGGBB" recolours the modern look to match the document. Other options: cellWidth (150), cellHeight (76), nodeWidth (112), nodeHeight (40), size, font, fill, color, ink; a node may give its own width, height, fill, color, bold, or box: [x, y, w, h] in points instead of row / col. Keep node text short (up to about 8 Chinese characters a line; "\n" breaks a line). Use this, not a picture drawn with matplotlib, for flow charts, process diagrams and structure diagrams in a document.
- insert_image {path, caption?, para?, where?, width?} — the picture is centred on a line of its own. width is in points, or "12cm". set_image {image: n | para, width?, height?, align?} resizes a picture that is already there.
- In insert_paragraphs, an item may also be a picture {path, caption?, width?} or a table {data, caption?, borders?}: it is put in right there, between the paragraphs before and after it — no paragraph numbers needed.
- Tables, pictures, their size and their borders are all done with these operations; do not use the mouse for them.
- Citations: write \\cite{1}, \\cite{2,5} or \\cite{3-6} in the text where the source is used; they become superscript [1] that jump to the reference. List the sources with insert_references {items:["Author. Title[M]. ...", ..], para?, where?} (numbered [1], [2].. in order; put it under a "参考文献" heading). A [1] typed by hand in the text, and a reference list typed as "[1] ..." paragraphs, are converted the same way.
- style_format {style, font?, latinFont?, size?, bold?, italic?, color?, align?, indentChars?, firstLineIndent?, spaceBefore?, spaceAfter?, lineSpacing? (1.5 | "20pt" | "at least 20pt"), pageBreakBefore?, numbering?: false} — change what a style looks like everywhere it is used (e.g. Heading 1 in 黑体 三号 without automatic numbers; Normal in 宋体 小四, 1.5 lines, 2-character first-line indent). Set the styles first, then write; this replaces fiddling with Word's style dialogs.
- fill_gaps {} — where a figure or table did not fit under the text before it and went to the next page, leaving the foot of that page empty, the paragraphs that follow it are brought before it (as a typesetter would); and a table from which the page break cut off a single row (its last row alone on the next page, or just its header and one row at the foot of the first) gets its rows together again. Run it once the document is written, before the final look.
- Figures you draw yourself (matplotlib etc.): draw each for the width it will have on the page (a full-width figure is about 6 in wide: figsize=(6, 3.6)), with type of 9–10.5 pt at that size, so it is not shrunk; one idea per figure rather than several panels in a row; leave room so labels, legends and lines do not overlap (constrained layout; look at the picture before you insert it); save as .svg (vector, stays sharp; set svg.fonttype to "path" for Chinese labels) or as .png at 300 dpi. Insert without "width" (it then keeps the size it was drawn at) or with the width it was drawn for. If the result says a figure is shown much smaller than it was drawn, draw it again for that size.
- Comments (native Word comments; the user reads them in the review pane): insert_comment {text, para?, expect?, find?, occurrence?, to?, author?, footnote?} — para..to covers whole paragraphs, find marks one phrase in them (para alone without find is refused: numbers may be stale), footnote pins the comment in that footnote. Read them with office_read (each comment lists the paragraph it sits on); edit_comment {comment, expect, text?, author?, initials?} and delete_comment {comment, expect} take the number office_read gave (or expect, its first words). Numbers change after a delete: read again.
- Links: add_hyperlink {find, text?, url | anchor, para?, expect?} — find becomes a link (text replaces what the link shows); anchor links inside the document and must name an existing bookmark. add_bookmark {name, para?, expect?, find?}; edit_hyperlink {link, expect, text?, url?, anchor?}; delete_hyperlink {link, expect} (the words stay). Links are listed by office_read as {i, para, text, url | anchor}.
- watermark {text, fontSize?, color?, angle?, clear?} — a pale word turned across the page behind the text on every section (draft marks, company marks); clear: true removes it. It is a header shape named DSHWatermark.
- page_setup {paper?: A4|A3|B5|Letter, orientation?, top?, bottom?, left?, right?} (margins in cm); page_numbers {align?, start?}; header {text}; page_break {para, expect} (that paragraph starts a new page); insert_toc {title?, levels?, para?, where?}; update_fields {}`,
	excel: `Excel ops. sheet = sheet name (default: the active sheet). A range may name its sheet: "产品表!A1:E13". Where an op takes "range" for a block of data and you leave it out, it means everything filled on the sheet, the first row being the headers; a column is named by its header ("金额"), or by its letter.
Reading: office_read shows the first 60 rows; for a longer table it adds a profile of every column (kinds of values, min / median / max, the distinct values with counts, and what looks wrong: stray spaces, numbers or dates written as text, mixed kinds, far-off values, empty and repeated rows). Work from the profile, do not page through thousands of rows; office_read {profile: true, range?} gives it for any block. Dates are shown as dates.
- write_range {range: top-left cell, values:[[..],..]} — a string starting with "=" is a formula; also takes the format_range fields.
- fill_formula {range: "F2:F500", formula: "=D2*E2"} — one formula down (or across) a range: write it for the first cell, relative references adjust; use $ to pin. add_column {header, formula | value, after?: column, numberFormat?, values?: true} — a new column at the right of the data (or after a column) with the formula filled down all data rows, written for row 2; values: true leaves the results instead of formulas; an existing header is filled again. Both report cells that show an error.
- Cleaning (each says how many cells or rows it changed, and for a column of few values what it holds afterwards):
  · clean {column | columns, trim?: true, noSpaces?, remove?: ["市"], case?: upper|lower|proper, map?: {"East": "华东", ..}, to?: number|date|text, year?} — tidy the texts of columns: spaces, pieces to drop, case, values to rename, and reading "¥1,299" / "89元" / "10%" as numbers or "2025/3/1" / "2025.03.01" / "20250301" / "3月1日" as dates. Cells it cannot read are left and listed.
  · replace {find, replace, column?, whole?, matchCase?}; fill_blanks {column | columns, with: value | "above" | "mean" | "median" | "mode"}
  · remove_duplicates {columns?: [..]} — rows with the same values in those columns (default: whole rows); the first is kept.
  · delete_rows {blank: true} (entirely empty rows) or {where: {column, is: blank | = | <> | > | < | >= | <= | contains | text | error, value?} | {column, values: [..]}}; also {row, count?}. insert_rows {row, count?}; insert_columns / delete_columns {column: "C", count?}.
  · sort {by: [{column, order?: asc|desc}, ..]}; filter {column, values: [..] | criteria: ">100", criteria2?, join?: and|or} hides rows (it does not delete them); clear_filter {}.
  · copy_range {range, to, values?: true} — e.g. keep the raw sheet and clean a copy: add_sheet {name: "清洗后"} then copy_range {range: "原始!A1:L1205", to: "清洗后!A1", values: true}.
- Analysis: pivot {source?, rows: [..], columns?: [..], values: [{field, fn?: sum|count|average|max|min, name?}], filters?, group?: {field: a date column, by: month|quarter|year}, sort?: {field, by?: a value name, order?}, to?: "汇总!A3", name?} — a pivot table (on a new sheet named by "name" unless "to" is given); the result is returned, no need to read it. For figures that should stay as plain cells use formulas (SUMIFS, COUNTIFS, AVERAGEIFS, XLOOKUP / VLOOKUP, TEXT(date,"yyyy-mm")) with fill_formula.
- Charts: add_chart {type: column | bar | stacked | stacked bar | line | pie | doughnut | scatter | scatter line | area | radar | combo, range (categories in its first column, a header row of series names) | categories: range + series: [{name: text or a cell, values: range, x?, chart?: line (for a combo), axis?: secondary, color?}], title?, xTitle?, yTitle?, labels?: true | "percent", legend?: bottom|right|top|none, numberFormat?, min?, max?, at?: cell, width?, height?, name?}. The result gives the cells the chart covers: place the next one clear of them (a chart of the default size covers about 9 columns x 21 rows). set_chart {chart: name, ..same fields} changes one; delete_chart {chart}. Chart plain cells, not a pivot table (a chart on a pivot table shows all its values). office_render {sheet, range} pictures cells with the charts over them; {sheet, chart: name} one chart.
- Presentation: format_range {range, bold?, italic?, size?, font?, color?, fill?, numberFormat?, align?, wrap?, border?, merge?, columnWidth?, rowHeight?}; autofit {range?}; freeze {cell: "A2"} (put it last in a batch: it is the one step that needs the sheet in front); table {range?, name?, style?} (banded rows and filter arrows); conditional_format {range, rule: > | < | >= | <= | = | <> | between | top | bottom | duplicate | blank | formula | colorscale | databar | clear, value?, value2?, fill?, color?}; validation {range, list: [..]}.
- add_sheet {name}; rename_sheet {sheet, name}; delete_sheet {sheet}
- Comments: set_comment {range, text, author?} (the little red mark in the corner), delete_comment {range}; office_read lists the notes inside the block as {cell, text}, and the sheets line says which sheets carry comments.
- Sheet protection: protect_sheet {password?, allowSelection?, allowFormattingCells?, allowFormattingColumns?, allowFormattingRows?, allowSorting?, allowFiltering?, allowInsertingHyperlinks?, allowInsertingRows?, allowDeletingRows?, userInterfaceOnly?, drawingObjects?} and unprotect_sheet {}. Formatting, sorting and filtering stay allowed by default; rows and columns cannot be inserted or deleted. Cells the author unlocked (format_range lock is not offered: use Python / a cell format when the user needs input cells) stay editable. The sheets line of office_read shows "protected". Protect the sheet LAST in a batch: the protection also blocks what the office_edit calls after it would change.
- Hyperlinks: add_hyperlink {range | cell, address, text?} — address is an http(s) URL, a mailto:, a file path, or "#Sheet2!A1" to jump inside the workbook; with a cell that already had a link it is replaced, and the cell shows the link (text gives the words to show).
- How to do data work well: read first and plan from the profile; keep the raw data (clean a copy on a new sheet); after cleaning read the profile again and check that nothing is left under LOOK; say in your answer what you changed and how many rows each step touched, and what you left alone and why (a far-off value may be real: ask or flag it rather than delete it silently); put results on their own sheets with clear headers, number formats and fitted columns; give every chart a title that states what it shows and axis titles with units; look at the result with office_render before you finish.`,
	ppt: `PowerPoint: build a deck out of DESIGNED SLIDES. One "slide" op makes one finished slide — laid out, coloured, set in type, with its page number, transition and entrance animation — from the content you give; you do not place boxes yourself. A whole deck is one or two office_edit calls.
- theme {name, background?, primary?, accent?, bg?, surface?, text?, muted?, titleFont?, bodyFont?, transition?: fade|push|wipe|split|none, animate?, section?: solid|side|band|number, corners?: soft|square|round|pill, cards?: soft|outline|shadow, marks?: square|dot|dash|triangle, chrome?: plain|beamer, author?, talk?} — set it FIRST, once. corners, cards and marks restyle the small parts every layout is built from, for the whole deck: the corners of cards and panels; the ground of a card (soft = a tinted fill, outline = a thin frame on the page colour, shadow = white with a soft shadow); the mark before a point. Choose them with the theme so decks do not all look alike — square corners with outline cards and dash marks read sober and academic, round corners with shadow cards read friendly, soft is the middle. A single card of a grid or canvas takes style: soft|outline|shadow of its own. chrome: beamer gives the deck the frame of a LaTeX Beamer talk (the Madrid look): every title on a band across the top with the kicker on a strip above it, a foot line with the speaker, the short title of the talk and the page, triangles before the points, theorems as titled blocks, the cover, section and closing pages with the title in a box in the middle. Use it ONLY when the user asks for Beamer, LaTeX or a "mathematical" classic look; it combines with any colour theme (academic or ink suit it; primary recolours it). author and talk are what the foot line says (left and middle); left out, they are taken from the cover's meta and title. section is how the section slides of the deck look (a full ground of colour; a panel at the left with the number; a band across; a huge faint number): leave it out and one is taken by chance, so decks differ. name: ink (white, navy + red; business reports, data), academic (pure white, deep blue with a restrained red, nothing behind the content; thesis defences, course-paper and research reports, conference talks — the default for anything academic), paper (warm white with a paper grain, serif titles; lessons, popular science, essays), sepia (aged paper, dark red + gold, serif; history, literature, the humanities), crimson (warm white, red + gold; Party and government, commemoration, festive), slate (cool grey with a fine grid; engineering, technology), ocean, forest, graphite (light) or night (near-black with gold; pitches, stories), chalk (blackboard green; lessons), plum (dark). Pick by the subject and the audience — a deck on a revolution, a dynasty or a writer is not a quarterly report: do not fall back on ink for everything; change single colours to match a brand. background: what lies behind the content — plain, paper (grain, darker edges), gradient, grid, dots, or the path of a picture laid faintly behind every slide (a period photograph, a texture, a map); each theme has its own default, set it when the subject asks for more atmosphere. An academic or formal report keeps a plain white ground: no paper, no gradient, no picture behind the slides unless the user asks.
- slide {kind, title, subtitle?, kicker?, note?, notes?, image?, callout?, at?, ...} — kicker: the small line above the title (the section, e.g. "二、方法"); note: the source line at the foot; notes: speaker notes; callout: one sentence set apart at the foot of the body; image: path of a picture. In every text **words** are set bold in the accent colour and $...$ is a formula.
  kinds and their own fields:
  · cover {title, subtitle?, kicker?, meta?, image?, style?: "full"} — the opening slide; closing {title, subtitle?, image?, style?} — the last one. A photograph stands at the right behind a fade; style "full" lays it over the whole slide with the words in white on it.
  · section {number?, title, subtitle?, image?, style?} — a divider before each part ("01", "02" ..), in the look the theme fixed.
  · agenda {points:[..]} — the numbered outline.
  · bullets {points:[text | {head, text}], image?} — at most 5 points; prefer {head, text}. With an image the slide is laid out around it: a wide drawing goes across the slide with the points in a row under it, others stand beside the text.
  · cards {cards:[{head, text, icon?, value?}]} — 2 to 6 parallel ideas side by side.
  · stats {stats:[{value, unit?, label, delta?, icon?}], points?} — 2 to 4 big numbers.
  · chart {chart, side?:[{head, text}]} — a chart drawn from your numbers in the deck's colours, with remarks beside it. chart is one of:
      {type: column|bar|line, categories:[..], series:[{name, values:[..]}], unit?, highlight?: a category, decimals?}
      {type: scatter, series:[{name?, points:[[x, y], ..]}], fit?: true (least-squares line with its equation and R²), xTitle?, yTitle?, xMin?, xMax?, yMin?, yMax?} — measurements over two number axes
      {type: curve, series:[{name?, x:[..], y:[..], dashed?, color?}], xTitle?, yTitle?, endValues?: true, lines?:[{y | x, text?}], annotations?:[{text, x, y}], yMin?, yMax?} — graphs of functions, training curves, simulations: give 20–100 points per curve (compute them; a short script that prints the numbers is fine). Each line is named at its end, the names kept apart; lines marks a threshold or a moment, annotations puts words at a point.
      {type: pie | donut, categories:[..], values:[..], unit?} — shares of a whole, at most 8 parts
  · table {data:[[header..],[row..]..]} — up to 8 rows; cells may hold $formulas$.
  · formula {formulas:[{latex, label?}], points?:[text | {head, text}], layout?: focus|side|stack} — 1 to 3 formulas with what they mean. Three arrangements: focus (ONE formula set large on a band, its remarks in columns underneath — the key result of a talk), side (formulas at the left, a remark beside each — give as many points as formulas and they pair up), stack (formulas over a list of remarks). Left out, one formula is focus, two with remarks are side, three stack. Do not make every formula slide the same: a derivation reads well as a grid with formula parts and arrows of text between them, a formula with a figure as a grid of [formula, image], a definition as theorem.
  · theorem {label?: "定理 2.1", name?, statement, proof?:[text | {head, text}], proofTitle?} — a definition, theorem, lemma or law set apart on a card, with the steps that lead to it (or what it means) underneath.
  · gallery {images:[{image, caption?, focus?, fit?}] (2–6), text?} — pictures side by side with captions: the people of a story, the panels of an experiment, results to compare. Upright pictures (portraits) stand in one row; a last row that is not full stands in the middle.
  · Photographs are cut to fill their place (split, gallery, cover, image, a photo beside bullets); of a tall picture the upper part is kept, where heads are. "focus": top | center | bottom | left | right on the slide (or on a gallery item) moves the cut; "fit": "contain" on a gallery item shows the picture whole. The result names the pictures that lost much: look at those slides.
  · summary {points:[{head, text}] (2–4)} — what to take away, numbered large; the first stands out.
  · code {code, language?, points?:[{head, text}]} — a listing on a dark panel with line numbers and coloured keywords, remarks beside it; 8–16 lines read best.
  · split {title, text?, points?, image, side?: left|right, kicker?} — half the slide is the picture, edge to edge; the words stand on the other half. A change of rhythm among content slides.
  · grid {title, rows:[[part, part, ..], [part], ..], edges?, callout?} — COMPOSE YOUR OWN SLIDE without coordinates: each row is a list of parts that stand side by side, the rows stand under one another, and the places, gaps and sizes are worked out for you so the body is filled evenly. A part is {type, span?: relative width (default 1), ..} with the types of a canvas item below (no box): text, heading, card, stat, chart, image, formula, table, bullets, steps, callout, quote, panel, node, icon, line, or space (an empty cell). {type: "stack", span?, items:[part, ..]} puts parts under one another inside one cell — a chart at the left with a heading, two stats and a remark stacked at its right; a picture beside numbered steps; a table over a quote. Headings, stats, formulas, callouts and short texts take the height they need and charts, pictures, tables, cards and stacks share the rest ("height": points, or "weight" on a row {cells:[..], weight} to change that). Nodes placed in a grid are joined with edges:[{from, to, kind?, label?}] at slide level. Use grid whenever the fixed kinds would repeat themselves or leave half a slide empty: it is how a deck gets pages that are not all rows of blocks.
  · canvas {title?, items:[{type, box:[x, y, w, h], ..}]} — exact placement, for a drawing no grid can hold: place items on the 960 x 540 canvas (the title takes the top 118; keep to x 48–912, y 130–494). Types: text {text, size?, bold?, color?, align?, valign?}, heading {text}, card {head, text, icon?, value?}, panel {fill?, text?, shape?: round|circle}, line / arrow {box: start x, y, then width and height as the run}, image {image, fit?: cover|contain}, formula {latex, label?}, stat {value, unit?, label}, chart {chart}, table {data}, icon {icon}, bullets {points}, steps {steps:[{head, text}]} (numbered, one under another), callout {text}, quote {text, by?}. Colours by role: primary, accent, text, muted, surface, line, or #RRGGBB. For block diagrams (a network architecture, a pipeline, a system) give the boxes an id — node {id, box, text, fill?} — and join them with edge {from, to, kind?: elbow|straight|curve, label?, dashed?, fromSide?, toSide?: top|bottom|left|right} (among the items, or all together in edges:[..] beside items): the connector joins the sides that face each other and carries an arrow; kind curve is a jump over what lies between (a skip connection, a feedback path) and arches above a row of boxes. Keep boxes of a diagram small (about 150 x 70) with 50 or more between them so the arrows show, and put the same top and height on boxes of one row. Each item comes in as its own animation step ("with": true joins it to the one before). Keep a grid: equal gaps, aligned edges; look at the result.
  · timeline {steps:[{when, head, text}]}; process {steps:[{head, text}]} — up to 5; compare {left:{head, points:[..]}, right:{head, points:[..]}}.
  · image {image, title, text?, caption?} — a photograph fills the slide with the words on it; a chart or diagram is shown whole, as large as the slide allows, under the title (a squarish one at the left with the text beside it). Figures drawn for a paper are not slide material as they are: a picture with several panels shows each one small — redraw what the deck can draw (chart, canvas) and give the rest one panel per slide. quote {text, by?}.
  icon: check clock search star chart list people target bolt flag globe lock book calendar document layers trend link code heart warning home idea pin money arrow cloud mail shield eye pie wave question info grid edit plus play.
- How to make a deck people want to look at:
  · Plan it as a talk: cover, agenda, then for each part a section slide and 2–4 content slides, a closing. 10–16 slides for a normal report.
  · Give every slide ONE message, and write the title as that message, a full statement ("客运量超过 2019 年高点"), not a topic ("客运情况").
  · Change the kind from slide to slide: never two bullets slides in a row, and no more than a third of the deck as bullets. Numbers go on stats or chart, parallel ideas on cards, a sequence on process or timeline, two sides on compare, a formula on formula, a definition or theorem on theorem, several figures on gallery, the conclusion on summary; vary how pictures stand (beside the text, split, image, a full cover).
  · A research talk or thesis defence: cover; background and the question (split or bullets); contributions (cards or summary); method (process, theorem, formula, or a canvas for the pipeline diagram); experiments (table for the setup, chart / scatter / curve drawn from the numbers, gallery for result figures); discussion (compare); conclusion (summary); references (bullets, small); thanks (closing). A course lecture: definitions and theorems on theorem slides, worked examples on process, intuition on image or curve.
  · A layout you built once (a canvas, or any slide) is reused with reuse_slide {from: n, texts: {..}}: the copy keeps the arrangement and takes new words.
  · Keep text short: a point is one line, a card text two. What you would say aloud goes into notes.
  · Fill the page and vary it. A slide whose content ends half way down looks unfinished: when a kind would hold only a little (a table of three rows, two remarks, three short cards), add what belongs with it — the figure beside the table, the conclusion as a callout, a stat — or build the page as a grid. In a deck of 15 slides at least three or four should be grids or canvases of your own, not the same card rows again.
  · Use pictures: a photograph on the cover and on an image slide or two lifts a deck more than anything else. Get them with your web / image tools (search and download, or generate) into the working folder, then pass the path; a chart or diagram you draw yourself is shown whole on a card.
  · DRAW IN THE DECK, NOT IN MATPLOTLIB. What the deck can draw from shapes looks far better on a slide than a pasted picture, stays sharp and editable, and takes the theme: bar / column / line charts, scatter plots with a fit, curves of functions and training runs (type curve, with the lines named at their ends), pie and ring charts, tables, formulas, code listings (kind code), block diagrams and pipelines (canvas with node and edge), timelines and processes. Make a picture file only for what none of these can show: a heat map, a 3-D surface, a photograph, an image from a paper, a plot with thousands of points.
  · Numbers you have belong on a chart slide (drawn in the deck's own colours and type), not on a picture of a chart. A drawing you do make yourself (a diagram, a plot of a function) must suit a slide: ONE idea per drawing, not several panels side by side; drawn for the place it goes (about 9 x 3.5 in across a slide, 5 x 4 in beside text) with type of 14 pt or more at that size, few words, thick lines, the deck's colours; nothing overlapping (check the picture itself before you use it). A drawing that matters gets a slide of its own (kind image). If the result says a figure is shown much smaller than it was drawn, act on it.
  · If the user has not said what the deck should look like (no template, no style, no audience to go by) and you have a tool for asking the user a question, ask once before you build: offer two to four of the themes with a word on each, and whether they have a template or example to follow. If they said to decide yourself, or there is no such tool, choose and say what you chose.
  · Look at the result with office_render {overview: true} and fix slides whose text is cut off or crowded (shorten the words, or split the slide).
- Working from a TEMPLATE or an EXAMPLE the user gives (a .pptx whose look the deck should have): do not rebuild its look by hand, work inside a copy of it.
  1. office_open {path: the new file, template: the user's file}, then office_render {overview: true} (again with slide numbers for the rest of a long template) to see every page, and office_read {slide: n} for the pages you will use: it lists each text and picture with its "#id".
  2. For each slide of your deck pick the template page that fits what it says (cover, contents, section, text + picture, three points, comparison, timeline, thanks ..) and make it with reuse_slide {from: n, texts: {"#id" or name: new text, ..}, images?: {"#id": path}, delete?: ["#id", ..], notes?} — a copy of page n with your words in place of its sample text, its look and animation kept. Replace EVERY sample text of the page (Lorem ipsum, "SAMPLE TITLE", xxxx) or delete the shape; keep your text about as long as the sample it replaces. New slides are added after the template's pages, so the numbers of the template pages stay valid.
  3. Where no page fits (a chart, big figures, a formula, a table): slide {kind, .., base: n, area: [left, top, width, height]} draws that kind on a copy of template page n (a page with little on it: a content page with its sample deleted, or a plain background), inside the free area you give in points; set theme {primary, accent, text, bg, titleFont..} first to the template's colours (read them off the render) so the two match. theme {base, area} makes this the default for every slide op — the way to use a template that is only background pictures.
  4. Finish with delete_slides {from: 1, to: last template page} so only your slides remain, render the overview and check that no sample text is left.
The operations below are for touching up a designed slide or for the rare slide no kind fits.

PowerPoint ops. slide = slide number; shape = a shape name or "#id" from office_read, or "title" / "body". Positions and sizes are in points. Formulas: write LaTeX between dollar signs in any slide text (title, body, text box), as in Word; they become native PowerPoint equations, inline with the text. A formula on its own gets a text box of its own.
- add_slide {layout?: title|title_content|two_content|title_only|blank|section, title?, body?, notes?, at?} — body: lines separated by newlines become bullets.
- set_text {slide, shape, text, size?, bold?, color?, align?} — a line that starts with a tab (or two spaces) is a bullet one level down.
- format_text {slide, shape, find?, font?, size?, bold?, italic?, underline?, color?, align?, bullets?, lineSpacing?, fit?: "shrink"|"grow"|"none"} — the whole text of the shape, or only the part that reads find.
- add_textbox {slide, text, left, top, width, height, size?, bold?, color?, align?, fill?, name?}
- add_image {slide, path, left, top, width?, height?, name?}
- add_smartart {slide, layout, items, left?, top?, width?, height?, colors?, look?, name?} — a native SmartArt diagram. layout: list | bullet_list | horizontal_list | process | chevron | arrows | timeline | steps | cycle | radial | hierarchy | org_chart | tree | pyramid | venn | matrix | funnel | target | balance | gear, or the name a layout has in PowerPoint. items: texts, or {text, children:[..]} for the levels below (the branches of a hierarchy, the bullets under a step). Prefer it to hand-drawn boxes and arrows for steps, cycles, structures and comparisons.
- add_table {slide, data:[[cell,..],..], left?, top?, width?, height?, size?, font?, align?, header?, name?} — a native table; formulas in its cells are set as text with real subscripts and powers.
- add_shape {slide, kind, left, top, width, height, text?, fill?, line?: colour|"none", lineWidth?, size?, bold?, color?, name?} — kind: rectangle | rounded | ellipse | diamond | triangle | arrow | arrow_left | arrow_up | arrow_down | chevron | pentagon | hexagon | star | callout | cloud | line | arrow_line (for a line, width and height are how far it runs).
- set_shape {slide, shape, left?, top?, width?, height?, fill?, name?}; delete_shape {slide, shape}
- reuse_slide {from, texts?, images?, delete?, at?, notes?}; delete_slides {from, to}; delete_slide {slide}; move_slide {slide, to} (to = the number the slide has afterwards; page numbers follow); duplicate_slide {slide, to?}; set_layout {slide, layout}; set_background {slide, color}; set_notes {slide, text}`
};
function createTools(host) {
	const { helper } = host;
	const tools = [];
	let current;
	const own = /* @__PURE__ */ new Set();
	const unseen = /* @__PURE__ */ new Set();
	const keyOf = (on) => `${on.app}:${fileName(on.doc).toLowerCase()}`;
	const told = /* @__PURE__ */ new Map();
	const sessionOf = (exec) => exec?.agent?.session?.id ?? "default";
	/** The operation list of an app, the first time a conversation works with it ('' afterwards). */
	const guide = (exec, kind) => {
		const session = sessionOf(exec);
		let seen = told.get(session);
		if (!seen) {
			seen = /* @__PURE__ */ new Set();
			told.set(session, seen);
		}
		if (seen.has(kind)) return "";
		seen.add(kind);
		return `\n\n${APP_NAMES[kind]} operations for office_edit (this list is given once; office_help {app: "${kind}"} gives it again):\n${GUIDES[kind]}`;
	};
	/** Pictures of the document as it is now: Word pages or PowerPoint slides first..last, or an Excel range. */
	const picture = async (on, first, last, extra, width) => {
		const folder = join(tmpdir(), "dsh-office");
		await mkdir(folder, { recursive: true });
		const lines = [];
		const images = [];
		let total;
		for (let number = first; number <= last; number++) {
			const out = join(folder, `render-${process.pid}-${Date.now()}-${number}.png`);
			try {
				const which = on.app === "word" ? { page: number } : on.app === "ppt" ? { slide: number } : {};
				const result = await helper.call("render", {
					...on,
					...extra,
					...which,
					out,
					width
				}, 5e4);
				images.push(toJson(await host.saveImage(await readFile(out), `${on.app}-${result.what.replace(/[^\w一-龥]+/g, "-")}.png`)));
				lines.push(`${APP_NAMES[on.app]} ${result.what}, ${result.width}×${result.height}.`);
				const of = / of (\d+)$/.exec(result.what);
				if (of) {
					total = Number(of[1]);
					last = Math.min(last, total);
				}
			} catch (error) {
				if (images.length === 0) throw error;
				if (on.app === "word") lines.push(error instanceof Error ? error.message : String(error));
				break;
			} finally {
				await rm(out, { force: true });
			}
		}
		return {
			lines,
			images,
			...total === void 0 ? {} : { total }
		};
	};
	/** The whole document at a glance: its pages (6 to a picture) or slides tiled with their numbers, up to 18. */
	const overview = async (on) => {
		if (on.app === "excel") return picture(on, 0, 0, {}, Math.min(host.settings().renderWidth, 1e3));
		const folder = join(tmpdir(), "dsh-office");
		await mkdir(folder, { recursive: true });
		const lines = [];
		const images = [];
		let total = 1;
		for (let from = 1; from <= Math.min(total, 18); from += 6) {
			const out = join(folder, `sheet-${process.pid}-${Date.now()}-${from}.png`);
			try {
				const result = await helper.call("render", {
					...on,
					sheet: true,
					from,
					to: from + 5,
					out,
					width: 640
				}, 11e4);
				total = result.total;
				images.push(toJson(await host.saveImage(await readFile(out), `${on.app}-overview-${from}.png`)));
				lines.push(`${APP_NAMES[on.app]} ${result.what}, tiled in one picture.`);
			} finally {
				await rm(out, { force: true });
			}
		}
		if (total > 18) lines.push(`Only the first 18 of ${total} are shown.`);
		return {
			lines,
			images
		};
	};
	/** The helper arguments naming the document; the app comes from doc / app only, never from a save-as path. */
	const target = (args) => {
		if (args.doc === void 0 && args.app === void 0) {
			if (!current) throw new Error("No document yet: call office_open first, or pass doc.");
			return current;
		}
		current = {
			app: appOf({
				...args.app === void 0 ? {} : { app: args.app },
				...args.doc === void 0 ? {} : { doc: args.doc }
			}),
			...args.doc === void 0 ? {} : { doc: args.doc }
		};
		return current;
	};
	const quietTold = /* @__PURE__ */ new Set();
	const quietNote = (exec) => {
		const session = sessionOf(exec);
		if (quietTold.has(session)) return "";
		quietTold.add(session);
		return "\n\n" + SILENT;
	};
	tools.push(defineTool({
		name: "office_open",
		description: "Open a .docx / .xlsx / .pptx file in Word, Excel or PowerPoint so the user watches the work and can edit alongside; a path that does not exist yet creates the file. If the file is already open it is used as it is. Call this before the other office_ tools.",
		parameters: {
			path: {
				type: "string",
				description: "Absolute path of the file. Omit to start an unsaved new document (then app is required)."
			},
			template: {
				type: "string",
				description: "Absolute path of a template or an example to start from: it is copied to path (which must not exist yet) and the copy is opened, the original left untouched. Use it when the user gives a template or a sample whose look the new document should have."
			},
			app: {
				type: "string",
				enum: [
					"word",
					"excel",
					"ppt"
				],
				description: "Needed only without path."
			},
			show: {
				type: "boolean",
				description: "true brings the document's window into view even in silent mode. Needed only before working on the window itself with mouse / keyboard tools."
			}
		},
		output,
		timeoutMs: 9e4,
		async execute(args, exec) {
			const { path, app, template, show } = args;
			if (path !== void 0 && !isAbsolute(path)) throw new Error("path must be an absolute path.");
			if (template !== void 0) {
				if (path === void 0) throw new Error("With template, give path: where the copy is to be saved.");
				if (!isAbsolute(template)) throw new Error("template must be an absolute path.");
				if (extname(template).toLowerCase() !== extname(path).toLowerCase()) throw new Error("template and path must be the same kind of file.");
				if (await access(path).then(() => true, () => false)) throw new Error(`${path} exists already: pick a path that does not, the template is copied there.`);
				await copyFile(template, path).catch((error) => {
					throw new Error(`Could not copy the template: ${error instanceof Error ? error.message : String(error)}`);
				});
			}
			const kind = appOf({
				...app === void 0 ? {} : { app },
				...path === void 0 ? {} : { path }
			});
			const { showOnOpen, follow, typing, card, silent } = host.settings();
			const quiet = silent && show !== true;
			const doc = await helper.call("open", quiet ? {
				app: kind,
				...path === void 0 ? {} : { path },
				silent: true,
				show: false
			} : {
				app: kind,
				...path === void 0 ? {} : { path },
				show: show === true ? true : showOnOpen,
				follow,
				typing,
				card: silent ? false : card
			}, 8e4);
			current = {
				app: kind,
				doc: doc.path ?? doc.name
			};
			if (doc.how !== "attached") own.add(keyOf(current));
			return { text: formatOpened(doc) + (quiet ? quietNote(exec) : "") + guide(exec, kind) };
		},
		presentCall: (args) => card(`打开 ${fileName(args.path) || "新文档"}`)
	}));
	tools.push(defineTool({
		name: "office_status",
		description: "List what is open in Word, Excel and PowerPoint right now, and which documents have unsaved changes.",
		parameters: {},
		output,
		isConcurrencySafe: () => true,
		async execute() {
			return { text: formatStatus(await helper.call("status", {}, 3e4)) };
		},
		presentCall: () => card("查看 Office 状态")
	}));
	tools.push(defineTool({
		name: "office_read",
		description: "Read an open document as text, exactly as it is now (including what the user just changed). Word: numbered paragraphs with heading levels and tables; table=N reads one table. Excel: the sheets, then the cells of a range (default: the used range of the active sheet) with their formulas. PowerPoint: every slide with its shapes (name, type, [left, top, width, height] in points, text).",
		parameters: {
			doc: DOC,
			app: APP,
			from: {
				type: "integer",
				description: "Word: first paragraph to list (default 1)."
			},
			count: {
				type: "integer",
				description: "Word: how many paragraphs (default 120, max 400)."
			},
			full: {
				type: "boolean",
				description: "Word / PowerPoint: full text instead of the first 100 / 200 characters of each paragraph / shape."
			},
			footnotes: {
				type: "boolean",
				description: "Word: include every native footnote, independent of the body paragraph range. Default: only notes anchored in the paragraphs read. Use full:true for complete note text."
			},
			table: {
				type: "integer",
				description: "Word: read the cells of table N."
			},
			sheet: {
				type: "string",
				description: "Excel: sheet name."
			},
			range: {
				type: "string",
				description: "Excel: A1-style range, e.g. A1:F40."
			},
			profile: {
				type: "boolean",
				description: "Excel: also describe every column of the block (kinds of values, range, distinct values, what looks wrong). Given by itself for a table too long to show whole."
			},
			slide: {
				type: "integer",
				description: "PowerPoint: only this slide."
			}
		},
		output,
		timeoutMs: 6e4,
		isConcurrencySafe: () => true,
		async execute(args, exec) {
			const input = args;
			const { doc: _doc, app: _app, ...rest } = input;
			const on = target(input), kind = on.app;
			const result = await helper.call("read", {
				...on,
				...rest
			}, 5e4);
			return { text: (kind === "word" ? formatWord(result) : kind === "excel" ? formatExcel(result) : formatPpt(result)) + guide(exec, kind) };
		},
		presentCall: (args) => card(`读取 ${fileName(args.doc) || "当前文档"}`)
	}));
	tools.push(defineTool({
		name: "office_edit",
		description: EDIT_CORE,
		parameters: {
			doc: DOC,
			app: APP,
			ops: {
				type: "array",
				required: true,
				description: "Operations to run in order, e.g. [{\"op\":\"set_text\",\"para\":3,\"expect\":\"Old start\",\"text\":\"New text\"}].",
				items: {
					type: "object",
					additionalProperties: true,
					properties: { op: {
						type: "string",
						required: true,
						description: "Operation name."
					} }
				}
			}
		},
		output,
		timeoutMs: 3e5,
		async execute(args, exec) {
			const input = args;
			if (!Array.isArray(input.ops) || input.ops.length === 0) throw new Error("ops must be a non-empty array.");
			if (input.ops.length > 200) throw new Error("At most 200 operations per call.");
			const { follow, typing, card, silent } = host.settings();
			const on = target(input);
			const math = formulasOf(input.ops);
			const result = await helper.call("edit", {
				...on,
				ops: input.ops,
				follow,
				typing,
				card: silent ? false : card,
				silent,
				...Object.keys(math).length > 0 ? { math } : {}
			}, 27e4);
			if (result.done.length > 0) unseen.add(keyOf(on));
			return { text: formatEdit(result) + guide(exec, on.app) };
		},
		presentCall: (args) => {
			const { doc, ops } = args;
			return card(`编辑 ${fileName(doc) || "当前文档"}${Array.isArray(ops) ? ` · ${ops.length} 项` : ""}`);
		}
	}));
	tools.push(defineTool({
		name: "office_help",
		description: "The operations office_edit takes for one app, with their fields: Word (text, formulas, tables, pictures, citations, styles, page setup), Excel (cells, formats, sheets, charts) or PowerPoint (slides, text, pictures, SmartArt, tables, shapes). Call it whenever you are not sure what an operation is called or what fields it takes.",
		parameters: { app: {
			type: "string",
			enum: [
				"word",
				"excel",
				"ppt"
			],
			required: true,
			description: "Which app."
		} },
		output,
		isConcurrencySafe: () => true,
		async execute(args, exec) {
			const kind = appOf({ app: args.app ?? "" });
			const session = sessionOf(exec);
			if (!told.has(session)) told.set(session, /* @__PURE__ */ new Set());
			told.get(session).add(kind);
			return { text: `${APP_NAMES[kind]} operations for office_edit:\n${GUIDES[kind]}\n\n${TAIL}` };
		},
		presentCall: (args) => card(`查看 ${APP_NAMES[appOf({ app: args.app ?? "word" })]} 操作说明`)
	}));
	tools.push(defineTool({
		name: "office_render",
		description: "Get a picture of how the open document really looks: a Word page, a PowerPoint slide (with \"to\": up to 4 in a row), an Excel range (default: the used range), or with overview: true the whole document tiled into one picture. Use it on every page you changed before you finish: it shows what office_read cannot, such as content that landed outside its frame, a table that is too wide, text that overflows. It pictures the document as Word itself lays it out, unsaved changes included, with the window in the background.",
		parameters: {
			doc: DOC,
			app: APP,
			page: {
				type: "integer",
				description: "Word: page number (default 1)."
			},
			slide: {
				type: "integer",
				description: "PowerPoint: slide number."
			},
			to: {
				type: "integer",
				description: "Word / PowerPoint: last page or slide, to get up to 4 in one call."
			},
			overview: {
				type: "boolean",
				description: "Word / PowerPoint: every page or slide tiled into one picture (6 pages per picture, numbered): the whole document at a glance, for the final check. Then look closer at a page with page / slide."
			},
			sheet: {
				type: "string",
				description: "Excel: sheet name."
			},
			range: {
				type: "string",
				description: "Excel: range to picture (charts lying over it are in the picture)."
			},
			chart: {
				type: "string",
				description: "Excel: name of one chart to picture, instead of a range."
			}
		},
		output,
		timeoutMs: 6e4,
		async execute(args, exec) {
			const input = { ...args };
			const on = target(input), kind = on.app;
			if (kind === "ppt" && input.slide === void 0 && typeof input.page === "number") input.slide = input.page;
			if (kind === "ppt" && input.slide === void 0 && input.overview !== true) throw new Error("slide is required for PowerPoint (or pass overview: true).");
			if (!await host.vision(exec)) return { text: "The current model cannot view images; check the result with office_read instead." };
			if (input.overview === true && kind !== "excel") {
				const all = await overview(on);
				unseen.delete(keyOf(on));
				const [image, ...more] = all.images;
				return {
					text: all.lines.join("\n"),
					...image === void 0 ? {} : { image },
					...more.length ? { more } : {}
				};
			}
			const { doc: _doc, app: _app, to, page: _page, slide: _slide, overview: _overview, ...rest } = input;
			const first = kind === "word" ? input.page ?? 1 : kind === "ppt" ? input.slide : 0;
			const last = kind === "excel" || to === void 0 ? first : Math.min(Math.max(to, first), first + 3);
			const { lines, images } = await picture(on, first, last, rest, host.settings().renderWidth);
			unseen.delete(keyOf(on));
			const [image, ...more] = images;
			return {
				text: lines.join("\n"),
				...image === void 0 ? {} : { image },
				...more.length ? { more } : {}
			};
		},
		presentCall: (args) => card(`查看 ${fileName(args.doc) || "当前文档"} 的排版`)
	}));
	tools.push(defineTool({
		name: "office_save",
		description: "Save the open document; if you have edited since you last looked at it, the reply also shows you its pages for a final check. With path: save a copy under another name or format (.pdf exports a PDF and keeps the document as it is; other extensions switch the document to the new file). The file as it stood on disk before this save is kept as a version in \".dsh-versions\" beside it (the last 50 per document), so any save can be undone with office_revert.",
		parameters: {
			doc: DOC,
			app: APP,
			path: {
				type: "string",
				description: "Absolute path to save as. Omit to save in place."
			},
			no_version: {
				type: "boolean",
				description: "true: do not keep the previous content as a version (for a routine save that need not be undoable)."
			},
			version_note: {
				type: "string",
				description: "A short word put in the version's file name saying why it was kept (\"before-rewrite\", \"draft2\"); defaults to \"save\"."
			}
		},
		output,
		timeoutMs: 12e4,
		async execute(args, exec) {
			const input = args;
			if (input.path !== void 0 && !isAbsolute(input.path)) throw new Error("path must be an absolute path.");
			const on = target(input);
			let tidied = "";
			if (on.app === "word" && unseen.has(keyOf(on)) && extname(input.path ?? "").toLowerCase() !== ".pdf") try {
				const said = (await helper.call("edit", {
					...on,
					ops: [{ op: "fill_gaps" }]
				}, 1e5)).done?.[0] ?? "";
				if (said !== "" && !/^no page/.test(said)) tidied = ` Before saving, ${said}.`;
			} catch {}
			const result = await helper.call("save", {
				...on,
				...input.path === void 0 ? {} : { path: input.path },
				...input.no_version === void 0 ? {} : { no_version: input.no_version },
				...input.version_note === void 0 ? {} : { version_note: input.version_note }
			}, 11e4);
			const moved = input.path !== void 0 && extname(input.path).toLowerCase() !== ".pdf";
			const wasUnseen = unseen.delete(keyOf(on));
			if (moved) {
				if (own.delete(keyOf(on))) own.add(keyOf({
					app: on.app,
					doc: result.path
				}));
				current = {
					app: on.app,
					doc: result.path
				};
			}
			const kept = result.version ? ` Previous content kept as a version (${fileName(result.version)}).` : "";
			const text = `Saved ${result.path} (${result.bytes} bytes).${moved ? " The open document is now this file." : ""}${kept}${tidied}`;
			if (!wasUnseen || !host.settings().finalCheck || !await host.vision(exec)) return { text };
			try {
				const now = moved ? {
					app: on.app,
					doc: result.path
				} : on;
				const shown = await overview(now);
				const [image, ...more] = shown.images;
				return {
					text: `${text}\nFinal check: you edited after you last looked at the document, so here is the whole of it as it is now. ${shown.lines.join(" ")}\nGo through every page before you tell the user it is done. If anything is off (content outside its frame, uneven fonts or spacing, formulas as plain text, bad tables, leftover hints, blank gaps, a nearly empty last page), look closer with office_render, fix it with office_edit and save again; if all is well, say so.`,
					...image === void 0 ? {} : { image },
					...more.length ? { more } : {}
				};
			} catch {
				return { text };
			}
		},
		presentCall: (args) => card(`保存 ${fileName(args.path ?? args.doc) || "当前文档"}`)
	}));
	tools.push(defineTool({
		name: "office_versions",
		description: "List the versions kept for a document: every save keeps the content as it stood before it in a hidden \".dsh-versions\" folder beside the file (the last 50). Use it to find the version to go back to with office_revert, or to tell the user what history exists.",
		parameters: {
			doc: DOC,
			app: APP,
			path: {
				type: "string",
				description: "Absolute path of the document whose versions to list (it need not be open)."
			}
		},
		output,
		timeoutMs: 3e4,
		async execute(args) {
			const input = args;
			let path = input.path;
			if (path === void 0) {
				if (input.doc === void 0) throw new Error("Pass path, or doc for a document that is open.");
				const info = (await helper.call("status", {}, 3e4)).flatMap((app) => app.documents).find((doc) => fileName(doc.path ?? doc.name).toLowerCase() === fileName(input.doc).toLowerCase());
				path = info?.path;
				if (path === void 0 || path === "") throw new Error(`The document as ${input.doc} has no file on disk yet, so it has no versions.`);
			}
			if (!isAbsolute(path)) throw new Error("path must be an absolute path.");
			const listed = await helper.call("versions", { path }, 3e4);
			if (listed.count === 0) return { text: `No versions of ${listed.path} yet: one is kept on every save. Versions go in ${listed.dir}.` };
			const lines = listed.items.map((item, n) => `${n + 1}. ${item.version} (${item.time}, ${item.bytes} bytes)${item.reason ? ` — ${item.reason}` : ""}`);
			return { text: `${listed.count} version(s) of ${listed.path}, newest first (the file as it stood before that save); ${listed.keep} are kept:\n${lines.join("\n")}\nGo back with office_revert (version: the stamp shown, or \"latest\").` };
		},
		presentCall: (args) => card(`版本 ${fileName(args.path ?? args.doc) || "文档"}`)
	}));
	tools.push(defineTool({
		name: "office_revert",
		description: "Put a document back to one of its versions: the version's content replaces the file, the open document is closed without saving and opened again, so the user sees the restored content and the next read returns it. By default the content being replaced is kept as a version first (keep: false to skip). Ask the user before going back if the current content matters.",
		parameters: {
			doc: DOC,
			app: APP,
			path: {
				type: "string",
				description: "Absolute path of the document to go back. Omit for the open one (pass doc)."
			},
			version: {
				type: "string",
				description: "The version stamp from office_versions (its start is enough), or \"latest\" for the most recent, \"oldest\" for the first."
			},
			keep: {
				type: "boolean",
				description: "false: do not keep the content being replaced as a version. Default true."
			}
		},
		output,
		timeoutMs: 12e4,
		async execute(args) {
			const input = args;
			const on = target(input);
			let docPath = input.path;
			if (docPath === void 0) {
				const info = (await helper.call("status", {}, 3e4)).flatMap((app) => app.documents).find((doc) => doc.app === on.app && fileName(doc.path ?? doc.name).toLowerCase() === fileName(on.doc).toLowerCase());
				docPath = info?.path;
				if (docPath === void 0 || docPath === "") throw new Error(`The document as ${on.doc} has no file on disk yet, so it has no versions.`);
			}
			if (!isAbsolute(docPath)) throw new Error("path must be an absolute path.");
			const result = await helper.call("revert", {
				app: on.app,
				path: docPath,
				...input.version === void 0 ? {} : { version: input.version },
				...input.keep === void 0 ? {} : { keep: input.keep }
			}, 11e4);
			unseen.delete(keyOf({ app: on.app, doc: fileName(docPath) }));
			current = {
				app: on.app,
				doc: docPath
			};
			const text = `Went back: ${docPath} now holds the version ${result.version}${input.keep === false ? "" : `; the content it replaced was kept as a version too`}. The document was closed and opened again so you see the restored content; read it with office_read before changing it again.`;
			if (!await host.vision) return { text };
			try {
				const shown = await overview({ app: on.app, doc: docPath });
				const [image, ...more] = shown.images;
				return {
					text,
					...image === void 0 ? {} : { image },
					...more.length ? { more } : {}
				};
			} catch {
				return { text };
			}
		},
		presentCall: (args) => card(`回退 ${fileName(args.path ?? args.doc) || "文档"}`)
	}));
	tools.push(defineTool({
		name: "office_close",
		description: "Close a document that YOU opened or created with office_open (a scratch file, a source you are done with). Documents the user already had open are theirs and are refused. Unsaved changes are refused unless you pass save.",
		parameters: {
			doc: DOC,
			app: APP,
			save: {
				type: "boolean",
				description: "true: save and close. false: close and drop unsaved changes."
			}
		},
		output,
		timeoutMs: 6e4,
		async execute(args) {
			const input = args;
			const on = target(input);
			if (!own.has(keyOf(on))) throw new Error("This document was open before you came to it, or was not opened by you: leave closing it to the user.");
			if (input.save === void 0) {
				const info = (await helper.call("status", {}, 3e4)).flatMap((app) => app.documents).find((doc) => doc.app === on.app && fileName(doc.path ?? doc.name).toLowerCase() === fileName(on.doc).toLowerCase());
				if (info && !info.saved) throw new Error("It has unsaved changes: pass save: true to keep them or save: false to drop them.");
			}
			await helper.call("close", {
				...on,
				save: input.save === true
			}, 5e4);
			own.delete(keyOf(on));
			unseen.delete(keyOf(on));
			if (current && keyOf(current) === keyOf(on)) current = void 0;
			return { text: `Closed ${fileName(on.doc)}.` };
		},
		presentCall: (args) => card(`关闭 ${fileName(args.doc) || "文档"}`)
	}));
	return tools;
}
//#endregion
//#region lib/types/index.js
const name = "@andrewveraside/dsh-wps-sidebar";
const inject = ["tools", "attachments"];
const STATUS_ROUTE = "/api/dsh-office/status";
const Config = z.object({
	suite: z.string().default("auto").volatile().i18n({
		"zh-CN": { $description: "办公套件：auto（自动：装了微软 Office 就用它，否则用 WPS）| office（微软 Office）| wps（WPS Office）。已经打开着的文档始终在它所在的软件里修改" },
		"en-US": { $description: "Office suite: auto (Microsoft Office when installed, else WPS) | office | wps. A document that is already open is always edited where it is" }
	}),
	follow: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "跟随 AI 正在修改的位置（自动滚动到那里）；关闭为静默执行，不动你的视图" },
		"en-US": { $description: "Scroll to where the AI is editing; off = silent, your view is left alone" }
	}),
	silent: z.boolean().default(false).volatile().i18n({
		"zh-CN": { $description: "静默模式：在后台完成，不弹出任何窗口和迷你卡片，一次写入（不逐字）；回合结束时自动保存并关闭后台打开的文档。开启后下面的“跟随”“逐字”“迷你卡片”不再起作用" },
		"en-US": { $description: "Silent mode: work in the background with no window and no card, text written at once; what was opened is saved and closed when the turn ends. Follow, typing and the card do not apply while it is on" }
	}),
	typing: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "文字像打字一样逐步写出，而不是整块出现（每段约多花 0.3 秒）" },
		"en-US": { $description: "Write text progressively like typing instead of all at once (about 0.3 s more per paragraph)" }
	}),
	card: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "AI 编辑时在屏幕右下角显示迷你卡片，可随时开关“查看”和“快速”（快速 = 不逐字写入）" },
		"en-US": { $description: "Show a small card at the bottom right while the AI edits, with switches for following and typing" }
	}),
	finalCheck: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "交付前整体视觉检查：保存时如果 AI 改完后还没看过成品，自动把各页渲染成图片给它检查" },
		"en-US": { $description: "Whole-document visual check: on save, show the AI every page if it has not looked since its last edits" }
	}),
	preferLive: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "优先在真实 Office 里操作：AI 想加载 DSH 自带的 office-docx / office-pptx 技能（用脚本改文件）时，先把它引回本插件；表格的 office-xlsx 技能照常加载" },
		"en-US": { $description: "Prefer the real Office app: when the AI reaches for the host's file-based office-docx / office-pptx skills, point it back to this plugin first; office-xlsx loads as usual" }
	}),
	showOnOpen: z.boolean().default(true).volatile().i18n({
		"zh-CN": { $description: "打开文档时把它的窗口带到前台" },
		"en-US": { $description: "Bring the document window to the front when it is opened" }
	}),
	renderWidth: z.natural().min(600).max(2e3).default(1100).volatile().i18n({
		"zh-CN": { $description: "给模型核对排版的图片宽度（像素）；越大越清晰、越费 token" },
		"en-US": { $description: "Width in pixels of the pictures the model checks layout with" }
	})
});
function apply(ctx, config = {}) {
	if (process.platform !== "win32") {
		ctx.logger.warn("dsh-office only supports Windows; no tools registered.");
		return;
	}
	const settings = () => resolveConfig(config);
	const client = new HelperClient((message) => ctx.logger.warn(message));
	const helper = {
		call: (cmd, args = {}, timeoutMs) => client.call(cmd, {
			suite: settings().suite,
			...args
		}, timeoutMs),
		onEvent: (listener) => client.onEvent(listener),
		dispose: () => client.dispose()
	};
	ctx.effect(() => () => helper.dispose(), "office: helper");
	const visionCache = /* @__PURE__ */ new Map();
	const supportsVision = (agent, signal) => {
		const route = agent?.session.requestHeader?.()?.config ?? agent?.options;
		const provider = route?.provider;
		const model = route?.model;
		const llm = ctx.get("llm");
		if (!provider || !model || typeof llm?.resolveModelInfo !== "function") return Promise.resolve(true);
		const key = `${provider}\u0000${model}`;
		let cached = visionCache.get(key);
		if (!cached) {
			cached = llm.resolveModelInfo(provider, model, signal).then((info) => info.inputModalities === void 0 || info.inputModalities.includes("image")).catch(() => {
				visionCache.delete(key);
				return true;
			});
			visionCache.set(key, cached);
		}
		return cached;
	};
	const tools = createTools({
		helper,
		settings,
		saveImage: (data, fileName) => ctx.attachments.saveImage({
			data,
			mediaType: "image/png",
			name: fileName
		}),
		vision: (raw) => {
			const exec = raw;
			return supportsVision(exec.agent, exec.signal);
		}
	});
	for (const tool of tools) ctx.effect(() => ctx.tools.register(tool), `office: tool ${tool.name}`);
	let installed;
	const installedApps = () => {
		installed ??= helper.call("status", {}, 3e4).then((apps) => new Set(apps.filter((app) => app.installed).map((app) => app.app))).catch(() => {
			installed = void 0;
			return /* @__PURE__ */ new Set();
		});
		return installed;
	};
	const working = /* @__PURE__ */ new WeakSet();
	let carded = false;
	ctx.on("tools/pre-execute", (exec, next) => {
		if (exec.name.startsWith("office_") && exec.agent) {
			working.add(exec.agent);
			carded = true;
		}
		return next();
	});
	const retire = (agent) => {
		if (!carded || typeof agent !== "object" || agent === null || !working.has(agent)) return;
		working.delete(agent);
		helper.call("card", {
			text: "AI 已完成",
			hold: false
		}, 5e3).catch(() => {});
		helper.call("settle", {}, 6e4).then((said) => {
			for (const line of said) ctx.logger.info(`office: ${line}`);
		}).catch(() => {});
	};
	ctx.on("agent/status", ({ agent, status }) => {
		if (status === "idle") retire(agent);
	});
	ctx.on("agent/disposed", ({ agent }) => {
		retire(agent);
	});
	const redirected = /* @__PURE__ */ new Set();
	ctx.on("tools/post-execute", async (exec, result, next) => {
		if (exec.name !== "skill" || result.isError || !settings().preferLive) return next();
		const skill = String(exec.arguments?.name ?? "");
		const app = FILE_SKILLS[skill];
		if (app === void 0) return next();
		const key = `${exec.agent?.session?.id ?? "default"}:${skill}`;
		if (redirected.has(key) || !(await installedApps()).has(app)) return next();
		redirected.add(key);
		return {
			kind: "accept",
			content: [{
				type: "text",
				text: redirectText(skill, app)
			}]
		};
	});
	ctx.inject(["webServer"], (webCtx) => {
		const handler = async (_req, res) => {
			let body;
			try {
				body = {
					ok: true,
					apps: await helper.call("status", {}, 3e4)
				};
			} catch (error) {
				body = {
					ok: false,
					error: error instanceof Error ? error.message : String(error)
				};
			}
			res.writeHead(200, {
				"content-type": "application/json; charset=utf-8",
				"cache-control": "no-store"
			});
			res.end(JSON.stringify(body));
		};
		webCtx.effect(() => webCtx.webServer.register({
			kind: "exact",
			path: STATUS_ROUTE,
			handler
		}), `office: ${STATUS_ROUTE}`);
	});
	ctx.inject(["systemPrompt"], (promptCtx) => {
		promptCtx.effect(() => promptCtx.systemPrompt.section({
			name: "office:guide",
			order: 73,
			text: promptText()
		}), "office: prompt");
	});
}
//#endregion
export { Config, DEFAULTS, ENTRY_ID, STATUS_ROUTE, apply, inject, name, resolveConfig };
