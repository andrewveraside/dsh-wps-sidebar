window.__ModuleLoader__.load({
	id: "@andrewveraside/dsh-wps-sidebar",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		//#region \0rolldown/runtime.js
		var __create = Object.create;
		var __defProp = Object.defineProperty;
		var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
		var __getOwnPropNames = Object.getOwnPropertyNames;
		var __getProtoOf = Object.getPrototypeOf;
		var __hasOwnProp = Object.prototype.hasOwnProperty;
		var __copyProps = (to, from, except, desc) => {
			if (from && typeof from === "object" || typeof from === "function") for (var keys = __getOwnPropNames(from), i = 0, n = keys.length, key; i < n; i++) {
				key = keys[i];
				if (!__hasOwnProp.call(to, key) && key !== except) __defProp(to, key, {
					get: ((k) => from[k]).bind(null, key),
					enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable
				});
			}
			return to;
		};
		var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(isNodeMode || !mod || !mod.__esModule || !__hasOwnProp.call(mod, "default") ? __defProp(target, "default", {
			value: mod,
			enumerable: true
		}) : target, mod));
		//#endregion
		let react = require("react");
		react = __toESM(react, 1);
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
		//#region lib/types/client/accent.js
		/**
		* The accent colour shared by the copylee DSH plugins (docs/ui-spec.md).
		*
		* One choice — terracotta orange by default, blue or black — is kept in the
		* browser under a key every plugin reads, and painted as two CSS variables on
		* `<body>`, where the Host keeps its own theme variables. A plugin's styles
		* only ever write `var(--cl-accent)` / `var(--cl-accent-ink)`, so changing the
		* colour in any plugin's settings recolours all of them at once.
		*
		* This file is identical in every plugin; change it in all of them together.
		*/
		const ACCENTS = {
			orange: {
				name: "陶土橙",
				accent: "#D97757",
				ink: "#FFFFFF"
			},
			blue: {
				name: "蓝色",
				accent: "#3D63E6",
				ink: "#FFFFFF"
			},
			black: {
				name: "黑色",
				accent: "var(--dsw-alias-label-primary, #1F1E1D)",
				ink: "var(--dsw-alias-bg-base, #FFFFFF)"
			}
		};
		const ACCENT_IDS = Object.keys(ACCENTS);
		const DEFAULT_ACCENT = "orange";
		/** Use these in styles; the fallbacks cover the moment before the first paint. */
		const ACCENT = `var(--cl-accent, ${ACCENTS[DEFAULT_ACCENT].accent})`;
		const ACCENT_INK = `var(--cl-accent-ink, ${ACCENTS[DEFAULT_ACCENT].ink})`;
		const KEY = "copylee.dsh.accent";
		const EVENT = "copylee-dsh-accent";
		function isAccent(value) {
			return typeof value === "string" && Object.hasOwn(ACCENTS, value);
		}
		/** The choice made in this browser, or null while none has been made. */
		function storedAccent() {
			try {
				const stored = localStorage.getItem(KEY);
				return isAccent(stored) ? stored : null;
			} catch {
				return null;
			}
		}
		function readAccent() {
			return storedAccent() ?? "orange";
		}
		function paint() {
			const colors = ACCENTS[readAccent()];
			document.body?.style.setProperty("--cl-accent", colors.accent);
			document.body?.style.setProperty("--cl-accent-ink", colors.ink);
		}
		/** Choose the accent for every plugin. */
		function writeAccent(accent) {
			try {
				localStorage.setItem(KEY, accent);
			} catch {}
			paint();
			window.dispatchEvent(new Event(EVENT));
		}
		/** Paint the current accent and keep it current. Returns the undo for plugin disposal. */
		function installAccent() {
			if (typeof document === "undefined") return () => {};
			paint();
			if (document.body === null) document.addEventListener("DOMContentLoaded", paint, { once: true });
			window.addEventListener(EVENT, paint);
			window.addEventListener("storage", paint);
			return () => {
				window.removeEventListener(EVENT, paint);
				window.removeEventListener("storage", paint);
			};
		}
		function useAccent() {
			const [accent, setAccent] = react.useState(readAccent);
			react.useEffect(() => {
				const sync = () => setAccent(readAccent());
				window.addEventListener(EVENT, sync);
				window.addEventListener("storage", sync);
				return () => {
					window.removeEventListener(EVENT, sync);
					window.removeEventListener("storage", sync);
				};
			}, []);
			return [accent, writeAccent];
		}
		/** A row of round swatches: one of them is always chosen. */
		function AccentPicker({ label = "强调色", hint = "", names }) {
			const [accent, choose] = useAccent();
			const h = react.createElement;
			const name = (id) => names?.[id] ?? ACCENTS[id].name;
			return h("div", { style: {
				display: "flex",
				alignItems: "center",
				gap: 10,
				fontSize: 13,
				flexWrap: "wrap"
			} }, h("span", { style: { fontWeight: 500 } }, label), h("span", {
				role: "radiogroup",
				"aria-label": label,
				style: {
					display: "inline-flex",
					gap: 8
				}
			}, ...ACCENT_IDS.map((id) => h("button", {
				key: id,
				type: "button",
				role: "radio",
				"aria-checked": accent === id,
				"aria-label": name(id),
				title: name(id),
				onClick: () => choose(id),
				style: {
					width: 18,
					height: 18,
					padding: 0,
					borderRadius: "50%",
					cursor: "pointer",
					background: ACCENTS[id].accent,
					border: "2px solid var(--dsw-alias-bg-base, #fff)",
					boxShadow: accent === id ? "0 0 0 2px var(--dsw-alias-label-primary, #1F1E1D)" : "0 0 0 1px var(--dsw-alias-border-l2, rgba(127,127,127,.35))"
				}
			}))), h("span", { style: {
				fontSize: 12,
				color: "var(--dsw-alias-label-tertiary, #888)"
			} }, hint ? `${name(accent)} · ${hint}` : name(accent)));
		}
		//#endregion
		//#region lib/types/client/nav-icon.js
		/**
		* Give the "Office" row in the DSH settings navigation a document icon
		* instead of the shell's fallback gear.
		*
		* `settings.section` registrations only carry `id`, `order` and `label`; the
		* shell picks icons for built-in ids only. So the row is marked by its label
		* and CSS swaps the glyph (Lucide file-text, drawn as a currentColor mask so
		* hover / active colors still apply).
		*/
		const MARKER = "data-dsh-office-settings-nav";
		const ICON_SVG = "%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z'/%3E%3Cpath d='M14 2v4a2 2 0 0 0 2 2h4'/%3E%3Cpath d='M10 9H8'/%3E%3Cpath d='M16 13H8'/%3E%3Cpath d='M16 17H8'/%3E%3C/svg%3E";
		const navIconCss = `
[${MARKER}] > svg:first-child { display: none; }
[${MARKER}]::before {
  content: '';
  flex: none;
  width: 16px;
  height: 16px;
  background: currentColor;
  -webkit-mask: url("data:image/svg+xml,${ICON_SVG}") center / contain no-repeat;
  mask: url("data:image/svg+xml,${ICON_SVG}") center / contain no-repeat;
}
`;
		/** Keep the marker on the settings-nav button labelled `label`; returns a disposer. */
		function registerNavIcon(label) {
			const style = document.createElement("style");
			style.setAttribute("data-dsh-office", "nav-icon");
			style.textContent = navIconCss;
			document.head.appendChild(style);
			let disposed = false;
			let frame = 0;
			const sync = () => {
				frame = 0;
				if (disposed) return;
				const dialogs = document.querySelectorAll("[role=\"dialog\"]");
				if (dialogs.length === 0) return;
				for (const dialog of dialogs) for (const button of dialog.querySelectorAll("nav button")) {
					const mine = button.textContent?.trim() === label;
					if (mine && !button.hasAttribute(MARKER)) button.setAttribute(MARKER, "");
					else if (!mine && button.hasAttribute(MARKER)) button.removeAttribute(MARKER);
				}
			};
			sync();
			const observer = new MutationObserver(() => {
				if (frame === 0) frame = requestAnimationFrame(sync);
			});
			observer.observe(document.body, {
				childList: true,
				subtree: true
			});
			return () => {
				disposed = true;
				if (frame !== 0) cancelAnimationFrame(frame);
				observer.disconnect();
				style.remove();
				document.querySelectorAll(`[${MARKER}]`).forEach((element) => element.removeAttribute(MARKER));
			};
		}
		//#endregion
		//#region lib/types/client/index.js
		/**
		* Browser half: the "Office" section in Settings. It reads the plugin's own
		* Loader entry through `remote.settings.describe()` and writes with
		* `remote.settings.mutate()`; every field is volatile and each change saves
		* itself. The live status (installed apps, open documents) comes from the
		* host route in ../index.ts.
		*/
		const STATUS_ROUTE = "/api/dsh-office/status";
		const APP_NAMES = {
			word: "Word",
			excel: "Excel",
			ppt: "PowerPoint"
		};
		const WPS_NAMES = {
			word: "WPS 文字",
			excel: "WPS 表格",
			ppt: "WPS 演示"
		};
		const SUITES = [
			["auto", "自动"],
			["office", "微软 Office"],
			["wps", "WPS Office"]
		];
		const WIDTHS = [
			800,
			1100,
			1400,
			1800
		];
		const h = react.createElement;
		const S = {
			page: {
				display: "grid",
				gap: 20,
				maxWidth: 880,
				paddingBottom: 32,
				color: "var(--dsw-alias-label-primary, inherit)"
			},
			title: {
				margin: 0,
				fontSize: 20,
				fontWeight: 600
			},
			subtitle: {
				margin: "6px 0 0",
				fontSize: 13,
				lineHeight: 1.6,
				color: "var(--dsw-alias-label-secondary, #666)"
			},
			card: {
				display: "grid",
				gap: 14,
				padding: 16,
				borderRadius: "var(--dsw-radius-lg, 12px)",
				border: "1px solid var(--dsw-alias-border-l2, rgba(127,127,127,.25))"
			},
			cardTitle: {
				margin: 0,
				fontSize: 15,
				fontWeight: 600
			},
			hint: {
				margin: 0,
				fontSize: 12,
				lineHeight: 1.55,
				color: "var(--dsw-alias-label-tertiary, #888)"
			},
			toggleRow: {
				display: "flex",
				gap: 12,
				alignItems: "flex-start",
				justifyContent: "space-between"
			},
			toggleText: {
				display: "grid",
				gap: 4,
				fontSize: 13,
				fontWeight: 500
			},
			error: {
				fontSize: 12,
				color: "var(--dsw-alias-state-error-primary, #d33)"
			},
			ok: {
				fontSize: 12,
				color: "var(--dsw-alias-state-success-primary, #2a2)"
			},
			statusLine: {
				display: "flex",
				gap: 8,
				alignItems: "baseline",
				fontSize: 13
			},
			dot: {
				width: 8,
				height: 8,
				borderRadius: 4,
				flex: "0 0 auto",
				alignSelf: "center"
			},
			chips: {
				display: "flex",
				gap: 8,
				flexWrap: "wrap"
			},
			secondary: {
				padding: "6px 12px",
				fontSize: 12,
				borderRadius: "var(--dsw-radius-md, 8px)",
				border: "1px solid var(--dsw-alias-border-l2, rgba(127,127,127,.3))",
				background: "transparent",
				color: "inherit",
				cursor: "pointer"
			}
		};
		function Switch({ checked, disabled, label, onChange }) {
			return h("button", {
				type: "button",
				role: "switch",
				"aria-checked": checked,
				"aria-label": label,
				disabled,
				onClick: () => onChange(!checked),
				style: {
					flex: "0 0 auto",
					width: 40,
					height: 22,
					borderRadius: 11,
					border: "none",
					padding: 2,
					cursor: disabled ? "default" : "pointer",
					background: checked ? ACCENT : "var(--dsw-alias-border-l3, rgba(127,127,127,.35))",
					opacity: disabled ? .5 : 1,
					transition: "background .15s"
				}
			}, h("span", { style: {
				display: "block",
				width: 18,
				height: 18,
				borderRadius: 9,
				background: checked ? ACCENT_INK : "#fff",
				transform: checked ? "translateX(18px)" : "none",
				transition: "transform .15s"
			} }));
		}
		function OfficeSection({ ctx }) {
			const [loaded, setLoaded] = react.useState(false);
			const [writable, setWritable] = react.useState(false);
			const [view, setView] = react.useState(void 0);
			const [settings, setSettings] = react.useState(DEFAULTS);
			const [message, setMessage] = react.useState(null);
			const [status, setStatus] = react.useState(null);
			const [checking, setChecking] = react.useState(false);
			const load = react.useCallback(async () => {
				try {
					const result = await ctx.remote.settings.describe();
					if (!result.ok) throw new Error(result.error.message);
					const found = result.value.namespaces.find((item) => item.ns === ENTRY_ID);
					setWritable(result.value.writable);
					setView(found);
					setSettings(resolveConfig(found?.value));
				} catch (error) {
					setMessage({
						kind: "error",
						text: `读取设置失败：${error.message}`
					});
				} finally {
					setLoaded(true);
				}
			}, [ctx]);
			const refreshStatus = react.useCallback(async () => {
				setChecking(true);
				try {
					const response = await fetch(STATUS_ROUTE, { credentials: "same-origin" });
					setStatus(await response.json());
				} catch {
					setStatus({
						ok: false,
						error: "无法连接到插件"
					});
				} finally {
					setChecking(false);
				}
			}, []);
			react.useEffect(() => {
				load();
				refreshStatus();
				return ctx.remote.$on?.("settings/document-updated", (ns) => {
					if (ns === "andrewveraside-wps-sidebar") load();
				}) ?? void 0;
			}, [
				ctx,
				load,
				refreshStatus
			]);
			const disabled = !writable || view === void 0;
			const change = async (key, value) => {
				if (view === void 0) return;
				setSettings((previous) => ({
					...previous,
					[key]: value
				}));
				try {
					const result = await ctx.remote.settings.mutate(ENTRY_ID, [{
						op: "set",
						path: [key],
						value
					}], view.revision);
					if (!result.ok) throw new Error(result.error.message);
					setMessage({
						kind: "ok",
						text: "已保存，下一次调用即生效"
					});
				} catch (error) {
					setMessage({
						kind: "error",
						text: `保存失败：${error.message}`
					});
				}
				await load();
			};
			if (!loaded) return h("div", { style: S.page }, "加载中…");
			return h("div", { style: S.page }, h("div", null, h("h2", { style: S.title }, "Office"), h("p", { style: S.subtitle }, "AI 直接在你电脑上打开的 Word、Excel、PowerPoint 里改文档：改动实时出现在窗口中，你可以同时编辑；窗口不必在前台。只有调用保存时才写入文件。")), view === void 0 ? h("div", { style: S.error }, "没有找到本插件的设置项（插件可能未启用）。") : null, view !== void 0 && !writable ? h("div", { style: S.error }, "当前设置为只读。") : null, h("section", { style: S.card }, h("div", { style: {
				display: "flex",
				justifyContent: "space-between",
				alignItems: "center"
			} }, h("h3", { style: S.cardTitle }, "应用状态"), h("button", {
				type: "button",
				style: S.secondary,
				disabled: checking,
				onClick: () => {
					refreshStatus();
				}
			}, checking ? "检查中…" : "重新检查")), status === null ? h("p", { style: S.hint }, "检查中…") : !status.ok ? h("div", { style: S.error }, `控制组件不可用：${status.error ?? ""}`) : (status.apps ?? []).map((app) => h("div", {
				key: `${app.suite ?? "office"}-${app.app}`,
				style: S.statusLine
			}, h("span", { style: {
				...S.dot,
				background: !app.installed ? "#999" : app.running ? "#2a2" : ACCENT
			} }), h("span", { style: {
				minWidth: 86,
				fontWeight: 500
			} }, (app.suite === "wps" ? WPS_NAMES : APP_NAMES)[app.app]), h("span", { style: {
				...S.hint,
				fontSize: 13
			} }, !app.installed ? "未安装" : !app.running ? "已安装，未运行（需要时自动启动）" : app.documents.length === 0 ? "运行中，没有打开的文档" : `运行中：${app.documents.map((doc) => doc.name + (doc.saved ? "" : "（未保存）")).join("、")}`))), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "办公套件", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "新打开的文件用哪个软件。“自动”：装了微软 Office 就用它，否则用 WPS。已经打开着的文档始终在它所在的软件里修改。WPS 文字里的公式是可编辑的原生公式；WPS 演示不支持插入公式，那里的公式是叠在文字上的矢量图片，不能编辑，移动文本框后需要让 AI 重写该文本。幻灯片里公式多的话，推荐用 PowerPoint。")), h("div", { style: {
				display: "flex",
				gap: 6,
				flexShrink: 0
			} }, SUITES.map(([value, label]) => h("button", {
				key: value,
				type: "button",
				disabled,
				onClick: () => {
					change("suite", value);
				},
				style: {
					...S.secondary,
					...settings.suite === value ? {
						background: ACCENT,
						color: "#fff",
						borderColor: ACCENT
					} : {}
				}
			}, label)))), h("p", { style: S.hint }, "需要本机装有微软 Office 或 WPS Office。Word 里 AI 的一批修改可以用一次 Ctrl+Z 撤回；Excel 和 PowerPoint 的修改无法用 Ctrl+Z 撤回。你正在单元格里输入或开着对话框时，AI 会等你。")), h("section", { style: S.card }, h("h3", { style: S.cardTitle }, "行为"), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "静默模式", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "在后台完成：不弹出 Word、Excel、PowerPoint 的窗口，不显示迷你卡片，文字一次写入。AI 照常渲染页面检查成品；回合结束时，后台打开的文档自动保存并关闭。你自己已经打开着的文档仍在原窗口里修改，只是不再滚动和弹卡片。开启后，下面的“跟随”“逐字写入”“迷你卡片”不再起作用。")), h(Switch, {
				checked: settings.silent,
				disabled,
				label: "静默模式",
				onChange: (value) => {
					change("silent", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "跟随 AI 的修改位置", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "AI 改到哪里，窗口就自动滚动到哪里（Word 滚到那一段，Excel 滚到那片单元格，PowerPoint 切到那一页），不用自己翻找。关闭即静默执行：AI 照常修改，但不动你的视图。想看时点迷你卡片上的“查看”：文档窗口被带到前台并开始跟随，按钮变成“查看中”；你切到别的窗口后它自动熄灭，再点一次回来。")), h(Switch, {
				checked: settings.follow,
				disabled,
				label: "跟随 AI 的修改位置",
				onChange: (value) => {
					change("follow", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "逐字写入", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "文字像打字一样逐步出现，而不是整块贴上去；表格逐行填入。每段约多花 0.3 秒，关闭后写得最快。")), h(Switch, {
				checked: settings.typing,
				disabled,
				label: "逐字写入",
				onChange: (value) => {
					change("typing", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "迷你卡片", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "AI 每次动手修改文档时，屏幕右下角显示一张小卡片：正在改哪个文档、改到第几项，以及“查看”“快速”两个开关（“快速”亮起时一次写入，熄灭时逐字写入），点一下立即生效（连正在进行的这一批也会跟着变）。卡片可拖动，不抢焦点；AI 这一轮结束时消失。在卡片上切换过的开关，本次运行期间以卡片为准。")), h(Switch, {
				checked: settings.card,
				disabled,
				label: "迷你卡片",
				onChange: (value) => {
					change("card", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "交付前整体视觉检查", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "AI 保存时，如果它改完之后还没看过成品，就自动把整份文档拼成带页码的总览图（一张图最多 6 页）交给它过一遍，发现问题先修再交付。会多用一些 token。")), h(Switch, {
				checked: settings.finalCheck,
				disabled,
				label: "交付前整体视觉检查",
				onChange: (value) => {
					change("finalCheck", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "优先在真实 Office 里操作", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "DSH 自带的 office-docx / office-pptx 技能是用脚本生成文件，你看不到过程，公式和排版也由脚本决定。开启后，AI 第一次想加载这些技能时会被引回本插件（对应的 Office 应用已安装时才拦）；它确有需要而再次请求时放行。表格不在此列：数据分析默认用代码和自带的 office-xlsx 技能，你指明要在 Excel 或 WPS 表格里做时，AI 才在窗口里操作。")), h(Switch, {
				checked: settings.preferLive,
				disabled,
				label: "优先在真实 Office 里操作",
				onChange: (value) => {
					change("preferLive", value);
				}
			})), h("div", { style: S.toggleRow }, h("div", { style: S.toggleText }, "打开文档时显示到前台", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "关闭后文档仍会打开，只是不抢你当前的窗口。")), h(Switch, {
				checked: settings.showOnOpen,
				disabled,
				label: "打开文档时显示到前台",
				onChange: (value) => {
					change("showOnOpen", value);
				}
			})), h("div", { style: {
				display: "grid",
				gap: 8
			} }, h("div", { style: S.toggleText }, "核对排版的图片宽度", h("span", { style: {
				...S.hint,
				fontWeight: 400
			} }, "AI 改完后会把页面 / 幻灯片渲染成图片检查排版。越大越清晰，也越费 token。")), h("div", { style: S.chips }, WIDTHS.map((width) => h("button", {
				key: width,
				type: "button",
				disabled,
				onClick: () => {
					change("renderWidth", width);
				},
				style: {
					...S.secondary,
					...settings.renderWidth === width ? {
						background: ACCENT,
						borderColor: ACCENT,
						color: ACCENT_INK
					} : {}
				}
			}, `${width} 像素${width === DEFAULTS.renderWidth ? "（默认）" : ""}`)))), message === null ? null : h("span", { style: message.kind === "ok" ? S.ok : S.error }, message.text)), h(AccentPicker, null));
		}
		const inject = [
			"slots",
			"remote",
			"remote.settings"
		];
		function apply(ctx) {
			ctx.effect(() => installAccent(), "office: accent colour");
			ctx.effect(() => registerNavIcon("Office"), "office: settings nav icon");
			ctx.slots.inject("settings.section", () => ctx.slots.register({
				name: "settings.section",
				id: ENTRY_ID,
				order: 62,
				label: () => "Office"
			}, () => h(OfficeSection, { ctx })));
		}
		//#endregion
		exports.apply = apply;
		exports.inject = inject;
		return module.exports;
	}
});
