// Client half of dsh-ms-office.
//
// Registers one section on the Plugins page detail for this plugin: a status
// dot, the autostart switch and the port field, all talking to the host half's
// settings bridge (lib/sidecar-host.mjs). It renders nothing for any other
// plugin's detail page — the page hands every registered section the subject it
// is showing, and this one only speaks for its own item.
window.__ModuleLoader__.load({
	id: "dsh-ms-office",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		let react_jsx_runtime = require("react/jsx-runtime");
		let react = require("react");

		const CONFIG = {
			pluginName: "dsh-ms-office",
			itemId: "dsh-ms-office",
			bridge: "/api/ms-office",
			defaultPort: 3000,
			navZh: "Office 桥",
			navEn: "Office bridge",
		};

		const BRIDGE = CONFIG.bridge;

		const zh = {
			nav: CONFIG.navZh,
			status: "状态",
			running: "运行中",
			stopped: "未运行",
			starting: "启动中",
			unhealthy: "服务异常",
			invalidPort: "端口须为 1024–65535 的整数，或留空使用默认值",
			port: "端口",
			portNow: (p) => "当前端口 " + p,
			autostart: "随 dsh 服务自动启动",
			portLabel: "端口（留空使用默认 " + CONFIG.defaultPort + "）",
			portHint: "改端口并保存后立即生效；若与其它程序冲突会自动落到空闲端口。Office 加载项的地址由服务启动时写入，改端口后需要关闭并重新打开 Office 文档。",
			save: "保存",
			saving: "保存中…",
			loading: "加载中…",
			loadError: "读取失败: {message}",
			saveError: "保存失败: {message}",
			saved: "已保存。",
		};
		const en = {
			nav: CONFIG.navEn,
			status: "Status",
			running: "Running",
			stopped: "Not running",
			starting: "Starting",
			unhealthy: "Service unhealthy",
			invalidPort: "Port must be an integer from 1024 to 65535, or empty for the default",
			port: "Port",
			portNow: (p) => "current port " + p,
			autostart: "Start with the dsh service",
			portLabel: "Port (empty = default " + CONFIG.defaultPort + ")",
			portHint: "Saving applies the port immediately; if it collides with another program a free port is picked. The add-in URL is written when the server starts, so reopen your Office documents after changing it.",
			save: "Save",
			saving: "Saving…",
			loading: "Loading…",
			loadError: "Failed to load: {message}",
			saveError: "Failed to save: {message}",
			saved: "Saved.",
		};

		function makeT() {
			const lang = (typeof navigator !== "undefined" && navigator.language || "zh").toLowerCase().startsWith("en") ? en : zh;
			return (key, vars) => {
				let text = lang[key];
				if (typeof text === "function") return text(vars && vars.p !== undefined ? vars.p : "");
				if (typeof text !== "string") return key;
				if (vars) for (const k of Object.keys(vars)) text = text.split("{" + k + "}").join(String(vars[k]));
				return text;
			};
		}

		async function api(path, init) {
			const res = await fetch(BRIDGE + path, init);
			const data = await res.json().catch(() => ({}));
			if (!res.ok || data.ok === false) throw new Error(data.error || "HTTP " + res.status);
			return data;
		}

		const S = {
			wrap: { display: "flex", flexDirection: "column", gap: 12, paddingTop: 4 },
			title: { margin: 0, fontSize: 14, fontWeight: 500, lineHeight: "20px" },
			row: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" },
			dot: { width: 8, height: 8, borderRadius: 999, flex: "none" },
			dotOn: { background: "var(--dsw-alias-state-success-primary, #2e7d32)" },
			dotOff: { background: "var(--dsw-alias-label-tertiary, #999)" },
			label: { fontSize: 13, color: "var(--dsw-alias-label-secondary)" },
			input: {
				border: "0.5px solid var(--dsw-alias-border-l4, #ccc)", borderRadius: 8,
				background: "var(--dsw-alias-bg-layer-3, #fff)", height: 36, width: 180,
				font: "inherit", color: "inherit", outline: "none", padding: "0 12px", fontSize: 13,
			},
			hint: { margin: 0, fontSize: 12, lineHeight: "18px", color: "var(--dsw-alias-label-tertiary, #888)" },
			btn: {
				border: 0, borderRadius: 8, height: 32, padding: "0 14px", fontSize: 13,
				fontWeight: 500, color: "#fff", background: "var(--dsw-alias-button-primary-fill, var(--dsw-alias-brand-primary, #4169e1))",
				cursor: "pointer", alignSelf: "flex-start",
			},
			btnDisabled: { opacity: 0.5, cursor: "default" },
			note: { margin: 0, fontSize: 12.5, color: "var(--dsw-alias-state-success-primary, #2e7d32)" },
			error: { margin: 0, fontSize: 12.5, color: "var(--dsw-alias-state-error-primary, #c62828)", overflowWrap: "anywhere" },
			checkbox: { width: 15, height: 15, accentColor: "var(--dsw-alias-brand-primary, #4169e1)" },
			field: { display: "flex", flexDirection: "column", gap: 4 },
			fieldLabel: { fontSize: 12, color: "var(--dsw-alias-label-secondary)" },
		};

		function SidecarSection(props) {
			const subject = props && props.subject;
			// The detail page renders every registered section with the open page's
			// subject; this section only speaks for its own plugin.
			if (!subject || subject.kind !== "item") return null;
			if (subject.id !== CONFIG.pluginName && subject.id !== CONFIG.itemId) return null;

			const t = makeT();
			const [loading, setLoading] = react.useState(true);
			const [saving, setSaving] = react.useState(false);
			const [error, setError] = react.useState(null);
			const [notice, setNotice] = react.useState(null);
			const [status, setStatus] = react.useState({ running: false, port: null });
			const [config, setConfig] = react.useState({ port: "", autostart: true });

			react.useEffect(() => {
				let cancelled = false;
				(async () => {
					try {
						const data = await api("/describe");
						if (cancelled) return;
						setStatus(data.status || { running: false, port: null });
						setConfig({
							port: data.config && data.config.port ? String(data.config.port) : "",
							autostart: !data.config || data.config.autostart !== false,
						});
					} catch (err) {
						if (!cancelled) setError(t("loadError", { message: err instanceof Error ? err.message : String(err) }));
					} finally {
						if (!cancelled) setLoading(false);
					}
				})();
				let refreshing = false;
				const timer = setInterval(async () => {
					if (refreshing) return;
					refreshing = true;
					try {
						const data = await api("/describe");
						if (!cancelled) setStatus(data.status || { running: false, port: null });
					} catch {
						if (!cancelled) setStatus({ running: false, phase: "unhealthy", port: null });
					} finally { refreshing = false; }
				}, 3000);
				return () => { cancelled = true; clearInterval(timer); };
			}, []);

			const save = async () => {
				setSaving(true); setError(null); setNotice(null);
				try {
					const raw = String(config.port).trim();
					if (raw && (!/^\d+$/.test(raw) || Number(raw) < 1024 || Number(raw) > 65535)) throw new Error(t("invalidPort"));
					const data = await api("/save", {
						method: "POST",
						headers: { "content-type": "application/json" },
						body: JSON.stringify({ config: { port: raw === "" ? null : Number(raw), autostart: !!config.autostart } }),
					});
					setStatus(data.status || { running: false, port: null });
					setNotice(t("saved"));
				} catch (err) {
					setError(t("saveError", { message: err instanceof Error ? err.message : String(err) }));
				} finally {
					setSaving(false);
				}
			};

			if (loading) {
				return react_jsx_runtime.jsx("section", { "data-sidecar-settings": CONFIG.pluginName, children: react_jsx_runtime.jsx("p", { style: S.hint, children: t("loading") }) });
			}

			const children = [
				react_jsx_runtime.jsx("h4", { style: S.title, children: t("nav") }),
				react_jsx_runtime.jsxs("div", { style: S.row, children: [
					react_jsx_runtime.jsx("span", { style: Object.assign({}, S.dot, status.running ? S.dotOn : S.dotOff) }),
					react_jsx_runtime.jsx("span", { style: S.label, children: (status.running ? t("running") : status.phase === "starting" ? t("starting") : status.phase === "unhealthy" ? t("unhealthy") : t("stopped")) + (status.port ? " · " + t("portNow", { p: status.port }) : "") }),
				] }),
				react_jsx_runtime.jsxs("label", { style: S.row, children: [
					react_jsx_runtime.jsx("input", { type: "checkbox", style: S.checkbox, checked: !!config.autostart, onChange: (e) => setConfig(Object.assign({}, config, { autostart: e.target.checked })) }),
					react_jsx_runtime.jsx("span", { style: S.label, children: t("autostart") }),
				] }),
				react_jsx_runtime.jsxs("div", { style: S.field, children: [
					react_jsx_runtime.jsx("label", { style: S.fieldLabel, children: t("portLabel") }),
					react_jsx_runtime.jsx("input", { type: "text", inputMode: "numeric", style: S.input, value: config.port, onChange: (e) => setConfig(Object.assign({}, config, { port: e.target.value })) }),
				] }),
				react_jsx_runtime.jsx("p", { style: S.hint, children: t("portHint") }),
			];

			if (status.lastError) children.push(react_jsx_runtime.jsx("p", { style: S.error, children: status.lastError }));
			if (notice) children.push(react_jsx_runtime.jsx("p", { style: S.note, children: notice }));
			if (error) children.push(react_jsx_runtime.jsx("p", { style: S.error, children: error }));
			children.push(
				react_jsx_runtime.jsx("button", { type: "button", disabled: saving, onClick: () => { void save(); }, style: Object.assign({}, S.btn, saving ? S.btnDisabled : null), children: saving ? t("saving") : t("save") }),
			);

			return react_jsx_runtime.jsx("section", { "data-sidecar-settings": CONFIG.pluginName, style: S.wrap, children: children });
		}

		function apply(ctx) {
			ctx.slots.inject("plugins.detail.section", () => ctx.slots.register(
				{ name: "plugins.detail.section", id: "sidecar-settings." + CONFIG.pluginName, order: 100 },
				SidecarSection,
			));
		}

		const inject = ["slots"];
		exports.apply = apply;
		exports.inject = inject;
		return module.exports;
	},
});
