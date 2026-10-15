/**
 * Sidecar host for dsh-ms-office.
 *
 * The plugin owns a local HTTP bridge (server/server.js) that Office add-in
 * task panes poll for commands. This module spawns it, waits for its health
 * endpoint, keeps it alive across crashes, points it at a usable port, and
 * tears it down on dispose.
 *
 * It also mounts the plugin's settings surface on the harness web server,
 * rendered by the client half on the Plugins page detail:
 *
 *   GET  <bridge>/describe → { ok, status, config }
 *   POST <bridge>/save     → { config: { port, autostart } }; applies live
 *
 * Config lives in `$DSH_HOME/<pluginName>/settings.json`. `port: null` means
 * "use the default, fall back to any free port"; an explicit port is probed the
 * same way. `autostart: false` keeps the sidecar down until the next save turns
 * it back on.
 *
 * The bundled skill is published into `~/.agents/skills` so the harness can
 * discover it.
 *
 * What is deliberately NOT here: a Windows Job Object. A hard kill of the dsh
 * process tree still orphans the sidecar — the dispose hook covers every
 * graceful path only.
 */

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { publishSkillTree } from "./skill-publisher.mjs";

const HEALTH_TIMEOUT_MS = 10_000;
const RESTART_DELAY_MS = 1_000;
/** Crash-restart budget per rolling minute. */
const RESTART_BUDGET = 5;

function dshHome() {
  const env = process.env.DSH_HOME;
  if (typeof env === "string" && env.trim()) return env.trim();
  return path.join(os.homedir(), ".dsh");
}

function log(sidecar, message) {
  console.log(`[${sidecar.pluginName}] ${message}`);
}

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return fallback;
  }
}

function writeJson(file, value) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function sendJson(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(body),
  });
  res.end(body);
}

async function readJsonBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 16 * 1024) throw new Error("settings request exceeds 16 KiB");
    chunks.push(chunk);
  }
  const body = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(Buffer.concat(chunks)) || "{}");
  if (!body || typeof body !== "object" || Array.isArray(body)) throw new Error("settings must be a JSON object");
  return body;
}

// ── settings ─────────────────────────────────────────────────────────────────

function settingsPath(sidecar) {
  return path.join(dshHome(), sidecar.pluginName, "settings.json");
}

function defaultSettings() {
  return { version: 1, port: null, autostart: true };
}

function readSettings(sidecar) {
  const raw = readJson(settingsPath(sidecar), {});
  const src = raw && typeof raw === "object" ? raw : {};
  const port = Number(src.port);
  return {
    ...defaultSettings(),
    port: Number.isInteger(port) && port >= 1024 && port <= 65535 ? port : null,
    autostart: src.autostart === undefined ? true : Boolean(src.autostart),
  };
}

function writeSettings(sidecar, settings) {
  writeJson(settingsPath(sidecar), { ...defaultSettings(), ...settings });
}

// ── process management ───────────────────────────────────────────────────────

/** Free port picked by the OS; used only when the chosen port is taken. */
function findFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      server.close(() => resolve(port));
    });
    server.on("error", reject);
  });
}

function waitHealth(sidecar, port, timeoutMs = HEALTH_TIMEOUT_MS) {
  const deadline = Date.now() + timeoutMs;
  return new Promise((resolve, reject) => {
    const attempt = async () => {
      if (sidecar.disposed) return reject(new Error("disposed"));
      if (!sidecar.child || sidecar.child.exitCode !== null) return reject(new Error("sidecar exited before becoming healthy"));
      try {
        const res = await fetch(`http://127.0.0.1:${port}${sidecar.healthPath}`, { signal: AbortSignal.timeout(1000) });
        if (res.ok) {
          const health = await res.json();
          if (health.plugin === sidecar.pluginName && health.pid === sidecar.pid) return resolve(true);
        }
      } catch {
        // not up yet
      }
      if (Date.now() > deadline) return reject(new Error(`health timeout on :${port}`));
      setTimeout(attempt, 250);
    };
    attempt();
  });
}

async function pickPort(sidecar) {
  const settings = readSettings(sidecar);
  const wanted = settings.port ?? sidecar.defaultPort;
  const busy = await new Promise((resolve) => {
    const probe = net.createServer();
    probe.once("error", () => resolve(true));
    probe.once("listening", () => probe.close(() => resolve(false)));
    probe.listen(wanted, "127.0.0.1");
  });
  if (!busy) return wanted;
  const port = await findFreePort();
  log(sidecar, `port ${wanted} is taken; falling back to ${port}`);
  return port;
}

function killChild(sidecar) {
  const child = sidecar.child;
  sidecar.child = null;
  sidecar.running = false;
  sidecar.pid = null;
  sidecar.phase = "stopped";
  clearTimeout(sidecar.restartTimer);
  sidecar.restartTimer = null;
  if (!child || !child.pid || child.exitCode !== null) return Promise.resolve();
  return new Promise((resolve) => {
    const timer = setTimeout(() => { try { child.kill('SIGKILL'); } catch {} }, 3000);
    child.once('exit', () => { clearTimeout(timer); resolve(); });
    if (process.platform === "win32") {
      const killer = spawn("taskkill", ["/F", "/T", "/PID", String(child.pid)], { stdio: "ignore", windowsHide: true });
      killer.on('error', (error) => { sidecar.lastError = error.message; try { child.kill(); } catch {} });
    } else {
      try { child.kill("SIGTERM"); } catch { clearTimeout(timer); resolve(); }
    }
  });
}

function startServer(sidecar) {
  if (sidecar.startPromise) return sidecar.startPromise;
  sidecar.startPromise = launchServer(sidecar).finally(() => { sidecar.startPromise = null; });
  return sidecar.startPromise;
}

async function launchServer(sidecar) {
  sidecar.starting = true;
  try {
    const port = await pickPort(sidecar);
    if (sidecar.disposed || !readSettings(sidecar).autostart) return;
    sidecar.currentPort = port;
    sidecar.requestedPort = readSettings(sidecar).port;
    sidecar.phase = "starting";
    sidecar.lastError = null;
    const child = spawn(process.execPath, [path.join(sidecar.pluginDir, sidecar.entry)], {
      env: { ...process.env, PORT: String(port) },
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    });
    sidecar.child = child;
    sidecar.running = false;
    sidecar.pid = child.pid;
    const capture = (data) => { sidecar.logs = (sidecar.logs + data.toString()).slice(-8000); };
    child.stdout.on("data", capture);
    child.stderr.on("data", capture);
    child.on("error", (error) => {
      if (sidecar.child !== child) return;
      sidecar.lastError = error.message;
      sidecar.phase = "unhealthy";
      sidecar.running = false;
    });
    log(sidecar, `sidecar pid ${child.pid} on 127.0.0.1:${port}`);

    child.on("exit", (code) => {
      if (sidecar.disposed || sidecar.child !== child) return;
      sidecar.child = null;
      sidecar.running = false;
      sidecar.pid = null;
      sidecar.phase = "unhealthy";
      sidecar.lastError = `sidecar exited (code ${code})`;
      log(sidecar, `sidecar exited (code ${code}); restarting`);
      const now = Date.now();
      sidecar.restarts = sidecar.restarts.filter((t) => now - t < 60_000);
      sidecar.restarts.push(now);
      if (sidecar.restarts.length > RESTART_BUDGET) {
        log(sidecar, "restart budget exhausted for this minute; giving up until next boot");
        return;
      }
      sidecar.restartTimer = setTimeout(() => {
        sidecar.restartTimer = null;
        if (!sidecar.disposed && readSettings(sidecar).autostart) {
          startServer(sidecar).catch((e) => log(sidecar, `restart failed: ${e.message}`));
        }
      }, RESTART_DELAY_MS);
    });

    try {
      await waitHealth(sidecar, port);
      if (sidecar.child === child && !sidecar.disposed) {
        sidecar.running = true;
        sidecar.phase = "healthy";
      }
    } catch (e) {
      if (sidecar.child === child) {
        sidecar.running = false;
        sidecar.phase = "unhealthy";
        sidecar.lastError = e.message;
      }
      log(sidecar, `sidecar unhealthy: ${e.message}`);
    }
  } finally {
    sidecar.starting = false;
  }
}

function stopServer(sidecar) {
  return killChild(sidecar);
}

/** Apply settings: (re)start or stop the sidecar to match the saved config. */
async function applySettings(sidecar) {
  const settings = readSettings(sidecar);
  if (!settings.autostart) {
    await stopServer(sidecar);
    if (sidecar.startPromise) await sidecar.startPromise;
    log(sidecar, "sidecar stopped (autostart disabled)");
    return;
  }
  if (sidecar.startPromise) await sidecar.startPromise;
  const portChanged = sidecar.child && settings.port !== sidecar.requestedPort;
  if (sidecar.running && !portChanged) return;
  await stopServer(sidecar);
  sidecar.restarts = [];
  await startServer(sidecar);
}

// ── skills ───────────────────────────────────────────────────────────────────

/**
 * Host UI language. Same rule as the rest of the project: `en*` means English,
 * anything else means Chinese. There is no clean host-side locale signal to read
 * (the harness page's locale is not reachable from a plugin), so the process
 * environment is the best available proxy and the fallback is Chinese.
 */
function hostLang() {
  const raw = process.env.LC_ALL || process.env.LC_MESSAGES || process.env.LANG || "";
  return String(raw).trim().toLowerCase().startsWith("en") ? "en" : "zh";
}

/**
 * Publish each `skills/<name>/` directory that has a SKILL.md.
 *
 * A skill is addressed by its directory name, so shipping one directory per
 * language would register two skills and the model would see the same capability
 * twice. Instead one skill is published and its body is picked by language:
 * `SKILL.md` is the Chinese source of truth, `SKILL.<lang>.md` is the variant.
 */
function publishSkills(sidecar) {
  const skillsDir = path.join(sidecar.pluginDir, "skills");
  const targetRoot = path.join(os.homedir(), ".agents", "skills");
  const lang = hostLang();
  try {
    for (const entry of fs.readdirSync(skillsDir, { withFileTypes: true })) {
      if (!entry.isDirectory() || entry.isSymbolicLink()) continue;
      const name = entry.name;
      const src = path.join(skillsDir, name);
      if (!fs.existsSync(path.join(src, "SKILL.md"))) continue;
      try {
        const result = publishSkillTree(src, path.join(targetRoot, name), { lang });
        if (result.changed) log(sidecar, `skill published → ~/.agents/skills/${name} (${lang})`);
      } catch (error) {
        log(sidecar, `skill publishing skipped (${name}): ${error.message}`);
      }
    }
  } catch (e) {
    log(sidecar, `skill publishing skipped: ${e.message}`);
  }
}

// ── settings bridge ──────────────────────────────────────────────────────────

function statusView(sidecar) {
  return {
    running: Boolean(sidecar.running),
    phase: sidecar.phase,
    port: sidecar.child ? sidecar.currentPort : null,
    pid: sidecar.pid || null,
    lastError: sidecar.lastError,
    logs: sidecar.logs,
    sideloadSupported: process.platform === "win32",
  };
}

function describe(sidecar) {
  return { ok: true, status: statusView(sidecar), config: readSettings(sidecar) };
}

function normalizeConfigPatch(sidecar, body) {
  if (!body.config || typeof body.config !== "object" || Array.isArray(body.config)) throw new Error("config must be an object");
  const patch = body.config;
  const next = readSettings(sidecar);
  if ("port" in patch) {
    const raw = patch.port;
    const n = Number(raw);
    if (raw === null || raw === "" || raw === 0 || raw === "0") next.port = null;
    else if (!Number.isInteger(n) || n < 1024 || n > 65535) {
      throw new Error("port must be 1024–65535, or empty for the default");
    } else next.port = n;
  }
  if ("autostart" in patch) {
    if (typeof patch.autostart !== "boolean") throw new Error("autostart must be a boolean");
    next.autostart = patch.autostart;
  }
  return next;
}

// ── Windows add-in (un)registration ───────────────────────────────────────────
// The bridge auto-registers the add-in in the WEF registry on every start. This
// endpoint is the in-product recovery path for when that automatic step fails:
// instead of asking the user to run sideload.ps1 by hand, the plugin detail
// page exposes a button that re-applies the same registry entry here.

const WEF_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";

function sideloadManifestPath(sidecar) {
  const runtime = path.join(dshHome(), sidecar.pluginName, "manifest.xml");
  if (fs.existsSync(runtime)) return runtime;
  return path.join(sidecar.pluginDir, "server", "manifest.xml");
}

function runSideload(sidecar, action) {
  if (process.platform !== "win32") {
    return { ok: false, error: "manual load is required on this platform (no WEF registry)" };
  }
  const manifest = sideloadManifestPath(sidecar);
  if (!fs.existsSync(manifest)) return { ok: false, error: `manifest not found: ${manifest}` };
  const content = fs.readFileSync(manifest, "utf8");
  const m = content.match(/<Id>\s*([0-9a-fA-F-]{36})\s*<\/Id>/);
  if (!m) return { ok: false, error: "manifest has no <Id> GUID" };
  const guid = m[1];
  const args =
    action === "unregister"
      ? ["delete", WEF_KEY, "/v", guid, "/f"]
      : ["add", WEF_KEY, "/v", guid, "/t", "REG_SZ", "/d", manifest, "/f"];
  const status = spawnSync("reg", args, { stdio: "ignore", windowsHide: true }).status;
  return {
    ok: status === 0,
    action,
    registered: action !== "unregister" && status === 0,
    ...(action !== "unregister" && status === 0 ? { manifest } : null),
  };
}

function registerRoutes(sidecar, sctx) {
  const bridge = sidecar.bridge;
  const disposers = [
    sctx.webServer.register({
      kind: "exact",
      path: `${bridge}/describe`,
      handler: async (_req, res) => {
        try {
          sendJson(res, 200, describe(sidecar));
        } catch (error) {
          sendJson(res, 500, { ok: false, error: error instanceof Error ? error.message : String(error) });
        }
      },
    }),
    sctx.webServer.register({
      kind: "exact",
      path: `${bridge}/save`,
      handler: async (req, res) => {
        try {
          const body = await readJsonBody(req);
          const operation = sidecar.settingsPromise.catch(() => {}).then(async () => {
            const next = normalizeConfigPatch(sidecar, body);
            writeSettings(sidecar, next);
            await applySettings(sidecar);
          });
          sidecar.settingsPromise = operation;
          await operation;
          sendJson(res, 200, describe(sidecar));
        } catch (error) {
          sendJson(res, 500, { ok: false, error: error instanceof Error ? error.message : String(error) });
        }
      },
    }),
    sctx.webServer.register({
      kind: "exact",
      path: `${bridge}/sideload`,
      handler: async (req, res) => {
        try {
          const body = await readJsonBody(req);
          const action = body && body.action;
          if (action !== "register" && action !== "unregister") {
            sendJson(res, 400, { ok: false, error: "action must be 'register' or 'unregister'" });
            return;
          }
          sendJson(res, 200, runSideload(sidecar, action));
        } catch (error) {
          sendJson(res, 500, { ok: false, error: error instanceof Error ? error.message : String(error) });
        }
      },
    }),
  ];
  return () => {
    for (const dispose of disposers) {
      try {
        dispose?.();
      } catch {
        /* ignore */
      }
    }
  };
}

// ── entry ────────────────────────────────────────────────────────────────────

/**
 * Start the sidecar, mount the settings bridge, register cleanup. Never
 * throws: a failed start is logged and retried by the crash-restart path.
 */
export function runSidecarHost(sidecar, ctx) {
  const state = {
    ...sidecar,
    pluginDir: path.dirname(path.dirname(fileURLToPath(import.meta.url))),
    child: null,
    currentPort: 0,
    running: false,
    pid: null,
    disposed: false,
    restarts: [],
    starting: false,
    startPromise: null,
    settingsPromise: Promise.resolve(),
    restartTimer: null,
    requestedPort: null,
    phase: "stopped",
    lastError: null,
    logs: "",
  };
  publishSkills(state);
  if (readSettings(state).autostart) {
    startServer(state).catch((e) => log(state, `start failed: ${e.message}`));
  } else {
    log(state, "autostart disabled — sidecar stays down until enabled in its settings");
  }

  try {
    ctx.inject(["webServer"], (sctx) => registerRoutes(state, sctx));
  } catch (e) {
    log(state, `settings bridge unavailable: ${e.message}`);
  }

  ctx.effect(
    () => () => {
      state.disposed = true;
      return killChild(state);
    },
    `${sidecar.pluginName}: sidecar process`,
  );
  return state;
}
