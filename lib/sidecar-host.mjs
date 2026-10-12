/**
 * Sidecar host for dsh-office-for-mso.
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

import { spawn } from "node:child_process";
import fs from "node:fs";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

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
  for await (const chunk of req) chunks.push(chunk);
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
  } catch {
    return {};
  }
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
      try {
        const res = await fetch(`http://127.0.0.1:${port}${sidecar.healthPath}`);
        if (res.ok) return resolve(true);
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
  if (!sidecar.child) return false;
  const pid = sidecar.child.pid;
  sidecar.child = null;
  sidecar.running = false;
  if (!pid) return false;
  if (process.platform === "win32") {
    spawn("taskkill", ["/F", "/T", "/PID", String(pid)], { stdio: "ignore", windowsHide: true });
  } else {
    try {
      process.kill(pid, "SIGTERM");
    } catch {
      // already gone
    }
  }
  return true;
}

async function startServer(sidecar) {
  if (sidecar.starting) return;
  sidecar.starting = true;
  try {
    const port = await pickPort(sidecar);
    sidecar.currentPort = port;
    sidecar.child = spawn(process.execPath, [path.join(sidecar.pluginDir, sidecar.entry)], {
      env: { ...process.env, PORT: String(port) },
      stdio: "ignore",
      windowsHide: true,
    });
    sidecar.running = true;
    sidecar.pid = sidecar.child.pid;
    log(sidecar, `sidecar pid ${sidecar.child.pid} on 127.0.0.1:${port}`);

    sidecar.child.on("exit", (code) => {
      if (sidecar.disposed || sidecar.child === null) return;
      sidecar.running = false;
      sidecar.pid = null;
      log(sidecar, `sidecar exited (code ${code}); restarting`);
      const now = Date.now();
      sidecar.restarts = sidecar.restarts.filter((t) => now - t < 60_000);
      sidecar.restarts.push(now);
      if (sidecar.restarts.length > RESTART_BUDGET) {
        log(sidecar, "restart budget exhausted for this minute; giving up until next boot");
        return;
      }
      setTimeout(() => {
        if (!sidecar.disposed && readSettings(sidecar).autostart) {
          startServer(sidecar).catch((e) => log(sidecar, `restart failed: ${e.message}`));
        }
      }, RESTART_DELAY_MS);
    });

    try {
      await waitHealth(sidecar, port);
    } catch (e) {
      log(sidecar, `sidecar unhealthy: ${e.message}`);
    }
  } finally {
    sidecar.starting = false;
  }
}

function stopServer(sidecar) {
  killChild(sidecar);
}

/** Apply settings: (re)start or stop the sidecar to match the saved config. */
async function applySettings(sidecar) {
  const settings = readSettings(sidecar);
  const wasRunning = sidecar.running;
  if (!settings.autostart) {
    if (wasRunning) {
      stopServer(sidecar);
      log(sidecar, "sidecar stopped (autostart disabled)");
    }
    return;
  }
  const portChanged = wasRunning && settings.port !== null && settings.port !== sidecar.currentPort;
  if (wasRunning && !portChanged) return;
  if (wasRunning) stopServer(sidecar);
  // Let the old process release the port before the replacement probes it.
  await new Promise((resolve) => setTimeout(resolve, 300));
  await startServer(sidecar);
}

// ── skills ───────────────────────────────────────────────────────────────────

function publishSkills(sidecar) {
  const skillsDir = path.join(sidecar.pluginDir, "skills");
  const targetRoot = path.join(os.homedir(), ".agents", "skills");
  try {
    for (const name of fs.readdirSync(skillsDir)) {
      const src = path.join(skillsDir, name);
      if (!fs.existsSync(path.join(src, "SKILL.md"))) continue;
      const dest = path.join(targetRoot, name);
      const skillMd = path.join(dest, "SKILL.md");
      if (fs.existsSync(skillMd) && fs.readFileSync(skillMd, "utf8") === fs.readFileSync(path.join(src, "SKILL.md"), "utf8")) {
        continue;
      }
      fs.mkdirSync(dest, { recursive: true });
      fs.cpSync(src, dest, { recursive: true });
      log(sidecar, `skill published → ~/.agents/skills/${name}`);
    }
  } catch (e) {
    log(sidecar, `skill publishing skipped: ${e.message}`);
  }
}

// ── settings bridge ──────────────────────────────────────────────────────────

function statusView(sidecar) {
  return {
    running: Boolean(sidecar.running),
    port: sidecar.running ? sidecar.currentPort : null,
    pid: sidecar.running ? sidecar.pid : null,
  };
}

function describe(sidecar) {
  return { ok: true, status: statusView(sidecar), config: readSettings(sidecar) };
}

function normalizeConfigPatch(sidecar, body) {
  const patch = body?.config && typeof body.config === "object" ? body.config : {};
  const next = readSettings(sidecar);
  if ("port" in patch) {
    const raw = patch.port;
    const n = Number(raw);
    if (raw === null || raw === "" || raw === 0 || raw === "0") next.port = null;
    else if (!Number.isInteger(n) || n < 1024 || n > 65535) {
      throw new Error("port must be 1024–65535, or empty for the default");
    } else next.port = n;
  }
  if ("autostart" in patch) next.autostart = Boolean(patch.autostart);
  return next;
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
          const next = normalizeConfigPatch(sidecar, body);
          writeSettings(sidecar, next);
          await applySettings(sidecar);
          sendJson(res, 200, describe(sidecar));
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
      killChild(state);
    },
    `${sidecar.pluginName}: sidecar process`,
  );
  return state;
}
