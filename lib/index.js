// Host half of dsh-ms-office: owns the bridge sidecar's whole lifecycle
// (spawn, health, port selection, restart, dispose) and mounts the small
// settings surface the client half renders on the Plugins page. The mechanics
// live in ./sidecar-host.mjs.
import { runSidecarHost } from "./sidecar-host.mjs";

export const inject = [];

export function apply(ctx) {
  runSidecarHost(
    {
      pluginName: "dsh-ms-office",
      defaultPort: 3000,
      healthPath: "/office/status",
      entry: "server/server.js",
      bridge: "/api/ms-office",
    },
    ctx,
  );
}
