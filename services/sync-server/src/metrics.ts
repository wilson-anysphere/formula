import { createRequire } from "node:module";

/**
 * `prom-client` is used for exposing Prometheus metrics in production, but we
 * want sync-server to remain runnable in lightweight/dev environments where the
 * dependency may not be installed (e.g. minimal agent sandboxes).
 *
 * `tsx` (used for dev/test) transpiles without typechecking, so a hard runtime
 * import would crash the server. Instead we best-effort load `prom-client` and
 * fall back to a no-op metrics implementation.
 */

type Counter<Labels extends string = string> = {
  inc: (labels?: Record<Labels, string>, value?: number) => void;
};

type Gauge<Labels extends string = string> = {
  set: (labelsOrValue: Record<Labels, string> | number, value?: number) => void;
  reset: () => void;
};

type Registry = {
  contentType: string;
  setDefaultLabels: (labels: Record<string, string>) => void;
  metrics: () => Promise<string>;
};

function loadPromClient(): {
  Registry: new () => Registry;
  Counter: new (opts: any) => Counter<any>;
  Gauge: new (opts: any) => Gauge<any>;
} | null {
  // prom-client is a CommonJS package. Using createRequire avoids ESM/CJS interop
  // edge cases (e.g. Vite/Vitest SSR) when sync-server source is imported from
  // other workspaces.
  const require = createRequire(import.meta.url);
  try {
    return require("prom-client");
  } catch (err) {
    const code = (err as NodeJS.ErrnoException)?.code;
    if (code === "ERR_MODULE_NOT_FOUND" || code === "MODULE_NOT_FOUND") return null;
    return null;
  }
}

function createNoopMetrics(): SyncServerMetrics {
  const noopCounter: Counter<string> = {
    inc: () => {
      // noop
    },
  };
  const noopGauge: Gauge<string> = {
    set: () => {
      // noop
    },
    reset: () => {
      // noop
    },
  };
  const registry: Registry = {
    // Match prom-client default.
    contentType: "text/plain; version=0.0.4; charset=utf-8",
    setDefaultLabels: () => {
      // noop
    },
    metrics: async () => "",
  };

  return {
    registry,
    wsConnectionsTotal: noopCounter,
    wsConnectionsCurrent: noopGauge,
    wsConnectionsRejectedTotal: noopCounter as Counter<"reason">,
    wsClosesTotal: noopCounter as Counter<"code">,
    wsMessagesRateLimitedTotal: noopCounter,
    wsMessagesTooLargeTotal: noopCounter,
    retentionSweepsTotal: noopCounter as Counter<"sweep">,
    retentionDocsPurgedTotal: noopCounter as Counter<"sweep">,
    retentionSweepErrorsTotal: noopCounter as Counter<"sweep">,
    persistenceInfo: noopGauge as Gauge<"backend" | "encryption">,
    setPersistenceInfo: () => {
      // noop
    },
    metricsText: async () => "",
  };
}

export type WsConnectionRejectionReason =
  | "rate_limit"
  | "auth_failure"
  | "tombstone"
  | "retention_purging";

export type RetentionSweepKind = "leveldb" | "tombstone";

export type SyncServerMetrics = {
  registry: Registry;

  wsConnectionsTotal: Counter<string>;
  wsConnectionsCurrent: Gauge<string>;
  wsConnectionsRejectedTotal: Counter<"reason">;

  wsClosesTotal: Counter<"code">;

  wsMessagesRateLimitedTotal: Counter<string>;
  wsMessagesTooLargeTotal: Counter<string>;

  retentionSweepsTotal: Counter<"sweep">;
  retentionDocsPurgedTotal: Counter<"sweep">;
  retentionSweepErrorsTotal: Counter<"sweep">;

  persistenceInfo: Gauge<"backend" | "encryption">;
  setPersistenceInfo: (params: {
    backend: "file" | "leveldb";
    encryptionEnabled: boolean;
  }) => void;

  metricsText: () => Promise<string>;
};

export function createSyncServerMetrics(): SyncServerMetrics {
  const promClient = loadPromClient();
  if (!promClient) return createNoopMetrics();

  const registry = new promClient.Registry();
  registry.setDefaultLabels({ service: "sync-server" });

  const wsConnectionsTotal = new promClient.Counter({
    name: "sync_server_ws_connections_total",
    help: "Total accepted WebSocket connections.",
    registers: [registry],
  });

  const wsConnectionsCurrent = new promClient.Gauge({
    name: "sync_server_ws_connections_current",
    help: "Current active WebSocket connections.",
    registers: [registry],
  });
  wsConnectionsCurrent.set(0);

  const wsConnectionsRejectedTotal = new promClient.Counter({
    name: "sync_server_ws_connections_rejected_total",
    help: "Total rejected WebSocket upgrade attempts.",
    labelNames: ["reason"],
    registers: [registry],
  });
  // Pre-initialize the known rejection reasons so dashboards don't need to handle
  // missing series vs. a literal 0 value.
  wsConnectionsRejectedTotal.inc({ reason: "rate_limit" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "auth_failure" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "tombstone" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "retention_purging" }, 0);

  const wsClosesTotal = new promClient.Counter({
    name: "sync_server_ws_closes_total",
    help: "Total WebSocket close events by close code.",
    labelNames: ["code"],
    registers: [registry],
  });
  // Pre-initialize common codes used by the server.
  for (const code of ["1000", "1003", "1006", "1008", "1009", "1011", "1013", "other"]) {
    wsClosesTotal.inc({ code }, 0);
  }

  const wsMessagesRateLimitedTotal = new promClient.Counter({
    name: "sync_server_ws_messages_rate_limited_total",
    help: "Total WebSocket messages rejected due to message rate limits.",
    registers: [registry],
  });

  const wsMessagesTooLargeTotal = new promClient.Counter({
    name: "sync_server_ws_messages_too_large_total",
    help: "Total WebSocket messages rejected due to message size limits.",
    registers: [registry],
  });

  const retentionSweepsTotal = new promClient.Counter({
    name: "sync_server_retention_sweeps_total",
    help: "Total retention sweeps executed.",
    labelNames: ["sweep"],
    registers: [registry],
  });
  retentionSweepsTotal.inc({ sweep: "leveldb" }, 0);
  retentionSweepsTotal.inc({ sweep: "tombstone" }, 0);

  const retentionDocsPurgedTotal = new promClient.Counter({
    name: "sync_server_retention_docs_purged_total",
    help: "Total documents purged by retention sweeps.",
    labelNames: ["sweep"],
    registers: [registry],
  });
  retentionDocsPurgedTotal.inc({ sweep: "leveldb" }, 0);
  retentionDocsPurgedTotal.inc({ sweep: "tombstone" }, 0);

  const retentionSweepErrorsTotal = new promClient.Counter({
    name: "sync_server_retention_sweep_errors_total",
    help: "Total retention sweep errors.",
    labelNames: ["sweep"],
    registers: [registry],
  });
  retentionSweepErrorsTotal.inc({ sweep: "leveldb" }, 0);
  retentionSweepErrorsTotal.inc({ sweep: "tombstone" }, 0);

  const persistenceInfo = new promClient.Gauge({
    name: "sync_server_persistence_info",
    help:
      "Persistence backend and at-rest encryption configuration (set to 1 for the active config).",
    labelNames: ["backend", "encryption"],
    registers: [registry],
  });

  const setPersistenceInfo = (params: {
    backend: "file" | "leveldb";
    encryptionEnabled: boolean;
  }) => {
    persistenceInfo.reset();
    persistenceInfo.set(
      {
        backend: params.backend,
        encryption: params.encryptionEnabled ? "on" : "off",
      },
      1
    );
  };

  return {
    registry,
    wsConnectionsTotal,
    wsConnectionsCurrent,
    wsConnectionsRejectedTotal,
    wsClosesTotal,
    wsMessagesRateLimitedTotal,
    wsMessagesTooLargeTotal,
    retentionSweepsTotal,
    retentionDocsPurgedTotal,
    retentionSweepErrorsTotal,
    persistenceInfo,
    setPersistenceInfo,
    metricsText: async () => await registry.metrics(),
  };
}
