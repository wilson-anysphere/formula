import promClient from "prom-client";

export type WsConnectionRejectionReason =
  | "rate_limit"
  | "auth_failure"
  | "tombstone"
  | "retention_purging";

export type RetentionSweepKind = "leveldb" | "tombstone";

export type SyncServerMetrics = {
  registry: promClient.Registry;

  wsConnectionsTotal: promClient.Counter<string>;
  wsConnectionsCurrent: promClient.Gauge<string>;
  wsConnectionsRejectedTotal: promClient.Counter<"reason">;

  wsMessagesRateLimitedTotal: promClient.Counter<string>;
  wsMessagesTooLargeTotal: promClient.Counter<string>;

  retentionSweepsTotal: promClient.Counter<"sweep">;
  retentionDocsPurgedTotal: promClient.Counter<"sweep">;
  retentionSweepErrorsTotal: promClient.Counter<"sweep">;

  persistenceInfo: promClient.Gauge<"backend" | "encryption">;
  setPersistenceInfo: (params: {
    backend: "file" | "leveldb";
    encryptionEnabled: boolean;
  }) => void;

  metricsText: () => Promise<string>;
};

export function createSyncServerMetrics(): SyncServerMetrics {
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
