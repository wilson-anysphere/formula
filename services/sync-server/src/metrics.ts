import { createRequire } from "node:module";

/**
 * `prom-client` is used for exposing Prometheus metrics in production, but we
 * want sync-server to remain runnable in lightweight/dev environments where the
 * dependency may not be installed (e.g. minimal agent sandboxes).
 *
 * `tsx` (used for dev/test) transpiles without typechecking, so a hard runtime
 * import would crash the server. Instead we best-effort load `prom-client` and
 * fall back to a small in-memory implementation.
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

type MetricKind = "counter" | "gauge";

type Sample = {
  labels: Record<string, string>;
  value: number;
};

function escapeLabelValue(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/"/g, '\\"');
}

function formatLabels(labels: Record<string, string>): string {
  const keys = Object.keys(labels);
  if (keys.length === 0) return "";
  keys.sort();
  const parts = keys.map((key) => `${key}="${escapeLabelValue(labels[key] ?? "")}"`);
  return `{${parts.join(",")}}`;
}

function sampleKey(labels: Record<string, string>): string {
  const keys = Object.keys(labels);
  if (keys.length === 0) return "";
  keys.sort();
  // Use a delimiter that's unlikely to appear in label values.
  return keys.map((key) => `${key}\u0000${labels[key] ?? ""}`).join("\u0001");
}

type SimpleMetric = {
  name: string;
  help: string;
  kind: MetricKind;
  samples: Map<string, Sample>;
};

function renderMetric(metric: SimpleMetric, defaultLabels: Record<string, string>): string {
  const lines: string[] = [];
  lines.push(`# HELP ${metric.name} ${metric.help}`);
  lines.push(`# TYPE ${metric.name} ${metric.kind}`);
  for (const sample of metric.samples.values()) {
    lines.push(
      `${metric.name}${formatLabels({ ...defaultLabels, ...sample.labels })} ${sample.value}`
    );
  }
  return lines.join("\n");
}

function createSimpleRegistry(): Registry & {
  registerMetric: (metric: SimpleMetric) => void;
} {
  const metrics: SimpleMetric[] = [];
  const defaultLabels: Record<string, string> = {};
  return {
    contentType: "text/plain; version=0.0.4; charset=utf-8",
    setDefaultLabels: (labels) => {
      Object.assign(defaultLabels, labels);
    },
    registerMetric: (metric) => {
      metrics.push(metric);
    },
    metrics: async () => {
      if (metrics.length === 0) return "";
      const body = metrics.map((metric) => renderMetric(metric, defaultLabels)).join("\n");
      return `${body}\n`;
    },
  };
}

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

function createFallbackMetrics(): SyncServerMetrics {
  const registry = createSimpleRegistry();
  registry.setDefaultLabels({ service: "sync-server" });

  const createCounter = <Labels extends string>(opts: {
    name: string;
    help: string;
    labelNames?: Labels[];
  }): Counter<Labels> => {
    const labelNames = (opts.labelNames ?? []) as string[];
    const metric: SimpleMetric = {
      name: opts.name,
      help: opts.help,
      kind: "counter",
      samples: new Map(),
    };
    registry.registerMetric(metric);

    // Match prom-client behavior: unlabelled metrics show up as `0` even before
    // they're incremented.
    if (labelNames.length === 0) {
      metric.samples.set("", { labels: {}, value: 0 });
    }

    const inc: Counter<Labels>["inc"] = (labels, value) => {
      const delta = typeof value === "number" ? value : 1;

      const normalizedLabels: Record<string, string> = {};
      const raw = (labels ?? {}) as Record<string, string>;
      for (const labelName of labelNames) {
        const labelValue = raw[labelName];
        if (typeof labelValue !== "string" || labelValue.length === 0) {
          return;
        }
        normalizedLabels[labelName] = labelValue;
      }

      const key = sampleKey(normalizedLabels);
      const current = metric.samples.get(key) ?? { labels: normalizedLabels, value: 0 };
      metric.samples.set(key, {
        labels: normalizedLabels,
        value: current.value + delta,
      });
    };

    return { inc };
  };

  const createGauge = <Labels extends string>(opts: {
    name: string;
    help: string;
    labelNames?: Labels[];
  }): Gauge<Labels> => {
    const labelNames = (opts.labelNames ?? []) as string[];
    const metric: SimpleMetric = {
      name: opts.name,
      help: opts.help,
      kind: "gauge",
      samples: new Map(),
    };
    registry.registerMetric(metric);

    const set: Gauge<Labels>["set"] = (labelsOrValue, value) => {
      if (typeof labelsOrValue === "number") {
        if (labelNames.length > 0) return;
        metric.samples.set("", { labels: {}, value: labelsOrValue });
        return;
      }

      const normalizedLabels: Record<string, string> = {};
      const raw = labelsOrValue as Record<string, string>;
      for (const labelName of labelNames) {
        const labelValue = raw[labelName];
        if (typeof labelValue !== "string" || labelValue.length === 0) {
          return;
        }
        normalizedLabels[labelName] = labelValue;
      }

      const v = typeof value === "number" ? value : 0;
      metric.samples.set(sampleKey(normalizedLabels), { labels: normalizedLabels, value: v });
    };

    const reset: Gauge<Labels>["reset"] = () => {
      metric.samples.clear();
    };

    return { set, reset };
  };

  const wsConnectionsTotal = createCounter({
    name: "sync_server_ws_connections_total",
    help: "Total accepted WebSocket connections.",
  });

  const wsConnectionsCurrent = createGauge({
    name: "sync_server_ws_connections_current",
    help: "Current active WebSocket connections.",
  });
  wsConnectionsCurrent.set(0);

  const wsConnectionsRejectedTotal = createCounter<"reason">({
    name: "sync_server_ws_connections_rejected_total",
    help: "Total rejected WebSocket upgrade attempts.",
    labelNames: ["reason"],
  });
  wsConnectionsRejectedTotal.inc({ reason: "rate_limit" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "auth_failure" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "tombstone" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "retention_purging" }, 0);
  wsConnectionsRejectedTotal.inc({ reason: "max_connections_per_doc" }, 0);

  const wsClosesTotal = createCounter<"code">({
    name: "sync_server_ws_closes_total",
    help: "Total WebSocket close events by close code.",
    labelNames: ["code"],
  });
  // Pre-initialize common codes used by the server.
  for (const code of ["1000", "1003", "1006", "1008", "1009", "1011", "1013", "other"]) {
    wsClosesTotal.inc({ code }, 0);
  }

  const wsMessagesRateLimitedTotal = createCounter({
    name: "sync_server_ws_messages_rate_limited_total",
    help: "Total WebSocket messages rejected due to message rate limits.",
  });

  const wsMessagesTooLargeTotal = createCounter({
    name: "sync_server_ws_messages_too_large_total",
    help: "Total WebSocket messages rejected due to message size limits.",
  });

  const wsReservedRootQuotaViolationsTotal = createCounter<"kind">({
    name: "sync_server_ws_reserved_root_quota_violations_total",
    help:
      "Total WebSocket messages rejected due to reserved-root history growth quotas.",
    labelNames: ["kind"],
  });
  wsReservedRootQuotaViolationsTotal.inc({ kind: "branching_commits" }, 0);
  wsReservedRootQuotaViolationsTotal.inc({ kind: "versions" }, 0);

  const retentionSweepsTotal = createCounter<"sweep">({
    name: "sync_server_retention_sweeps_total",
    help: "Total retention sweeps executed.",
    labelNames: ["sweep"],
  });
  retentionSweepsTotal.inc({ sweep: "leveldb" }, 0);
  retentionSweepsTotal.inc({ sweep: "tombstone" }, 0);

  const retentionDocsPurgedTotal = createCounter<"sweep">({
    name: "sync_server_retention_docs_purged_total",
    help: "Total documents purged by retention sweeps.",
    labelNames: ["sweep"],
  });
  retentionDocsPurgedTotal.inc({ sweep: "leveldb" }, 0);
  retentionDocsPurgedTotal.inc({ sweep: "tombstone" }, 0);

  const retentionSweepErrorsTotal = createCounter<"sweep">({
    name: "sync_server_retention_sweep_errors_total",
    help: "Total retention sweep errors.",
    labelNames: ["sweep"],
  });
  retentionSweepErrorsTotal.inc({ sweep: "leveldb" }, 0);
  retentionSweepErrorsTotal.inc({ sweep: "tombstone" }, 0);

  const processResidentMemoryBytes = createGauge({
    name: "sync_server_process_resident_memory_bytes",
    help: "Resident set size (RSS) memory used by the process in bytes.",
  });
  processResidentMemoryBytes.set(0);

  const processHeapUsedBytes = createGauge({
    name: "sync_server_process_heap_used_bytes",
    help: "V8 heap used in bytes.",
  });
  processHeapUsedBytes.set(0);

  const processHeapTotalBytes = createGauge({
    name: "sync_server_process_heap_total_bytes",
    help: "V8 heap total size in bytes.",
  });
  processHeapTotalBytes.set(0);

  const eventLoopDelayMs = createGauge({
    name: "sync_server_event_loop_delay_ms",
    help: "Event loop delay (p99) sampled over the last collection interval, in milliseconds.",
  });
  eventLoopDelayMs.set(0);

  const persistenceInfo = createGauge<"backend" | "encryption">({
    name: "sync_server_persistence_info",
    help:
      "Persistence backend and at-rest encryption configuration (set to 1 for the active config).",
    labelNames: ["backend", "encryption"],
  });

  const setPersistenceInfo: SyncServerMetrics["setPersistenceInfo"] = (params) => {
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
    wsReservedRootQuotaViolationsTotal,
    retentionSweepsTotal,
    retentionDocsPurgedTotal,
    retentionSweepErrorsTotal,
    processResidentMemoryBytes,
    processHeapUsedBytes,
    processHeapTotalBytes,
    eventLoopDelayMs,
    persistenceInfo,
    setPersistenceInfo,
    metricsText: async () => await registry.metrics(),
  };
}

export type WsConnectionRejectionReason =
  | "rate_limit"
  | "auth_failure"
  | "tombstone"
  | "retention_purging"
  | "max_connections_per_doc";

export type RetentionSweepKind = "leveldb" | "tombstone";

export type SyncServerMetrics = {
  registry: Registry;

  wsConnectionsTotal: Counter<string>;
  wsConnectionsCurrent: Gauge<string>;
  wsConnectionsRejectedTotal: Counter<"reason">;

  wsClosesTotal: Counter<"code">;

  wsMessagesRateLimitedTotal: Counter<string>;
  wsMessagesTooLargeTotal: Counter<string>;
  wsReservedRootQuotaViolationsTotal: Counter<"kind">;

  retentionSweepsTotal: Counter<"sweep">;
  retentionDocsPurgedTotal: Counter<"sweep">;
  retentionSweepErrorsTotal: Counter<"sweep">;

  processResidentMemoryBytes: Gauge<string>;
  processHeapUsedBytes: Gauge<string>;
  processHeapTotalBytes: Gauge<string>;
  eventLoopDelayMs: Gauge<string>;

  persistenceInfo: Gauge<"backend" | "encryption">;
  setPersistenceInfo: (params: {
    backend: "file" | "leveldb";
    encryptionEnabled: boolean;
  }) => void;

  metricsText: () => Promise<string>;
};

export function createSyncServerMetrics(): SyncServerMetrics {
  const promClient = loadPromClient();
  if (!promClient) return createFallbackMetrics();

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
  wsConnectionsRejectedTotal.inc({ reason: "max_connections_per_doc" }, 0);

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

  const wsReservedRootQuotaViolationsTotal = new promClient.Counter({
    name: "sync_server_ws_reserved_root_quota_violations_total",
    help:
      "Total WebSocket messages rejected due to reserved-root history growth quotas.",
    labelNames: ["kind"],
    registers: [registry],
  });
  wsReservedRootQuotaViolationsTotal.inc({ kind: "branching_commits" }, 0);
  wsReservedRootQuotaViolationsTotal.inc({ kind: "versions" }, 0);

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

  const processResidentMemoryBytes = new promClient.Gauge({
    name: "sync_server_process_resident_memory_bytes",
    help: "Resident set size (RSS) memory used by the process in bytes.",
    registers: [registry],
  });
  processResidentMemoryBytes.set(0);

  const processHeapUsedBytes = new promClient.Gauge({
    name: "sync_server_process_heap_used_bytes",
    help: "V8 heap used in bytes.",
    registers: [registry],
  });
  processHeapUsedBytes.set(0);

  const processHeapTotalBytes = new promClient.Gauge({
    name: "sync_server_process_heap_total_bytes",
    help: "V8 heap total size in bytes.",
    registers: [registry],
  });
  processHeapTotalBytes.set(0);

  const eventLoopDelayMs = new promClient.Gauge({
    name: "sync_server_event_loop_delay_ms",
    help: "Event loop delay (p99) sampled over the last collection interval, in milliseconds.",
    registers: [registry],
  });
  eventLoopDelayMs.set(0);

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
    wsReservedRootQuotaViolationsTotal,
    retentionSweepsTotal,
    retentionDocsPurgedTotal,
    retentionSweepErrorsTotal,
    processResidentMemoryBytes,
    processHeapUsedBytes,
    processHeapTotalBytes,
    eventLoopDelayMs,
    persistenceInfo,
    setPersistenceInfo,
    metricsText: async () => await registry.metrics(),
  };
}
