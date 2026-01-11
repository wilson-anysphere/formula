import { Counter, Gauge, Histogram, Registry } from "prom-client";
import type { FastifyInstance, FastifyRequest } from "fastify";
import { SpanStatusCode, trace } from "@opentelemetry/api";
import { SEMATTRS_DB_OPERATION, SEMATTRS_DB_SYSTEM } from "@opentelemetry/semantic-conventions";

export type ApiMetrics = {
  registry: Registry;
  httpRequestsTotal: Counter<"method" | "route" | "status">;
  httpRequestDurationSeconds: Histogram<"method" | "route" | "status">;
  dbQueriesTotal: Counter<"operation" | "status">;
  dbQueryDurationSeconds: Histogram<"operation" | "status">;
  authFailuresTotal: Counter<"reason">;
  rateLimitedTotal: Counter<"route" | "reason">;
  dataResidencyBlockedTotal: Counter<"operation">;
  syncTokenIntrospectFailuresTotal: Counter;
  siemBatchesTotal: Counter<"status">;
  siemEventsTotal: Counter<"status">;
  siemBatchDurationSeconds: Histogram;
  siemExportLagSeconds: Gauge;
};

const kReqStart = Symbol("api.metrics.start");
const kDbInstrumented = Symbol("api.db.instrumented");
const kDbRootInstrumented = Symbol("api.db.rootInstrumented");

function routeLabel(request: FastifyRequest): string {
  const route = (request.routeOptions as { url?: string } | undefined)?.url;
  if (typeof route === "string" && route.length > 0) return route;
  return "unknown";
}

function extractDbOperation(args: unknown[]): string {
  const first = args[0] as unknown;
  let text: string | null = null;

  if (typeof first === "string") {
    text = first;
  } else if (first && typeof first === "object" && "text" in first && typeof (first as any).text === "string") {
    text = (first as any).text as string;
  }

  if (!text) return "unknown";
  const match = text.trim().match(/^([a-zA-Z]+)/);
  return match ? match[1]!.toUpperCase() : "unknown";
}

export function createMetrics(): ApiMetrics {
  const registry = new Registry();

  const httpRequestsTotal = new Counter({
    name: "http_requests_total",
    help: "HTTP requests processed by the API",
    labelNames: ["method", "route", "status"],
    registers: [registry]
  });

  const httpRequestDurationSeconds = new Histogram({
    name: "http_request_duration_seconds",
    help: "HTTP request latency (seconds)",
    labelNames: ["method", "route", "status"],
    buckets: [0.005, 0.01, 0.025, 0.05, 0.1, 0.25, 0.5, 1, 2.5, 5, 10],
    registers: [registry]
  });

  const dbQueriesTotal = new Counter({
    name: "db_queries_total",
    help: "Database queries executed by the API",
    labelNames: ["operation", "status"],
    registers: [registry]
  });

  const dbQueryDurationSeconds = new Histogram({
    name: "db_query_duration_seconds",
    help: "Database query latency (seconds)",
    labelNames: ["operation", "status"],
    buckets: [0.001, 0.0025, 0.005, 0.01, 0.025, 0.05, 0.1, 0.25, 0.5, 1, 2.5],
    registers: [registry]
  });

  const authFailuresTotal = new Counter({
    name: "auth_failures_total",
    help: "Authentication failures",
    labelNames: ["reason"],
    registers: [registry]
  });

  const rateLimitedTotal = new Counter({
    name: "rate_limited_total",
    help: "Requests rejected by API rate limiting",
    labelNames: ["route", "reason"],
    registers: [registry]
  });

  const dataResidencyBlockedTotal = new Counter({
    name: "data_residency_blocked_total",
    help: "Operations blocked by data residency policy",
    labelNames: ["operation"],
    registers: [registry]
  });

  const syncTokenIntrospectFailuresTotal = new Counter({
    name: "sync_token_introspect_failures_total",
    help: "Sync token introspection failures",
    registers: [registry]
  });

  const siemBatchesTotal = new Counter({
    name: "siem_batches_total",
    help: "SIEM export batches processed",
    labelNames: ["status"],
    registers: [registry]
  });

  const siemEventsTotal = new Counter({
    name: "siem_events_total",
    help: "SIEM export events processed",
    labelNames: ["status"],
    registers: [registry]
  });

  const siemBatchDurationSeconds = new Histogram({
    name: "siem_batch_duration_seconds",
    help: "SIEM export batch duration (seconds)",
    buckets: [0.01, 0.025, 0.05, 0.1, 0.25, 0.5, 1, 2.5, 5, 10, 30],
    registers: [registry]
  });

  const siemExportLagSeconds = new Gauge({
    name: "siem_export_lag_seconds",
    help: "Approximate SIEM export lag (seconds)",
    registers: [registry]
  });

  return {
    registry,
    httpRequestsTotal,
    httpRequestDurationSeconds,
    dbQueriesTotal,
    dbQueryDurationSeconds,
    authFailuresTotal,
    rateLimitedTotal,
    dataResidencyBlockedTotal,
    syncTokenIntrospectFailuresTotal,
    siemBatchesTotal,
    siemEventsTotal,
    siemBatchDurationSeconds,
    siemExportLagSeconds
  };
}

export function registerMetrics(app: FastifyInstance, metrics: ApiMetrics): void {
  app.get("/metrics", async (_request, reply) => {
    reply.header("content-type", metrics.registry.contentType);
    return metrics.registry.metrics();
  });

  app.addHook("onRequest", (request, _reply, done) => {
    (request as any)[kReqStart] = process.hrtime.bigint();
    done();
  });

  app.addHook("onResponse", (request, reply, done) => {
    const start = (request as any)[kReqStart] as bigint | undefined;
    if (!start) return done();

    const durationSeconds = Number(process.hrtime.bigint() - start) / 1e9;
    const labels = {
      method: request.method,
      route: routeLabel(request),
      status: String(reply.statusCode)
    };

    metrics.httpRequestsTotal.inc(labels);
    metrics.httpRequestDurationSeconds.observe(labels, durationSeconds);
    done();
  });
}

export function instrumentDb<T extends { query: (...args: any[]) => any }>(
  db: T,
  metrics: ApiMetrics
): T {
  if ((db as any)[kDbRootInstrumented]) return db;
  (db as any)[kDbRootInstrumented] = true;

  const tracer = trace.getTracer("api.db");

  const patchQueryable = (queryable: any) => {
    if (!queryable || (queryable as any)[kDbInstrumented]) return;
    (queryable as any)[kDbInstrumented] = true;

    const originalQuery = queryable.query?.bind(queryable);
    if (typeof originalQuery !== "function") return;

    queryable.query = async (...args: any[]) => {
      const operation = extractDbOperation(args);
      const start = process.hrtime.bigint();

      return tracer.startActiveSpan(
        "db.query",
        {
          attributes: {
            [SEMATTRS_DB_SYSTEM]: "postgresql",
            [SEMATTRS_DB_OPERATION]: operation
          }
        },
        async (span) => {
          try {
            const result = await originalQuery(...args);

            const durationSeconds = Number(process.hrtime.bigint() - start) / 1e9;
            metrics.dbQueriesTotal.inc({ operation, status: "ok" });
            metrics.dbQueryDurationSeconds.observe({ operation, status: "ok" }, durationSeconds);

            span.setStatus({ code: SpanStatusCode.OK });
            return result;
          } catch (err) {
            const durationSeconds = Number(process.hrtime.bigint() - start) / 1e9;
            metrics.dbQueriesTotal.inc({ operation, status: "error" });
            metrics.dbQueryDurationSeconds.observe({ operation, status: "error" }, durationSeconds);

            span.recordException(err as Error);
            span.setStatus({ code: SpanStatusCode.ERROR });
            throw err;
          } finally {
            span.end();
          }
        }
      );
    };
  };

  patchQueryable(db);

  const originalConnect = (db as any).connect?.bind(db);
  if (typeof originalConnect === "function") {
    (db as any).connect = async (...args: any[]) => {
      const client = await originalConnect(...args);
      patchQueryable(client);
      return client;
    };
  }

  return db;
}
