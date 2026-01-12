import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { Writable } from "node:stream";
import { trace } from "@opentelemetry/api";
import Fastify from "fastify";
import { createLogger } from "../observability/logger";
import { createMetrics, instrumentDb } from "../observability/metrics";
import { initOpenTelemetry } from "../observability/otel";
import { genRequestId, registerRequestId } from "../observability/request-id";
import { InMemorySpanExporter } from "@opentelemetry/sdk-trace-base";

describe("observability: request-id, log correlation, db spans", () => {
  let db: Pool;
  let app: any;
  const logs: any[] = [];
  const exporter = new InMemorySpanExporter();
  let otelShutdown: (() => Promise<void>) | null = null;

  beforeAll(async () => {
    // This suite intentionally builds a minimal Fastify instance (instead of the
    // full API app) to keep `beforeAll` fast and avoid hook timeouts under
    // full-suite Vitest contention.
    otelShutdown = initOpenTelemetry({ serviceName: "api-test", spanExporter: exporter }).shutdown;

    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();

    const metrics = createMetrics();
    instrumentDb(db, metrics);

    let buffer = "";
    const stream = new Writable({
      write(chunk, _encoding, callback) {
        buffer += chunk.toString();
        let idx = buffer.indexOf("\n");
        while (idx !== -1) {
          const line = buffer.slice(0, idx).trim();
          buffer = buffer.slice(idx + 1);
          if (line) {
            try {
              logs.push(JSON.parse(line));
            } catch {
              // Defensive: if logging output is ever non-JSON, keep it for debugging.
              logs.push(line);
            }
          }
          idx = buffer.indexOf("\n");
        }
        callback();
      }
    });

    const logger = createLogger({ level: "info", stream });

    app = Fastify({
      loggerInstance: logger as any,
      genReqId: genRequestId,
      requestIdLogLabel: "requestId"
    });
    app.decorate("db", db);
    registerRequestId(app);

    app.get("/health", async () => ({ status: "ok" }));
    app.get("/_test/log", async (request: any) => {
      request.log.info("test_log");
      return { ok: true };
    });

    app.get("/_test/db", async (request: any) => {
      await request.server.db.query("SELECT 1");
      request.log.info("test_db_query_done");
      return { ok: true };
    });

    await app.ready();
  });

  afterAll(async () => {
    await app?.close?.();
    await db?.end?.();
    await otelShutdown?.();
  });

  it("sets x-request-id and respects incoming x-request-id", async () => {
    const res = await app.inject({ method: "GET", url: "/health" });
    expect(res.statusCode).toBe(200);
    expect(res.headers["x-request-id"]).toBeTypeOf("string");

    const res2 = await app.inject({
      method: "GET",
      url: "/health",
      headers: { "x-request-id": "client-request-id-123" }
    });
    expect(res2.statusCode).toBe(200);
    expect(res2.headers["x-request-id"]).toBe("client-request-id-123");
  });

  it("includes traceId/spanId/requestId in logs", async () => {
    logs.length = 0;
    exporter.reset();

    const tracer = trace.getTracer("api-test");
    const { requestId, spanContext } = await tracer.startActiveSpan("test.request", async (span) => {
      try {
        const res = await app.inject({ method: "GET", url: "/_test/log" });
        return { requestId: res.headers["x-request-id"], spanContext: span.spanContext() };
      } finally {
        span.end();
      }
    });

    // Pino writes asynchronously; in contended runs the log line may arrive a
    // tick later. Poll briefly to avoid flakes.
    const deadline = Date.now() + 1000;
    let entry: any | undefined;
    while (!entry && Date.now() < deadline) {
      entry = logs.find((line) => typeof line === "object" && line && line.msg === "test_log");
      if (!entry) await new Promise((resolve) => setTimeout(resolve, 5));
    }

    expect(entry).toBeTruthy();
    expect(entry.requestId).toBe(requestId);
    expect(entry.traceId).toBe(spanContext.traceId);
    // The active span during route execution may be the span we started here,
    // or (if request instrumentation is active) a child span created by the
    // Fastify/HTTP instrumentations. Either way, logs should include a valid
    // span id correlated to the same trace.
    expect(entry.spanId).toMatch(/^[0-9a-f]{16}$/);
  });

  it("creates db.query spans", async () => {
    logs.length = 0;
    exporter.reset();

    const res = await app.inject({ method: "GET", url: "/_test/db" });
    expect(res.statusCode).toBe(200);

    const spans = exporter.getFinishedSpans();
    expect(spans.some((span) => span.name === "db.query")).toBe(true);
  });
});
