import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { Writable } from "node:stream";
import { trace } from "@opentelemetry/api";
import type { AppConfig } from "../config";
import { createLogger } from "../observability/logger";
import { initOpenTelemetry } from "../observability/otel";
import { InMemorySpanExporter } from "@opentelemetry/sdk-trace-base";
import { deriveSecretStoreKey } from "../secrets/secretStore";

describe("observability: request-id, log correlation, db spans", () => {
  let db: Pool;
  let config: AppConfig;
  let app: any;
  const logs: string[] = [];
  const exporter = new InMemorySpanExporter();

  beforeAll(async () => {
    initOpenTelemetry({ serviceName: "api-test", spanExporter: exporter });

    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();

    config = {
      port: 0,
      databaseUrl: "postgres://unused",
      publicBaseUrl: "http://localhost",
      publicBaseUrlHostAllowlist: ["localhost"],
      trustProxy: false,
      sessionCookieName: "formula_session",
      sessionTtlSeconds: 60 * 60,
      cookieSecure: false,
      corsAllowedOrigins: [],
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      secretStoreKeys: {
        currentKeyId: "legacy",
        keys: { legacy: deriveSecretStoreKey("test-secret-store-key") }
      },
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null,
      oidcAuthStateCleanupIntervalMs: null
    };

    const stream = new Writable({
      write(chunk, _encoding, callback) {
        logs.push(chunk.toString());
        callback();
      }
    });

    const logger = createLogger({ level: "info", stream });

    // Ensure OTel is initialized before importing Fastify/app.
    const { buildApp } = await import("../app");
    app = buildApp({ db, config, logger });

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

    const entry = logs
      .map((line) => JSON.parse(line) as any)
      .find((line) => line.msg === "test_log");

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
