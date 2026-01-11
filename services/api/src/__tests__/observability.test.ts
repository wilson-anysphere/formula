import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { Writable } from "node:stream";
import type { AppConfig } from "../config";
import { createLogger } from "../observability/logger";
import { initOpenTelemetry } from "../observability/otel";
import { InMemorySpanExporter } from "@opentelemetry/sdk-trace-base";
import { deriveSecretStoreKey } from "../secrets/secretStore";

describe("observability: request-id, log correlation, db spans", () => {
  let db: Pool;
  let config: AppConfig;
  let app: any;
  let baseUrl: string;
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
      retentionSweepIntervalMs: null
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

    await app.listen({ port: 0, host: "127.0.0.1" });
    const addr = app.server.address();
    if (!addr || typeof addr === "string") throw new Error("expected server to listen on a tcp port");
    baseUrl = `http://127.0.0.1:${addr.port}`;
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("sets x-request-id and respects incoming x-request-id", async () => {
    const res = await fetch(`${baseUrl}/health`);
    expect(res.status).toBe(200);
    expect(res.headers.get("x-request-id")).toBeTypeOf("string");

    const res2 = await fetch(`${baseUrl}/health`, {
      headers: { "x-request-id": "client-request-id-123" }
    });
    expect(res2.status).toBe(200);
    expect(res2.headers.get("x-request-id")).toBe("client-request-id-123");
  });

  it("includes traceId/spanId/requestId in logs", async () => {
    logs.length = 0;
    exporter.reset();

    const res = await fetch(`${baseUrl}/_test/log`);
    const requestId = res.headers.get("x-request-id");

    const entry = logs
      .map((line) => JSON.parse(line) as any)
      .find((line) => line.msg === "test_log");

    expect(entry).toBeTruthy();
    expect(entry.requestId).toBe(requestId);
    expect(entry.traceId).toMatch(/^[0-9a-f]{32}$/);
    expect(entry.spanId).toMatch(/^[0-9a-f]{16}$/);
  });

  it("creates db.query spans", async () => {
    logs.length = 0;
    exporter.reset();

    const res = await fetch(`${baseUrl}/_test/db`);
    expect(res.status).toBe(200);

    const spans = exporter.getFinishedSpans();
    expect(spans.some((span) => span.name === "db.query")).toBe(true);
  });
});
