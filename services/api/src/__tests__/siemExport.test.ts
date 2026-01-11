import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import http, { type IncomingMessage } from "node:http";
import crypto from "node:crypto";
import { runMigrations } from "../db/migrations";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { createMetrics } from "../observability/metrics";
import { DbSiemConfigProvider, type EnabledSiemOrg, type SiemConfigProvider } from "../siem/configProvider";
import type { SiemEndpointConfig } from "../siem/types";
import { deriveSecretStoreKey, putSecret, type SecretStoreKeyring } from "../secrets/secretStore";
import { SiemExportWorker } from "../siem/worker";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

type CapturedRequest = {
  headers: IncomingMessage["headers"];
  body: string;
  statusCode: number;
};

async function startSiemServer(options: { failTimes?: number } = {}): Promise<{
  url: string;
  requests: CapturedRequest[];
  close: () => Promise<void>;
}> {
  const requests: CapturedRequest[] = [];
  let remainingFailures = options.failTimes ?? 0;

  const server = http.createServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    req.on("end", () => {
      const statusCode = remainingFailures > 0 ? 500 : 200;
      if (remainingFailures > 0) remainingFailures -= 1;

      requests.push({
        headers: req.headers,
        body: Buffer.concat(chunks).toString("utf8"),
        statusCode
      });

      res.writeHead(statusCode, { "content-type": "text/plain" });
      res.end(statusCode === 200 ? "ok" : "retry");
    });
  });

  await new Promise<void>((resolve) => {
    server.listen(0, "127.0.0.1", () => resolve());
  });
  const address = server.address();
  if (!address || typeof address === "string") throw new Error("expected server to listen on tcp");

  return {
    url: `http://127.0.0.1:${address.port}/ingest`,
    requests,
    close: async () => {
      await new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      });
    }
  };
}

class StaticConfigProvider implements SiemConfigProvider {
  constructor(private readonly orgs: EnabledSiemOrg[]) {}
  async listEnabledOrgs(): Promise<EnabledSiemOrg[]> {
    return this.orgs;
  }
}

async function insertOrg(db: Pool, orgId: string): Promise<void> {
  await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Acme"]);
  await db.query("INSERT INTO org_settings (org_id) VALUES ($1)", [orgId]);
}

async function insertAuditEvent(options: {
  db: Pool;
  table: "audit_log" | "audit_log_archive";
  id: string;
  orgId: string;
  createdAt: Date;
  eventType: string;
}): Promise<void> {
  const base = {
    id: options.id,
    orgId: options.orgId,
    createdAt: options.createdAt,
    eventType: options.eventType
  };

  if (options.table === "audit_log") {
    await options.db.query(
      `
        INSERT INTO audit_log (id, org_id, event_type, resource_type, success, details, created_at)
        VALUES ($1, $2, $3, 'organization', true, '{}'::jsonb, $4)
      `,
      [base.id, base.orgId, base.eventType, base.createdAt]
    );
    return;
  }

  await options.db.query(
    `
      INSERT INTO audit_log_archive (id, org_id, event_type, resource_type, success, details, created_at)
      VALUES ($1, $2, $3, 'organization', true, '{}'::jsonb, $4)
    `,
    [base.id, base.orgId, base.eventType, base.createdAt]
  );
}

function idempotencyKeyFor(ids: string[]): string {
  return crypto.createHash("sha256").update(ids.join(","), "utf8").digest("hex");
}

describe("SIEM export worker", () => {
  let db: Pool;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });
  });

  afterAll(async () => {
    await db.end();
  });

  it("exports audit events (including archive) and advances cursor", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const firstId = crypto.randomUUID();
    const secondId = crypto.randomUUID();
    const t0 = new Date("2025-01-01T00:00:00.000Z");
    const t1 = new Date("2025-01-01T00:00:01.000Z");

    await insertAuditEvent({
      db,
      table: "audit_log_archive",
      id: firstId,
      orgId,
      createdAt: t0,
      eventType: "test.archived"
    });
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: secondId,
      orgId,
      createdAt: t1,
      eventType: "test.live"
    });

    const siem = await startSiemServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 2, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(1);
      const payload = JSON.parse(siem.requests[0]!.body) as any[];
      expect(payload.map((e) => e.id)).toEqual([firstId, secondId]);
      for (const exported of payload) {
        expect(exported).toHaveProperty("actor");
        expect(exported.details).not.toHaveProperty("__audit");
      }

      const idempotency = siem.requests[0]!.headers["idempotency-key"];
      expect(idempotency).toBe(idempotencyKeyFor([firstId, secondId]));

      const state = await db.query(
        "SELECT last_created_at, last_event_id, consecutive_failures FROM org_siem_export_state WHERE org_id = $1",
        [orgId]
      );
      expect(state.rowCount).toBe(1);
      expect(new Date(state.rows[0].last_created_at).toISOString()).toBe(t1.toISOString());
      expect(state.rows[0].last_event_id).toBe(secondId);
      expect(state.rows[0].consecutive_failures).toBe(0);
    } finally {
      await siem.close();
    }
  });

  it("strips __audit metadata from details and exports actor/correlation/resourceName", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const id = crypto.randomUUID();
    const timestamp = "2025-01-01T03:00:00.000Z";

    const event = createAuditEvent({
      id,
      timestamp,
      eventType: "test.meta",
      actor: { type: "user", id: "user_1" },
      context: {
        orgId,
        userEmail: "user@example.com",
        ipAddress: "203.0.113.5",
        userAgent: "UnitTest/1.0"
      },
      resource: { type: "document", id: "doc_1", name: "Q1 Plan" },
      success: true,
      details: { title: "Q1 Plan" },
      correlation: { requestId: "req_123", traceId: "trace_abc" }
    });

    await writeAuditEvent(db, event);

    const raw = await db.query("SELECT details FROM audit_log WHERE id = $1", [id]);
    expect(raw.rowCount).toBe(1);
    const storedDetails =
      typeof raw.rows[0].details === "string" ? JSON.parse(raw.rows[0].details) : (raw.rows[0].details as any);
    expect(storedDetails).toHaveProperty("__audit");

    const siem = await startSiemServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 2, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(1);
      const payload = JSON.parse(siem.requests[0]!.body) as any[];
      expect(payload).toHaveLength(1);
      expect(payload[0].id).toBe(id);
      expect(payload[0].timestamp).toBe(timestamp);
      expect(payload[0].details).toEqual({ title: "Q1 Plan" });
      expect(payload[0].details).not.toHaveProperty("__audit");
      expect(payload[0].actor).toEqual({ type: "user", id: "user_1" });
      expect(payload[0].correlation).toEqual({ requestId: "req_123", traceId: "trace_abc" });
      expect(payload[0].resource).toEqual({ type: "document", id: "doc_1", name: "Q1 Plan" });
    } finally {
      await siem.close();
    }
  });

  it("does not include __audit metadata in CEF/LEEF exports", async () => {
    const formats: Array<"cef" | "leef"> = ["cef", "leef"];

    for (const format of formats) {
      const orgId = crypto.randomUUID();
      await insertOrg(db, orgId);

      const id = crypto.randomUUID();
      const timestamp = "2025-01-01T05:00:00.000Z";

      await writeAuditEvent(
        db,
        createAuditEvent({
          id,
          timestamp,
          eventType: "test.siem_format",
          actor: { type: "user", id: "user_1" },
          context: {
            orgId,
            userEmail: "user@example.com",
            ipAddress: "203.0.113.5",
            userAgent: "UnitTest/1.0"
          },
          resource: { type: "document", id: "doc_1", name: "Q1 Plan" },
          success: true,
          details: { title: "Q1 Plan" },
          correlation: { requestId: "req_123", traceId: "trace_abc" }
        })
      );

      const siem = await startSiemServer();
      try {
        const config: SiemEndpointConfig = {
          endpointUrl: siem.url,
          format,
          idempotencyKeyHeader: "Idempotency-Key",
          retry: { maxAttempts: 2, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
        };

        const metrics = createMetrics();
        const worker = new SiemExportWorker({
          db,
          configProvider: new StaticConfigProvider([{ orgId, config }]),
          metrics,
          logger: console,
          pollIntervalMs: 0
        });

        await worker.tick();

        expect(siem.requests).toHaveLength(1);
        expect(siem.requests[0]!.headers["content-type"]).toBe("text/plain");
        expect(siem.requests[0]!.body).not.toContain("__audit");
        if (format === "cef") {
          expect(siem.requests[0]!.body).toContain("CEF:0|");
        } else {
          expect(siem.requests[0]!.body).toContain("LEEF:2.0|");
        }
      } finally {
        await siem.close();
      }
    }
  });

  it("retries transient failures and eventually succeeds", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T01:00:00.000Z"),
      eventType: "test.retry"
    });

    const siem = await startSiemServer({ failTimes: 2 });
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(3);
      expect(siem.requests.map((r) => r.statusCode)).toEqual([500, 500, 200]);
      for (const req of siem.requests) {
        expect(req.headers["idempotency-key"]).toBe(idempotencyKeyFor([eventId]));
      }

      const state = await db.query(
        "SELECT last_event_id, consecutive_failures FROM org_siem_export_state WHERE org_id = $1",
        [orgId]
      );
      expect(state.rows[0].last_event_id).toBe(eventId);
      expect(state.rows[0].consecutive_failures).toBe(0);
    } finally {
      await siem.close();
    }
  });

  it("fails fast when certificate pinning is enabled but pins are missing", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    await db.query("UPDATE org_settings SET certificate_pinning_enabled = true WHERE org_id = $1", [orgId]);

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T03:00:00.000Z"),
      eventType: "test.pinning_missing"
    });

    const config: SiemEndpointConfig = {
      endpointUrl: "https://example.invalid/ingest",
      format: "json",
      retry: { maxAttempts: 5, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
    };

    const metrics = createMetrics();
    const worker = new SiemExportWorker({
      db,
      configProvider: new StaticConfigProvider([{ orgId, config }]),
      metrics,
      logger: { info: () => {}, warn: () => {}, error: () => {} },
      pollIntervalMs: 0
    });

    await worker.tick();

    const state = await db.query(
      "SELECT last_event_id, consecutive_failures, last_error FROM org_siem_export_state WHERE org_id = $1",
      [orgId]
    );
    expect(state.rowCount).toBe(1);
    expect(state.rows[0].last_event_id).toBeNull();
    expect(state.rows[0].consecutive_failures).toBe(1);
    expect(String(state.rows[0].last_error)).toContain("certificatePins must be non-empty");
  });

  it("skips exports when disabled_until is in the future", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T02:00:00.000Z"),
      eventType: "test.disabled"
    });

    const disabledUntil = new Date(Date.now() + 60_000);
    await db.query(
      `
        INSERT INTO org_siem_export_state (org_id, disabled_until)
        VALUES ($1, $2)
        ON CONFLICT (org_id) DO UPDATE SET disabled_until = EXCLUDED.disabled_until
      `,
      [orgId, disabledUntil]
    );

    const siem = await startSiemServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 1, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(0);
    } finally {
      await siem.close();
    }
  });

  it("resolves SIEM auth secrets from the encrypted secret store", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T02:00:00.000Z"),
      eventType: "test.secret"
    });

    const siem = await startSiemServer();
    try {
      const encryptionSecret = "test-secret-store-key";
      const keyring: SecretStoreKeyring = {
        currentKeyId: "legacy",
        keys: { legacy: deriveSecretStoreKey(encryptionSecret) }
      };
      const secretName = `siem:${orgId}:headerValue:authorization`;
      await putSecret(db, keyring, secretName, "Splunk supersecret-token");

      const storedConfig: SiemEndpointConfig = {
        endpointUrl: siem.url,
        format: "json",
        auth: {
          type: "header",
          name: "Authorization",
          value: { secretRef: secretName }
        },
        retry: { maxAttempts: 1, baseDelayMs: 1, maxDelayMs: 5, jitter: false }
      };

      await db.query(
        `
          INSERT INTO org_siem_configs (org_id, enabled, config)
          VALUES ($1, true, $2)
        `,
        [orgId, JSON.stringify(storedConfig)]
      );

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new DbSiemConfigProvider(db, keyring, console),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(1);
      expect(siem.requests[0]!.headers["authorization"]).toBe("Splunk supersecret-token");
    } finally {
      await siem.close();
    }
  });

  it("blocks exports when SIEM dataRegion violates data residency policy", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);
    await db.query(
      "UPDATE org_settings SET data_residency_region = 'eu', allow_cross_region_processing = false WHERE org_id = $1",
      [orgId]
    );

    const eventId = crypto.randomUUID();
    await insertAuditEvent({
      db,
      table: "audit_log",
      id: eventId,
      orgId,
      createdAt: new Date("2025-01-01T04:00:00.000Z"),
      eventType: "test.residency_violation"
    });

    const siem = await startSiemServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        dataRegion: "us",
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 1, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: { info: () => {}, warn: () => {}, error: () => {} },
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(0);

      const blocked = await db.query(
        "SELECT 1 FROM audit_log WHERE org_id = $1 AND event_type = 'org.data_residency.blocked'",
        [orgId]
      );
      expect(blocked.rowCount).toBe(1);
    } finally {
      await siem.close();
    }
  });

  it("exports audit events inserted via writeAuditEvent (ingestion-compatible)", async () => {
    const orgId = crypto.randomUUID();
    await insertOrg(db, orgId);

    const userId = crypto.randomUUID();
    await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [userId, "siem@example.com", "SIEM"]);

    const event = createAuditEvent({
      eventType: "client.audit_ingested",
      actor: { type: "user", id: userId },
      context: {
        orgId,
        userId,
        userEmail: "siem@example.com",
        ipAddress: "203.0.113.5",
        userAgent: "UnitTest/siem"
      },
      resource: { type: "document", id: "doc_123", name: "Doc" },
      success: true,
      details: { token: "super-secret" }
    });
    await writeAuditEvent(db, event);

    const siem = await startSiemServer();
    try {
      const config: SiemEndpointConfig = {
        endpointUrl: siem.url,
        dataRegion: "us",
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        retry: { maxAttempts: 1, baseDelayMs: 1, maxDelayMs: 1, jitter: false }
      };

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new StaticConfigProvider([{ orgId, config }]),
        metrics,
        logger: { info: () => {}, warn: () => {}, error: () => {} },
        pollIntervalMs: 0
      });

      await worker.tick();

      expect(siem.requests).toHaveLength(1);
      const payload = JSON.parse(siem.requests[0]!.body) as any[];
      const exported = payload.find((e) => e.eventType === "client.audit_ingested");
      expect(exported).toBeTruthy();
      expect(exported.details).toMatchObject({ token: "[REDACTED]" });
    } finally {
      await siem.close();
    }
  });
});
