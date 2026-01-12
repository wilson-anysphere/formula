import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import http from "node:http";
import crypto from "node:crypto";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { createMetrics } from "../observability/metrics";
import { deriveSecretStoreKey } from "../secrets/secretStore";
import { DbSiemConfigProvider } from "../siem/configProvider";
import { SiemExportWorker } from "../siem/worker";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

function extractCookie(setCookieHeader: string | string[] | undefined): string {
  if (!setCookieHeader) throw new Error("missing set-cookie header");
  const raw = Array.isArray(setCookieHeader) ? setCookieHeader[0] : setCookieHeader;
  return raw.split(";")[0];
}

async function startSiemServer(): Promise<{
  url: string;
  requests: Array<{ body: string }>;
  close: () => Promise<void>;
}> {
  const requests: Array<{ body: string }> = [];

  const server = http.createServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
    req.on("end", () => {
      requests.push({ body: Buffer.concat(chunks).toString("utf8") });
      res.writeHead(200, { "content-type": "text/plain" });
      res.end("ok");
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

async function insertAuditEvent(options: {
  db: Pool;
  id: string;
  orgId: string;
  createdAt: Date;
  eventType: string;
}): Promise<void> {
  await options.db.query(
    `
      INSERT INTO audit_log (id, org_id, event_type, resource_type, success, details, created_at)
      VALUES ($1, $2, $3, 'organization', true, '{}'::jsonb, $4)
    `,
    [options.id, options.orgId, options.eventType, options.createdAt]
  );
}

describe("SIEM data residency enforcement", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

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

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("rejects SIEM configs outside allowed regions when allow_cross_region_processing=false", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "residency-admin@example.com",
        password: "password1234",
        name: "Admin",
        orgName: "Acme"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const setResidency = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { dataResidencyRegion: "eu", allowCrossRegionProcessing: false }
    });
    expect(setResidency.statusCode).toBe(200);

    const blocked = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: { endpointUrl: "https://siem.example.com/ingest", dataRegion: "us" }
    });
    expect(blocked.statusCode).toBe(400);
    expect((blocked.json() as any).error).toBe("invalid_request");
  });

  it("accepts SIEM configs within allowed regions when allow_cross_region_processing=false", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "residency-admin-2@example.com",
        password: "password1234",
        name: "Admin2",
        orgName: "Acme2"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const setResidency = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { dataResidencyRegion: "eu", allowCrossRegionProcessing: false }
    });
    expect(setResidency.statusCode).toBe(200);

    const ok = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: { endpointUrl: "https://siem.example.com/ingest", dataRegion: "eu" }
    });
    expect(ok.statusCode).toBe(200);
    expect((ok.json() as any).config.dataRegion).toBe("eu");

    // Cleanup so later worker-based tests don't try to export to a non-existent endpoint.
    const deleted = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(deleted.statusCode).toBe(204);
  });

  it("disables exports when org policy changes to block SIEM dataRegion", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "residency-admin-3@example.com",
        password: "password1234",
        name: "Admin3",
        orgName: "Acme3"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    // Org is EU, but cross-region processing initially allowed.
    const setResidency = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { dataResidencyRegion: "eu", allowCrossRegionProcessing: true }
    });
    expect(setResidency.statusCode).toBe(200);

    const siem = await startSiemServer();
    try {
      const putConfig = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/siem`,
        headers: { cookie },
        payload: { endpointUrl: siem.url, dataRegion: "us", format: "json" }
      });
      expect(putConfig.statusCode).toBe(200);

      const t0 = new Date("2025-01-01T00:00:00.000Z");
      const firstId = crypto.randomUUID();
      await insertAuditEvent({ db, id: firstId, orgId, createdAt: t0, eventType: "test.first" });

      const metrics = createMetrics();
      const worker = new SiemExportWorker({
        db,
        configProvider: new DbSiemConfigProvider(db, config.secretStoreKeys, console),
        metrics,
        logger: console,
        pollIntervalMs: 0
      });

      await worker.tick();
      expect(siem.requests).toHaveLength(1);

      // Policy flips to disallow cross-region processing; existing config is now non-compliant.
      const blockCrossRegion = await app.inject({
        method: "PATCH",
        url: `/orgs/${orgId}/settings`,
        headers: { cookie },
        payload: { allowCrossRegionProcessing: false }
      });
      expect(blockCrossRegion.statusCode).toBe(200);

      const t1 = new Date("2025-01-01T00:00:01.000Z");
      const secondId = crypto.randomUUID();
      await insertAuditEvent({ db, id: secondId, orgId, createdAt: t1, eventType: "test.second" });

      await worker.tick();
      // Still only one request: the second tick should block before sending.
      expect(siem.requests).toHaveLength(1);

      const configRow = await db.query("SELECT enabled FROM org_siem_configs WHERE org_id = $1", [orgId]);
      expect(configRow.rowCount).toBe(1);
      expect(configRow.rows[0].enabled).toBe(false);

      const stateRow = await db.query("SELECT last_error FROM org_siem_export_state WHERE org_id = $1", [orgId]);
      expect(stateRow.rowCount).toBe(1);
      expect(stateRow.rows[0].last_error).toBeTypeOf("string");
      expect(String(stateRow.rows[0].last_error)).toContain("siem.export.send");

      const blockedEvent = await db.query(
        "SELECT event_type, success FROM audit_log WHERE org_id = $1 AND event_type = 'org.data_residency.blocked'",
        [orgId]
      );
      expect(blockedEvent.rowCount).toBe(1);
      expect(blockedEvent.rows[0].success).toBe(false);
    } finally {
      await siem.close();
    }
  });
});
