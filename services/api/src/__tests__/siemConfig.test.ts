import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";

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

function parsePgJson(value: unknown): any {
  if (typeof value === "string") return JSON.parse(value);
  return value;
}

describe("SIEM config", () => {
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
      sessionCookieName: "formula_session",
      sessionTtlSeconds: 60 * 60,
      cookieSecure: false,
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("upserts a SIEM config, masks secrets on read, and deletes cleanly", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "siem-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "SIEM Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const put = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: {
        endpointUrl: "https://example.invalid/ingest",
        format: "json",
        idempotencyKeyHeader: "Idempotency-Key",
        auth: { type: "bearer", token: "supersecret" }
      }
    });
    expect(put.statusCode).toBe(200);
    const putBody = put.json() as any;
    expect(putBody.config.endpointUrl).toBe("https://example.invalid/ingest");
    expect(putBody.config.auth.token).toBe("***");

    const stored = await db.query("SELECT enabled, config FROM org_siem_configs WHERE org_id = $1", [orgId]);
    expect(stored.rowCount).toBe(1);
    expect(stored.rows[0]!.enabled).toBe(true);
    const storedConfig = parsePgJson(stored.rows[0]!.config);
    expect(storedConfig.auth.token).toBe("supersecret");

    const get = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(get.statusCode).toBe(200);
    const getBody = get.json() as any;
    expect(getBody.config.auth.token).toBe("***");

    const del = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(del.statusCode).toBe(204);

    const afterDelete = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(afterDelete.statusCode).toBe(404);

    const audit = await db.query(
      "SELECT event_type FROM audit_log WHERE org_id = $1 AND event_type LIKE 'org.siem_config.%'",
      [orgId]
    );
    const eventTypes = audit.rows.map((row) => row.event_type as string);
    expect(eventTypes).toContain("org.siem_config.updated");
    expect(eventTypes).toContain("org.siem_config.deleted");
  });
});

