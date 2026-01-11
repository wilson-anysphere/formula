import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";

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

describe("SIEM config routes", () => {
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

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("stores SIEM config with encrypted secrets (CRUD)", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "admin@example.com",
        password: "password1234",
        name: "Admin",
        orgName: "Acme"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const token = "supersecret-token";
    const putRes = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: {
        enabled: true,
        config: {
          endpointUrl: "https://example.invalid/siem",
          format: "json",
          auth: {
            type: "header",
            name: "Authorization",
            value: `Splunk ${token}`
          }
        }
      }
    });
    expect(putRes.statusCode).toBe(200);

    const stored = await db.query("SELECT enabled, config FROM org_siem_configs WHERE org_id = $1", [orgId]);
    expect(stored.rowCount).toBe(1);
    expect(stored.rows[0]!.enabled).toBe(true);

    const rawConfig = JSON.stringify(stored.rows[0]!.config);
    expect(rawConfig).not.toContain(token);

    const secretName = `siem:${orgId}:headerValue:authorization`;
    const secretRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [secretName]);
    expect(secretRow.rowCount).toBe(1);
    const encrypted = secretRow.rows[0]!.encrypted_value as string;
    expect(encrypted).toMatch(/^v2:legacy:/);
    expect(encrypted).not.toContain(token);

    const getRes = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(getRes.statusCode).toBe(200);
    const getBody = getRes.json() as any;
    expect(getBody.enabled).toBe(true);
    expect(getBody.config.auth).toEqual({ type: "header", name: "Authorization", value: "***" });

    const delRes = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie }
    });
    expect(delRes.statusCode).toBe(204);

    const remainingConfig = await db.query("SELECT 1 FROM org_siem_configs WHERE org_id = $1", [orgId]);
    expect(remainingConfig.rowCount).toBe(0);
    const remainingSecrets = await db.query("SELECT 1 FROM secrets WHERE name = $1", [secretName]);
    expect(remainingSecrets.rowCount).toBe(0);
  });
});
