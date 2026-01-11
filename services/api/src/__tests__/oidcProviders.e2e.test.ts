import path from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
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

describe("OIDC provider admin APIs", () => {
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

  it("creates, lists, gets, and deletes providers; secrets are encrypted and never returned", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "oidc-admin@example.com",
        password: "password1234",
        name: "OIDC Admin",
        orgName: "OIDC Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const providerId = "okta";
    const clientSecret = "super-secret-value";

    const putRes = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/oidc/providers/${providerId}`,
      headers: { cookie },
      payload: {
        issuerUrl: "http://issuer.example.com",
        clientId: "client-123",
        scopes: ["email", "profile"],
        enabled: true,
        clientSecret
      }
    });
    expect(putRes.statusCode).toBe(200);
    expect(putRes.json()).toMatchObject({
      provider: {
        providerId,
        issuerUrl: "http://issuer.example.com",
        clientId: "client-123",
        enabled: true
      },
      clientSecretConfigured: true
    });
    expect(JSON.stringify(putRes.json())).not.toContain(clientSecret);

    const providerRow = await db.query(
      "SELECT issuer_url, client_id, scopes, enabled FROM org_oidc_providers WHERE org_id = $1 AND provider_id = $2",
      [orgId, providerId]
    );
    expect(providerRow.rowCount).toBe(1);
    expect(providerRow.rows[0]).toMatchObject({
      issuer_url: "http://issuer.example.com",
      client_id: "client-123",
      enabled: true
    });

    const scopes = providerRow.rows[0]!.scopes;
    const scopesArray = Array.isArray(scopes) ? scopes : JSON.parse(String(scopes));
    expect(scopesArray).toContain("openid");

    const secretRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [
      `oidc:${orgId}:${providerId}`
    ]);
    expect(secretRow.rowCount).toBe(1);
    const encrypted = String(secretRow.rows[0]!.encrypted_value);
    expect(encrypted).not.toContain(clientSecret);

    const listRes = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc/providers`,
      headers: { cookie }
    });
    expect(listRes.statusCode).toBe(200);
    const listBody = listRes.json() as any;
    expect(listBody.providers).toHaveLength(1);
    expect(listBody.providers[0]).toMatchObject({
      providerId,
      issuerUrl: "http://issuer.example.com",
      clientId: "client-123",
      enabled: true,
      clientSecretConfigured: true
    });
    expect(listBody.providers[0].clientSecret).toBeUndefined();

    const getRes = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc/providers/${providerId}`,
      headers: { cookie }
    });
    expect(getRes.statusCode).toBe(200);
    const getBody = getRes.json() as any;
    expect(getBody).toMatchObject({
      provider: { providerId, issuerUrl: "http://issuer.example.com", clientId: "client-123", enabled: true },
      clientSecretConfigured: true
    });
    expect(JSON.stringify(getBody)).not.toContain(clientSecret);

    const delRes = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/oidc/providers/${providerId}`,
      headers: { cookie }
    });
    expect(delRes.statusCode).toBe(200);
    expect(delRes.json()).toMatchObject({ ok: true });

    const providerAfter = await db.query("SELECT 1 FROM org_oidc_providers WHERE org_id = $1 AND provider_id = $2", [
      orgId,
      providerId
    ]);
    expect(providerAfter.rowCount).toBe(0);

    const secretAfter = await db.query("SELECT 1 FROM secrets WHERE name = $1", [`oidc:${orgId}:${providerId}`]);
    expect(secretAfter.rowCount).toBe(0);
  });
});
