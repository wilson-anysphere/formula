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

describe("API keys", () => {
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

  it("creates an API key and can use it to call /me (no raw key persisted)", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "api-key-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Key Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const allowApiKeys = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { allowedAuthMethods: ["password", "api_key"] }
    });
    expect(allowApiKeys.statusCode).toBe(200);

    const createKey = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/api-keys`,
      headers: { cookie },
      payload: { name: "ci" }
    });
    expect(createKey.statusCode).toBe(200);
    const createdBody = createKey.json() as any;
    const apiKeyId = createdBody.apiKey.id as string;
    const rawKey = createdBody.key as string;
    expect(rawKey).toMatch(/^api_[0-9a-f-]{36}\./i);

    const keyRow = await db.query("SELECT key_hash FROM api_keys WHERE id = $1", [apiKeyId]);
    expect(keyRow.rowCount).toBe(1);
    const keyHash = keyRow.rows[0]!.key_hash as string;
    const secret = rawKey.split(".")[1]!;
    expect(keyHash).not.toContain(secret);
    expect(keyHash).not.toBe(rawKey);

    const me = await app.inject({
      method: "GET",
      url: "/me",
      headers: { authorization: `Bearer ${rawKey}` }
    });
    expect(me.statusCode).toBe(200);
    const meBody = me.json() as any;
    expect(meBody.apiKey).toMatchObject({ id: apiKeyId, orgId, name: "ci" });
    expect(meBody.organizations).toHaveLength(1);
    expect(meBody.organizations[0]).toMatchObject({ id: orgId, name: "Key Org" });

    const usedRow = await db.query("SELECT last_used_at FROM api_keys WHERE id = $1", [apiKeyId]);
    expect(usedRow.rows[0]!.last_used_at).toBeTruthy();

    const audit = await db.query("SELECT event_type FROM audit_log WHERE resource_id = $1 ORDER BY created_at ASC", [
      apiKeyId
    ]);
    const eventTypes = audit.rows.map((row) => row.event_type as string);
    expect(eventTypes).toContain("org.api_key.created");
    expect(eventTypes).toContain("auth.api_key_used");
  });

  it("revoked API keys fail authentication", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "revoke-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Revoke Org"
      }
    });
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { allowedAuthMethods: ["password", "api_key"] }
    });

    const createKey = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/api-keys`,
      headers: { cookie },
      payload: { name: "automation" }
    });
    const createdBody = createKey.json() as any;
    const apiKeyId = createdBody.apiKey.id as string;
    const rawKey = createdBody.key as string;

    const revoke = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/api-keys/${apiKeyId}`,
      headers: { cookie }
    });
    expect(revoke.statusCode).toBe(200);

    const me = await app.inject({
      method: "GET",
      url: "/me",
      headers: { authorization: `Bearer ${rawKey}` }
    });
    expect(me.statusCode).toBe(401);
  });

  it("blocks API keys when org allowed_auth_methods excludes api_key", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "policy-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Policy Org"
      }
    });
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const createKey = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/api-keys`,
      headers: { cookie },
      payload: { name: "blocked" }
    });
    expect(createKey.statusCode).toBe(200);
    const rawKey = (createKey.json() as any).key as string;

    const me = await app.inject({
      method: "GET",
      url: "/me",
      headers: { authorization: `Bearer ${rawKey}` }
    });
    expect(me.statusCode).toBe(403);
    expect((me.json() as any).error).toBe("auth_method_not_allowed");
  });

  it("enforces org IP allowlists for API keys", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "ip-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "IP Org"
      }
    });
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: {
        allowedAuthMethods: ["password", "api_key"],
        ipAllowlist: ["10.0.0.0/8"]
      }
    });

    const createKey = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/api-keys`,
      remoteAddress: "10.1.2.3",
      headers: { cookie },
      payload: { name: "ip-test" }
    });
    const rawKey = (createKey.json() as any).key as string;

    const blocked = await app.inject({
      method: "GET",
      url: "/me",
      remoteAddress: "203.0.113.5",
      headers: { authorization: `Bearer ${rawKey}` }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("ip_not_allowed");

    const allowed = await app.inject({
      method: "GET",
      url: "/me",
      remoteAddress: "10.1.2.3",
      headers: { authorization: `Bearer ${rawKey}` }
    });
    expect(allowed.statusCode).toBe(200);
  });
});
