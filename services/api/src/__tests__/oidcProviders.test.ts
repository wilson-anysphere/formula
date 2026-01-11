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

function parseJsonValue(value: unknown): any {
  if (!value) return null;
  if (typeof value === "object") return value;
  if (typeof value === "string") return JSON.parse(value);
  return null;
}

describe("OIDC provider administration APIs", () => {
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

  it(
    "allows org admins to CRUD providers and never returns client secrets",
    async () => {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
      payload: {
        email: "oidc-admin-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "OIDC Admin Org"
      }
    });
    expect(ownerRegister.statusCode).toBe(200);
    const orgId = (ownerRegister.json() as any).organization.id as string;
    const cookie = extractCookie(ownerRegister.headers["set-cookie"]);

    const listEmpty = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie }
    });
    expect(listEmpty.statusCode).toBe(200);
    expect((listEmpty.json() as any).providers).toEqual([]);

    const secret = "super-secret-value";
    const createRes = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/oidc-providers/okta`,
      headers: { cookie },
      payload: {
        issuerUrl: "https://issuer.example.com",
        clientId: "client-123",
        clientSecret: secret,
        scopes: ["email", "openid", "email", "profile"],
        enabled: true
      }
    });
    expect(createRes.statusCode).toBe(200);
    const created = createRes.json() as any;
    expect(created.provider).toMatchObject({
      providerId: "okta",
      issuerUrl: "https://issuer.example.com",
      clientId: "client-123",
      enabled: true
    });
    expect(created.provider.scopes).toEqual(["openid", "email", "profile"]);
    expect(created.clientSecretConfigured).toBe(true);
    expect(created.secretConfigured).toBe(true);
    expect(JSON.stringify(created)).not.toContain(secret);

    const listAfterCreate = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie }
    });
    expect(listAfterCreate.statusCode).toBe(200);
    expect((listAfterCreate.json() as any).providers).toMatchObject([
      {
        providerId: "okta",
        issuerUrl: "https://issuer.example.com",
        clientId: "client-123",
        enabled: true,
        secretConfigured: true,
        clientSecretConfigured: true
      }
    ]);
    expect(JSON.stringify(listAfterCreate.json())).not.toContain(secret);

    const updateRes = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/oidc-providers/okta`,
      headers: { cookie },
      payload: {
        issuerUrl: "https://issuer.example.com",
        clientId: "client-456",
        enabled: false
      }
    });
    expect(updateRes.statusCode).toBe(200);
    const updated = updateRes.json() as any;
    expect(updated.provider).toMatchObject({
      providerId: "okta",
      issuerUrl: "https://issuer.example.com",
      clientId: "client-456",
      enabled: false
    });
    expect(updated.clientSecretConfigured).toBe(true);
    expect(updated.secretConfigured).toBe(true);

    const deleteRes = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/oidc-providers/okta`,
      headers: { cookie }
    });
    expect(deleteRes.statusCode).toBe(200);
    expect((deleteRes.json() as any).ok).toBe(true);

    const listAfterDelete = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie }
    });
    expect(listAfterDelete.statusCode).toBe(200);
    expect((listAfterDelete.json() as any).providers).toEqual([]);

    const secretRow = await db.query("SELECT 1 FROM secrets WHERE name = $1", [`oidc:${orgId}:okta`]);
    expect(secretRow.rowCount).toBe(0);

    const audit = await db.query(
      "SELECT event_type, details FROM audit_log WHERE org_id = $1 AND event_type = 'admin.integration_added' ORDER BY created_at DESC LIMIT 1",
      [orgId]
    );
    expect(audit.rowCount).toBe(1);
    const details = parseJsonValue(audit.rows[0]!.details);
    expect(JSON.stringify(details)).not.toContain(secret);
    expect(details).toMatchObject({ type: "oidc", providerId: "okta" });
    },
    15_000
  );

  it(
    "rejects non-admin org members",
    async () => {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
      payload: {
        email: "oidc-admin2-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "OIDC Admin Org 2"
      }
    });
    const orgId = (ownerRegister.json() as any).organization.id as string;
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);

    const memberRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "oidc-member@example.com",
        password: "password1234",
        name: "Member",
        orgName: "Member Org"
      }
    });
    const memberCookie = extractCookie(memberRegister.headers["set-cookie"]);
    const memberId = (memberRegister.json() as any).user.id as string;

    await db.query("INSERT INTO org_members (org_id, user_id, role) VALUES ($1, $2, 'member')", [
      orgId,
      memberId
    ]);

    const listRes = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie: memberCookie }
    });
    expect(listRes.statusCode).toBe(403);

    // Owner still works.
    const ownerList = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie: ownerCookie }
    });
    expect(ownerList.statusCode).toBe(200);
    },
    10_000
  );

  it("rejects invalid orgId path params", async () => {
    const registerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "oidc-invalid-orgid@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "OIDC Invalid OrgId"
      }
    });
    expect(registerRes.statusCode).toBe(200);
    const cookie = extractCookie(registerRes.headers["set-cookie"]);

    const res = await app.inject({
      method: "GET",
      url: "/orgs/not-a-uuid/oidc-providers",
      headers: { cookie }
    });
    expect(res.statusCode).toBe(400);
    expect((res.json() as any).error).toBe("invalid_request");
  });
});
