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

describe("internal: sync token introspection", () => {
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
      oidcAuthStateCleanupIntervalMs: null,
      internalAdminToken: "internal-admin-token"
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it(
    "returns inactive when the issuing session is revoked",
    async () => {
      const suffix = Math.random().toString(16).slice(2);
      const email = `introspect-session-${suffix}@example.com`;

    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email,
        password: "password1234",
        name: "User",
        orgName: "Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const body = register.json() as any;
    const userId = body.user.id as string;
    const orgId = body.organization.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const tokenRes = await app.inject({
      method: "POST",
      url: `/docs/${docId}/sync-token`,
      headers: { cookie }
    });
    expect(tokenRes.statusCode).toBe(200);
    const token = (tokenRes.json() as any).token as string;

    const introspectActive = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token, docId, clientIp: "203.0.113.1", userAgent: "vitest" }
    });
    expect(introspectActive.statusCode).toBe(200);
    expect(introspectActive.json()).toMatchObject({ active: true, userId, orgId, role: "owner" });

    const logout = await app.inject({
      method: "POST",
      url: "/auth/logout",
      headers: { cookie }
    });
    expect(logout.statusCode).toBe(200);

    const introspectRevoked = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token, docId, clientIp: "203.0.113.1", userAgent: "vitest" }
    });
    expect(introspectRevoked.statusCode).toBe(200);
    expect(introspectRevoked.json()).toMatchObject({
      active: false,
      reason: "session_revoked",
      userId,
      orgId,
      role: "owner"
    });
    },
    20_000
  );

  it(
    "returns inactive when document membership is removed",
    async () => {
      const suffix = Math.random().toString(16).slice(2);
      const ownerEmail = `introspect-owner-${suffix}@example.com`;
      const memberEmail = `introspect-member-${suffix}@example.com`;

    const ownerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: ownerEmail,
        password: "password1234",
        name: "Owner",
        orgName: "Org"
      }
    });
    expect(ownerRes.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRes.headers["set-cookie"]);
    const orgId = (ownerRes.json() as any).organization.id as string;

    const memberRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: memberEmail,
        password: "password1234",
        name: "Member"
      }
    });
    expect(memberRes.statusCode).toBe(200);
    const memberCookie = extractCookie(memberRes.headers["set-cookie"]);
    const memberId = (memberRes.json() as any).user.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const invite = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: ownerCookie },
      payload: { email: memberEmail, role: "editor" }
    });
    expect(invite.statusCode).toBe(200);

    const tokenRes = await app.inject({
      method: "POST",
      url: `/docs/${docId}/sync-token`,
      headers: { cookie: memberCookie }
    });
    expect(tokenRes.statusCode).toBe(200);
    const token = (tokenRes.json() as any).token as string;

    await db.query("DELETE FROM document_members WHERE document_id = $1 AND user_id = $2", [
      docId,
      memberId
    ]);

    const introspect = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token, docId, clientIp: "203.0.113.2", userAgent: "vitest" }
    });
    expect(introspect.statusCode).toBe(200);
    expect(introspect.json()).toMatchObject({
      active: false,
      reason: "not_member",
      userId: memberId,
      orgId,
      role: "editor"
    });
    },
    20_000
  );
});
