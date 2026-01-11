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

function getCounterValue(metricsText: string, name: string): number {
  const match = metricsText.match(new RegExp(`^${name}\\s+(\\d+(?:\\.\\d+)?)$`, "m"));
  if (!match) throw new Error(`missing counter ${name}`);
  return Number.parseFloat(match[1]!);
}

describe("internal sync token introspection", () => {
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
      secretStoreKey: "test-secret-store-key",
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null,
      internalAdminToken: "test-internal-admin-token"
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it(
    "introspects a sync token and reflects current document role",
    async () => {
    const aliceRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "alice-introspect@example.com",
        password: "password1234",
        name: "Alice",
        orgName: "Acme"
      }
    });
    expect(aliceRegister.statusCode).toBe(200);
    const aliceCookie = extractCookie(aliceRegister.headers["set-cookie"]);
    const orgId = (aliceRegister.json() as any).organization.id as string;

    const bobRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "bob-introspect@example.com",
        password: "password1234",
        name: "Bob"
      }
    });
    expect(bobRegister.statusCode).toBe(200);
    const bobId = (bobRegister.json() as any).user.id as string;
    const bobCookie = extractCookie(bobRegister.headers["set-cookie"]);

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: aliceCookie },
      payload: { orgId, title: "Q1 Plan" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const inviteBob = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: aliceCookie },
      payload: { email: "bob-introspect@example.com", role: "editor" }
    });
    expect(inviteBob.statusCode).toBe(200);

    const syncTokenRes = await app.inject({
      method: "POST",
      url: `/docs/${docId}/sync-token`,
      headers: { cookie: bobCookie }
    });
    expect(syncTokenRes.statusCode).toBe(200);
    const syncToken = (syncTokenRes.json() as any).token as string;

    const introspectRes = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token: syncToken, docId }
    });
    expect(introspectRes.statusCode).toBe(200);
    expect(introspectRes.json()).toMatchObject({
      ok: true,
      userId: bobId,
      orgId,
      role: "editor"
    });

    await db.query(
      "UPDATE document_members SET role = 'viewer' WHERE document_id = $1 AND user_id = $2",
      [docId, bobId]
    );

    const roleUpdatedRes = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token: syncToken, docId }
    });
    expect(roleUpdatedRes.statusCode).toBe(200);
    expect(roleUpdatedRes.json()).toMatchObject({ ok: true, role: "viewer" });

    await db.query("DELETE FROM document_members WHERE document_id = $1 AND user_id = $2", [docId, bobId]);

    const membershipRemovedRes = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token: syncToken, docId }
    });
    expect(membershipRemovedRes.statusCode).toBe(403);
    expect(membershipRemovedRes.json()).toEqual({ ok: false, error: "forbidden" });

    const metricsRes = await app.inject({ method: "GET", url: "/metrics" });
    const failures = getCounterValue(metricsRes.body, "sync_token_introspect_failures_total");
    expect(failures).toBeGreaterThanOrEqual(1);
    },
    15_000
  );

  it(
    "rejects introspection for revoked sessions",
    async () => {
    const aliceRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "alice-introspect-2@example.com",
        password: "password1234",
        name: "Alice 2",
        orgName: "Acme 2"
      }
    });
    expect(aliceRegister.statusCode).toBe(200);
    const aliceCookie = extractCookie(aliceRegister.headers["set-cookie"]);
    const orgId = (aliceRegister.json() as any).organization.id as string;

    const bobRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "bob-introspect-2@example.com",
        password: "password1234",
        name: "Bob 2"
      }
    });
    expect(bobRegister.statusCode).toBe(200);
    const bobCookie = extractCookie(bobRegister.headers["set-cookie"]);

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: aliceCookie },
      payload: { orgId, title: "Q2 Plan" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const inviteBob = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: aliceCookie },
      payload: { email: "bob-introspect-2@example.com", role: "editor" }
    });
    expect(inviteBob.statusCode).toBe(200);

    const syncTokenRes = await app.inject({
      method: "POST",
      url: `/docs/${docId}/sync-token`,
      headers: { cookie: bobCookie }
    });
    expect(syncTokenRes.statusCode).toBe(200);
    const syncToken = (syncTokenRes.json() as any).token as string;

    const logoutRes = await app.inject({
      method: "POST",
      url: "/auth/logout",
      headers: { cookie: bobCookie }
    });
    expect(logoutRes.statusCode).toBe(200);

    const revokedRes = await app.inject({
      method: "POST",
      url: "/internal/sync/introspect",
      headers: { "x-internal-admin-token": config.internalAdminToken! },
      payload: { token: syncToken, docId }
    });
    expect(revokedRes.statusCode).toBe(403);
    expect(revokedRes.json()).toEqual({ ok: false, error: "forbidden" });

    const metricsRes = await app.inject({ method: "GET", url: "/metrics" });
    const failures = getCounterValue(metricsRes.body, "sync_token_introspect_failures_total");
    expect(failures).toBeGreaterThanOrEqual(2);
    },
    15_000
  );
});
