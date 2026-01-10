import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import WebSocket from "ws";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { createSyncServer } from "../../../sync/src/server";

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

describe("API e2e: auth + RBAC + sync token", () => {
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
      retentionSweepIntervalMs: null
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("creates users, creates doc, invites user, issues sync token, and sync server accepts it", async () => {
    const aliceRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "alice@example.com",
        password: "password1234",
        name: "Alice",
        orgName: "Acme"
      }
    });
    expect(aliceRegister.statusCode).toBe(200);
    const aliceCookie = extractCookie(aliceRegister.headers["set-cookie"]);
    const aliceBody = aliceRegister.json() as any;
    const orgId = aliceBody.organization.id as string;

    const bobRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "bob@example.com",
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
      payload: { email: "bob@example.com", role: "editor" }
    });
    expect(inviteBob.statusCode).toBe(200);

    const syncTokenRes = await app.inject({
      method: "POST",
      url: `/docs/${docId}/sync-token`,
      headers: { cookie: bobCookie }
    });
    expect(syncTokenRes.statusCode).toBe(200);
    const syncToken = (syncTokenRes.json() as any).token as string;
    expect(syncToken).toBeTypeOf("string");

    const syncServer = createSyncServer({ port: 0, syncTokenSecret: config.syncTokenSecret });
    const syncPort = await syncServer.listen();

    const ws = new WebSocket(`ws://localhost:${syncPort}/${docId}?token=${encodeURIComponent(syncToken)}`);
    const firstMessage = await new Promise<string>((resolve, reject) => {
      ws.on("message", (data) => resolve(data.toString()));
      ws.on("error", reject);
    });
    const msg = JSON.parse(firstMessage) as any;
    expect(msg).toMatchObject({ type: "connected", docId, userId: bobId, role: "editor" });

    ws.close();
    await syncServer.close();
  });

  it("enforces document share permission (viewer cannot invite)", async () => {
    const ownerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Org"
      }
    });
    const ownerCookie = extractCookie(ownerRes.headers["set-cookie"]);
    const orgId = (ownerRes.json() as any).organization.id as string;

    await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "viewer@example.com",
        password: "password1234",
        name: "Viewer"
      }
    });
    await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "third@example.com",
        password: "password1234",
        name: "Third"
      }
    });

    const docRes = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Doc" }
    });
    const docId = (docRes.json() as any).document.id as string;

    await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: ownerCookie },
      payload: { email: "viewer@example.com", role: "viewer" }
    });

    const viewerLogin = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email: "viewer@example.com", password: "password1234" }
    });
    const viewerCookie = extractCookie(viewerLogin.headers["set-cookie"]);

    const forbiddenInvite = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: viewerCookie },
      payload: { email: "third@example.com", role: "viewer" }
    });

    expect(forbiddenInvite.statusCode).toBe(403);
  });
});

