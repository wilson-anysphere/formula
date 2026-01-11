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

  it("creates share links (public/private), redeems public link, and exposes permissions", async () => {
    const ownerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "link-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Link Org"
      }
    });
    const ownerCookie = extractCookie(ownerRes.headers["set-cookie"]);
    const orgId = (ownerRes.json() as any).organization.id as string;

    const privateUserRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "private-user@example.com",
        password: "password1234",
        name: "Private"
      }
    });
    const privateUserCookie = extractCookie(privateUserRes.headers["set-cookie"]);

    const publicUserRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "public-user@example.com",
        password: "password1234",
        name: "Public"
      }
    });
    const publicUserCookie = extractCookie(publicUserRes.headers["set-cookie"]);
    const publicUserId = (publicUserRes.json() as any).user.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Shared Doc" }
    });
    const docId = (createDoc.json() as any).document.id as string;

    const privateLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "private", role: "viewer" }
    });
    expect(privateLink.statusCode).toBe(200);
    const privateToken = (privateLink.json() as any).shareLink.token as string;

    const privateRedeem = await app.inject({
      method: "POST",
      url: `/share-links/${privateToken}/redeem`,
      headers: { cookie: privateUserCookie }
    });
    expect(privateRedeem.statusCode).toBe(403);

    const publicLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" }
    });
    expect(publicLink.statusCode).toBe(200);
    const publicToken = (publicLink.json() as any).shareLink.token as string;

    const publicRedeem = await app.inject({
      method: "POST",
      url: `/share-links/${publicToken}/redeem`,
      headers: { cookie: publicUserCookie }
    });
    expect(publicRedeem.statusCode).toBe(200);
    expect((publicRedeem.json() as any).role).toBe("viewer");

    const rangePerm = await app.inject({
      method: "POST",
      url: `/docs/${docId}/range-permissions`,
      headers: { cookie: ownerCookie },
      payload: {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        permissionType: "read",
        allowedUserEmail: "public-user@example.com"
      }
    });
    expect(rangePerm.statusCode).toBe(200);

    const permissions = await app.inject({
      method: "GET",
      url: `/docs/${docId}/permissions`,
      headers: { cookie: publicUserCookie }
    });
    expect(permissions.statusCode).toBe(200);
    const body = permissions.json() as any;
    expect(body.permissions.role).toBe("viewer");
    expect(body.permissions.rangeRestrictions.length).toBe(1);
    expect(body.permissions.rangeRestrictions[0]).toMatchObject({
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0
    });
    expect(body.permissions.rangeRestrictions[0].readAllowlist).toContain(publicUserId);
  });

  it("creates, lists, fetches, deletes document versions and enforces RBAC", async () => {
    const ownerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "versions-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Versions Org"
      }
    });
    expect(ownerRes.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRes.headers["set-cookie"]);
    const ownerBody = ownerRes.json() as any;
    const ownerId = ownerBody.user.id as string;
    const orgId = ownerBody.organization.id as string;

    const viewerRes = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "versions-viewer@example.com",
        password: "password1234",
        name: "Viewer"
      }
    });
    expect(viewerRes.statusCode).toBe(200);
    const viewerCookie = extractCookie(viewerRes.headers["set-cookie"]);

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Versions Doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const inviteViewer = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: ownerCookie },
      payload: { email: "versions-viewer@example.com", role: "viewer" }
    });
    expect(inviteViewer.statusCode).toBe(200);

    const bytes = Buffer.from("hello versions");
    const dataBase64 = bytes.toString("base64");

    const invalidBase64 = await app.inject({
      method: "POST",
      url: `/docs/${docId}/versions`,
      headers: { cookie: ownerCookie },
      payload: { dataBase64: "not base64!!!" }
    });
    expect(invalidBase64.statusCode).toBe(400);
    expect((invalidBase64.json() as any).error).toBe("invalid_request");

    const createVersion = await app.inject({
      method: "POST",
      url: `/docs/${docId}/versions`,
      headers: { cookie: ownerCookie },
      payload: { description: "v1", dataBase64 }
    });
    expect(createVersion.statusCode).toBe(200);
    const created = createVersion.json() as any;
    expect(created.version.description).toBe("v1");
    expect(created.version.sizeBytes).toBe(bytes.length);
    const versionId = created.version.id as string;

    const listAsViewer = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions`,
      headers: { cookie: viewerCookie }
    });
    expect(listAsViewer.statusCode).toBe(200);
    const listed = listAsViewer.json() as any;
    expect(listed.versions).toHaveLength(1);
    expect(listed.versions[0]).toMatchObject({
      id: versionId,
      createdBy: ownerId,
      description: "v1",
      sizeBytes: bytes.length
    });

    const fetchAsViewer = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions/${versionId}`,
      headers: { cookie: viewerCookie }
    });
    expect(fetchAsViewer.statusCode).toBe(200);
    const fetched = (fetchAsViewer.json() as any).version as any;
    expect(fetched.createdBy).toBe(ownerId);
    expect(Buffer.from(fetched.dataBase64, "base64").equals(bytes)).toBe(true);

    const viewerCannotCreate = await app.inject({
      method: "POST",
      url: `/docs/${docId}/versions`,
      headers: { cookie: viewerCookie },
      payload: { dataBase64 }
    });
    expect(viewerCannotCreate.statusCode).toBe(403);

    const viewerCannotDelete = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}/versions/${versionId}`,
      headers: { cookie: viewerCookie }
    });
    expect(viewerCannotDelete.statusCode).toBe(403);

    const deleteDoc = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}`,
      headers: { cookie: ownerCookie }
    });
    expect(deleteDoc.statusCode).toBe(200);

    const createAfterDocDeleted = await app.inject({
      method: "POST",
      url: `/docs/${docId}/versions`,
      headers: { cookie: ownerCookie },
      payload: { dataBase64 }
    });
    expect(createAfterDocDeleted.statusCode).toBe(403);
    expect((createAfterDocDeleted.json() as any).error).toBe("doc_deleted");

    const deleteAsOwner = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}/versions/${versionId}`,
      headers: { cookie: ownerCookie }
    });
    expect(deleteAsOwner.statusCode).toBe(200);

    const listAfterDelete = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions`,
      headers: { cookie: ownerCookie }
    });
    expect(listAfterDelete.statusCode).toBe(200);
    expect((listAfterDelete.json() as any).versions).toHaveLength(0);
  });
});
