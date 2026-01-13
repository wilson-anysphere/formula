import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import WebSocket from "ws";
import jwt from "jsonwebtoken";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";
import { createLogger } from "../../../sync-server/src/logger";
import { createSyncServer } from "../../../sync-server/src/server";
import type { SyncServerConfig } from "../../../sync-server/src/config";

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
    "creates users, creates doc, invites user, issues sync token, and sync server accepts it",
    async () => {
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

      const decoded = jwt.verify(syncToken, config.syncTokenSecret, {
        audience: "formula-sync"
      }) as any;
      expect(decoded).toMatchObject({ sub: bobId, docId, orgId, role: "editor" });

      const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-api-e2e-"));
      const syncServerConfig: SyncServerConfig = {
        host: "127.0.0.1",
        port: 0,
        trustProxy: false,
        gc: true,
        tls: null,
        metrics: { public: true },
        dataDir,
        disableDataDirLock: false,
        persistence: {
          backend: "file",
          compactAfterUpdates: 50,
          leveldbDocNameHashing: false,
          encryption: { mode: "off" }
        },
        auth: {
          mode: "jwt-hs256",
          secret: config.syncTokenSecret,
          audience: "formula-sync",
          // Mirror the sync-server production defaults for JWT hardening.
          requireSub: true,
          requireExp: true
        },
        enforceRangeRestrictions: false,
        introspection: null,
        internalAdminToken: null,
        retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 0 },
        limits: {
          maxConnections: 100,
          maxConnectionsPerIp: 100,
          maxConnectionsPerDoc: 0,
          maxConnAttemptsPerWindow: 500,
          connAttemptWindowMs: 60_000,
          maxMessageBytes: 2 * 1024 * 1024,
          maxMessagesPerWindow: 5_000,
          messageWindowMs: 10_000,
          maxAwarenessStateBytes: 64 * 1024,
          maxAwarenessEntries: 10,
          maxMessagesPerDocWindow: 10_000,
          docMessageWindowMs: 10_000,
          maxBranchingCommitsPerDoc: 0,
          maxVersionsPerDoc: 0
        },
        logLevel: "silent"
      };

      const syncServer = createSyncServer(syncServerConfig, createLogger("silent"));
      const { port: syncPort } = await syncServer.start();

      const ws = new WebSocket(`ws://127.0.0.1:${syncPort}/${docId}?token=${encodeURIComponent(syncToken)}`);
      try {
        await new Promise<void>((resolve, reject) => {
          ws.once("open", () => resolve());
          ws.once("error", reject);
        });

        // `y-websocket` servers don't proactively send a sync frame on connect; the
        // client initiates by sending SyncStep1. Send the minimal valid state
        // vector (length=1, value=0) to trigger a server response.
        ws.send(Buffer.from([0, 0, 1, 0]));

        const firstMessage = await new Promise<WebSocket.RawData>((resolve, reject) => {
          const timeout = setTimeout(
            () => reject(new Error("Timed out waiting for sync server message")),
            10_000
          );
          timeout.unref?.();
          ws.once("message", (data) => {
            clearTimeout(timeout);
            resolve(data);
          });
          ws.once("error", (err) => {
            clearTimeout(timeout);
            reject(err);
          });
        });

        const firstMessageBytes = Buffer.isBuffer(firstMessage)
          ? firstMessage
          : firstMessage instanceof ArrayBuffer
            ? Buffer.from(firstMessage)
            : Array.isArray(firstMessage)
              ? Buffer.concat(firstMessage)
              : Buffer.from(firstMessage as any);
        expect(firstMessageBytes.byteLength).toBeGreaterThan(0);
      } finally {
        ws.close();
        await syncServer.stop();
        // Give the sync server a moment to flush file persistence writes after the
        // last client disconnects (y-websocket does not await writeState()).
        await new Promise((resolve) => setTimeout(resolve, 250));
        // Retry cleanup to avoid flaking on transient ENOTEMPTY races.
        await rm(dataDir, {
          recursive: true,
          force: true,
          maxRetries: 10,
          retryDelay: 25
        });
      }
    },
    20_000
  );

  it(
    "enforces document share permission (viewer cannot invite)",
    async () => {
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
    },
    20_000
  );

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

    const invalidRangePermEndRow = await app.inject({
      method: "POST",
      url: `/docs/${docId}/range-permissions`,
      headers: { cookie: ownerCookie },
      payload: {
        sheetName: "Sheet1",
        startRow: 1,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        permissionType: "read",
        allowedUserEmail: "public-user@example.com"
      }
    });
    expect(invalidRangePermEndRow.statusCode).toBe(400);
    expect((invalidRangePermEndRow.json() as any).error).toBe("invalid_request");

    const invalidRangePermEndCol = await app.inject({
      method: "POST",
      url: `/docs/${docId}/range-permissions`,
      headers: { cookie: ownerCookie },
      payload: {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 1,
        endRow: 0,
        endCol: 0,
        permissionType: "read",
        allowedUserEmail: "public-user@example.com"
      }
    });
    expect(invalidRangePermEndCol.statusCode).toBe(400);
    expect((invalidRangePermEndCol.json() as any).error).toBe("invalid_request");

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
    const rangePermissionId = (rangePerm.json() as any).id as string;
    expect(rangePermissionId).toBeTypeOf("string");

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

    const deleteRangePerm = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}/range-permissions/${rangePermissionId}`,
      headers: { cookie: ownerCookie }
    });
    expect(deleteRangePerm.statusCode).toBe(200);

    const permissionsAfterDelete = await app.inject({
      method: "GET",
      url: `/docs/${docId}/permissions`,
      headers: { cookie: publicUserCookie }
    });
    expect(permissionsAfterDelete.statusCode).toBe(200);
    const afterBody = permissionsAfterDelete.json() as any;
    expect(afterBody.permissions.role).toBe("viewer");
    expect(afterBody.permissions.rangeRestrictions).toHaveLength(0);
  }, 30_000);

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

    // Ensure encryption-at-rest is enabled for this org so the test verifies we
    // are not storing document version blobs in plaintext.
    await db.query("UPDATE org_settings SET cloud_encryption_at_rest = true WHERE org_id = $1", [orgId]);

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

    const rawStored = await db.query(
      "SELECT data, data_ciphertext, data_encrypted_dek FROM document_versions WHERE id = $1",
      [versionId]
    );
    expect(rawStored.rows[0].data).toBeNull();
    expect(rawStored.rows[0].data_ciphertext).toBeTypeOf("string");
    expect(rawStored.rows[0].data_encrypted_dek).toBeTypeOf("string");

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
  }, 30_000);
});
