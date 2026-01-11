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

const ORG_DLP_POLICY = {
  version: 1,
  allowDocumentOverrides: true,
  rules: {
    "sharing.externalLink": { maxAllowed: "Internal" },
  },
};

describe("API e2e: DLP enforcement on external share links", () => {
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
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("blocks public share link creation when document classification exceeds policy and audits the decision", async () => {
    const ownerRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-enforce-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "DLP Enforcement Org",
      },
    });
    expect(ownerRegister.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
    const orgId = (ownerRegister.json() as any).organization.id as string;

    const editorRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-enforce-editor@example.com",
        password: "password1234",
        name: "Editor",
      },
    });
    expect(editorRegister.statusCode).toBe(200);

    const editorLogin = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email: "dlp-enforce-editor@example.com", password: "password1234" },
    });
    expect(editorLogin.statusCode).toBe(200);
    const editorCookie = extractCookie(editorLogin.headers["set-cookie"]);

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "DLP-enforced doc" },
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const inviteEditor = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: ownerCookie },
      payload: { email: "dlp-enforce-editor@example.com", role: "editor" },
    });
    expect(inviteEditor.statusCode).toBe(200);

    const putOrgPolicy = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/dlp-policy`,
      headers: { cookie: ownerCookie },
      payload: { policy: ORG_DLP_POLICY },
    });
    expect(putOrgPolicy.statusCode).toBe(200);

    const putRestrictedClassification = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: editorCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Restricted", labels: [] },
      },
    });
    expect(putRestrictedClassification.statusCode).toBe(200);

    const blockedShareLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" },
    });
    expect(blockedShareLink.statusCode).toBe(403);
    expect(blockedShareLink.json()).toMatchObject({ error: "dlp_blocked" });

    const audit = await db.query(
      "SELECT event_type, details FROM audit_log WHERE event_type = 'dlp.blocked' ORDER BY created_at DESC LIMIT 1"
    );
    expect(audit.rowCount).toBe(1);
    expect(audit.rows[0]!.details).toMatchObject({
      action: "sharing.externalLink",
      docId,
      classification: { level: "Restricted" },
      maxAllowed: "Internal",
    });

    const putInternalClassification = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: editorCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Internal", labels: [] },
      },
    });
    expect(putInternalClassification.statusCode).toBe(200);

    const allowedShareLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" },
    });
    expect(allowedShareLink.statusCode).toBe(200);
    expect((allowedShareLink.json() as any).shareLink).toMatchObject({ visibility: "public", role: "viewer" });
  });
});

