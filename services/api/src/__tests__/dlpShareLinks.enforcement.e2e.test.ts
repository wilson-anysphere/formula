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

describe("API e2e: DLP enforcement on public share links", () => {
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

  it("blocks creating a public share link when DLP policy disallows external links", async () => {
    const ownerRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-link-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "DLP Link Org"
      }
    });
    expect(ownerRegister.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
    const orgId = (ownerRegister.json() as any).organization.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "DLP doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const putOrgPolicy = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/dlp-policy`,
      headers: { cookie: ownerCookie },
      payload: {
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            "sharing.externalLink": { maxAllowed: "Internal" }
          }
        }
      }
    });
    expect(putOrgPolicy.statusCode).toBe(200);

    const putClassification = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Confidential", labels: [] }
      }
    });
    expect(putClassification.statusCode).toBe(200);

    const evaluate = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie: ownerCookie },
      payload: { action: "sharing.externalLink" }
    });
    expect(evaluate.statusCode).toBe(200);
    expect(evaluate.json()).toMatchObject({
      decision: "block",
      reasonCode: "dlp.blockedByPolicy",
      classification: { level: "Confidential" },
      maxAllowed: "Internal"
    });

    const createLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" }
    });
    expect(createLink.statusCode).toBe(403);
    expect(createLink.json()).toMatchObject({ error: "dlp_blocked" });

    const links = await db.query("SELECT id FROM document_share_links WHERE document_id = $1", [docId]);
    expect(links.rowCount).toBe(0);

    const audit = await db.query(
      "SELECT event_type, resource_id, details FROM audit_log WHERE event_type = 'dlp.blocked' ORDER BY created_at DESC LIMIT 1"
    );
    expect(audit.rowCount).toBe(1);
    expect(audit.rows[0]).toMatchObject({
      event_type: "dlp.blocked",
      resource_id: docId
    });
    expect((audit.rows[0] as any).details).toMatchObject({
      action: "sharing.externalLink",
      docId,
      classification: { level: "Confidential" },
      maxAllowed: "Internal",
      reasonCode: "dlp.blockedByPolicy"
    });
  }, 20_000);

  it("falls back to the default org policy when org_dlp_policies is missing", async () => {
    const ownerRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-default-policy-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "DLP Default Policy Org"
      }
    });
    expect(ownerRegister.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
    const orgId = (ownerRegister.json() as any).organization.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "DLP default policy doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const putClassification = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Confidential", labels: [] }
      }
    });
    expect(putClassification.statusCode).toBe(200);

    const evaluate = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie: ownerCookie },
      payload: { action: "sharing.externalLink" }
    });
    expect(evaluate.statusCode).toBe(200);
    expect(evaluate.json()).toMatchObject({
      decision: "block",
      reasonCode: "dlp.blockedByPolicy",
      classification: { level: "Confidential" },
      maxAllowed: "Internal"
    });

    const createLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" }
    });
    expect(createLink.statusCode).toBe(403);
    expect(createLink.json()).toMatchObject({ error: "dlp_blocked" });

    const links = await db.query("SELECT id FROM document_share_links WHERE document_id = $1", [docId]);
    expect(links.rowCount).toBe(0);

    const audit = await db.query(
      "SELECT event_type, resource_id, details FROM audit_log WHERE event_type = 'dlp.blocked' AND resource_id = $1 ORDER BY created_at DESC LIMIT 1",
      [docId]
    );
    expect(audit.rowCount).toBe(1);
    expect(audit.rows[0]).toMatchObject({
      event_type: "dlp.blocked",
      resource_id: docId
    });
    expect((audit.rows[0] as any).details).toMatchObject({
      action: "sharing.externalLink",
      docId,
      classification: { level: "Confidential" },
      maxAllowed: "Internal",
      reasonCode: "dlp.blockedByPolicy"
    });
  }, 20_000);

  it("blocks redeeming a public share link for a now-confidential document (external user)", async () => {
    const ownerRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-redeem-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "DLP Redeem Org"
      }
    });
    expect(ownerRegister.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
    const ownerOrgId = (ownerRegister.json() as any).organization.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId: ownerOrgId, title: "DLP redeem doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const putOrgPolicy = await app.inject({
      method: "PUT",
      url: `/orgs/${ownerOrgId}/dlp-policy`,
      headers: { cookie: ownerCookie },
      payload: {
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            "sharing.externalLink": { maxAllowed: "Internal" }
          }
        }
      }
    });
    expect(putOrgPolicy.statusCode).toBe(200);

    const setInternal = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Internal", labels: [] }
      }
    });
    expect(setInternal.statusCode).toBe(200);

    const createLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" }
    });
    expect(createLink.statusCode).toBe(200);
    const token = (createLink.json() as any).shareLink.token as string;
    expect(token).toBeTypeOf("string");

    // Later, the document is reclassified more restrictively. Redemption should be
    // re-evaluated for external users even if the link already exists.
    const setConfidential = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: {
        selector: { scope: "document", documentId: docId },
        classification: { level: "Confidential", labels: [] }
      }
    });
    expect(setConfidential.statusCode).toBe(200);

    const externalRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-redeem-external@example.com",
        password: "password1234",
        name: "External User"
      }
    });
    expect(externalRegister.statusCode).toBe(200);
    const externalCookie = extractCookie(externalRegister.headers["set-cookie"]);
    const externalUserId = (externalRegister.json() as any).user.id as string;

    const redeem = await app.inject({
      method: "POST",
      url: `/share-links/${token}/redeem`,
      headers: { cookie: externalCookie }
    });
    expect(redeem.statusCode).toBe(403);
    expect(redeem.json()).toMatchObject({ error: "dlp_blocked" });

    const orgMembership = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
      ownerOrgId,
      externalUserId
    ]);
    expect(orgMembership.rowCount).toBe(0);

    const docMembership = await db.query("SELECT 1 FROM document_members WHERE document_id = $1 AND user_id = $2", [
      docId,
      externalUserId
    ]);
    expect(docMembership.rowCount).toBe(0);

    const audit = await db.query(
      "SELECT event_type, resource_id, details FROM audit_log WHERE event_type = 'dlp.blocked' AND resource_id = $1 ORDER BY created_at DESC LIMIT 1",
      [docId]
    );
    expect(audit.rowCount).toBe(1);
    expect(audit.rows[0]).toMatchObject({ event_type: "dlp.blocked", resource_id: docId });
    expect((audit.rows[0] as any).details).toMatchObject({
      action: "sharing.externalLink",
      docId,
      classification: { level: "Confidential" },
      maxAllowed: "Internal",
      reasonCode: "dlp.blockedByPolicy"
    });
  }, 20_000);
});
