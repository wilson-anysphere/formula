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

const DEFAULT_DLP_POLICY = {
  version: 1,
  allowDocumentOverrides: true,
  rules: {
    "sharing.externalLink": { maxAllowed: "Internal" },
    "export.csv": { maxAllowed: "Confidential" },
    "export.pdf": { maxAllowed: "Confidential" },
    "export.xlsx": { maxAllowed: "Confidential" },
    "clipboard.copy": { maxAllowed: "Confidential" },
    "connector.external": { maxAllowed: "Internal" },
    "ai.cloudProcessing": {
      maxAllowed: "Confidential",
      allowRestrictedContent: false,
      redactDisallowed: true
    }
  }
};

describe("API e2e: DLP policy + classification endpoints", () => {
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
    "stores org policy, doc policy overrides, and document classifications",
    async () => {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "dlp-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "DLP Org"
        }
      });
      expect(ownerRegister.statusCode).toBe(200);
      const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
      const orgId = (ownerRegister.json() as any).organization.id as string;

      const editorRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "dlp-editor@example.com",
          password: "password1234",
          name: "Editor"
        }
      });
      expect(editorRegister.statusCode).toBe(200);

      const createDoc = await app.inject({
        method: "POST",
        url: "/docs",
        headers: { cookie: ownerCookie },
        payload: { orgId, title: "DLP doc" }
      });
      expect(createDoc.statusCode).toBe(200);
      const docId = (createDoc.json() as any).document.id as string;

      const inviteEditor = await app.inject({
        method: "POST",
        url: `/docs/${docId}/invite`,
        headers: { cookie: ownerCookie },
        payload: { email: "dlp-editor@example.com", role: "editor" }
      });
      expect(inviteEditor.statusCode).toBe(200);

      const editorLogin = await app.inject({
        method: "POST",
        url: "/auth/login",
        payload: { email: "dlp-editor@example.com", password: "password1234" }
      });
      expect(editorLogin.statusCode).toBe(200);
      const editorCookie = extractCookie(editorLogin.headers["set-cookie"]);

      const putOrgPolicy = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/dlp-policy`,
        headers: { cookie: ownerCookie },
        payload: { policy: DEFAULT_DLP_POLICY }
      });
      expect(putOrgPolicy.statusCode).toBe(200);

      const getOrgPolicy = await app.inject({
        method: "GET",
        url: `/orgs/${orgId}/dlp-policy`,
        headers: { cookie: editorCookie }
      });
      expect(getOrgPolicy.statusCode).toBe(200);
      expect(getOrgPolicy.json()).toMatchObject({ policy: DEFAULT_DLP_POLICY });

      const putDocPolicyForbidden = await app.inject({
        method: "PUT",
        url: `/docs/${docId}/dlp-policy`,
        headers: { cookie: editorCookie },
        payload: { policy: DEFAULT_DLP_POLICY }
      });
      expect(putDocPolicyForbidden.statusCode).toBe(403);

      const putDocPolicy = await app.inject({
        method: "PUT",
        url: `/docs/${docId}/dlp-policy`,
        headers: { cookie: ownerCookie },
        payload: { policy: DEFAULT_DLP_POLICY }
      });
      expect(putDocPolicy.statusCode).toBe(200);

      const putClassification = await app.inject({
        method: "PUT",
        url: `/docs/${docId}/classifications`,
        headers: { cookie: editorCookie },
        payload: {
          selector: {
            scope: "range",
            documentId: docId,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }
          },
          classification: {
            level: "Restricted",
            labels: ["PII"]
          }
        }
      });
      expect(putClassification.statusCode).toBe(200);

      const listClassifications = await app.inject({
        method: "GET",
        url: `/docs/${docId}/classifications`,
        headers: { cookie: ownerCookie }
      });
      expect(listClassifications.statusCode).toBe(200);
      expect((listClassifications.json() as any).classifications).toHaveLength(1);

      const selectorKey = "range:" + `${docId}:Sheet1:` + "0,0:1,1";
      const deleteClassification = await app.inject({
        method: "DELETE",
        url: `/docs/${docId}/classifications/${encodeURIComponent(selectorKey)}`,
        headers: { cookie: editorCookie }
      });
      expect(deleteClassification.statusCode).toBe(200);

      const listAfterDelete = await app.inject({
        method: "GET",
        url: `/docs/${docId}/classifications`,
        headers: { cookie: ownerCookie }
      });
      expect(listAfterDelete.statusCode).toBe(200);
      expect((listAfterDelete.json() as any).classifications).toHaveLength(0);
    },
    20_000
  );

  it("resolves effective classification with selector precedence", async () => {
    const ownerRegister = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "dlp-resolve-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "DLP Resolve Org"
      }
    });
    expect(ownerRegister.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerRegister.headers["set-cookie"]);
    const orgId = (ownerRegister.json() as any).organization.id as string;

    const putOrgPolicy = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/dlp-policy`,
      headers: { cookie: ownerCookie },
      payload: { policy: DEFAULT_DLP_POLICY }
    });
    expect(putOrgPolicy.statusCode).toBe(200);

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "DLP resolve doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const rangeSelector = {
      scope: "range",
      documentId: docId,
      sheetId: "Sheet1",
      range: { start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }
    };
    const putRange = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: { selector: rangeSelector, classification: { level: "Restricted", labels: ["PII"] } }
    });
    expect(putRange.statusCode).toBe(200);

    const cellSelector = {
      scope: "cell",
      documentId: docId,
      sheetId: "Sheet1",
      row: 0,
      col: 0
    };
    const putCell = await app.inject({
      method: "PUT",
      url: `/docs/${docId}/classifications`,
      headers: { cookie: ownerCookie },
      payload: { selector: cellSelector, classification: { level: "Internal", labels: ["Mask"] } }
    });
    expect(putCell.statusCode).toBe(200);

    const resolveCell = await app.inject({
      method: "POST",
      url: `/docs/${docId}/classifications/resolve`,
      headers: { cookie: ownerCookie },
      payload: { selector: cellSelector }
    });
    expect(resolveCell.statusCode).toBe(200);
    expect(resolveCell.json()).toMatchObject({
      classification: { level: "Internal", labels: ["Mask"] },
      source: { scope: "cell", selectorKey: `cell:${docId}:Sheet1:0,0` }
    });

    const resolveInRange = await app.inject({
      method: "POST",
      url: `/docs/${docId}/classifications/resolve`,
      headers: { cookie: ownerCookie },
      payload: {
        selector: {
          scope: "cell",
          documentId: docId,
          sheetId: "Sheet1",
          row: 0,
          col: 1
        }
      }
    });
    expect(resolveInRange.statusCode).toBe(200);
    expect(resolveInRange.json()).toMatchObject({
      classification: { level: "Restricted", labels: ["PII"] },
      source: { scope: "range", selectorKey: `range:${docId}:Sheet1:0,0:2,2` }
    });

    const evaluateCopy = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie: ownerCookie },
      payload: {
        action: "clipboard.copy",
        selector: {
          scope: "cell",
          documentId: docId,
          sheetId: "Sheet1",
          row: 0,
          col: 1
        }
      }
    });
    expect(evaluateCopy.statusCode).toBe(200);
    expect(evaluateCopy.json()).toMatchObject({
      decision: "block",
      classification: { level: "Restricted", labels: ["PII"] },
      maxAllowed: "Confidential"
    });

    const evaluateRange = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie: ownerCookie },
      payload: {
        action: "clipboard.copy",
        selector: {
          scope: "range",
          documentId: docId,
          sheetId: "Sheet1",
          range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }
        }
      }
    });
    expect(evaluateRange.statusCode).toBe(200);
    expect(evaluateRange.json()).toMatchObject({
      decision: "block",
      classification: { level: "Restricted", labels: ["Mask", "PII"] },
      maxAllowed: "Confidential"
    });
  });
});
