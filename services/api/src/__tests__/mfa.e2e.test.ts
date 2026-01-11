import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { authenticator } from "otplib";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { totpSecretName } from "../auth/mfa";
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

describe("MFA e2e: encrypted TOTP secrets + recovery codes + org enforcement", () => {
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

  it("stores TOTP secrets in the encrypted secret store (not users)", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "mfa-secret@example.com", password: "password1234", name: "MFA Secret", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const userId = (reg.json() as any).user.id as string;

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;
    expect(secret).toBeTypeOf("string");

    const userRow = await db.query("SELECT mfa_totp_enabled, mfa_totp_secret_legacy FROM users WHERE id = $1", [
      userId
    ]);
    expect(userRow.rowCount).toBe(1);
    expect(userRow.rows[0].mfa_totp_enabled).toBe(false);
    expect(userRow.rows[0].mfa_totp_secret_legacy).toBeNull();

    const stored = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [totpSecretName(userId)]);
    expect(stored.rowCount).toBe(1);
    expect(stored.rows[0].encrypted_value).toBeTypeOf("string");
    expect(String(stored.rows[0].encrypted_value)).toMatch(/^v(1|2):/);
    expect(String(stored.rows[0].encrypted_value)).not.toBe(secret);

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const enabled = await db.query("SELECT mfa_totp_enabled, mfa_totp_secret_legacy FROM users WHERE id = $1", [
      userId
    ]);
    expect(enabled.rows[0].mfa_totp_enabled).toBe(true);
    expect(enabled.rows[0].mfa_totp_secret_legacy).toBeNull();
  });

  it("requires a TOTP code at login when MFA is enabled", async () => {
    const email = "mfa-login@example.com";
    const password = "password1234";

    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email, password, name: "MFA Login", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const missingMfa = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email, password }
    });
    expect(missingMfa.statusCode).toBe(401);
    expect((missingMfa.json() as any).error).toBe("mfa_required");

    const wrongMfa = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email, password, mfaCode: "000000" }
    });
    expect(wrongMfa.statusCode).toBe(401);
    expect((wrongMfa.json() as any).error).toBe("mfa_required");

    const ok = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email, password, mfaCode: authenticator.generate(secret) }
    });
    expect(ok.statusCode).toBe(200);
    expect(ok.headers["set-cookie"]).toBeTruthy();
  });

  it("generates recovery codes and allows them to be used once for login", async () => {
    const email = "mfa-recovery@example.com";
    const password = "password1234";

    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email, password, name: "MFA Recovery", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const userId = (reg.json() as any).user.id as string;

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const regen = await app.inject({
      method: "POST",
      url: "/auth/mfa/recovery-codes/regenerate",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(regen.statusCode).toBe(200);
    const codes = (regen.json() as any).codes as string[];
    expect(Array.isArray(codes)).toBe(true);
    expect(codes).toHaveLength(10);

    const stored = await db.query(
      "SELECT code_hash, used_at FROM user_mfa_recovery_codes WHERE user_id = $1 ORDER BY created_at ASC",
      [userId]
    );
    expect(stored.rowCount).toBe(10);
    expect(stored.rows.every((row) => row.used_at == null)).toBe(true);
    expect(stored.rows.every((row) => typeof row.code_hash === "string" && row.code_hash.startsWith("sha256:"))).toBe(
      true
    );

    const first = codes[0]!;
    expect(stored.rows.some((row) => row.code_hash === first)).toBe(false);

    const loginViaRecovery = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email, password, mfaRecoveryCode: first }
    });
    expect(loginViaRecovery.statusCode).toBe(200);

    const usedCount = await db.query(
      "SELECT COUNT(*)::int AS c FROM user_mfa_recovery_codes WHERE user_id = $1 AND used_at IS NOT NULL",
      [userId]
    );
    expect(Number(usedCount.rows[0].c)).toBe(1);

    const reuse = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: { email, password, mfaRecoveryCode: first }
    });
    expect(reuse.statusCode).toBe(401);
    expect((reuse.json() as any).error).toBe("mfa_required");
  });

  it("enforces org require_mfa on sensitive org settings endpoints", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "org-mfa@example.com", password: "password1234", name: "Org Admin", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const userId = body.user.id as string;
    const orgId = body.organization.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { allowPublicLinks: false }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { allowPublicLinks: false }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).ok).toBe(true);

    const user = await db.query("SELECT mfa_totp_enabled FROM users WHERE id = $1", [userId]);
    expect(user.rows[0].mfa_totp_enabled).toBe(true);
  });

  it("enforces org require_mfa on share-link management endpoints", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "share-link-mfa@example.com", password: "password1234", name: "Share Admin", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Test doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie },
      payload: {}
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie },
      payload: {}
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).shareLink?.token).toBeTypeOf("string");
  });

  it("enforces org require_mfa on document admin actions", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "doc-admin-mfa@example.com", password: "password1234", name: "Doc Admin", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Admin doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}`,
      headers: { cookie }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}`,
      headers: { cookie }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).ok).toBe(true);
  });

  it("enforces org require_mfa on document metadata endpoints", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "doc-read-mfa@example.com", password: "password1234", name: "Doc Read", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Read doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "GET",
      url: `/docs/${docId}`,
      headers: { cookie }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "GET",
      url: `/docs/${docId}`,
      headers: { cookie }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).document.id).toBe(docId);
  });

  it("enforces org require_mfa on document version export endpoints", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "doc-versions-mfa@example.com", password: "password1234", name: "Doc Versions", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Versions doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    const bytes = Buffer.from("hello versions");
    const dataBase64 = bytes.toString("base64");
    const createVersion = await app.inject({
      method: "POST",
      url: `/docs/${docId}/versions`,
      headers: { cookie },
      payload: { description: "v1", dataBase64 }
    });
    expect(createVersion.statusCode).toBe(200);
    const versionId = (createVersion.json() as any).version.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions/${versionId}`,
      headers: { cookie }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions/${versionId}`,
      headers: { cookie }
    });
    expect(allowed.statusCode).toBe(200);
    const fetched = (allowed.json() as any).version as any;
    expect(Buffer.from(fetched.dataBase64, "base64").equals(bytes)).toBe(true);
  });

  it("enforces org require_mfa on range permission management endpoints", async () => {
    const email = "range-perms-mfa@example.com";
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email, password: "password1234", name: "Range Perms", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Range perms doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "POST",
      url: `/docs/${docId}/range-permissions`,
      headers: { cookie },
      payload: {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 1,
        endCol: 1,
        permissionType: "read",
        allowedUserEmail: email
      }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "POST",
      url: `/docs/${docId}/range-permissions`,
      headers: { cookie },
      payload: {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 1,
        endCol: 1,
        permissionType: "read",
        allowedUserEmail: email
      }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).ok).toBe(true);
    expect((allowed.json() as any).id).toBeTypeOf("string");
  });

  it("enforces org require_mfa on share-link redemption", async () => {
    const ownerReg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "share-redeem-owner@example.com", password: "password1234", name: "Owner", orgName: "Org" }
    });
    expect(ownerReg.statusCode).toBe(200);
    const ownerCookie = extractCookie(ownerReg.headers["set-cookie"]);
    const ownerBody = ownerReg.json() as any;
    const orgId = ownerBody.organization.id as string;

    const memberReg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "share-redeem-member@example.com", password: "password1234", name: "Member", orgName: "Other" }
    });
    expect(memberReg.statusCode).toBe(200);
    const memberCookie = extractCookie(memberReg.headers["set-cookie"]);

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Share redeem doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    const invite = await app.inject({
      method: "POST",
      url: `/docs/${docId}/invite`,
      headers: { cookie: ownerCookie },
      payload: { email: "share-redeem-member@example.com", role: "viewer" }
    });
    expect(invite.statusCode).toBe(200);

    const link = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: {}
    });
    expect(link.statusCode).toBe(200);
    const token = (link.json() as any).shareLink.token as string;
    expect(token).toBeTypeOf("string");

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "POST",
      url: `/share-links/${token}/redeem`,
      headers: { cookie: memberCookie }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie: memberCookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie: memberCookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "POST",
      url: `/share-links/${token}/redeem`,
      headers: { cookie: memberCookie }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).ok).toBe(true);
    expect((allowed.json() as any).documentId).toBe(docId);
  });

  it("enforces org require_mfa on audit ingest", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "audit-ingest-mfa@example.com", password: "password1234", name: "Audit", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/audit`,
      headers: { cookie },
      payload: { eventType: "client.test" }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/audit`,
      headers: { cookie },
      payload: { eventType: "client.test" }
    });
    expect(allowed.statusCode).toBe(202);
    expect((allowed.json() as any).id).toBeTypeOf("string");
  });

  it("enforces org require_mfa on DLP evaluation endpoints", async () => {
    const reg = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: { email: "dlp-mfa@example.com", password: "password1234", name: "DLP", orgName: "Org" }
    });
    expect(reg.statusCode).toBe(200);
    const cookie = extractCookie(reg.headers["set-cookie"]);
    const body = reg.json() as any;
    const orgId = body.organization.id as string;

    const createdDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "DLP doc" }
    });
    expect(createdDoc.statusCode).toBe(200);
    const docId = (createdDoc.json() as any).document.id as string;

    await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

    const blocked = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie },
      payload: { action: "sharing.externalLink" }
    });
    expect(blocked.statusCode).toBe(403);
    expect((blocked.json() as any).error).toBe("mfa_required");

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const secret = (setup.json() as any).secret as string;

    const confirm = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie },
      payload: { code: authenticator.generate(secret) }
    });
    expect(confirm.statusCode).toBe(200);

    const allowed = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie },
      payload: { action: "sharing.externalLink" }
    });
    expect(allowed.statusCode).toBe(200);
    expect((allowed.json() as any).decision).toBeTruthy();
  });
});
