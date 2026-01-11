import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { authenticator } from "otplib";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  return path.resolve(here, "../../migrations");
}

function extractCookie(setCookieHeader: string | string[] | undefined): string {
  if (!setCookieHeader) throw new Error("missing set-cookie header");
  const raw = Array.isArray(setCookieHeader) ? setCookieHeader[0] : setCookieHeader;
  return raw.split(";")[0];
}

describe("MFA hardening: encrypted TOTP + recovery codes + org enforcement", () => {
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

  it("stores new TOTP secrets in the encrypted secret store (no plaintext in users table)", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "mfa-encrypted@example.com",
        password: "password1234",
        name: "Encrypted"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const userId = (register.json() as any).user.id as string;

    const setup = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie }
    });
    expect(setup.statusCode).toBe(200);
    const { secret } = setup.json() as any;
    expect(secret).toBeTypeOf("string");

    const userRow = await db.query("SELECT mfa_totp_secret_legacy FROM users WHERE id = $1", [userId]);
    expect(userRow.rowCount).toBe(1);
    expect(userRow.rows[0]?.mfa_totp_secret_legacy).toBeNull();

    const secretRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [`mfa:totp:${userId}`]);
    expect(secretRow.rowCount).toBe(1);
    expect(secretRow.rows[0]?.encrypted_value).toMatch(/^v2:legacy:/);
    expect(secretRow.rows[0]?.encrypted_value).not.toBe(secret);
  }, 30_000);

  it("migrates legacy plaintext TOTP secrets into the encrypted store on login", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "mfa-legacy@example.com",
        password: "password1234",
        name: "Legacy"
      }
    });
    expect(register.statusCode).toBe(200);
    const userId = (register.json() as any).user.id as string;

    const legacySecret = authenticator.generateSecret();
    await db.query("UPDATE users SET mfa_totp_secret_legacy = $1, mfa_totp_enabled = true WHERE id = $2", [
      legacySecret,
      userId
    ]);
    await db.query("DELETE FROM secrets WHERE name = $1", [`mfa:totp:${userId}`]);

    const login = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: {
        email: "mfa-legacy@example.com",
        password: "password1234",
        mfaCode: authenticator.generate(legacySecret)
      }
    });
    expect(login.statusCode).toBe(200);

    const userRow = await db.query("SELECT mfa_totp_secret_legacy FROM users WHERE id = $1", [userId]);
    expect(userRow.rows[0]?.mfa_totp_secret_legacy).toBeNull();

    const secretRow = await db.query("SELECT encrypted_value FROM secrets WHERE name = $1", [`mfa:totp:${userId}`]);
    expect(secretRow.rowCount).toBe(1);
  }, 30_000);

  it("supports recovery code login and ensures codes cannot be reused", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "mfa-recovery@example.com",
        password: "password1234",
        name: "Recovery"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);

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

    const regen = await app.inject({
      method: "POST",
      url: "/auth/mfa/recovery-codes/regenerate",
      headers: { cookie },
      payload: { password: "password1234" }
    });
    expect(regen.statusCode).toBe(200);
    const codes = (regen.json() as any).codes as string[];
    expect(Array.isArray(codes)).toBe(true);
    expect(codes.length).toBeGreaterThan(0);

    const meta = await app.inject({ method: "GET", url: "/auth/mfa/recovery-codes", headers: { cookie } });
    expect(meta.statusCode).toBe(200);
    expect((meta.json() as any).remaining).toBe(codes.length);

    const recoveryCode = codes[0]!;

    const loginViaRecovery = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: {
        email: "mfa-recovery@example.com",
        password: "password1234",
        mfaRecoveryCode: recoveryCode
      }
    });
    expect(loginViaRecovery.statusCode).toBe(200);
    const loginCookie = extractCookie(loginViaRecovery.headers["set-cookie"]);

    const metaAfter = await app.inject({
      method: "GET",
      url: "/auth/mfa/recovery-codes",
      headers: { cookie: loginCookie }
    });
    expect(metaAfter.statusCode).toBe(200);
    expect((metaAfter.json() as any).remaining).toBe(codes.length - 1);

    const loginReuse = await app.inject({
      method: "POST",
      url: "/auth/login",
      payload: {
        email: "mfa-recovery@example.com",
        password: "password1234",
        mfaRecoveryCode: recoveryCode
      }
    });
    expect(loginReuse.statusCode).toBe(401);
    expect((loginReuse.json() as any).error).toBe("mfa_required");
  }, 30_000);

  it("enforces org require_mfa consistently on protected doc endpoints", async () => {
    const owner = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "mfa-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "MFA Org"
      }
    });
    expect(owner.statusCode).toBe(200);
    const ownerCookie = extractCookie(owner.headers["set-cookie"]);
    const orgId = (owner.json() as any).organization.id as string;

    const created = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie: ownerCookie },
      payload: { orgId, title: "Private Doc" }
    });
    expect(created.statusCode).toBe(200);
    const docId = (created.json() as any).document.id as string;

    const enablePolicy = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie: ownerCookie },
      payload: { requireMfa: true }
    });
    expect(enablePolicy.statusCode).toBe(200);

    const blockedRead = await app.inject({
      method: "GET",
      url: `/docs/${docId}`,
      headers: { cookie: ownerCookie }
    });
    expect(blockedRead.statusCode).toBe(403);
    expect((blockedRead.json() as any).error).toBe("mfa_required");

    const blockedShareLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "private", role: "viewer" }
    });
    expect(blockedShareLink.statusCode).toBe(403);
    expect((blockedShareLink.json() as any).error).toBe("mfa_required");

    const setupOwner = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie: ownerCookie }
    });
    const ownerSecret = (setupOwner.json() as any).secret as string;

    const confirmOwner = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie: ownerCookie },
      payload: { code: authenticator.generate(ownerSecret) }
    });
    expect(confirmOwner.statusCode).toBe(200);

    const shareLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie: ownerCookie },
      payload: { visibility: "public", role: "viewer" }
    });
    expect(shareLink.statusCode).toBe(200);
    const token = (shareLink.json() as any).shareLink.token as string;

    const bob = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "mfa-bob@example.com",
        password: "password1234",
        name: "Bob"
      }
    });
    expect(bob.statusCode).toBe(200);
    const bobCookie = extractCookie(bob.headers["set-cookie"]);

    const blockedRedeem = await app.inject({
      method: "POST",
      url: `/share-links/${token}/redeem`,
      headers: { cookie: bobCookie }
    });
    expect(blockedRedeem.statusCode).toBe(403);
    expect((blockedRedeem.json() as any).error).toBe("mfa_required");

    const setupBob = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/setup",
      headers: { cookie: bobCookie }
    });
    const bobSecret = (setupBob.json() as any).secret as string;

    const confirmBob = await app.inject({
      method: "POST",
      url: "/auth/mfa/totp/confirm",
      headers: { cookie: bobCookie },
      payload: { code: authenticator.generate(bobSecret) }
    });
    expect(confirmBob.statusCode).toBe(200);

    const redeemed = await app.inject({
      method: "POST",
      url: `/share-links/${token}/redeem`,
      headers: { cookie: bobCookie }
    });
    expect(redeemed.statusCode).toBe(200);
  }, 30_000);
});
