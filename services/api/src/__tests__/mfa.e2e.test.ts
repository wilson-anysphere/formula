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
      headers: { cookie }
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
});
