import crypto from "node:crypto";
import { authenticator } from "otplib";
import type { Pool } from "pg";
import type { AuthenticatedUser } from "./sessions";

authenticator.options = { window: 1 };

export function generateTotpSecret(): string {
  return authenticator.generateSecret();
}

export function buildOtpAuthUrl(params: { issuer: string; accountName: string; secret: string }): string {
  return authenticator.keyuri(params.accountName, params.issuer, params.secret);
}

export function verifyTotpCode(secret: string, code: string): boolean {
  return authenticator.check(code, secret);
}

export function totpSecretName(userId: string): string {
  return `mfa:totp:${userId}`;
}

export async function isMfaEnforcedForOrg(pool: Pool, orgId: string): Promise<boolean> {
  const result = await pool.query("SELECT require_mfa FROM org_settings WHERE org_id = $1", [orgId]);
  if (result.rowCount !== 1) return false;
  return Boolean(result.rows[0]?.require_mfa);
}

export async function requireOrgMfaSatisfied(pool: Pool, orgId: string, user: AuthenticatedUser): Promise<boolean> {
  if (!(await isMfaEnforcedForOrg(pool, orgId))) return true;
  return Boolean(user.mfaTotpEnabled);
}

export function generateRecoveryCode(): string {
  // Keep codes URL-safe and copy/paste friendly.
  return crypto.randomBytes(10).toString("base64url");
}

export function hashRecoveryCode(code: string, saltHex?: string): string {
  const salt = saltHex ?? crypto.randomBytes(16).toString("hex");
  const digest = crypto.createHash("sha256").update(salt, "utf8").update(code, "utf8").digest("hex");
  return `sha256:${salt}:${digest}`;
}

export function verifyRecoveryCode(code: string, storedHash: string): boolean {
  const [algo, salt, digest] = storedHash.split(":");
  if (algo !== "sha256" || !salt || !digest) return false;

  const computed = crypto.createHash("sha256").update(salt, "utf8").update(code, "utf8").digest("hex");
  try {
    const a = Buffer.from(digest, "hex");
    const b = Buffer.from(computed, "hex");
    if (a.length !== b.length) return false;
    return crypto.timingSafeEqual(a, b);
  } catch {
    return false;
  }
}
