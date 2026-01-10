import { authenticator } from "otplib";
import type { Pool } from "pg";

export function generateTotpSecret(): string {
  return authenticator.generateSecret();
}

export function buildOtpAuthUrl(params: {
  issuer: string;
  accountName: string;
  secret: string;
}): string {
  return authenticator.keyuri(params.accountName, params.issuer, params.secret);
}

export function verifyTotpCode(secret: string, code: string): boolean {
  return authenticator.check(code, secret);
}

export async function isMfaEnforcedForOrg(pool: Pool, orgId: string): Promise<boolean> {
  const result = await pool.query("SELECT require_mfa FROM org_settings WHERE org_id = $1", [orgId]);
  if (result.rowCount !== 1) return false;
  return Boolean(result.rows[0]?.require_mfa);
}

