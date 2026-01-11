import crypto from "node:crypto";
import type { Pool } from "pg";
import { hashApiKeySecret, verifyApiKeySecret } from "./apiKeys";

export type ScimAuthMethod = "scim";

export type ScimTokenParts = {
  orgId: string;
  secret: string;
};

export function generateScimToken(orgId: string): { token: string; secret: string } {
  // Keep secrets URL/header safe; base64url avoids `/+` characters.
  const secret = crypto.randomBytes(32).toString("base64url");
  return { secret, token: `scim_${orgId}.${secret}` };
}

export function parseScimToken(token: string): ScimTokenParts | null {
  if (!token.startsWith("scim_")) return null;
  const rest = token.slice("scim_".length);
  const dot = rest.indexOf(".");
  if (dot <= 0) return null;
  const orgId = rest.slice(0, dot);
  const secret = rest.slice(dot + 1);
  if (!orgId || !secret) return null;
  // Avoid adding dependencies; validate UUID shape loosely.
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(orgId)) return null;
  return { orgId, secret };
}

export function hashScimTokenSecret(secret: string, saltHex?: string): string {
  return hashApiKeySecret(secret, saltHex);
}

export function verifyScimTokenSecret(secret: string, storedHash: string): boolean {
  return verifyApiKeySecret(secret, storedHash);
}

export interface ScimAuthSuccess {
  orgId: string;
}

export type ScimAuthFailure = { statusCode: 401; error: "unauthorized" };

export type ScimAuthResult = { ok: true; value: ScimAuthSuccess } | { ok: false; value: ScimAuthFailure };

export async function authenticateScimToken(pool: Pool, rawToken: string): Promise<ScimAuthResult> {
  const parsed = parseScimToken(rawToken);
  if (!parsed) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };

  const res = await pool.query(
    `
      SELECT org_id, token_hash, revoked_at
      FROM org_scim_tokens
      WHERE org_id = $1
      LIMIT 1
    `,
    [parsed.orgId]
  );

  if (res.rowCount !== 1) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };

  const row = res.rows[0] as { org_id: string; token_hash: string; revoked_at: Date | null };
  if (row.revoked_at) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };
  if (!verifyScimTokenSecret(parsed.secret, row.token_hash)) {
    return { ok: false, value: { statusCode: 401, error: "unauthorized" } };
  }

  return { ok: true, value: { orgId: row.org_id } };
}

