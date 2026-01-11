import crypto from "node:crypto";
import net from "node:net";
import type { Pool } from "pg";
import type { AuthenticatedUser } from "./sessions";

export type ApiKeyAuthMethod = "api_key";

export interface ApiKeyInfo {
  id: string;
  orgId: string;
  name: string;
  createdBy: string;
  createdAt: Date;
  lastUsedAt: Date | null;
  revokedAt: Date | null;
}

export type ApiKeyTokenParts = {
  id: string;
  secret: string;
};

export function generateApiKeyToken(): { apiKeyId: string; token: string; secret: string } {
  const apiKeyId = crypto.randomUUID();
  // Keep secrets URL/header safe; base64url avoids `/+` characters.
  const secret = crypto.randomBytes(32).toString("base64url");
  return { apiKeyId, secret, token: `api_${apiKeyId}.${secret}` };
}

export function parseApiKeyToken(token: string): ApiKeyTokenParts | null {
  if (!token.startsWith("api_")) return null;
  const rest = token.slice("api_".length);
  const dot = rest.indexOf(".");
  if (dot <= 0) return null;
  const id = rest.slice(0, dot);
  const secret = rest.slice(dot + 1);
  if (!id || !secret) return null;
  // Avoid adding dependencies; validate UUID shape loosely.
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) return null;
  return { id, secret };
}

export function hashApiKeySecret(secret: string, saltHex?: string): string {
  const salt = saltHex ?? crypto.randomBytes(16).toString("hex");
  const digest = crypto.createHash("sha256").update(salt, "utf8").update(secret, "utf8").digest("hex");
  return `sha256:${salt}:${digest}`;
}

export function verifyApiKeySecret(secret: string, storedHash: string): boolean {
  const [algo, salt, digest] = storedHash.split(":");
  if (algo !== "sha256" || !salt || !digest) return false;
  const computed = crypto.createHash("sha256").update(salt, "utf8").update(secret, "utf8").digest("hex");
  try {
    const a = Buffer.from(digest, "hex");
    const b = Buffer.from(computed, "hex");
    if (a.length !== b.length) return false;
    return crypto.timingSafeEqual(a, b);
  } catch {
    return false;
  }
}

function parseStringArray(value: unknown): string[] {
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value) as unknown;
      if (Array.isArray(parsed)) value = parsed;
    } catch {
      // Fall through.
    }
  }
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is string => typeof entry === "string" && entry.length > 0);
}

function stripIpZone(value: string): string {
  const idx = value.indexOf("%");
  return idx === -1 ? value : value.slice(0, idx);
}

type ParsedIp = { kind: 4 | 6; value: bigint };

function parseIpv4ToBigInt(ip: string): bigint | null {
  const parts = ip.split(".");
  if (parts.length !== 4) return null;
  let value = 0n;
  for (const part of parts) {
    if (!/^\d{1,3}$/.test(part)) return null;
    const n = Number(part);
    if (!Number.isInteger(n) || n < 0 || n > 255) return null;
    value = (value << 8n) | BigInt(n);
  }
  return value;
}

function expandIpv6Hextets(ip: string): number[] | null {
  const normalized = stripIpZone(ip).toLowerCase();
  if (normalized.length === 0) return null;

  const splitOnce = normalized.split("::");
  if (splitOnce.length > 2) return null;

  const head = splitOnce[0] ? splitOnce[0].split(":") : [];
  const tail = splitOnce.length === 2 && splitOnce[1] ? splitOnce[1].split(":") : [];

  const parsePart = (part: string): number[] | null => {
    if (part.includes(".")) {
      const v4 = parseIpv4ToBigInt(part);
      if (v4 == null) return null;
      const hi = Number((v4 >> 16n) & 0xffffn);
      const lo = Number(v4 & 0xffffn);
      return [hi, lo];
    }

    if (part.length === 0) return null;
    if (!/^[0-9a-f]{1,4}$/.test(part)) return null;
    const n = Number.parseInt(part, 16);
    if (!Number.isInteger(n) || n < 0 || n > 0xffff) return null;
    return [n];
  };

  const headNums: number[] = [];
  for (const part of head) {
    if (part.length === 0) return null;
    const parsed = parsePart(part);
    if (!parsed) return null;
    headNums.push(...parsed);
  }

  const tailNums: number[] = [];
  for (const part of tail) {
    if (part.length === 0) return null;
    const parsed = parsePart(part);
    if (!parsed) return null;
    tailNums.push(...parsed);
  }

  const total = headNums.length + tailNums.length;
  if (total > 8) return null;

  const zerosToInsert = splitOnce.length === 2 ? 8 - total : 0;
  if (splitOnce.length === 1 && total !== 8) return null;

  return [...headNums, ...Array.from({ length: zerosToInsert }, () => 0), ...tailNums];
}

function parseIp(value: string): ParsedIp | null {
  const cleaned = stripIpZone(value);
  const kind = net.isIP(cleaned);
  if (kind === 4) {
    const v4 = parseIpv4ToBigInt(cleaned);
    return v4 == null ? null : { kind: 4, value: v4 };
  }
  if (kind === 6) {
    const hextets = expandIpv6Hextets(cleaned);
    if (!hextets) return null;
    // Normalize IPv4-mapped IPv6 (::ffff:a.b.c.d) down to IPv4 when possible so allowlists don't
    // need to guess which representation Fastify will provide.
    const isV4Mapped =
      hextets.slice(0, 5).every((n) => n === 0) && (hextets[5] === 0xffff || hextets[5] === 0);
    if (isV4Mapped && hextets[5] === 0xffff) {
      const v4 = (BigInt(hextets[6]!) << 16n) | BigInt(hextets[7]!);
      return { kind: 4, value: v4 };
    }

    let out = 0n;
    for (const hextet of hextets) {
      out = (out << 16n) | BigInt(hextet);
    }
    return { kind: 6, value: out };
  }
  return null;
}

function ipMatchesCidr(client: ParsedIp, cidr: string): boolean {
  const [addr, prefixRaw] = cidr.split("/");
  if (!addr || !prefixRaw) return false;
  const base = parseIp(addr);
  if (!base || base.kind !== client.kind) return false;

  const bits = client.kind === 4 ? 32 : 128;
  const prefix = Number(prefixRaw);
  if (!Number.isInteger(prefix) || prefix < 0 || prefix > bits) return false;

  const shift = BigInt(bits - prefix);
  if (shift === BigInt(bits)) return true;
  return (base.value >> shift) === (client.value >> shift);
}

export function isClientIpAllowed(clientIp: string | null, ipAllowlist: unknown): boolean {
  const entries = parseStringArray(ipAllowlist);
  if (entries.length === 0) return true;
  if (!clientIp) return false;
  const parsedClient = parseIp(clientIp);
  if (!parsedClient) return false;

  for (const entry of entries) {
    if (entry.includes("/")) {
      if (ipMatchesCidr(parsedClient, entry)) return true;
      continue;
    }
    const parsedEntry = parseIp(entry);
    if (!parsedEntry) continue;
    if (parsedEntry.kind === parsedClient.kind && parsedEntry.value === parsedClient.value) return true;
  }

  return false;
}

export interface ApiKeyAuthSuccess {
  apiKey: ApiKeyInfo;
  user: AuthenticatedUser;
}

export type ApiKeyAuthFailure =
  | { statusCode: 401; error: "unauthorized" }
  | { statusCode: 403; error: "auth_method_not_allowed" }
  | { statusCode: 403; error: "ip_not_allowed" };

export type ApiKeyAuthResult = { ok: true; value: ApiKeyAuthSuccess } | { ok: false; value: ApiKeyAuthFailure };

export async function authenticateApiKey(
  pool: Pool,
  rawToken: string,
  options: { clientIp: string | null }
): Promise<ApiKeyAuthResult> {
  const parsed = parseApiKeyToken(rawToken);
  if (!parsed) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };

  const result = await pool.query(
    `
      SELECT
        ak.id AS api_key_id,
        ak.org_id AS api_key_org_id,
        ak.name AS api_key_name,
        ak.key_hash AS api_key_hash,
        ak.created_by AS api_key_created_by,
        ak.created_at AS api_key_created_at,
        ak.last_used_at AS api_key_last_used_at,
        ak.revoked_at AS api_key_revoked_at,
        u.id AS user_id,
        u.email AS user_email,
        u.name AS user_name,
        u.mfa_totp_enabled AS user_mfa_totp_enabled,
        os.allowed_auth_methods AS org_allowed_auth_methods,
        os.ip_allowlist AS org_ip_allowlist
      FROM api_keys ak
      JOIN users u ON u.id = ak.created_by
      JOIN org_settings os ON os.org_id = ak.org_id
      WHERE ak.id = $1
      LIMIT 1
    `,
    [parsed.id]
  );

  if (result.rowCount !== 1) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };

  const row = result.rows[0] as {
    api_key_id: string;
    api_key_org_id: string;
    api_key_name: string;
    api_key_hash: string;
    api_key_created_by: string;
    api_key_created_at: Date;
    api_key_last_used_at: Date | null;
    api_key_revoked_at: Date | null;
    user_id: string;
    user_email: string;
    user_name: string;
    user_mfa_totp_enabled: boolean;
    org_allowed_auth_methods: unknown;
    org_ip_allowlist: unknown;
  };

  if (row.api_key_revoked_at) return { ok: false, value: { statusCode: 401, error: "unauthorized" } };
  if (!verifyApiKeySecret(parsed.secret, row.api_key_hash)) {
    return { ok: false, value: { statusCode: 401, error: "unauthorized" } };
  }

  const allowedAuthMethods = parseStringArray(row.org_allowed_auth_methods);
  if (!allowedAuthMethods.includes("api_key")) {
    return { ok: false, value: { statusCode: 403, error: "auth_method_not_allowed" } };
  }

  if (!isClientIpAllowed(options.clientIp, row.org_ip_allowlist)) {
    return { ok: false, value: { statusCode: 403, error: "ip_not_allowed" } };
  }

  await pool.query("UPDATE api_keys SET last_used_at = now() WHERE id = $1", [row.api_key_id]);

  return {
    ok: true,
    value: {
      apiKey: {
        id: row.api_key_id,
        orgId: row.api_key_org_id,
        name: row.api_key_name,
        createdBy: row.api_key_created_by,
        createdAt: new Date(row.api_key_created_at),
        lastUsedAt: row.api_key_last_used_at ? new Date(row.api_key_last_used_at) : null,
        revokedAt: row.api_key_revoked_at ? new Date(row.api_key_revoked_at) : null
      },
      user: {
        id: row.user_id,
        email: row.user_email,
        name: row.user_name,
        mfaTotpEnabled: row.user_mfa_totp_enabled
      }
    }
  };
}
