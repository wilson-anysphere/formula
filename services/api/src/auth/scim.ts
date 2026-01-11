import crypto from "node:crypto";

export type ScimAuthMethod = "scim";

export interface ScimTokenInfo {
  id: string;
  orgId: string;
  name: string;
  createdBy: string;
  createdAt: Date;
  lastUsedAt: Date | null;
  revokedAt: Date | null;
}

export type ScimTokenParts = {
  id: string;
  secret: string;
};

export function generateScimToken(): { tokenId: string; token: string; secret: string } {
  const tokenId = crypto.randomUUID();
  // Keep secrets URL/header safe; base64url avoids `/+` characters.
  const secret = crypto.randomBytes(32).toString("base64url");
  return { tokenId, secret, token: `scim_${tokenId}.${secret}` };
}

export function parseScimToken(token: string): ScimTokenParts | null {
  if (!token.startsWith("scim_")) return null;
  const rest = token.slice("scim_".length);
  const dot = rest.indexOf(".");
  if (dot <= 0) return null;
  const id = rest.slice(0, dot);
  const secret = rest.slice(dot + 1);
  if (!id || !secret) return null;
  // Avoid adding dependencies; validate UUID shape loosely.
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) return null;
  return { id, secret };
}

export function hashScimTokenSecret(secret: string, saltHex?: string): string {
  const salt = saltHex ?? crypto.randomBytes(16).toString("hex");
  const digest = crypto.createHash("sha256").update(salt, "utf8").update(secret, "utf8").digest("hex");
  return `sha256:${salt}:${digest}`;
}

export function verifyScimTokenSecret(secret: string, storedHash: string): boolean {
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

