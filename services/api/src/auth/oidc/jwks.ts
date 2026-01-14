import crypto from "node:crypto";
import { fetch as undiciFetch } from "undici";

export type Jwk = {
  kty: string;
  kid?: string;
  use?: string;
  alg?: string;
  n?: string;
  e?: string;
  crv?: string;
  x?: string;
  y?: string;
};

type CacheEntry = { fetchedAt: number; keys: Jwk[] };

const jwksCache = new Map<string, CacheEntry>();
const JWKS_TTL_MS = 10 * 60 * 1000;

export async function getJwksKeys(jwksUri: string): Promise<Jwk[]> {
  const cached = jwksCache.get(jwksUri);
  const now = Date.now();
  if (cached && now - cached.fetchedAt < JWKS_TTL_MS) return cached.keys;

  const res = await undiciFetch(jwksUri, { signal: AbortSignal.timeout(5000) });
  if (!res.ok) throw new Error(`OIDC JWKS fetch failed (${res.status})`);
  const json = (await res.json()) as { keys?: unknown };
  if (!Array.isArray(json.keys)) throw new Error("OIDC JWKS response missing keys");

  const keys = json.keys.filter((k) => k && typeof k === "object") as Jwk[];
  jwksCache.set(jwksUri, { fetchedAt: now, keys });
  return keys;
}

export function jwkToPublicKey(jwk: Jwk): crypto.KeyObject {
  // `crypto.createPublicKey` supports JWK inputs for RSA/ECDSA keys.
  // We intentionally accept the entire JWK payload (incl. kid/use/alg).
  return crypto.createPublicKey({ key: jwk as any, format: "jwk" });
}
