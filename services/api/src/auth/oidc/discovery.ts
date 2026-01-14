import { fetch as undiciFetch } from "undici";

export type OidcDiscoveryDocument = {
  issuer: string;
  authorization_endpoint: string;
  token_endpoint: string;
  jwks_uri: string;
};

type CacheEntry = { fetchedAt: number; doc: OidcDiscoveryDocument };

const discoveryCache = new Map<string, CacheEntry>();
const DISCOVERY_TTL_MS = 10 * 60 * 1000;

function normalizeIssuerUrl(issuerUrl: string): string {
  return issuerUrl.endsWith("/") ? issuerUrl.slice(0, -1) : issuerUrl;
}

function requireString(obj: Record<string, unknown>, key: keyof OidcDiscoveryDocument): string {
  const value = obj[key];
  if (typeof value !== "string" || value.length === 0) throw new Error(`OIDC discovery missing ${key}`);
  return value;
}

export async function getOidcDiscovery(issuerUrl: string): Promise<OidcDiscoveryDocument> {
  const normalized = normalizeIssuerUrl(issuerUrl);
  const cached = discoveryCache.get(normalized);
  const now = Date.now();
  if (cached && now - cached.fetchedAt < DISCOVERY_TTL_MS) return cached.doc;

  const discoveryUrl = new URL("/.well-known/openid-configuration", normalized).toString();
  // TODO(data-residency): OIDC discovery/token exchange is an outbound integration.
  // Enforce org data residency once we have a strategy to map IdP endpoints to regions.
  const res = await undiciFetch(discoveryUrl, { signal: AbortSignal.timeout(5000) });
  if (!res.ok) {
    throw new Error(`OIDC discovery failed (${res.status})`);
  }
  const json = (await res.json()) as Record<string, unknown>;

  const doc: OidcDiscoveryDocument = {
    issuer: requireString(json, "issuer"),
    authorization_endpoint: requireString(json, "authorization_endpoint"),
    token_endpoint: requireString(json, "token_endpoint"),
    jwks_uri: requireString(json, "jwks_uri")
  };

  discoveryCache.set(normalized, { fetchedAt: now, doc });
  return doc;
}
