import type { FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { fetch as undiciFetch } from "undici";
import jwt, { type Algorithm } from "jsonwebtoken";
import type { Pool, PoolClient } from "pg";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../../audit/audit";
import { createSession } from "../sessions";
import { withTransaction } from "../../db/tx";
import { getClientIp, getUserAgent } from "../../http/request-meta";
import { getSecret } from "../../secrets/secretStore";
import { getOidcDiscovery } from "./discovery";
import { getJwksKeys, jwkToPublicKey, type Jwk } from "./jwks";
import { randomBase64Url, sha256Base64Url } from "./pkce";

type OrgOidcProviderRow = {
  org_id: string;
  provider_id: string;
  issuer_url: string;
  client_id: string;
  scopes: unknown;
  enabled: boolean;
};

type OidcAuthStateRow = {
  state: string;
  org_id: string;
  provider_id: string;
  nonce: string;
  pkce_verifier: string;
  redirect_uri: string;
  created_at: Date;
};

type OrgAuthSettingsRow = {
  allowed_auth_methods: unknown;
  require_mfa: boolean;
};

export const OIDC_AUTH_STATE_TTL_MS = 10 * 60 * 1000;

type Queryable = Pick<Pool, "query"> | Pick<PoolClient, "query">;

export async function cleanupOidcAuthStates(db: Queryable): Promise<number> {
  const cutoff = new Date(Date.now() - OIDC_AUTH_STATE_TTL_MS);
  const res = await db.query("DELETE FROM oidc_auth_states WHERE created_at < $1", [cutoff]);
  return typeof res?.rowCount === "number" ? res.rowCount : 0;
}

function parseStringArray(value: unknown): string[] {
  if (!value) return [];
  if (Array.isArray(value)) return value.filter((v) => typeof v === "string");
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value) as unknown;
      if (Array.isArray(parsed)) return parsed.filter((v) => typeof v === "string");
    } catch {
      // fall through
    }
  }
  return [];
}

function ensureOpenIdScope(scopes: string[]): string[] {
  const normalized = scopes.map((s) => s.trim()).filter((s) => s.length > 0);
  const unique = new Set(normalized);
  unique.delete("openid");
  return ["openid", ...Array.from(unique)];
}

function extractEmail(claims: Record<string, unknown>): string | null {
  const candidates = [claims.email, claims.preferred_username, claims.upn];
  for (const value of candidates) {
    if (typeof value === "string" && value.trim().length > 0) return value.trim().toLowerCase();
  }
  return null;
}

function extractName(claims: Record<string, unknown>, email: string): string {
  const name = claims.name;
  if (typeof name === "string" && name.trim().length > 0) return name.trim();
  const local = email.split("@")[0];
  return local && local.length > 0 ? local : "User";
}

function isValidProviderId(value: string): boolean {
  return /^[a-z0-9_-]{1,64}$/.test(value);
}

function isValidOrgId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function extractHost(request: FastifyRequest): string {
  // Only trust forwarded headers when the API is configured to trust the proxy.
  const trustProxy = Boolean(request.server.config.trustProxy);
  const xfHost = trustProxy ? request.headers["x-forwarded-host"] : undefined;
  const hostValue =
    typeof xfHost === "string" && xfHost.length > 0
      ? xfHost.split(",")[0]!.trim()
      : typeof request.headers.host === "string"
        ? request.headers.host
        : "localhost";
  return hostValue.length > 0 ? hostValue : "localhost";
}

function extractProto(request: FastifyRequest): string {
  // Only trust forwarded headers when the API is configured to trust the proxy.
  const trustProxy = Boolean(request.server.config.trustProxy);
  const xfProto = trustProxy ? request.headers["x-forwarded-proto"] : undefined;
  const proto =
    typeof xfProto === "string" && xfProto.length > 0 ? xfProto.split(",")[0]!.trim() : request.protocol;
  return proto === "https" ? "https" : "http";
}

function externalBaseUrl(request: FastifyRequest): string {
  const configured = request.server.config.publicBaseUrl;
  if (typeof configured === "string" && configured.length > 0) return configured;

  if (!request.server.config.trustProxy) {
    throw new Error("PUBLIC_BASE_URL is required when trustProxy is disabled");
  }

  const host = extractHost(request);
  if (!isHostAllowed(host, request.server.config.publicBaseUrlHostAllowlist)) {
    throw new Error("Untrusted host for OIDC redirect URI");
  }

  return `${extractProto(request)}://${host}`;
}

function isHostAllowed(host: string, allowlist: string[]): boolean {
  let parsedHost: URL;
  try {
    parsedHost = new URL(`http://${host}`);
  } catch {
    return false;
  }

  const hostLower = parsedHost.host.toLowerCase();
  const hostnameLower = parsedHost.hostname.toLowerCase();

  for (const entry of allowlist) {
    const trimmed = entry.trim().toLowerCase();
    if (!trimmed) continue;

    try {
      const parsedEntry = new URL(`http://${trimmed}`);
      // If the entry has an explicit port, require an exact host match.
      if (parsedEntry.port) {
        if (parsedEntry.host.toLowerCase() === hostLower) return true;
        continue;
      }
      // Otherwise treat it as a hostname allowlist entry (port-agnostic).
      if (parsedEntry.hostname.toLowerCase() === hostnameLower) return true;
    } catch {
      // Fall back to a raw hostname comparison for entries that aren't parseable
      // as host specs (should be rare).
      if (trimmed === hostnameLower) return true;
    }
  }

  return false;
}

function buildRedirectUri(baseUrl: string, pathname: string): string {
  const base = new URL(baseUrl);
  // `PUBLIC_BASE_URL` should be a stable origin/prefix, not a full request URL.
  // Strip any search/hash so we don't accidentally embed it into redirect_uri.
  base.search = "";
  base.hash = "";
  // Ensure the base URL behaves like a directory for URL resolution so optional
  // path prefixes (e.g. https://example.com/api/) are preserved.
  if (!base.pathname.endsWith("/")) base.pathname = `${base.pathname}/`;
  const relative = pathname.startsWith("/") ? pathname.slice(1) : pathname;
  return new URL(relative, base).toString();
}

function oidcCallbackPath(orgId: string, providerId: string): string {
  return `/auth/oidc/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`;
}

function redirectUriAllowed(request: FastifyRequest, redirectUri: string, orgId: string, providerId: string): boolean {
  let parsed: URL;
  try {
    parsed = new URL(redirectUri);
  } catch {
    return false;
  }

  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") return false;

  const callbackPath = oidcCallbackPath(orgId, providerId);
  // Allow optional base path prefixes (e.g. https://example.com/api) by only checking
  // that the callback suffix matches.
  if (!parsed.pathname.endsWith(callbackPath)) return false;

  const configured = request.server.config.publicBaseUrl;
  if (typeof configured === "string" && configured.length > 0) {
    try {
      return redirectUri === buildRedirectUri(configured, callbackPath);
    } catch {
      return false;
    }
  }

  // Dev/test fallback: without a configured base URL we only allow callback hosts
  // that are explicitly allowlisted *and* the server is configured to trust the proxy.
  if (!request.server.config.trustProxy) return false;
  return isHostAllowed(parsed.host, request.server.config.publicBaseUrlHostAllowlist);
}

async function loadOrgSettings(
  request: FastifyRequest,
  orgId: string
): Promise<{ allowedAuthMethods: string[]; requireMfa: boolean } | null> {
  const res = await request.server.db.query<OrgAuthSettingsRow>(
    "SELECT allowed_auth_methods, require_mfa FROM org_settings WHERE org_id = $1",
    [orgId]
  );
  if (res.rowCount !== 1) return null;
  const row = res.rows[0] as OrgAuthSettingsRow;
  return {
    allowedAuthMethods: parseStringArray(row.allowed_auth_methods),
    requireMfa: Boolean(row.require_mfa)
  };
}

async function loadOrgProvider(
  request: FastifyRequest,
  orgId: string,
  providerId: string
): Promise<{
  issuerUrl: string;
  clientId: string;
  scopes: string[];
  enabled: boolean;
} | null> {
  const res = await request.server.db.query<OrgOidcProviderRow>(
    `
      SELECT org_id, provider_id, issuer_url, client_id, scopes, enabled
      FROM org_oidc_providers
      WHERE org_id = $1 AND provider_id = $2
      LIMIT 1
    `,
    [orgId, providerId]
  );
  if (res.rowCount !== 1) return null;
  const row = res.rows[0] as OrgOidcProviderRow;
  return {
    issuerUrl: String(row.issuer_url),
    clientId: String(row.client_id),
    scopes: parseStringArray(row.scopes),
    enabled: Boolean(row.enabled)
  };
}

function tokenClaimsIndicateMfa(claims: Record<string, unknown>): boolean {
  const amr = claims.amr;
  if (Array.isArray(amr)) {
    const normalized = amr.filter((v) => typeof v === "string").map((v) => v.toLowerCase());
    if (normalized.includes("mfa")) return true;
    if (normalized.includes("otp")) return true;
    if (normalized.includes("totp")) return true;
  }

  const acr = claims.acr;
  if (typeof acr === "string" && acr.toLowerCase().includes("mfa")) return true;
  return false;
}

async function exchangeCodeForTokens(options: {
  tokenEndpoint: string;
  code: string;
  redirectUri: string;
  clientId: string;
  clientSecret: string;
  codeVerifier: string;
}): Promise<{ idToken: string }> {
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    code: options.code,
    redirect_uri: options.redirectUri,
    client_id: options.clientId,
    client_secret: options.clientSecret,
    code_verifier: options.codeVerifier
  }).toString();

  // TODO(data-residency): OIDC token exchange is an outbound integration.
  // Enforce org data residency once we have a strategy to map IdP endpoints to regions.
  const res = await undiciFetch(options.tokenEndpoint, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
    signal: AbortSignal.timeout(5000)
  });

  if (!res.ok) {
    throw new Error(`OIDC token exchange failed (${res.status})`);
  }

  const json = (await res.json()) as Record<string, unknown>;
  const idToken = json.id_token;
  if (typeof idToken !== "string" || idToken.length === 0) {
    throw new Error("OIDC token response missing id_token");
  }

  return { idToken };
}

async function verifyIdToken(options: {
  idToken: string;
  issuer: string;
  audience: string;
  jwksUri: string;
  expectedNonce: string;
}): Promise<Record<string, unknown>> {
  const allowedAlgorithms: Algorithm[] = ["RS256", "RS384", "RS512", "ES256", "ES384", "ES512"];
  const decoded = jwt.decode(options.idToken, { complete: true }) as
    | { header?: Record<string, unknown> }
    | null;
  const header = decoded?.header ?? null;
  const kid = typeof header?.kid === "string" ? header.kid : undefined;
  const alg = typeof header?.alg === "string" ? header.alg : undefined;
  if (!alg) throw new Error("OIDC id_token missing alg");
  if (!allowedAlgorithms.includes(alg as Algorithm)) {
    throw new Error(`OIDC id_token alg not allowed: ${alg}`);
  }

  const keys = await getJwksKeys(options.jwksUri);
  let jwk: Jwk | undefined;
  if (kid) jwk = keys.find((k) => k.kid === kid);
  if (!jwk && keys.length === 1) jwk = keys[0];
  if (!jwk) throw new Error("OIDC signing key not found");

  const publicKey = jwkToPublicKey(jwk);
  const claims = jwt.verify(options.idToken, publicKey, {
    algorithms: allowedAlgorithms,
    issuer: options.issuer,
    audience: options.audience
  }) as unknown;

  if (!claims || typeof claims !== "object") throw new Error("OIDC id_token invalid");
  const record = claims as Record<string, unknown>;

  if (record.nonce !== options.expectedNonce) throw new Error("OIDC nonce mismatch");
  return record;
}

async function writeOidcFailureAudit(options: {
  request: FastifyRequest;
  orgId: string | null;
  providerId: string | null;
  userId?: string | null;
  userEmail?: string | null;
  errorCode: string;
  errorMessage?: string;
}): Promise<void> {
  const actor = options.userId
    ? { type: "user", id: options.userId }
    : options.userEmail
      ? { type: "anonymous", id: options.userEmail }
      : { type: "anonymous", id: `oidc:${options.providerId ?? "unknown"}` };

  const event = createAuditEvent({
    eventType: "auth.login_failed",
    actor,
    context: {
      orgId: options.orgId,
      userId: options.userId ?? null,
      userEmail: options.userEmail ?? null,
      ipAddress: getClientIp(options.request),
      userAgent: getUserAgent(options.request)
    },
    resource: { type: "user", id: options.userId ?? null },
    success: false,
    error: { code: options.errorCode, message: options.errorMessage },
    details: { method: "oidc", provider: options.providerId }
  });

  try {
    await writeAuditEvent(options.request.server.db, event);
  } catch (err) {
    // The audit log has a FK to organizations; if the org id is invalid, we still
    // want a best-effort record of the failed login attempt.
    if (options.orgId) {
      try {
        await writeAuditEvent(
          options.request.server.db,
          createAuditEvent({
            ...event,
            // Overwrite just the context.orgId; preserve other auto-generated fields.
            context: { ...event.context, orgId: null }
          })
        );
      } catch {
        // Ignore audit failures; authentication code paths must not fail closed
        // due to observability plumbing.
      }
      return;
    }
    // If orgId was already null, ignore the failure.
  }
}

export async function oidcStart(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;
  if (!isValidOrgId(orgId) || !isValidProviderId(providerId)) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  // Best-effort cleanup so `oidc_auth_states` does not grow unbounded when sweeps are disabled.
  try {
    await cleanupOidcAuthStates(request.server.db);
  } catch (err) {
    request.server.log.warn({ err }, "oidc_auth_state_cleanup_failed");
  }

  const provider = await loadOrgProvider(request, orgId, providerId);
  if (!provider) {
    reply.code(404).send({ error: "provider_not_found" });
    return;
  }
  if (!provider.enabled) {
    reply.code(403).send({ error: "provider_disabled" });
    return;
  }

  const settings = await loadOrgSettings(request, orgId);
  if (!settings) {
    reply.code(404).send({ error: "org_not_found" });
    return;
  }
  if (!settings.allowedAuthMethods.includes("oidc")) {
    reply.code(403).send({ error: "auth_method_not_allowed" });
    return;
  }

  // Ensure the provider has a client secret configured before starting the flow.
  // Otherwise we'd redirect the user to the IdP only to fail the callback during
  // token exchange.
  const secretName = `oidc:${orgId}:${providerId}`;
  try {
    const secret = await getSecret(request.server.db, request.server.config.secretStoreKeys, secretName);
    if (!secret) {
      request.server.metrics.authFailuresTotal.inc({ reason: "oidc_not_configured" });
      await writeOidcFailureAudit({
        request,
        orgId,
        providerId,
        errorCode: "oidc_not_configured",
        errorMessage: "OIDC client secret not configured"
      });
      reply.code(500).send({ error: "oidc_not_configured" });
      return;
    }
  } catch (err) {
    request.server.log.warn({ err, orgId, providerId }, "oidc_client_secret_load_failed");
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_not_configured" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_not_configured",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(500).send({ error: "oidc_not_configured" });
    return;
  }

  let discovery;
  try {
    discovery = await getOidcDiscovery(provider.issuerUrl);
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_discovery_failed" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_discovery_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(502).send({ error: "oidc_discovery_failed" });
    return;
  }

  const state = randomBase64Url(32);
  const nonce = randomBase64Url(32);
  const pkceVerifier = randomBase64Url(32);
  const pkceChallenge = sha256Base64Url(pkceVerifier);

  let redirectUri: string;
  const callbackPath = oidcCallbackPath(orgId, providerId);
  try {
    redirectUri = buildRedirectUri(externalBaseUrl(request), callbackPath);
  } catch (err) {
    request.server.log.warn({ err }, "oidc_redirect_uri_base_url_invalid");
    reply.code(500).send({ error: "oidc_not_configured" });
    return;
  }

  try {
    await request.server.db.query(
      `
        INSERT INTO oidc_auth_states (state, org_id, provider_id, nonce, pkce_verifier, redirect_uri)
        VALUES ($1, $2, $3, $4, $5, $6)
      `,
      [state, orgId, providerId, nonce, pkceVerifier, redirectUri]
    );
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_state_store_failed" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_state_store_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(500).send({ error: "oidc_state_store_failed" });
    return;
  }

  const authUrl = new URL(discovery.authorization_endpoint);
  authUrl.search = new URLSearchParams({
    response_type: "code",
    client_id: provider.clientId,
    redirect_uri: redirectUri,
    scope: ensureOpenIdScope(provider.scopes).join(" "),
    state,
    nonce,
    code_challenge: pkceChallenge,
    code_challenge_method: "S256"
  }).toString();

  reply.redirect(authUrl.toString());
}

const CallbackQuery = z.object({
  code: z.string().min(1),
  state: z.string().min(1),
  error: z.string().optional(),
  error_description: z.string().optional()
});

export async function oidcCallback(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;
  if (!isValidOrgId(orgId) || !isValidProviderId(providerId)) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  const parsed = CallbackQuery.safeParse(request.query);
  if (!parsed.success) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  if (parsed.data.error) {
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_error" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_error",
      errorMessage: parsed.data.error_description ?? parsed.data.error
    });
    reply.code(401).send({ error: "oidc_error" });
    return;
  }

  const authRes = await request.server.db.query<OidcAuthStateRow>(
    `
      DELETE FROM oidc_auth_states
      WHERE state = $1
      RETURNING state, org_id, provider_id, nonce, pkce_verifier, redirect_uri, created_at
    `,
    [parsed.data.state]
  );

  if (authRes.rowCount !== 1) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_state" });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode: "invalid_state" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  const authState = authRes.rows[0] as OidcAuthStateRow;
  if (authState.org_id !== orgId || authState.provider_id !== providerId) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_state" });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode: "invalid_state" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  const ageMs = Date.now() - new Date(authState.created_at).getTime();
  if (!Number.isFinite(ageMs) || ageMs > OIDC_AUTH_STATE_TTL_MS) {
    request.server.metrics.authFailuresTotal.inc({ reason: "state_expired" });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode: "state_expired" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  if (!redirectUriAllowed(request, authState.redirect_uri, orgId, providerId)) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_redirect_uri" });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode: "invalid_redirect_uri" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  const provider = await loadOrgProvider(request, orgId, providerId);
  if (!provider) {
    reply.code(404).send({ error: "provider_not_found" });
    return;
  }
  if (!provider.enabled) {
    reply.code(403).send({ error: "provider_disabled" });
    return;
  }

  const settings = await loadOrgSettings(request, orgId);
  if (!settings) {
    reply.code(404).send({ error: "org_not_found" });
    return;
  }
  if (!settings.allowedAuthMethods.includes("oidc")) {
    reply.code(403).send({ error: "auth_method_not_allowed" });
    return;
  }

  const secretName = `oidc:${orgId}:${providerId}`;
  const clientSecret = await getSecret(request.server.db, request.server.config.secretStoreKeys, secretName);
  if (!clientSecret) {
    reply.code(500).send({ error: "oidc_not_configured" });
    return;
  }

  let discovery;
  try {
    discovery = await getOidcDiscovery(provider.issuerUrl);
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_discovery_failed" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_discovery_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(502).send({ error: "oidc_discovery_failed" });
    return;
  }

  let idToken: string;
  try {
    const tokens = await exchangeCodeForTokens({
      tokenEndpoint: discovery.token_endpoint,
      code: parsed.data.code,
      redirectUri: authState.redirect_uri,
      clientId: provider.clientId,
      clientSecret,
      codeVerifier: authState.pkce_verifier
    });
    idToken = tokens.idToken;
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "oidc_token_exchange_failed" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "oidc_token_exchange_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(401).send({ error: "oidc_token_exchange_failed" });
    return;
  }

  let claims: Record<string, unknown>;
  try {
    claims = await verifyIdToken({
      idToken,
      issuer: discovery.issuer,
      audience: provider.clientId,
      jwksUri: discovery.jwks_uri,
      expectedNonce: authState.nonce
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "OIDC token invalid";
    const errorCode = message.includes("nonce") ? "invalid_nonce" : "invalid_id_token";
    request.server.metrics.authFailuresTotal.inc({ reason: errorCode });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode, errorMessage: message });
    reply.code(401).send({ error: errorCode });
    return;
  }

  const mfaSatisfied = tokenClaimsIndicateMfa(claims);
  if (settings.requireMfa && !mfaSatisfied) {
    request.server.metrics.authFailuresTotal.inc({ reason: "mfa_required" });
    await writeOidcFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "mfa_required",
      errorMessage: "Organization requires MFA"
    });
    reply.code(401).send({ error: "mfa_required" });
    return;
  }

  const subject = typeof claims.sub === "string" && claims.sub.length > 0 ? claims.sub : null;
  const email = extractEmail(claims);
  if (!subject || !email) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_claims" });
    await writeOidcFailureAudit({ request, orgId, providerId, errorCode: "invalid_claims" });
    reply.code(401).send({ error: "invalid_claims" });
    return;
  }

  const name = extractName(claims, email);

  const now = new Date();
  const expiresAt = new Date(now.getTime() + request.server.config.sessionTtlSeconds * 1000);

  const { userId, sessionId, token } = await withTransaction(request.server.db, async (client) => {
    const existingIdentity = await client.query<{ user_id: string }>(
      `
        SELECT user_id
        FROM user_identities
        WHERE org_id = $1 AND provider = $2 AND subject = $3
        LIMIT 1
      `,
      [orgId, providerId, subject]
    );

    let userId: string;

    if (existingIdentity.rowCount === 1) {
      userId = String((existingIdentity.rows[0] as any).user_id);
      await client.query(
        `
          UPDATE user_identities
          SET email = $4
          WHERE org_id = $1 AND provider = $2 AND subject = $3
        `,
        [orgId, providerId, subject, email]
      );
    } else {
      const existingUser = await client.query<{ id: string }>(
        "SELECT id FROM users WHERE email = $1 LIMIT 1",
        [email]
      );

      if (existingUser.rowCount === 1) {
        userId = String((existingUser.rows[0] as any).id);
      } else {
        userId = crypto.randomUUID();
        await client.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [userId, email, name]);
      }

      await client.query(
        `
          INSERT INTO user_identities (user_id, provider, subject, email, org_id)
          VALUES ($1, $2, $3, $4, $5)
          ON CONFLICT (org_id, provider, subject)
          DO UPDATE SET email = EXCLUDED.email
        `,
        [userId, providerId, subject, email, orgId]
      );
    }

    await client.query(
      `
        INSERT INTO org_members (org_id, user_id, role)
        VALUES ($1, $2, 'member')
        ON CONFLICT (org_id, user_id) DO NOTHING
      `,
      [orgId, userId]
    );

    const session = await createSession(client, {
      userId,
      expiresAt,
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request),
      mfaSatisfied
    });
    return { userId, sessionId: session.sessionId, token: session.token };
  });

  reply.setCookie(request.server.config.sessionCookieName, token, {
    path: "/",
    httpOnly: true,
    sameSite: "lax",
    secure: request.server.config.cookieSecure
  });

  await writeAuditEvent(
    request.server.db,
    createAuditEvent({
      eventType: "auth.login",
      actor: { type: "user", id: userId },
      context: {
        orgId,
        userId,
        userEmail: email,
        sessionId,
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      },
      resource: { type: "session", id: sessionId },
      success: true,
      details: { method: "oidc", provider: providerId }
    })
  );

  reply.send({ user: { id: userId, email, name }, orgId });
}
