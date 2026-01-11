import type { FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { SAML, ValidateInResponseTo, type CacheItem, type CacheProvider } from "@node-saml/node-saml";
import type { Pool } from "pg";
import { createAuditEvent, writeAuditEvent } from "../../audit/audit";
import { withTransaction } from "../../db/tx";
import { getClientIp, getUserAgent } from "../../http/request-meta";
import { randomBase64Url } from "../oidc/pkce";
import { createSession } from "../sessions";

type OrgSamlProviderRow = {
  org_id: string;
  provider_id: string;
  entry_point: string;
  issuer: string;
  idp_cert_pem: string;
  want_assertions_signed: boolean;
  want_response_signed: boolean;
  attribute_mapping: unknown;
  enabled: boolean;
};

type OrgAuthSettingsRow = {
  allowed_auth_methods: unknown;
  require_mfa: boolean;
};

type SamlAuthStateRow = {
  state: string;
  org_id: string;
  provider_id: string;
  created_at: Date;
};

type AttributeMapping = {
  email: string;
  name: string;
  groups?: string;
};

const AUTH_STATE_TTL_MS = 10 * 60 * 1000;

type RequestCacheRow = { value: string; created_at: Date };

type DbClient = Pick<Pool, "query">;

function createSamlRequestCacheProvider(db: DbClient): CacheProvider {
  return {
    async saveAsync(key: string, value: string): Promise<CacheItem | null> {
      await db.query(
        `
          INSERT INTO saml_request_cache (id, value)
          VALUES ($1, $2)
          ON CONFLICT (id)
          DO UPDATE SET value = EXCLUDED.value, created_at = now()
        `,
        [key, value]
      );
      return { value, createdAt: Date.now() };
    },
    async getAsync(key: string): Promise<string | null> {
      const res = await db.query<RequestCacheRow>(
        `
          SELECT value, created_at
          FROM saml_request_cache
          WHERE id = $1
          LIMIT 1
        `,
        [key]
      );
      if (res.rowCount !== 1) return null;
      const row = res.rows[0] as RequestCacheRow;
      const ageMs = Date.now() - new Date(row.created_at).getTime();
      if (!Number.isFinite(ageMs) || ageMs > AUTH_STATE_TTL_MS) {
        await db.query("DELETE FROM saml_request_cache WHERE id = $1", [key]);
        return null;
      }
      return String(row.value);
    },
    async removeAsync(key: string | null): Promise<string | null> {
      if (!key) return null;
      const res = await db.query<{ value: string }>(
        `
          DELETE FROM saml_request_cache
          WHERE id = $1
          RETURNING value
        `,
        [key]
      );
      if (res.rowCount !== 1) return null;
      return String((res.rows[0] as any).value);
    }
  };
}

function isProd(): boolean {
  return process.env.NODE_ENV === "production";
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

function parseAttributeMapping(value: unknown): AttributeMapping | null {
  if (!value) return null;

  let obj: unknown = value;
  if (typeof value === "string") {
    try {
      obj = JSON.parse(value) as unknown;
    } catch {
      return null;
    }
  }

  if (!obj || typeof obj !== "object") return null;
  const record = obj as Record<string, unknown>;
  const email = typeof record.email === "string" ? record.email.trim() : null;
  const name = typeof record.name === "string" ? record.name.trim() : null;
  const groupsRaw = typeof record.groups === "string" ? record.groups.trim() : undefined;
  const groups = groupsRaw && groupsRaw.length > 0 ? groupsRaw : undefined;
  if (!email || email.length === 0) return null;
  if (!name || name.length === 0) return null;
  return { email, name, groups };
}

function extractPublicBaseUrl(request: FastifyRequest): string | null {
  const configured = request.server.config.publicBaseUrl;
  if (configured && configured.trim().length > 0) {
    try {
      const url = new URL(configured);
      if (isProd() && url.protocol !== "https:") return null;
      return url.toString().replace(/\/+$/, "");
    } catch {
      return null;
    }
  }

  if (isProd()) return null;

  const host = typeof request.headers.host === "string" && request.headers.host.length > 0 ? request.headers.host : null;
  if (!host) return null;
  const proto = request.protocol === "https" ? "https" : "http";
  return `${proto}://${host}`.replace(/\/+$/, "");
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
  entryPoint: string;
  issuer: string;
  idpCertPem: string;
  wantAssertionsSigned: boolean;
  wantResponseSigned: boolean;
  attributeMapping: AttributeMapping;
  enabled: boolean;
} | null> {
  const res = await request.server.db.query<OrgSamlProviderRow>(
    `
      SELECT org_id, provider_id, entry_point, issuer, idp_cert_pem,
             want_assertions_signed, want_response_signed, attribute_mapping, enabled
      FROM org_saml_providers
      WHERE org_id = $1 AND provider_id = $2
      LIMIT 1
    `,
    [orgId, providerId]
  );
  if (res.rowCount !== 1) return null;
  const row = res.rows[0] as OrgSamlProviderRow;
  const mapping = parseAttributeMapping(row.attribute_mapping);
  if (!mapping) return null;
  return {
    entryPoint: String(row.entry_point),
    issuer: String(row.issuer),
    idpCertPem: String(row.idp_cert_pem),
    wantAssertionsSigned: Boolean(row.want_assertions_signed),
    wantResponseSigned: Boolean(row.want_response_signed),
    attributeMapping: mapping,
    enabled: Boolean(row.enabled)
  };
}

function buildSaml(options: {
  entryPoint: string;
  issuer: string;
  callbackUrl: string;
  idpCertPem: string;
  wantAssertionsSigned: boolean;
  wantResponseSigned: boolean;
  cacheProvider: CacheProvider;
}): SAML {
  return new SAML({
    entryPoint: options.entryPoint,
    issuer: options.issuer,
    callbackUrl: options.callbackUrl,
    // Audience must match the SP issuer for most IdPs.
    audience: options.issuer,
    idpCert: options.idpCertPem,
    wantAssertionsSigned: options.wantAssertionsSigned,
    wantAuthnResponseSigned: options.wantResponseSigned,
    // Allow small clock skew for NotBefore/NotOnOrAfter checks.
    acceptedClockSkewMs: 5 * 60 * 1000,
    // Validate (and consume) InResponseTo when the IdP includes it so assertions
    // cannot be replayed against a fresh RelayState.
    validateInResponseTo: ValidateInResponseTo.ifPresent,
    requestIdExpirationPeriodMs: AUTH_STATE_TTL_MS,
    cacheProvider: options.cacheProvider
  });
}

function firstStringValue(value: unknown): string | null {
  if (typeof value === "string" && value.trim().length > 0) return value.trim();
  if (Array.isArray(value)) {
    for (const item of value) {
      if (typeof item === "string" && item.trim().length > 0) return item.trim();
    }
  }
  return null;
}

function extractAttribute(profile: Record<string, unknown>, key: string): string | null {
  const direct = firstStringValue(profile[key]);
  if (direct) return direct;
  const attributes = profile.attributes;
  if (attributes && typeof attributes === "object") {
    const nested = firstStringValue((attributes as Record<string, unknown>)[key]);
    if (nested) return nested;
  }
  return null;
}

function extractAttributeValues(profile: Record<string, unknown>, key: string): string[] {
  const raw = profile[key] ?? (profile.attributes && (profile.attributes as any)[key]);
  if (!raw) return [];
  if (typeof raw === "string") return raw.trim().length > 0 ? [raw.trim()] : [];
  if (Array.isArray(raw)) return raw.filter((v) => typeof v === "string" && v.trim().length > 0).map((v) => v.trim());
  return [];
}

function extractEmail(profile: Record<string, unknown>, mapping: AttributeMapping): string | null {
  const attr = extractAttribute(profile, mapping.email);
  if (attr && attr.includes("@")) return attr.trim().toLowerCase();

  // Some IdPs use common attribute names; fallback to the most common ones.
  const fallbacks = [profile.email, profile.mail, profile.upn];
  for (const value of fallbacks) {
    const found = firstStringValue(value);
    if (found && found.includes("@")) return found.trim().toLowerCase();
  }
  return null;
}

function extractName(profile: Record<string, unknown>, mapping: AttributeMapping, email: string): string {
  const attr = extractAttribute(profile, mapping.name);
  if (attr) return attr.trim();

  const displayName = firstStringValue(profile.displayName ?? profile.cn ?? profile.name);
  if (displayName) return displayName.trim();

  const local = email.split("@")[0];
  return local && local.length > 0 ? local : "User";
}

function samlIndicatesMfa(profile: Record<string, unknown>): boolean {
  const candidates = [
    profile.authnContextClassRef,
    profile.authnContextClassRefValue,
    profile.authnContext,
    profile.authnContextClass
  ];

  const getAssertionXml = (profile as any).getAssertionXml;
  if (typeof getAssertionXml === "function") {
    try {
      candidates.push(getAssertionXml.call(profile));
    } catch {
      // ignore
    }
  }

  for (const candidate of candidates) {
    const value = firstStringValue(candidate);
    if (!value) continue;
    const normalized = value.toLowerCase();
    if (normalized.includes("mfa")) return true;
    if (normalized.includes("otp")) return true;
    if (normalized.includes("totp")) return true;
    if (normalized.includes("timesynctoken")) return true;
    if (normalized.includes("refeds.org/profile/mfa")) return true;
  }
  return false;
}

async function writeSamlFailureAudit(options: {
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
      : { type: "anonymous", id: `saml:${options.providerId ?? "unknown"}` };

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
    details: { method: "saml", provider: options.providerId }
  });

  try {
    await writeAuditEvent(options.request.server.db, event);
  } catch {
    // Ignore audit failures; auth code paths must not fail closed due to observability.
  }
}

export async function samlStart(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;

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
  if (!settings.allowedAuthMethods.includes("saml")) {
    reply.code(403).send({ error: "auth_method_not_allowed" });
    return;
  }

  const baseUrl = extractPublicBaseUrl(request);
  if (!baseUrl) {
    reply.code(500).send({ error: "public_base_url_required" });
    return;
  }

  const callbackUrl = `${baseUrl}/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`;

  const cacheProvider = createSamlRequestCacheProvider(request.server.db);
  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    issuer: provider.issuer,
    callbackUrl,
    idpCertPem: provider.idpCertPem,
    wantAssertionsSigned: provider.wantAssertionsSigned,
    wantResponseSigned: provider.wantResponseSigned,
    cacheProvider
  });

  const relayState = randomBase64Url(32);
  let url: string;
  try {
    url = await saml.getAuthorizeUrlAsync(relayState, undefined, {} as any);
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "saml_authorize_url_failed" });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "saml_authorize_url_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(502).send({ error: "saml_authorize_url_failed" });
    return;
  }

  try {
    await request.server.db.query(
      `
        INSERT INTO saml_auth_states (state, org_id, provider_id)
        VALUES ($1, $2, $3)
      `,
      [relayState, orgId, providerId]
    );
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "saml_state_store_failed" });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "saml_state_store_failed",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(500).send({ error: "saml_state_store_failed" });
    return;
  }

  reply.redirect(url);
}

export async function samlMetadata(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;

  const provider = await loadOrgProvider(request, orgId, providerId);
  if (!provider) {
    reply.code(404).send({ error: "provider_not_found" });
    return;
  }

  const baseUrl = extractPublicBaseUrl(request);
  if (!baseUrl) {
    reply.code(500).send({ error: "public_base_url_required" });
    return;
  }

  const callbackUrl = `${baseUrl}/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`;

  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    issuer: provider.issuer,
    callbackUrl,
    idpCertPem: provider.idpCertPem,
    wantAssertionsSigned: provider.wantAssertionsSigned,
    wantResponseSigned: provider.wantResponseSigned,
    cacheProvider: createSamlRequestCacheProvider(request.server.db)
  });

  const xml = saml.generateServiceProviderMetadata(null, null);
  reply.header("content-type", "application/xml; charset=utf-8");
  reply.send(xml);
}

const CallbackBody = z
  .object({
    SAMLResponse: z.string().min(1),
    RelayState: z.string().min(1)
  })
  .passthrough();

export async function samlCallback(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;

  const parsed = CallbackBody.safeParse(request.body);
  if (!parsed.success) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  const stateRes = await request.server.db.query<SamlAuthStateRow>(
    `
      DELETE FROM saml_auth_states
      WHERE state = $1
      RETURNING state, org_id, provider_id, created_at
    `,
    [parsed.data.RelayState]
  );

  if (stateRes.rowCount !== 1) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_state" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_state" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  const authState = stateRes.rows[0] as SamlAuthStateRow;
  if (authState.org_id !== orgId || authState.provider_id !== providerId) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_state" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_state" });
    reply.code(401).send({ error: "invalid_state" });
    return;
  }

  const ageMs = Date.now() - new Date(authState.created_at).getTime();
  if (!Number.isFinite(ageMs) || ageMs > AUTH_STATE_TTL_MS) {
    request.server.metrics.authFailuresTotal.inc({ reason: "state_expired" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "state_expired" });
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
  if (!settings.allowedAuthMethods.includes("saml")) {
    reply.code(403).send({ error: "auth_method_not_allowed" });
    return;
  }

  const baseUrl = extractPublicBaseUrl(request);
  if (!baseUrl) {
    reply.code(500).send({ error: "public_base_url_required" });
    return;
  }
  const callbackUrl = `${baseUrl}/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`;

  const cacheProvider = createSamlRequestCacheProvider(request.server.db);
  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    issuer: provider.issuer,
    callbackUrl,
    idpCertPem: provider.idpCertPem,
    wantAssertionsSigned: provider.wantAssertionsSigned,
    wantResponseSigned: provider.wantResponseSigned,
    cacheProvider
  });

  let profile: Record<string, unknown>;
  try {
    const container: Record<string, string> = {
      SAMLResponse: parsed.data.SAMLResponse,
      RelayState: parsed.data.RelayState
    };

    const result = await saml.validatePostResponseAsync(container);
    if (!result.profile || typeof result.profile !== "object") {
      throw new Error("SAML response missing profile");
    }
    profile = result.profile as unknown as Record<string, unknown>;

    const inResponseTo = (result.profile as any).inResponseTo;
    if (typeof inResponseTo === "string" && inResponseTo.length > 0) {
      try {
        await cacheProvider.removeAsync(inResponseTo);
      } catch {
        // ignore cache cleanup failures; replay protection is best-effort
      }
    }
  } catch (err) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_saml_response" });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: "invalid_saml_response",
      errorMessage: err instanceof Error ? err.message : undefined
    });
    reply.code(401).send({ error: "invalid_saml_response" });
    return;
  }

  const subject = firstStringValue(profile.nameID ?? profile.nameId ?? profile.subject);
  if (!subject) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_claims" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_claims" });
    reply.code(401).send({ error: "invalid_claims" });
    return;
  }

  const email = extractEmail(profile, provider.attributeMapping);
  if (!email) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_claims" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_claims" });
    reply.code(401).send({ error: "invalid_claims" });
    return;
  }

  if (settings.requireMfa && !samlIndicatesMfa(profile)) {
    request.server.metrics.authFailuresTotal.inc({ reason: "mfa_required" });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      userEmail: email,
      errorCode: "mfa_required",
      errorMessage: "Organization requires MFA"
    });
    reply.code(401).send({ error: "mfa_required" });
    return;
  }

  const name = extractName(profile, provider.attributeMapping, email);

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
      const existingUser = await client.query<{ id: string }>("SELECT id FROM users WHERE email = $1 LIMIT 1", [email]);
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
      userAgent: getUserAgent(request)
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
      details: { method: "saml", provider: providerId }
    })
  );

  reply.send({ user: { id: userId, email, name }, orgId });
}
