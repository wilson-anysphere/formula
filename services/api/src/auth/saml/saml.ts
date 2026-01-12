import type { FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { SAML, ValidateInResponseTo, type CacheItem, type CacheProvider } from "@node-saml/node-saml";
import type { Pool, PoolClient } from "pg";
import { createAuditEvent, writeAuditEvent } from "../../audit/audit";
import { withTransaction } from "../../db/tx";
import { getClientIp, getUserAgent } from "../../http/request-meta";
import { randomBase64Url } from "../oidc/pkce";
import { createSession } from "../sessions";

type OrgSamlProviderRow = {
  org_id: string;
  provider_id: string;
  idp_entry_point: string;
  sp_entity_id: string;
  idp_issuer: string | null;
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
  email?: string;
  name?: string;
  groups?: string;
};

const CLOCK_SKEW_MS = 5 * 60 * 1000;
const AUTH_STATE_TTL_MS = 10 * 60 * 1000;
// Hard cap for decoded SAMLResponse XML size. The callback route body limit should
// be configured to allow base64 overhead but still keep the request bounded.
const MAX_SAML_RESPONSE_BYTES = 1024 * 1024;

type RequestCacheRow = { value: string; created_at: Date };

type Queryable = Pick<Pool, "query"> | Pick<PoolClient, "query">;

export async function cleanupSamlAuthStates(db: Queryable): Promise<number> {
  const cutoff = new Date(Date.now() - AUTH_STATE_TTL_MS);
  const res = await db.query("DELETE FROM saml_auth_states WHERE created_at < $1", [cutoff]);
  return typeof res?.rowCount === "number" ? res.rowCount : 0;
}

export async function cleanupSamlRequestCache(db: Queryable): Promise<number> {
  const cutoff = new Date(Date.now() - AUTH_STATE_TTL_MS);
  const res = await db.query("DELETE FROM saml_request_cache WHERE created_at < $1", [cutoff]);
  return typeof res?.rowCount === "number" ? res.rowCount : 0;
}

export async function cleanupSamlAssertionReplays(db: Queryable): Promise<number> {
  const now = new Date();
  const res = await db.query("DELETE FROM saml_assertion_replays WHERE expires_at < $1", [now]);
  return typeof res?.rowCount === "number" ? res.rowCount : 0;
}

function createSamlRequestCacheProvider(db: Queryable): CacheProvider {
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

function parseAttributeMapping(value: unknown): AttributeMapping {
  if (!value) return {};

  let obj: unknown = value;
  if (typeof value === "string") {
    try {
      obj = JSON.parse(value) as unknown;
    } catch {
      return {};
    }
  }

  if (!obj || typeof obj !== "object") return {};
  const record = obj as Record<string, unknown>;
  const mapping: AttributeMapping = {};
  if (typeof record.email === "string" && record.email.trim().length > 0) mapping.email = record.email.trim();
  if (typeof record.name === "string" && record.name.trim().length > 0) mapping.name = record.name.trim();
  if (typeof record.groups === "string" && record.groups.trim().length > 0) mapping.groups = record.groups.trim();
  return mapping;
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
      if (trimmed === hostnameLower) return true;
    }
  }

  return false;
}

function externalBaseUrl(request: FastifyRequest): string {
  const configured = request.server.config.publicBaseUrl;
  if (typeof configured === "string" && configured.length > 0) return configured;

  if (!request.server.config.trustProxy) {
    throw new Error("PUBLIC_BASE_URL is required when trustProxy is disabled");
  }

  const host = extractHost(request);
  if (!isHostAllowed(host, request.server.config.publicBaseUrlHostAllowlist)) {
    throw new Error("Untrusted host for SAML callback URL");
  }

  return `${extractProto(request)}://${host}`;
}

function buildRedirectUri(baseUrl: string, pathname: string): string {
  const base = new URL(baseUrl);
  base.search = "";
  base.hash = "";
  if (!base.pathname.endsWith("/")) base.pathname = `${base.pathname}/`;
  const relative = pathname.startsWith("/") ? pathname.slice(1) : pathname;
  return new URL(relative, base).toString();
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
  spEntityId: string;
  idpIssuer: string | null;
  idpCertPem: string;
  wantAssertionsSigned: boolean;
  wantResponseSigned: boolean;
  attributeMapping: AttributeMapping;
  enabled: boolean;
} | null> {
  const res = await request.server.db.query<OrgSamlProviderRow>(
    `
      SELECT org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
             want_assertions_signed, want_response_signed, attribute_mapping, enabled
      FROM org_saml_providers
      WHERE org_id = $1 AND provider_id = $2
      LIMIT 1
    `,
    [orgId, providerId]
  );
  if (res.rowCount !== 1) return null;
  const row = res.rows[0] as OrgSamlProviderRow;
  return {
    entryPoint: String(row.idp_entry_point),
    spEntityId: String(row.sp_entity_id),
    idpIssuer: row.idp_issuer ? String(row.idp_issuer) : null,
    idpCertPem: String(row.idp_cert_pem),
    wantAssertionsSigned: Boolean(row.want_assertions_signed),
    wantResponseSigned: Boolean(row.want_response_signed),
    attributeMapping: parseAttributeMapping(row.attribute_mapping),
    enabled: Boolean(row.enabled)
  };
}

function buildSaml(options: {
  entryPoint: string;
  spEntityId: string;
  callbackUrl: string;
  idpCertPem: string;
  wantAssertionsSigned: boolean;
  wantResponseSigned: boolean;
  cacheProvider: CacheProvider;
}): SAML {
  const certBlockRegex = /-----BEGIN CERTIFICATE-----[\s\S]*?-----END CERTIFICATE-----/g;
  const idpCertBlocks = [...options.idpCertPem.matchAll(certBlockRegex)].map((match) => match[0].trim());
  const idpCert = idpCertBlocks.length > 1 ? idpCertBlocks : options.idpCertPem;

  return new SAML({
    entryPoint: options.entryPoint,
    issuer: options.spEntityId,
    callbackUrl: options.callbackUrl,
    // Audience must match the SP issuer for most IdPs.
    audience: options.spEntityId,
    idpCert,
    wantAssertionsSigned: options.wantAssertionsSigned,
    wantAuthnResponseSigned: options.wantResponseSigned,
    acceptedClockSkewMs: CLOCK_SKEW_MS,
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

function extractEmail(profile: Record<string, unknown>, mapping: AttributeMapping, subject: string | null): string | null {
  if (mapping.email) {
    const attr = extractAttribute(profile, mapping.email);
    if (attr && attr.includes("@")) return attr.trim().toLowerCase();
  }

  // Some IdPs use common attribute names; fallback to the most common ones.
  const fallbacks = [profile.email, profile.mail, profile.upn];
  for (const value of fallbacks) {
    const found = firstStringValue(value);
    if (found && found.includes("@")) return found.trim().toLowerCase();
  }

  if (subject && subject.includes("@")) return subject.trim().toLowerCase();

  return null;
}

function extractName(profile: Record<string, unknown>, mapping: AttributeMapping, email: string, subject: string): string {
  if (mapping.name) {
    const attr = extractAttribute(profile, mapping.name);
    if (attr) return attr.trim();
  }

  const displayName = firstStringValue(profile.displayName ?? profile.cn ?? profile.name);
  if (displayName) return displayName.trim();

  // Some IdPs send a stable opaque NameID; avoid showing it in the UI if it isn't email-like.
  if (subject.includes("@")) {
    const local = subject.split("@")[0];
    if (local && local.length > 0) return local;
  }

  const local = email.split("@")[0];
  return local && local.length > 0 ? local : "User";
}

function assertionRecipientsMatch(profile: Record<string, unknown>, callbackUrl: string): boolean {
  const getAssertion = (profile as any).getAssertion;
  if (typeof getAssertion !== "function") return true;

  let assertionDoc: any;
  try {
    assertionDoc = getAssertion.call(profile);
  } catch {
    return true;
  }

  const assertion = assertionDoc?.Assertion;
  if (!assertion) return true;

  const subjects = Array.isArray(assertion.Subject) ? assertion.Subject : [];
  let sawRecipient = false;

  for (const subject of subjects) {
    const confirmations = Array.isArray(subject?.SubjectConfirmation) ? subject.SubjectConfirmation : [];
    for (const confirmation of confirmations) {
      const dataNodes = Array.isArray(confirmation?.SubjectConfirmationData) ? confirmation.SubjectConfirmationData : [];
      for (const node of dataNodes) {
        const recipient = node?.$?.Recipient;
        if (typeof recipient === "string" && recipient.trim().length > 0) {
          sawRecipient = true;
          if (recipient.trim() === callbackUrl) return true;
        }
      }
    }
  }

  // If the assertion did not include a Recipient, don't fail; some IdPs omit it.
  return !sawRecipient;
}

function responseDestinationMatches(profile: Record<string, unknown>, callbackUrl: string): boolean {
  const getSamlResponseXml = (profile as any).getSamlResponseXml;
  if (typeof getSamlResponseXml !== "function") return true;

  let xml: string;
  try {
    xml = String(getSamlResponseXml.call(profile));
  } catch {
    return true;
  }

  if (!xml) return true;

  // Validate the Destination attribute on the top-level Response when present.
  // When only the Assertion is signed, Destination is not covered by the XML
  // signature, so we must treat it as untrusted input.
  const match = xml.match(/<\s*(?:[A-Za-z0-9_]+:)?Response\b[^>]*\bDestination\s*=\s*"([^"]+)"/);
  if (!match) return true;
  return match[1] === callbackUrl;
}

function extractAuthnContextClassRefs(profile: Record<string, unknown>): string[] {
  const refs: string[] = [];
  const candidates = [profile.authnContextClassRef, profile.authnContextClassRefValue, profile.authnContext];

  for (const candidate of candidates) {
    const value = firstStringValue(candidate);
    if (value) refs.push(value);
  }

  const getAssertionXml = (profile as any).getAssertionXml;
  if (typeof getAssertionXml === "function") {
    try {
      const xml = String(getAssertionXml.call(profile));
      const regex =
        /<\s*(?:[A-Za-z0-9_]+:)?AuthnContextClassRef\b[^>]*>([^<]+)<\s*\/\s*(?:[A-Za-z0-9_]+:)?AuthnContextClassRef\s*>/gi;
      for (const match of xml.matchAll(regex)) {
        const ref = match[1]?.trim();
        if (ref) refs.push(ref);
      }
    } catch {
      // ignore
    }
  }

  return Array.from(new Set(refs));
}

function samlIndicatesMfa(profile: Record<string, unknown>): boolean {
  for (const ref of extractAuthnContextClassRefs(profile)) {
    const normalized = ref.trim().toLowerCase();
    if (!normalized) continue;
    // Conservative rules: accept only known MFA AuthnContextClassRef values.
    if (normalized.includes("refeds.org/profile/mfa")) return true;
    if (normalized.includes("timesynctoken")) return true;
    if (normalized.includes("smartcard")) return true;
    if (normalized.includes("twofactor")) return true;
    if (normalized.includes("totp")) return true;
    if (normalized.includes("otp")) return true;
    if (normalized.includes("mfa")) return true;
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

class SamlValidationError extends Error {
  readonly code: string;

  constructor(code: string, message?: string) {
    super(message ?? code);
    this.code = code;
  }
}

function decodeSamlXml(base64: string): string {
  const trimmed = base64.trim();
  // Reject overly large inputs before decoding to keep memory usage bounded.
  const estimatedBytes = Math.floor((trimmed.length * 3) / 4);
  if (!Number.isFinite(estimatedBytes) || estimatedBytes > MAX_SAML_RESPONSE_BYTES) {
    throw new SamlValidationError("response_too_large", "SAMLResponse exceeded maximum size");
  }
  // Note: Buffer.from(..., "base64") is permissive and ignores many invalid
  // characters. Validate the alphabet and ensure a stable round-trip so malformed
  // inputs fail fast.
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(trimmed)) {
    throw new SamlValidationError("invalid_response", "SAMLResponse was not valid base64");
  }

  const decoded = Buffer.from(trimmed, "base64");
  const normalizedInput = trimmed.replace(/=+$/, "");
  const normalizedRoundTrip = decoded.toString("base64").replace(/=+$/, "");
  if (normalizedInput !== normalizedRoundTrip) {
    throw new SamlValidationError("invalid_response", "SAMLResponse was not valid base64");
  }

  const xml = decoded.toString("utf8");

  if (!xml || xml.trim().length === 0) {
    throw new SamlValidationError("invalid_response", "SAMLResponse decoded to an empty document");
  }

  return xml;
}

function preflightSamlResponseXml(xml: string): void {
  if (/<\s*!doctype/i.test(xml) || /<\s*!entity/i.test(xml)) {
    throw new SamlValidationError("invalid_response", "SAMLResponse contains a forbidden DOCTYPE/ENTITY");
  }

  // Defense-in-depth against signature wrapping: reject responses with multiple assertions.
  const assertionRegex = /<\s*(?:[A-Za-z0-9_]+:)?Assertion\b/g;
  let assertionCount = 0;
  while (assertionRegex.exec(xml)) {
    assertionCount += 1;
    if (assertionCount > 1) break;
  }
  if (assertionCount !== 1) {
    throw new SamlValidationError("invalid_response", `expected exactly 1 Assertion (got ${assertionCount})`);
  }
}

function extractAssertionId(xml: string): string | null {
  const match = xml.match(/<\s*(?:[A-Za-z0-9_]+:)?Assertion\b[^>]*\bID\s*=\s*"([^"]+)"/);
  if (!match) return null;
  const value = match[1]?.trim();
  return value && value.length > 0 ? value : null;
}

function parseSamlTimeMs(value: string | null): number | null {
  if (!value) return null;
  const ms = Date.parse(value);
  return Number.isFinite(ms) ? ms : null;
}

function extractAssertionExpiryMs(xml: string): number | null {
  const conditions = xml.match(/<\s*(?:[A-Za-z0-9_]+:)?Conditions\b[^>]*\bNotOnOrAfter\s*=\s*"([^"]+)"/);
  const expiry = parseSamlTimeMs(conditions?.[1] ?? null);
  if (expiry !== null) return expiry;

  const subjectExpiry = xml.match(
    /<\s*(?:[A-Za-z0-9_]+:)?SubjectConfirmationData\b[^>]*\bNotOnOrAfter\s*=\s*"([^"]+)"/
  );
  return parseSamlTimeMs(subjectExpiry?.[1] ?? null);
}

function extractAssertionNotBeforeMs(xml: string): number | null {
  const conditions = xml.match(/<\s*(?:[A-Za-z0-9_]+:)?Conditions\b[^>]*\bNotBefore\s*=\s*"([^"]+)"/);
  const notBefore = parseSamlTimeMs(conditions?.[1] ?? null);
  if (notBefore !== null) return notBefore;

  const subjectNotBefore = xml.match(/<\s*(?:[A-Za-z0-9_]+:)?SubjectConfirmationData\b[^>]*\bNotBefore\s*=\s*"([^"]+)"/);
  return parseSamlTimeMs(subjectNotBefore?.[1] ?? null);
}

function classifyTimestampValidation(xml: string): SamlValidationError | null {
  const now = Date.now();

  const notBefore = extractAssertionNotBeforeMs(xml);
  if (notBefore !== null && now + CLOCK_SKEW_MS < notBefore) {
    return new SamlValidationError("assertion_not_yet_valid", "SAML assertion not yet valid");
  }

  const expiry = extractAssertionExpiryMs(xml);
  if (expiry !== null && now - CLOCK_SKEW_MS >= expiry) {
    return new SamlValidationError("assertion_expired", "SAML assertion expired");
  }

  return null;
}

function mapSamlValidationError(err: unknown): SamlValidationError {
  if (err instanceof SamlValidationError) return err;

  const message = err instanceof Error ? err.message : undefined;
  if (!message) return new SamlValidationError("invalid_saml_response");

  const lower = message.toLowerCase();
  if (lower.includes("signature") || lower.includes("digest")) {
    return new SamlValidationError("invalid_signature", message);
  }
  if (lower.includes("audience")) {
    return new SamlValidationError("invalid_audience", message);
  }
  if (lower.includes("expired") || lower.includes("notonorafter")) {
    return new SamlValidationError("assertion_expired", message);
  }
  if (lower.includes("notbefore") || lower.includes("not yet valid")) {
    return new SamlValidationError("assertion_not_yet_valid", message);
  }
  if (lower.includes("no valid subject confirmation")) {
    return new SamlValidationError("assertion_expired", message);
  }
  if (lower.includes("inresponseto")) {
    return new SamlValidationError("invalid_in_response_to", message);
  }

  return new SamlValidationError("invalid_saml_response", message);
}

export async function samlStart(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const params = request.params as { orgId: string; provider: string };
  const orgId = params.orgId;
  const providerId = params.provider;
  if (!isValidOrgId(orgId) || !isValidProviderId(providerId)) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  // Best-effort cleanup so SAML state/cache tables do not grow unbounded when sweeps are disabled.
  try {
    await cleanupSamlAuthStates(request.server.db);
    await cleanupSamlRequestCache(request.server.db);
    await cleanupSamlAssertionReplays(request.server.db);
  } catch (err) {
    request.server.log.warn({ err }, "saml_auth_state_cleanup_failed");
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
  if (!provider.idpIssuer) {
    reply.code(500).send({ error: "provider_not_configured" });
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

  let callbackUrl: string;
  try {
    callbackUrl = buildRedirectUri(
      externalBaseUrl(request),
      `/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`
    );
  } catch (err) {
    request.server.log.warn({ err }, "saml_callback_url_base_url_invalid");
    reply.code(500).send({ error: "saml_not_configured" });
    return;
  }

  const cacheProvider = createSamlRequestCacheProvider(request.server.db);
  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    spEntityId: provider.spEntityId,
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
  if (!isValidOrgId(orgId) || !isValidProviderId(providerId)) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

  const provider = await loadOrgProvider(request, orgId, providerId);
  if (!provider) {
    reply.code(404).send({ error: "provider_not_found" });
    return;
  }
  if (!provider.idpIssuer) {
    reply.code(500).send({ error: "provider_not_configured" });
    return;
  }

  let callbackUrl: string;
  try {
    callbackUrl = buildRedirectUri(
      externalBaseUrl(request),
      `/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`
    );
  } catch (err) {
    request.server.log.warn({ err }, "saml_metadata_base_url_invalid");
    reply.code(500).send({ error: "saml_not_configured" });
    return;
  }

  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    spEntityId: provider.spEntityId,
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
  if (!isValidOrgId(orgId) || !isValidProviderId(providerId)) {
    reply.code(400).send({ error: "invalid_request" });
    return;
  }

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
  if (!provider.idpIssuer) {
    reply.code(500).send({ error: "provider_not_configured" });
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

  let callbackUrl: string;
  try {
    callbackUrl = buildRedirectUri(
      externalBaseUrl(request),
      `/auth/saml/${encodeURIComponent(orgId)}/${encodeURIComponent(providerId)}/callback`
    );
  } catch (err) {
    request.server.log.warn({ err }, "saml_callback_base_url_invalid");
    reply.code(500).send({ error: "saml_not_configured" });
    return;
  }

  const cacheProvider = createSamlRequestCacheProvider(request.server.db);
  const saml = buildSaml({
    entryPoint: provider.entryPoint,
    spEntityId: provider.spEntityId,
    callbackUrl,
    idpCertPem: provider.idpCertPem,
    wantAssertionsSigned: provider.wantAssertionsSigned,
    wantResponseSigned: provider.wantResponseSigned,
    cacheProvider
  });

  let decodedXml: string;
  const samlResponse = parsed.data.SAMLResponse.replace(/ /g, "+").replace(/\s+/g, "");
  try {
    decodedXml = decodeSamlXml(samlResponse);
    preflightSamlResponseXml(decodedXml);
  } catch (err) {
    const validation = mapSamlValidationError(err);
    request.server.metrics.authFailuresTotal.inc({ reason: validation.code });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: validation.code,
      errorMessage: validation.message
    });
    reply.code(401).send({ error: validation.code });
    return;
  }

  let profile: Record<string, unknown>;
  try {
    const container: Record<string, string> = {
      SAMLResponse: samlResponse,
      RelayState: parsed.data.RelayState
    };

    const result = await saml.validatePostResponseAsync(container);
    if (!result.profile || typeof result.profile !== "object") {
      throw new SamlValidationError("invalid_saml_response", "SAML response missing profile");
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

    const issuer = firstStringValue((profile as any).issuer);
    if (!issuer || issuer !== provider.idpIssuer) {
      throw new SamlValidationError(
        "invalid_issuer",
        `SAML issuer mismatch (expected ${provider.idpIssuer}, got ${issuer ?? "missing"})`
      );
    }

    if (!assertionRecipientsMatch(profile, callbackUrl)) {
      throw new SamlValidationError("invalid_recipient", "SAML assertion recipient mismatch");
    }

    if (!responseDestinationMatches(profile, callbackUrl)) {
      throw new SamlValidationError("invalid_destination", "SAML response destination mismatch");
    }
  } catch (err) {
    const validation = mapSamlValidationError(err);
    request.server.metrics.authFailuresTotal.inc({ reason: validation.code });
    await writeSamlFailureAudit({
      request,
      orgId,
      providerId,
      errorCode: validation.code,
      errorMessage: validation.message
    });
    reply.code(401).send({ error: validation.code });
    return;
  }

  const subject = firstStringValue(profile.nameID ?? profile.nameId ?? profile.subject);
  if (!subject) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_claims" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_claims" });
    reply.code(401).send({ error: "invalid_claims" });
    return;
  }

  const email = extractEmail(profile, provider.attributeMapping, subject);
  if (!email || !z.string().email().safeParse(email).success) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_claims" });
    await writeSamlFailureAudit({ request, orgId, providerId, errorCode: "invalid_claims" });
    reply.code(401).send({ error: "invalid_claims" });
    return;
  }

  const mfaSatisfied = samlIndicatesMfa(profile);
  if (settings.requireMfa && !mfaSatisfied) {
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

  const assertionId = extractAssertionId(decodedXml);
  if (!assertionId) {
    request.server.metrics.authFailuresTotal.inc({ reason: "invalid_saml_response" });
    await writeSamlFailureAudit({ request, orgId, providerId, userEmail: email, errorCode: "invalid_saml_response" });
    reply.code(401).send({ error: "invalid_saml_response" });
    return;
  }

  const assertionExpiry = extractAssertionExpiryMs(decodedXml);
  const expiresAtForReplay = new Date(
    assertionExpiry !== null && Number.isFinite(assertionExpiry) ? assertionExpiry : Date.now() + AUTH_STATE_TTL_MS
  );
  try {
    await request.server.db.query(
      `
        INSERT INTO saml_assertion_replays (assertion_id, org_id, provider_id, expires_at)
        VALUES ($1, $2, $3, $4)
      `,
      [assertionId, orgId, providerId, expiresAtForReplay]
    );
  } catch (err) {
    const code = (err as any)?.code as string | undefined;
    if (code !== "23505") throw err;
    request.server.metrics.authFailuresTotal.inc({ reason: "replay_detected" });
    await writeSamlFailureAudit({ request, orgId, providerId, userEmail: email, errorCode: "replay_detected" });
    reply.code(401).send({ error: "replay_detected" });
    return;
  }

  const name = extractName(profile, provider.attributeMapping, email, subject);
  const expiresAt = new Date(Date.now() + request.server.config.sessionTtlSeconds * 1000);

  const providerKey = `saml:${providerId}`;

  const { userId, sessionId, token } = await withTransaction(request.server.db, async (client) => {
    const existingIdentity = await client.query<{ user_id: string }>(
      `
        SELECT user_id
        FROM user_identities
        WHERE org_id = $1 AND provider = $2 AND subject = $3
        LIMIT 1
      `,
      [orgId, providerKey, subject]
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
        [orgId, providerKey, subject, email]
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
        [userId, providerKey, subject, email, orgId]
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
      details: { method: "saml", provider: providerId }
    })
  );

  reply.send({ user: { id: userId, email, name }, orgId });
}
