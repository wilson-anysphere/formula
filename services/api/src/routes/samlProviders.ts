import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

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
  created_at: Date;
  updated_at: Date;
};

type AttributeMapping = {
  email?: string;
  name?: string;
  groups?: string;
};

function isValidProviderId(value: string): boolean {
  return /^[a-z0-9_-]{1,64}$/.test(value);
}

function isValidOrgId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function pemFromBase64Certificate(base64: string): string {
  const cleaned = base64.replace(/\s+/g, "");
  if (!cleaned) throw new Error("invalid_certificate");
  const lines = cleaned.match(/.{1,64}/g);
  if (!lines) throw new Error("invalid_certificate");
  return `-----BEGIN CERTIFICATE-----\n${lines.join("\n")}\n-----END CERTIFICATE-----`;
}

function normalizePemCertificateBlock(block: string): string {
  const normalized = block.replace(/\r\n|\r/g, "\n").trim();
  const header = "-----BEGIN CERTIFICATE-----";
  const footer = "-----END CERTIFICATE-----";
  if (!normalized.startsWith(header) || !normalized.endsWith(footer)) throw new Error("invalid_certificate");

  const body = normalized.slice(header.length, normalized.length - footer.length).replace(/\s+/g, "");
  return pemFromBase64Certificate(body);
}

function normalizeCertificatePem(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) throw new Error("invalid_certificate");

  const certBlockRegex = /-----BEGIN CERTIFICATE-----[\s\S]*?-----END CERTIFICATE-----/g;

  const blocks = [...trimmed.matchAll(certBlockRegex)].map((match) => match[0]);
  const pemBlocks: string[] =
    blocks.length > 0 ? blocks.map((block) => normalizePemCertificateBlock(block)) : [pemFromBase64Certificate(trimmed)];

  // Ensure we didn't accidentally accept other PEM blocks (e.g. private keys) in the input.
  if (blocks.length > 0) {
    const leftover = trimmed.replace(certBlockRegex, "").trim();
    if (leftover.length > 0) throw new Error("invalid_certificate");
  }

  try {
    for (const pemBlock of pemBlocks) {
      // Throws if invalid.
      new crypto.X509Certificate(pemBlock);
    }
  } catch {
    throw new Error("invalid_certificate");
  }

  return pemBlocks.join("\n");
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
  const mapping: AttributeMapping = {};
  if (typeof record.email === "string" && record.email.trim().length > 0) mapping.email = record.email.trim();
  if (typeof record.name === "string" && record.name.trim().length > 0) mapping.name = record.name.trim();
  if (typeof record.groups === "string" && record.groups.trim().length > 0) mapping.groups = record.groups.trim();
  return Object.keys(mapping).length > 0 ? mapping : null;
}

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, "");
}

function validateHttpsUrl(value: string, requireHttps: boolean): string {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("invalid_url");
  }

  const proto = url.protocol.toLowerCase();
  if (proto !== "https:" && proto !== "http:") throw new Error("invalid_url");
  if (requireHttps && proto !== "https:") throw new Error("https_required");
  return trimTrailingSlash(url.toString());
}

function validateIssuer(value: string, requireHttps: boolean): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) throw new Error("invalid_issuer");

  // Issuers are often URLs, but may also be URNs / other non-HTTP URI schemes.
  // Only enforce scheme rules when it is an HTTP(S) URL.
  try {
    const url = new URL(trimmed);
    const proto = url.protocol.toLowerCase();
    if (proto === "https:" || proto === "http:") {
      if (requireHttps && proto !== "https:") throw new Error("https_required");
    }
  } catch {
    // Not a URL; allow (e.g. URN).
  }
  return trimmed;
}

async function requireOrgAdmin(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
  if (request.authOrgId && request.authOrgId !== orgId) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const membership = await request.server.db.query<{ role: OrgRole }>(
    "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const role = membership.rows[0]!.role;
  if (!isOrgAdmin(role)) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return { role };
}

const AttributeMappingBody = z
  .object({
    email: z.string().min(1).optional(),
    name: z.string().min(1).optional(),
    groups: z.string().min(1).optional()
  })
  .strict();

const PutProviderLegacyBody = z.object({
  entryPoint: z.string().min(1),
  issuer: z.string().min(1),
  idpIssuer: z.string().min(1),
  idpCertPem: z.string().min(1),
  wantAssertionsSigned: z.boolean().optional(),
  wantResponseSigned: z.boolean().optional(),
  attributeMapping: AttributeMappingBody.optional().nullable(),
  enabled: z.boolean().optional()
});

const PutProviderBody = z.object({
  idpEntryPoint: z.string().min(1),
  spEntityId: z.string().min(1),
  idpIssuer: z.string().min(1),
  idpCertPem: z.string().min(1),
  wantAssertionsSigned: z.boolean().optional(),
  wantResponseSigned: z.boolean().optional(),
  enabled: z.boolean().optional(),
  attributeMapping: AttributeMappingBody.optional().nullable()
});

type ProviderConfig = {
  idpEntryPoint: string;
  spEntityId: string;
  idpIssuer: string;
  idpCertPem: string;
  wantAssertionsSigned: boolean;
  wantResponseSigned: boolean;
  enabled: boolean;
  attributeMapping: AttributeMapping | null;
};

function normalizeMapping(mapping: AttributeMapping | null | undefined): AttributeMapping | null {
  if (!mapping) return null;
  const normalized: AttributeMapping = {};
  if (typeof mapping.email === "string" && mapping.email.trim().length > 0) normalized.email = mapping.email.trim();
  if (typeof mapping.name === "string" && mapping.name.trim().length > 0) normalized.name = mapping.name.trim();
  if (typeof mapping.groups === "string" && mapping.groups.trim().length > 0) normalized.groups = mapping.groups.trim();
  return Object.keys(normalized).length > 0 ? normalized : null;
}

async function listProviders(app: FastifyInstance, orgId: string): Promise<OrgSamlProviderRow[]> {
  const providers = await app.db.query<OrgSamlProviderRow>(
    `
      SELECT org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
             want_assertions_signed, want_response_signed, attribute_mapping,
             enabled, created_at, updated_at
      FROM org_saml_providers
      WHERE org_id = $1
      ORDER BY provider_id ASC
    `,
    [orgId]
  );
  return providers.rows as OrgSamlProviderRow[];
}

function providerToLegacyResponse(row: OrgSamlProviderRow): Record<string, unknown> {
  return {
    orgId: row.org_id,
    providerId: row.provider_id,
    entryPoint: row.idp_entry_point,
    issuer: row.sp_entity_id,
    idpIssuer: row.idp_issuer,
    idpCertPem: row.idp_cert_pem,
    wantAssertionsSigned: Boolean(row.want_assertions_signed),
    wantResponseSigned: Boolean(row.want_response_signed),
    attributeMapping: parseAttributeMapping(row.attribute_mapping),
    enabled: Boolean(row.enabled),
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function providerToResponse(row: OrgSamlProviderRow): Record<string, unknown> {
  return {
    providerId: row.provider_id,
    idpEntryPoint: row.idp_entry_point,
    idpIssuer: row.idp_issuer,
    idpCertPem: row.idp_cert_pem,
    spEntityId: row.sp_entity_id,
    wantAssertionsSigned: Boolean(row.want_assertions_signed),
    wantResponseSigned: Boolean(row.want_response_signed),
    enabled: Boolean(row.enabled),
    attributeMapping: parseAttributeMapping(row.attribute_mapping),
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

export function registerSamlProviderRoutes(app: FastifyInstance): void {
  const requireSessionMfa = async (request: FastifyRequest, reply: FastifyReply, orgId: string) => {
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
      reply.code(403).send({ error: "mfa_required" });
      return false;
    }
    return true;
  };

  const upsertProvider = async (
    request: FastifyRequest,
    orgId: string,
    providerId: string,
    config: ProviderConfig
  ): Promise<OrgSamlProviderRow> => {
    const existing = await app.db.query<OrgSamlProviderRow>(
      `
        SELECT org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
               want_assertions_signed, want_response_signed, attribute_mapping,
               enabled, created_at, updated_at
        FROM org_saml_providers
        WHERE org_id = $1 AND provider_id = $2
        LIMIT 1
      `,
      [orgId, providerId]
    );

    const upserted = await app.db.query<OrgSamlProviderRow>(
      `
        INSERT INTO org_saml_providers (
          org_id,
          provider_id,
          idp_entry_point,
          sp_entity_id,
          idp_issuer,
          idp_cert_pem,
          want_assertions_signed,
          want_response_signed,
          attribute_mapping,
          enabled
        )
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9::jsonb,$10)
        ON CONFLICT (org_id, provider_id)
        DO UPDATE SET
          idp_entry_point = EXCLUDED.idp_entry_point,
          sp_entity_id = EXCLUDED.sp_entity_id,
          idp_issuer = EXCLUDED.idp_issuer,
          idp_cert_pem = EXCLUDED.idp_cert_pem,
          want_assertions_signed = EXCLUDED.want_assertions_signed,
          want_response_signed = EXCLUDED.want_response_signed,
          attribute_mapping = EXCLUDED.attribute_mapping,
          enabled = EXCLUDED.enabled,
          updated_at = now()
        RETURNING org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
                  want_assertions_signed, want_response_signed, attribute_mapping,
                  enabled, created_at, updated_at
      `,
      [
        orgId,
        providerId,
        config.idpEntryPoint,
        config.spEntityId,
        config.idpIssuer,
        config.idpCertPem,
        config.wantAssertionsSigned,
        config.wantResponseSigned,
        config.attributeMapping ? JSON.stringify(config.attributeMapping) : null,
        config.enabled
      ]
    );

    const before = existing.rowCount === 1 ? providerToResponse(existing.rows[0]!) : null;
    const after = providerToResponse(upserted.rows[0]!);

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: existing.rowCount === 1 ? "org.saml_provider.updated" : "org.saml_provider.created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "saml_provider", id: providerId },
        success: true,
        details: { before, after }
      })
    );
    return upserted.rows[0]!;
  };

  const deleteProvider = async (
    request: FastifyRequest,
    reply: FastifyReply,
    orgId: string,
    providerId: string
  ): Promise<void> => {
    const deleted = await app.db.query<OrgSamlProviderRow>(
      `
        DELETE FROM org_saml_providers
        WHERE org_id = $1 AND provider_id = $2
        RETURNING org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
                  want_assertions_signed, want_response_signed, attribute_mapping,
                  enabled, created_at, updated_at
      `,
      [orgId, providerId]
    );

    if (deleted.rowCount !== 1) {
      reply.code(404).send({ error: "provider_not_found" });
      return;
    }

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.saml_provider.deleted",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "saml_provider", id: providerId },
        success: true,
        details: { before: providerToResponse(deleted.rows[0]!) }
      })
    );

    reply.send({ ok: true });
  };

  // Legacy endpoints kept for compatibility.
  const preHandlers = [requireAuth, enforceOrgIpAllowlistFromParams];

  app.get("/orgs/:orgId/saml/providers", { preHandler: preHandlers }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    const providers = await listProviders(app, orgId);
    return reply.send({ providers: providers.map((row) => providerToLegacyResponse(row)) });
  });

  app.get("/orgs/:orgId/saml/providers/:providerId", { preHandler: preHandlers }, async (request, reply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    const providerRes = await app.db.query<OrgSamlProviderRow>(
      `
        SELECT org_id, provider_id, idp_entry_point, sp_entity_id, idp_issuer, idp_cert_pem,
               want_assertions_signed, want_response_signed, attribute_mapping,
               enabled, created_at, updated_at
        FROM org_saml_providers
        WHERE org_id = $1 AND provider_id = $2
        LIMIT 1
      `,
      [orgId, providerId]
    );
    if (providerRes.rowCount !== 1) return reply.code(404).send({ error: "provider_not_found" });
    return reply.send({ provider: providerToLegacyResponse(providerRes.rows[0]!) });
  });

  app.put("/orgs/:orgId/saml/providers/:providerId", { preHandler: preHandlers }, async (request, reply) => {
    const params = request.params as { orgId: string; providerId: string };
    const orgId = params.orgId;
    const providerId = params.providerId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    const parsed = PutProviderLegacyBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const requireHttps = process.env.NODE_ENV === "production";
    let idpEntryPoint: string;
    let spEntityId: string;
    let idpIssuer: string;
    let idpCertPem: string;
    try {
      idpEntryPoint = validateHttpsUrl(parsed.data.entryPoint, requireHttps);
      spEntityId = validateIssuer(parsed.data.issuer, requireHttps);
      idpIssuer = validateIssuer(parsed.data.idpIssuer, requireHttps);
      idpCertPem = normalizeCertificatePem(parsed.data.idpCertPem);
    } catch (err) {
      const code = err instanceof Error ? err.message : "invalid_request";
      return reply.code(400).send({ error: code });
    }

    const wantAssertionsSigned = parsed.data.wantAssertionsSigned ?? true;
    const wantResponseSigned = parsed.data.wantResponseSigned ?? false;
    const enabled = parsed.data.enabled ?? true;
    const attributeMapping = normalizeMapping(parsed.data.attributeMapping ?? null);

    const row = await upsertProvider(request, orgId, providerId, {
      idpEntryPoint,
      spEntityId,
      idpIssuer,
      idpCertPem,
      wantAssertionsSigned,
      wantResponseSigned,
      enabled,
      attributeMapping
    });

    reply.send({ provider: providerToLegacyResponse(row) });
  });

  app.delete("/orgs/:orgId/saml/providers/:providerId", { preHandler: preHandlers }, async (request, reply) => {
    const params = request.params as { orgId: string; providerId: string };
    const orgId = params.orgId;
    const providerId = params.providerId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    await deleteProvider(request, reply, orgId, providerId);
  });

  // Required endpoints (per task spec).
  app.get("/orgs/:orgId/saml-providers", { preHandler: preHandlers }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    const providers = await listProviders(app, orgId);
    return reply.send({ providers: providers.map((row) => providerToResponse(row)) });
  });

  app.put("/orgs/:orgId/saml-providers/:providerId", { preHandler: preHandlers }, async (request, reply) => {
    const params = request.params as { orgId: string; providerId: string };
    const orgId = params.orgId;
    const providerId = params.providerId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    const parsed = PutProviderBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const requireHttps = process.env.NODE_ENV === "production";
    let idpEntryPoint: string;
    let spEntityId: string;
    let idpIssuer: string;
    let idpCertPem: string;
    try {
      idpEntryPoint = validateHttpsUrl(parsed.data.idpEntryPoint, requireHttps);
      spEntityId = validateIssuer(parsed.data.spEntityId, requireHttps);
      idpIssuer = validateIssuer(parsed.data.idpIssuer, requireHttps);
      idpCertPem = normalizeCertificatePem(parsed.data.idpCertPem);
    } catch (err) {
      const code = err instanceof Error ? err.message : "invalid_request";
      return reply.code(400).send({ error: code });
    }

    const wantAssertionsSigned = parsed.data.wantAssertionsSigned ?? true;
    const wantResponseSigned = parsed.data.wantResponseSigned ?? false;
    const enabled = parsed.data.enabled ?? true;
    const attributeMapping = normalizeMapping(parsed.data.attributeMapping ?? null);

    const row = await upsertProvider(request, orgId, providerId, {
      idpEntryPoint,
      spEntityId,
      idpIssuer,
      idpCertPem,
      wantAssertionsSigned,
      wantResponseSigned,
      enabled,
      attributeMapping
    });

    reply.send({ provider: providerToResponse(row) });
  });

  app.delete("/orgs/:orgId/saml-providers/:providerId", { preHandler: preHandlers }, async (request, reply) => {
    const params = request.params as { orgId: string; providerId: string };
    const orgId = params.orgId;
    const providerId = params.providerId;
    if (!isValidOrgId(orgId)) return reply.code(400).send({ error: "invalid_request" });
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (!(await requireSessionMfa(request, reply, orgId))) return;

    await deleteProvider(request, reply, orgId, providerId);
  });
}
