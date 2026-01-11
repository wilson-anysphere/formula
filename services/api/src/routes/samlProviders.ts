import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

type OrgSamlProviderRow = {
  org_id: string;
  provider_id: string;
  entry_point: string;
  issuer: string;
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
  email: string;
  name: string;
  groups?: string;
};

function isValidProviderId(value: string): boolean {
  return /^[a-z0-9_-]{1,64}$/.test(value);
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
  const email = typeof record.email === "string" ? record.email : null;
  const name = typeof record.name === "string" ? record.name : null;
  const groups = typeof record.groups === "string" ? record.groups : undefined;
  if (!email || email.trim().length === 0) return null;
  if (!name || name.trim().length === 0) return null;
  return { email: email.trim(), name: name.trim(), groups: groups?.trim() };
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

  // Issuers may be URNs; only enforce scheme rules when it parses as a URL.
  try {
    const url = new URL(trimmed);
    const proto = url.protocol.toLowerCase();
    if (proto !== "https:" && proto !== "http:") throw new Error("invalid_issuer");
    if (requireHttps && proto !== "https:") throw new Error("https_required");
    return trimmed;
  } catch {
    return trimmed;
  }
}

async function requireOrgAdmin(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
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

const PutProviderBody = z.object({
  entryPoint: z.string().min(1),
  issuer: z.string().min(1),
  idpIssuer: z.string().min(1).optional(),
  idpCertPem: z.string().min(1),
  wantAssertionsSigned: z.boolean().optional(),
  wantResponseSigned: z.boolean().optional(),
  attributeMapping: z.object({
    email: z.string().min(1),
    name: z.string().min(1),
    groups: z.string().min(1).optional()
  }),
  enabled: z.boolean().optional()
});

export function registerSamlProviderRoutes(app: FastifyInstance): void {
  app.get("/orgs/:orgId/saml/providers", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const providers = await app.db.query<OrgSamlProviderRow>(
      `
        SELECT org_id, provider_id, entry_point, issuer, idp_cert_pem,
               idp_issuer, want_assertions_signed, want_response_signed, attribute_mapping,
               enabled, created_at, updated_at
        FROM org_saml_providers
        WHERE org_id = $1
        ORDER BY provider_id ASC
      `,
      [orgId]
    );

    return reply.send({
      providers: providers.rows.map((row) => ({
        orgId: row.org_id,
        providerId: row.provider_id,
        entryPoint: row.entry_point,
        issuer: row.issuer,
        idpIssuer: row.idp_issuer,
        idpCertPem: row.idp_cert_pem,
        wantAssertionsSigned: Boolean(row.want_assertions_signed),
        wantResponseSigned: Boolean(row.want_response_signed),
        attributeMapping: parseAttributeMapping(row.attribute_mapping) ?? row.attribute_mapping,
        enabled: Boolean(row.enabled),
        createdAt: row.created_at,
        updatedAt: row.updated_at
      }))
    });
  });

  app.get(
    "/orgs/:orgId/saml/providers/:providerId",
    { preHandler: requireAuth },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string; providerId: string }).orgId;
      const providerId = (request.params as { orgId: string; providerId: string }).providerId;
      if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

      const member = await requireOrgAdmin(request, reply, orgId);
      if (!member) return;

      const providerRes = await app.db.query<OrgSamlProviderRow>(
        `
          SELECT org_id, provider_id, entry_point, issuer, idp_issuer, idp_cert_pem,
                 want_assertions_signed, want_response_signed, attribute_mapping,
                 enabled, created_at, updated_at
          FROM org_saml_providers
          WHERE org_id = $1 AND provider_id = $2
          LIMIT 1
        `,
        [orgId, providerId]
      );
      if (providerRes.rowCount !== 1) return reply.code(404).send({ error: "provider_not_found" });

      const row = providerRes.rows[0]!;
      return reply.send({
        provider: {
          orgId: row.org_id,
          providerId: row.provider_id,
          entryPoint: row.entry_point,
          issuer: row.issuer,
          idpIssuer: row.idp_issuer,
          idpCertPem: row.idp_cert_pem,
          wantAssertionsSigned: Boolean(row.want_assertions_signed),
          wantResponseSigned: Boolean(row.want_response_signed),
          attributeMapping: parseAttributeMapping(row.attribute_mapping) ?? row.attribute_mapping,
          enabled: Boolean(row.enabled),
          createdAt: row.created_at,
          updatedAt: row.updated_at
        }
      });
    }
  );

  app.put(
    "/orgs/:orgId/saml/providers/:providerId",
    { preHandler: requireAuth },
    async (request, reply) => {
      const params = request.params as { orgId: string; providerId: string };
      const orgId = params.orgId;
      const providerId = params.providerId;
      if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

      const member = await requireOrgAdmin(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const parsed = PutProviderBody.safeParse(request.body);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      const requireHttps = process.env.NODE_ENV === "production";

      let entryPoint: string;
      let issuer: string;
      let idpIssuer: string | null = null;
      let idpCertPem: string;
      try {
        entryPoint = validateHttpsUrl(parsed.data.entryPoint, requireHttps);
        issuer = validateIssuer(parsed.data.issuer, requireHttps);
        if (parsed.data.idpIssuer !== undefined) {
          idpIssuer = validateIssuer(parsed.data.idpIssuer, requireHttps);
        }
        idpCertPem = normalizeCertificatePem(parsed.data.idpCertPem);
      } catch (err) {
        const code = err instanceof Error ? err.message : "invalid_request";
        return reply.code(400).send({ error: code });
      }

      const attributeMapping = parseAttributeMapping(parsed.data.attributeMapping);
      if (!attributeMapping) return reply.code(400).send({ error: "invalid_attribute_mapping" });

      const wantAssertionsSigned = parsed.data.wantAssertionsSigned ?? true;
      const wantResponseSigned = parsed.data.wantResponseSigned ?? true;
      const enabled = parsed.data.enabled ?? true;

      const existed = await app.db.query("SELECT 1 FROM org_saml_providers WHERE org_id = $1 AND provider_id = $2", [
        orgId,
        providerId
      ]);

      const upserted = await app.db.query<OrgSamlProviderRow>(
        `
          INSERT INTO org_saml_providers (
            org_id,
            provider_id,
            entry_point,
            issuer,
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
            entry_point = EXCLUDED.entry_point,
            issuer = EXCLUDED.issuer,
            idp_issuer = EXCLUDED.idp_issuer,
            idp_cert_pem = EXCLUDED.idp_cert_pem,
            want_assertions_signed = EXCLUDED.want_assertions_signed,
            want_response_signed = EXCLUDED.want_response_signed,
            attribute_mapping = EXCLUDED.attribute_mapping,
            enabled = EXCLUDED.enabled,
            updated_at = now()
          RETURNING org_id, provider_id, entry_point, issuer, idp_issuer, idp_cert_pem,
                    want_assertions_signed, want_response_signed, attribute_mapping,
                    enabled, created_at, updated_at
        `,
        [
          orgId,
          providerId,
          entryPoint,
          issuer,
          idpIssuer,
          idpCertPem,
          wantAssertionsSigned,
          wantResponseSigned,
          JSON.stringify(attributeMapping),
          enabled
        ]
      );

      const eventType = existed.rowCount === 1 ? "admin.integration_updated" : "admin.integration_added";
      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType,
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "integration", id: providerId, name: "saml" },
          success: true,
          details: { type: "saml", providerId }
        })
      );

      const row = upserted.rows[0]!;
      return reply.send({
        provider: {
          orgId: row.org_id,
          providerId: row.provider_id,
          entryPoint: row.entry_point,
          issuer: row.issuer,
          idpIssuer: row.idp_issuer,
          idpCertPem: row.idp_cert_pem,
          wantAssertionsSigned: Boolean(row.want_assertions_signed),
          wantResponseSigned: Boolean(row.want_response_signed),
          attributeMapping: parseAttributeMapping(row.attribute_mapping) ?? row.attribute_mapping,
          enabled: Boolean(row.enabled),
          createdAt: row.created_at,
          updatedAt: row.updated_at
        }
      });
    }
  );

  app.delete(
    "/orgs/:orgId/saml/providers/:providerId",
    { preHandler: requireAuth },
    async (request, reply) => {
      const params = request.params as { orgId: string; providerId: string };
      const orgId = params.orgId;
      const providerId = params.providerId;
      if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

      const member = await requireOrgAdmin(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const deleted = await app.db.query(
        `
          DELETE FROM org_saml_providers
          WHERE org_id = $1 AND provider_id = $2
          RETURNING provider_id
        `,
        [orgId, providerId]
      );

      if (deleted.rowCount !== 1) return reply.code(404).send({ error: "provider_not_found" });

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "admin.integration_removed",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "integration", id: providerId, name: "saml" },
          success: true,
          details: { type: "saml", providerId }
        })
      );

      return reply.send({ ok: true });
    }
  );
}
