import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { withTransaction } from "../db/tx";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { deleteSecret, putSecret } from "../secrets/secretStore";
import { requireAuth } from "./auth";

type OrgOidcProviderRow = {
  provider_id: string;
  issuer_url: string;
  client_id: string;
  scopes: unknown;
  enabled: boolean;
  created_at: Date;
};

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
  if (!normalized.includes("openid")) normalized.unshift("openid");
  return Array.from(new Set(normalized));
}

function isValidProviderId(value: string): boolean {
  return /^[a-z0-9_-]{1,64}$/.test(value);
}

function oidcSecretName(orgId: string, providerId: string): string {
  return `oidc:${orgId}:${providerId}`;
}

function validateIssuerUrl(value: string): boolean {
  let parsed: URL;
  try {
    parsed = new URL(value);
  } catch {
    return false;
  }
  if (process.env.NODE_ENV === "production" && parsed.protocol !== "https:") return false;
  return parsed.protocol === "https:" || parsed.protocol === "http:";
}

async function requireOrgAdminForOidcProviders(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
  const membership = await request.server.db.query(
    "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const role = membership.rows[0].role as OrgRole;
  if (!isOrgAdmin(role)) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return { role };
}

export function registerOidcProviderRoutes(app: FastifyInstance): void {
  const UpsertBody = z.object({
    issuerUrl: z.string().min(1),
    clientId: z.string().min(1),
    scopes: z.array(z.string().min(1)).default([]),
    enabled: z.boolean(),
    clientSecret: z.string().min(1).optional()
  });

  app.get("/orgs/:orgId/oidc/providers", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const providersRes = await app.db.query<OrgOidcProviderRow>(
      `
        SELECT provider_id, issuer_url, client_id, scopes, enabled, created_at
        FROM org_oidc_providers
        WHERE org_id = $1
        ORDER BY provider_id ASC
      `,
      [orgId]
    );

    const prefix = `oidc:${orgId}:`;
    const secretsRes = await app.db.query<{ name: string }>("SELECT name FROM secrets WHERE name LIKE $1", [
      `${prefix}%`
    ]);
    const configured = new Set(secretsRes.rows.map((row) => String(row.name)));

    return reply.send({
      providers: providersRes.rows.map((row) => {
        const providerId = String(row.provider_id);
        return {
          providerId,
          issuerUrl: String(row.issuer_url),
          clientId: String(row.client_id),
          scopes: ensureOpenIdScope(parseStringArray(row.scopes)),
          enabled: Boolean(row.enabled),
          createdAt: (row.created_at as Date).toISOString(),
          clientSecretConfigured: configured.has(`${prefix}${providerId}`)
        };
      })
    });
  });

  app.get("/orgs/:orgId/oidc/providers/:providerId", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const providerRes = await app.db.query<OrgOidcProviderRow>(
      `
        SELECT provider_id, issuer_url, client_id, scopes, enabled, created_at
        FROM org_oidc_providers
        WHERE org_id = $1 AND provider_id = $2
        LIMIT 1
      `,
      [orgId, providerId]
    );
    if (providerRes.rowCount !== 1) return reply.code(404).send({ error: "provider_not_found" });
    const row = providerRes.rows[0] as OrgOidcProviderRow;

    const secretName = oidcSecretName(orgId, providerId);
    const secretRes = await app.db.query("SELECT 1 FROM secrets WHERE name = $1", [secretName]);

    return reply.send({
      provider: {
        providerId,
        issuerUrl: String(row.issuer_url),
        clientId: String(row.client_id),
        scopes: ensureOpenIdScope(parseStringArray(row.scopes)),
        enabled: Boolean(row.enabled),
        createdAt: (row.created_at as Date).toISOString()
      },
      clientSecretConfigured: secretRes.rowCount === 1
    });
  });

  app.put("/orgs/:orgId/oidc/providers/:providerId", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const parsedBody = UpsertBody.safeParse(request.body);
    if (!parsedBody.success) return reply.code(400).send({ error: "invalid_request" });

    const issuerUrl = parsedBody.data.issuerUrl.trim();
    if (!validateIssuerUrl(issuerUrl)) return reply.code(400).send({ error: "invalid_request" });

    const scopes = ensureOpenIdScope(parsedBody.data.scopes);
    const clientId = parsedBody.data.clientId.trim();
    if (clientId.length === 0) return reply.code(400).send({ error: "invalid_request" });
    const enabled = Boolean(parsedBody.data.enabled);
    const clientSecret = parsedBody.data.clientSecret?.trim();
    if (parsedBody.data.clientSecret !== undefined && (!clientSecret || clientSecret.length === 0)) {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const secretName = oidcSecretName(orgId, providerId);

    const txResult = await withTransaction(app.db, async (client) => {
      const existingRes = await client.query("SELECT 1 FROM org_oidc_providers WHERE org_id = $1 AND provider_id = $2", [
        orgId,
        providerId
      ]);
      const existed = existingRes.rowCount === 1;

      await client.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
          ON CONFLICT (org_id, provider_id)
          DO UPDATE SET issuer_url = EXCLUDED.issuer_url, client_id = EXCLUDED.client_id, scopes = EXCLUDED.scopes, enabled = EXCLUDED.enabled
        `,
        [orgId, providerId, issuerUrl, clientId, JSON.stringify(scopes), enabled]
      );

      if (clientSecret) {
        await putSecret(client, request.server.config.secretStoreKeys, secretName, clientSecret);
      }

      const secretConfigured = clientSecret
        ? true
        : (await client.query("SELECT 1 FROM secrets WHERE name = $1", [secretName])).rowCount === 1;

      const providerRes = await client.query<OrgOidcProviderRow>(
        `
          SELECT provider_id, issuer_url, client_id, scopes, enabled, created_at
          FROM org_oidc_providers
          WHERE org_id = $1 AND provider_id = $2
          LIMIT 1
        `,
        [orgId, providerId]
      );

      const row = providerRes.rows[0] as OrgOidcProviderRow;
      return { existed, secretConfigured, row };
    });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: txResult.existed ? "admin.integration_updated" : "admin.integration_added",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "oidc_provider", id: providerId, name: providerId },
        success: true,
        details: { type: "oidc", providerId }
      })
    );

    return reply.send({
      provider: {
        providerId,
        issuerUrl: String(txResult.row.issuer_url),
        clientId: String(txResult.row.client_id),
        scopes: ensureOpenIdScope(parseStringArray(txResult.row.scopes)),
        enabled: Boolean(txResult.row.enabled),
        createdAt: (txResult.row.created_at as Date).toISOString()
      },
      clientSecretConfigured: txResult.secretConfigured
    });
  });

  app.delete(
    "/orgs/:orgId/oidc/providers/:providerId",
    { preHandler: requireAuth },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string; providerId: string }).orgId;
      const providerId = (request.params as { orgId: string; providerId: string }).providerId;
      if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

      const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const secretName = oidcSecretName(orgId, providerId);

      const deleted = await withTransaction(app.db, async (client) => {
        const res = await client.query(
          "DELETE FROM org_oidc_providers WHERE org_id = $1 AND provider_id = $2 RETURNING provider_id",
          [orgId, providerId]
        );
        if (res.rowCount !== 1) return false;
        await deleteSecret(client, secretName);
        return true;
      });

      if (!deleted) return reply.code(404).send({ error: "provider_not_found" });

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "admin.integration_removed",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id ?? null,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "oidc_provider", id: providerId, name: providerId },
          success: true,
          details: { type: "oidc", providerId }
        })
      );

      return reply.send({ ok: true });
    }
  );
}
