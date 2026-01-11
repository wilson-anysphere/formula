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

const DEFAULT_SCOPES = ["openid", "email", "profile"] as const;

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

function isValidProviderId(value: string): boolean {
  return /^[a-z0-9_-]{1,64}$/.test(value);
}

function isValidOrgId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function oidcSecretName(orgId: string, providerId: string): string {
  return `oidc:${orgId}:${providerId}`;
}

function normalizeIssuerUrl(value: string, options: { allowHttp: boolean; allowLocalhost: boolean }): string {
  const raw = value.trim();
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error("issuerUrl must be a valid URL");
  }

  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("issuerUrl must start with https:// (or http:// in development)");
  }
  if (url.protocol === "http:" && !options.allowHttp) {
    throw new Error("issuerUrl must use https:// in production");
  }

  if (url.username || url.password) {
    throw new Error("issuerUrl must not include credentials");
  }
  if (url.search || url.hash) {
    throw new Error("issuerUrl must not include query parameters or fragments");
  }

  const hostname = url.hostname.toLowerCase();
  const isLocal =
    hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1" || hostname.endsWith(".localhost");
  if (isLocal && !options.allowLocalhost) {
    throw new Error("issuerUrl must not use localhost in production");
  }

  // Strip trailing slashes for stability.
  url.pathname = url.pathname.replace(/\/+$/, "");
  const pathname = url.pathname === "/" ? "" : url.pathname;
  return `${url.origin}${pathname}`;
}

async function requireOrgAdminForOidcProviders(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
  if (!isValidOrgId(orgId)) {
    reply.code(400).send({ error: "invalid_request" });
    return null;
  }
  if (request.authOrgId && request.authOrgId !== orgId) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
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
    scopes: z.array(z.string().min(1)).optional(),
    enabled: z.boolean().optional(),
    clientSecret: z.string().min(1).optional()
  });

  const listProviders = async (request: FastifyRequest, reply: FastifyReply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
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
        const secretConfigured = configured.has(`${prefix}${providerId}`);
        return {
          providerId,
          issuerUrl: String(row.issuer_url),
          clientId: String(row.client_id),
          scopes: ensureOpenIdScope(parseStringArray(row.scopes)),
          enabled: Boolean(row.enabled),
          createdAt: (row.created_at as Date).toISOString(),
          clientSecretConfigured: secretConfigured,
          secretConfigured
        };
      })
    });
  };

  app.get("/orgs/:orgId/oidc/providers", { preHandler: requireAuth }, listProviders);
  app.get("/orgs/:orgId/oidc-providers", { preHandler: requireAuth }, listProviders);

  const getProvider = async (request: FastifyRequest, reply: FastifyReply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
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
    const secretConfigured = secretRes.rowCount === 1;

    return reply.send({
      provider: {
        providerId,
        issuerUrl: String(row.issuer_url),
        clientId: String(row.client_id),
        scopes: ensureOpenIdScope(parseStringArray(row.scopes)),
        enabled: Boolean(row.enabled),
        createdAt: (row.created_at as Date).toISOString()
      },
      clientSecretConfigured: secretConfigured,
      secretConfigured
    });
  };

  app.get("/orgs/:orgId/oidc/providers/:providerId", { preHandler: requireAuth }, getProvider);
  app.get("/orgs/:orgId/oidc-providers/:providerId", { preHandler: requireAuth }, getProvider);

  const putProvider = async (request: FastifyRequest, reply: FastifyReply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const parsedBody = UpsertBody.safeParse(request.body);
    if (!parsedBody.success) return reply.code(400).send({ error: "invalid_request" });

    const isProd = process.env.NODE_ENV === "production";
    let issuerUrl: string;
    try {
      issuerUrl = normalizeIssuerUrl(parsedBody.data.issuerUrl, {
        allowHttp: !isProd,
        allowLocalhost: !isProd
      });
    } catch {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const clientId = parsedBody.data.clientId.trim();
    if (clientId.length === 0) return reply.code(400).send({ error: "invalid_request" });

    const clientSecret = parsedBody.data.clientSecret?.trim();
    if (parsedBody.data.clientSecret !== undefined && (!clientSecret || clientSecret.length === 0)) {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const secretName = oidcSecretName(orgId, providerId);

    const txResult = await withTransaction(app.db, async (client) => {
      const existingRes = await client.query<Pick<OrgOidcProviderRow, "scopes" | "enabled">>(
        "SELECT scopes, enabled FROM org_oidc_providers WHERE org_id = $1 AND provider_id = $2",
        [orgId, providerId]
      );

      const existed = existingRes.rowCount === 1;
      const existingRow = existed ? existingRes.rows[0]! : null;

      const scopesInput =
        parsedBody.data.scopes ?? (existingRow ? parseStringArray(existingRow.scopes) : [...DEFAULT_SCOPES]);
      const scopes = ensureOpenIdScope(scopesInput);

      const enabled = parsedBody.data.enabled ?? (existingRow ? Boolean(existingRow.enabled) : true);

      await client.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
          ON CONFLICT (org_id, provider_id)
          DO UPDATE SET issuer_url = EXCLUDED.issuer_url, client_id = EXCLUDED.client_id, scopes = EXCLUDED.scopes, enabled = EXCLUDED.enabled
        `,
        [orgId, providerId, issuerUrl, clientId, JSON.stringify(scopes), enabled]
      );

      if (parsedBody.data.clientSecret !== undefined) {
        await putSecret(client, request.server.config.secretStoreKeys, secretName, clientSecret!);
      }

      const secretConfigured = parsedBody.data.clientSecret !== undefined
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
        resource: { type: "integration", id: providerId, name: "oidc" },
        success: true,
        details: { type: "oidc", providerId }
      })
    );

    const secretConfigured = txResult.secretConfigured;
    return reply.send({
      provider: {
        providerId,
        issuerUrl: String(txResult.row.issuer_url),
        clientId: String(txResult.row.client_id),
        scopes: ensureOpenIdScope(parseStringArray(txResult.row.scopes)),
        enabled: Boolean(txResult.row.enabled),
        createdAt: (txResult.row.created_at as Date).toISOString()
      },
      clientSecretConfigured: secretConfigured,
      secretConfigured
    });
  };

  app.put("/orgs/:orgId/oidc/providers/:providerId", { preHandler: requireAuth }, putProvider);
  app.put("/orgs/:orgId/oidc-providers/:providerId", { preHandler: requireAuth }, putProvider);

  const deleteProvider = async (request: FastifyRequest, reply: FastifyReply) => {
    const orgId = (request.params as { orgId: string; providerId: string }).orgId;
    const providerId = (request.params as { orgId: string; providerId: string }).providerId;
    if (!isValidProviderId(providerId)) return reply.code(400).send({ error: "invalid_request" });

    const member = await requireOrgAdminForOidcProviders(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
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
        resource: { type: "integration", id: providerId, name: "oidc" },
        success: true,
        details: { type: "oidc", providerId }
      })
    );

    return reply.send({ ok: true });
  };

  app.delete("/orgs/:orgId/oidc/providers/:providerId", { preHandler: requireAuth }, deleteProvider);
  app.delete("/orgs/:orgId/oidc-providers/:providerId", { preHandler: requireAuth }, deleteProvider);
}
