import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { generateScimToken, hashScimTokenSecret } from "../auth/scim";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

async function requireOrgAdminForScimTokenManagement(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
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

export function registerScimAdminRoutes(app: FastifyInstance): void {
  const CreateBody = z.object({
    name: z.string().min(1).max(100)
  });

  app.post(
    "/orgs/:orgId/scim/tokens",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const member = await requireOrgAdminForScimTokenManagement(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const parsedBody = CreateBody.safeParse(request.body);
      if (!parsedBody.success) return reply.code(400).send({ error: "invalid_request" });

      const name = parsedBody.data.name.trim();
      if (name.length === 0) return reply.code(400).send({ error: "invalid_request" });

      const { tokenId, token, secret } = generateScimToken();
      const tokenHash = hashScimTokenSecret(secret);

      const inserted = await app.db.query(
        `
          INSERT INTO org_scim_tokens (id, org_id, name, token_hash, created_by)
          VALUES ($1, $2, $3, $4, $5)
          ON CONFLICT (org_id, name) DO NOTHING
          RETURNING created_at
        `,
        [tokenId, orgId, name, tokenHash, request.user!.id]
      );

      if (inserted.rowCount !== 1) return reply.code(409).send({ error: "name_in_use" });

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "org.scim_token.created",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id ?? null,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "scim_token", id: tokenId, name },
          success: true,
          details: {}
        })
      );

      return reply.send({ id: tokenId, name, token });
    }
  );

  app.get(
    "/orgs/:orgId/scim/tokens",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const member = await requireOrgAdminForScimTokenManagement(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const res = await app.db.query(
        `
          SELECT id, name, created_by, created_at, last_used_at, revoked_at
          FROM org_scim_tokens
          WHERE org_id = $1
          ORDER BY created_at DESC
        `,
        [orgId]
      );

      return reply.send({
        tokens: res.rows.map((row) => ({
          id: row.id as string,
          orgId,
          name: row.name as string,
          createdBy: row.created_by as string,
          createdAt: (row.created_at as Date).toISOString(),
          lastUsedAt: row.last_used_at ? (row.last_used_at as Date).toISOString() : null,
          revokedAt: row.revoked_at ? (row.revoked_at as Date).toISOString() : null
        }))
      });
    }
  );

  app.delete(
    "/orgs/:orgId/scim/tokens/:tokenId",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string; tokenId: string }).orgId;
      const tokenId = (request.params as { orgId: string; tokenId: string }).tokenId;
      const member = await requireOrgAdminForScimTokenManagement(request, reply, orgId);
      if (!member) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const res = await app.db.query(
        `
          UPDATE org_scim_tokens
          SET revoked_at = COALESCE(revoked_at, now())
          WHERE id = $1 AND org_id = $2
          RETURNING name
        `,
        [tokenId, orgId]
      );

      if (res.rowCount !== 1) return reply.code(404).send({ error: "scim_token_not_found" });
      const name = res.rows[0]!.name as string;

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "org.scim_token.revoked",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id ?? null,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "scim_token", id: tokenId, name },
          success: true,
          details: {}
        })
      );

      return reply.send({ ok: true });
    }
  );
}

