import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { generateScimToken, hashScimTokenSecret } from "../auth/scimTokens";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

async function requireOrgAdminForScimToken(
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

export function registerScimAdminRoutes(app: FastifyInstance): void {
  app.post("/orgs/:orgId/scim/token", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForScimToken(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const existing = await app.db.query("SELECT revoked_at FROM org_scim_tokens WHERE org_id = $1", [orgId]);
    const rotated = existing.rowCount === 1 && !existing.rows[0]!.revoked_at;

    const { token, secret } = generateScimToken(orgId);
    const tokenHash = hashScimTokenSecret(secret);

    await app.db.query(
      `
        INSERT INTO org_scim_tokens (org_id, token_hash)
        VALUES ($1, $2)
        ON CONFLICT (org_id)
        DO UPDATE SET token_hash = EXCLUDED.token_hash, created_at = now(), revoked_at = null
      `,
      [orgId, tokenHash]
    );

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.scim.token_created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "org_scim_token", id: orgId },
        success: true,
        details: { rotated }
      })
    );

    return reply.send({ token });
  });

  app.delete("/orgs/:orgId/scim/token", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForScimToken(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const res = await app.db.query(
      `
        UPDATE org_scim_tokens
        SET revoked_at = COALESCE(revoked_at, now())
        WHERE org_id = $1
        RETURNING revoked_at
      `,
      [orgId]
    );
    if (res.rowCount !== 1) return reply.code(404).send({ error: "scim_token_not_found" });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.scim.token_revoked",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "org_scim_token", id: orgId },
        success: true,
        details: {}
      })
    );

    return reply.send({ ok: true });
  });
}
