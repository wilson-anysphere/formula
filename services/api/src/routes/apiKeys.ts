import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { generateApiKeyToken, hashApiKeySecret } from "../auth/apiKeys";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

async function requireOrgAdminForKeyManagement(
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

export function registerApiKeyRoutes(app: FastifyInstance): void {
  const CreateBody = z.object({
    name: z.string().min(1).max(100)
  });

  app.post("/orgs/:orgId/api-keys", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForKeyManagement(request, reply, orgId);
    if (!member) return;

    const parsedBody = CreateBody.safeParse(request.body);
    if (!parsedBody.success) return reply.code(400).send({ error: "invalid_request" });

    const name = parsedBody.data.name.trim();
    if (name.length === 0) return reply.code(400).send({ error: "invalid_request" });

    const { apiKeyId, token, secret } = generateApiKeyToken();
    const keyHash = hashApiKeySecret(secret);

    const inserted = await app.db.query(
      `
        INSERT INTO api_keys (id, org_id, name, key_hash, created_by)
        VALUES ($1, $2, $3, $4, $5)
        ON CONFLICT (org_id, name) DO NOTHING
        RETURNING created_at
      `,
      [apiKeyId, orgId, name, keyHash, request.user!.id]
    );

    if (inserted.rowCount !== 1) return reply.code(409).send({ error: "name_in_use" });
    const createdAt = inserted.rows[0]!.created_at as Date;

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.api_key.created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "api_key", id: apiKeyId, name },
        success: true,
        details: {}
      })
    );

    return reply.send({
      apiKey: {
        id: apiKeyId,
        orgId,
        name,
        createdBy: request.user!.id,
        createdAt: createdAt.toISOString()
      },
      key: token
    });
  });

  app.get("/orgs/:orgId/api-keys", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdminForKeyManagement(request, reply, orgId);
    if (!member) return;

    const res = await app.db.query(
      `
        SELECT id, name, created_by, created_at, last_used_at, revoked_at
        FROM api_keys
        WHERE org_id = $1
        ORDER BY created_at DESC
      `,
      [orgId]
    );

    return reply.send({
      apiKeys: res.rows.map((row) => ({
        id: row.id as string,
        orgId,
        name: row.name as string,
        createdBy: row.created_by as string,
        createdAt: (row.created_at as Date).toISOString(),
        lastUsedAt: row.last_used_at ? (row.last_used_at as Date).toISOString() : null,
        revokedAt: row.revoked_at ? (row.revoked_at as Date).toISOString() : null
      }))
    });
  });

  app.delete("/orgs/:orgId/api-keys/:id", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string; id: string }).orgId;
    const apiKeyId = (request.params as { orgId: string; id: string }).id;
    const member = await requireOrgAdminForKeyManagement(request, reply, orgId);
    if (!member) return;

    const res = await app.db.query(
      `
        UPDATE api_keys
        SET revoked_at = COALESCE(revoked_at, now())
        WHERE id = $1 AND org_id = $2
        RETURNING name, revoked_at
      `,
      [apiKeyId, orgId]
    );

    if (res.rowCount !== 1) return reply.code(404).send({ error: "api_key_not_found" });
    const name = res.rows[0]!.name as string;

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.api_key.revoked",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "api_key", id: apiKeyId, name },
        success: true,
        details: {}
      })
    );

    return reply.send({ ok: true });
  });
}
