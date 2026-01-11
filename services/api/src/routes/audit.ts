import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { auditLogRowToAuditEvent, redactAuditEvent, serializeBatch } from "@formula/audit-core";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

async function requireOrgAdminRole(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<OrgRole | null> {
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
  return role;
}

export function registerAuditRoutes(app: FastifyInstance): void {
  const AuditQuery = z.object({
    start: z.string().datetime().optional(),
    end: z.string().datetime().optional(),
    eventType: z.string().optional(),
    userId: z.string().uuid().optional(),
    resourceId: z.string().optional(),
    success: z.enum(["true", "false"]).optional(),
    limit: z.string().optional(),
    offset: z.string().optional()
  });
  const AuditExportQuery = AuditQuery.extend({
    format: z.enum(["json", "cef", "leef"]).optional()
  });

  app.get(
    "/orgs/:orgId/audit",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const role = await requireOrgAdminRole(request, reply, orgId);
      if (!role) return;

      const parsed = AuditQuery.safeParse(request.query);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      const where: string[] = ["org_id = $1"];
      const values: unknown[] = [orgId];
      const add = (clause: string, value: unknown) => {
        values.push(value);
        where.push(clause.replace("$", `$${values.length}`));
      };

      if (parsed.data.start) add("created_at >= $", new Date(parsed.data.start));
      if (parsed.data.end) add("created_at <= $", new Date(parsed.data.end));
      if (parsed.data.eventType) add("event_type = $", parsed.data.eventType);
      if (parsed.data.userId) add("user_id = $", parsed.data.userId);
      if (parsed.data.resourceId) add("resource_id = $", parsed.data.resourceId);
      if (parsed.data.success) add("success = $", parsed.data.success === "true" ? true : false);

      const limit = parsed.data.limit ? Math.min(Number(parsed.data.limit), 1000) : 100;
      const offset = parsed.data.offset ? Math.max(Number(parsed.data.offset), 0) : 0;

      values.push(limit);
      values.push(offset);

      const columns =
        "id, org_id, user_id, user_email, event_type, resource_type, resource_id, ip_address, user_agent, session_id, success, error_code, error_message, details, created_at";
      const whereSql = where.join(" AND ");

      const result = await app.db.query(
        `
          SELECT ${columns}
          FROM (
            SELECT ${columns}
            FROM audit_log
            WHERE ${whereSql}
            UNION ALL
            SELECT ${columns}
            FROM audit_log_archive
            WHERE ${whereSql}
          ) audit
          ORDER BY created_at DESC
          LIMIT $${values.length - 1}
          OFFSET $${values.length}
        `,
        values
      );

      const events = result.rows.map((row) => redactAuditEvent(auditLogRowToAuditEvent(row as any)));
      return { events };
    }
  );

  app.get(
    "/orgs/:orgId/audit/export",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const role = await requireOrgAdminRole(request, reply, orgId);
      if (!role) return;

      const parsed = AuditExportQuery.safeParse(request.query);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      const where: string[] = ["org_id = $1"];
      const values: unknown[] = [orgId];
      const add = (clause: string, value: unknown) => {
        values.push(value);
        where.push(clause.replace("$", `$${values.length}`));
      };

      if (parsed.data.start) add("created_at >= $", new Date(parsed.data.start));
      if (parsed.data.end) add("created_at <= $", new Date(parsed.data.end));
      if (parsed.data.eventType) add("event_type = $", parsed.data.eventType);
      if (parsed.data.userId) add("user_id = $", parsed.data.userId);
      if (parsed.data.resourceId) add("resource_id = $", parsed.data.resourceId);
      if (parsed.data.success) add("success = $", parsed.data.success === "true" ? true : false);

      const columns =
        "id, org_id, user_id, user_email, event_type, resource_type, resource_id, ip_address, user_agent, session_id, success, error_code, error_message, details, created_at";
      const whereSql = where.join(" AND ");

      const result = await app.db.query(
        `
          SELECT ${columns}
          FROM (
            SELECT ${columns}
            FROM audit_log
            WHERE ${whereSql}
            UNION ALL
            SELECT ${columns}
            FROM audit_log_archive
            WHERE ${whereSql}
          ) audit
          ORDER BY created_at ASC
        `,
        values
      );

      const events = result.rows.map((row) => auditLogRowToAuditEvent(row as any));
      const format = parsed.data.format ?? "json";
      const { contentType, body } = serializeBatch(events, { format });

      reply.header("content-type", contentType);
      reply.send(body);
    }
  );
}
