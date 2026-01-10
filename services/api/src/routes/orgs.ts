import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { writeAuditEvent } from "../audit/audit";
import { validateDlpPolicy } from "../dlp/dlp";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

async function requireOrgMember(
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
  return { role: membership.rows[0].role as OrgRole };
}

export function registerOrgRoutes(app: FastifyInstance): void {
  app.get("/orgs", { preHandler: requireAuth }, async (request) => {
    const result = await app.db.query(
      `
        SELECT o.id, o.name, om.role
        FROM organizations o
        JOIN org_members om ON om.org_id = o.id
        WHERE om.user_id = $1
        ORDER BY o.created_at ASC
      `,
      [request.user!.id]
    );
    return {
      organizations: result.rows.map((row) => ({
        id: row.id as string,
        name: row.name as string,
        role: row.role as string
      }))
    };
  });

  app.get("/orgs/:orgId", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgMember(request, reply, orgId);
    if (!member) return;

    const org = await app.db.query("SELECT id, name FROM organizations WHERE id = $1", [orgId]);
    const settings = await app.db.query("SELECT * FROM org_settings WHERE org_id = $1", [orgId]);

    return {
      organization: org.rows[0],
      role: member.role,
      settings: settings.rows[0]
    };
  });

  const PatchSettingsBody = z.object({
    requireMfa: z.boolean().optional(),
    allowedAuthMethods: z.array(z.string()).optional(),
    ipAllowlist: z.array(z.string()).optional(),
    allowExternalSharing: z.boolean().optional(),
    allowPublicLinks: z.boolean().optional(),
    defaultPermission: z.enum(["viewer", "commenter", "editor"]).optional(),
    aiEnabled: z.boolean().optional(),
    aiDataProcessingConsent: z.boolean().optional(),
    dataResidencyRegion: z.string().min(1).optional(),
    allowCrossRegionProcessing: z.boolean().optional(),
    aiProcessingRegion: z.string().min(1).nullable().optional(),
    auditLogRetentionDays: z.number().int().positive().optional(),
    documentVersionRetentionDays: z.number().int().positive().optional(),
    deletedDocumentRetentionDays: z.number().int().positive().optional()
  });

  app.patch("/orgs/:orgId/settings", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgMember(request, reply, orgId);
    if (!member) return;
    if (!isOrgAdmin(member.role)) return reply.code(403).send({ error: "forbidden" });

    const parsed = PatchSettingsBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const updates = parsed.data;
    const sets: string[] = [];
    const values: unknown[] = [];
    const addSet = (sql: string, value: unknown) => {
      values.push(value);
      sets.push(`${sql} = $${values.length}`);
    };

    if (updates.requireMfa !== undefined) addSet("require_mfa", updates.requireMfa);
    if (updates.allowedAuthMethods !== undefined)
      addSet("allowed_auth_methods", JSON.stringify(updates.allowedAuthMethods));
    if (updates.ipAllowlist !== undefined) addSet("ip_allowlist", JSON.stringify(updates.ipAllowlist));
    if (updates.allowExternalSharing !== undefined) addSet("allow_external_sharing", updates.allowExternalSharing);
    if (updates.allowPublicLinks !== undefined) addSet("allow_public_links", updates.allowPublicLinks);
    if (updates.defaultPermission !== undefined) addSet("default_permission", updates.defaultPermission);
    if (updates.aiEnabled !== undefined) addSet("ai_enabled", updates.aiEnabled);
    if (updates.aiDataProcessingConsent !== undefined)
      addSet("ai_data_processing_consent", updates.aiDataProcessingConsent);
    if (updates.dataResidencyRegion !== undefined) addSet("data_residency_region", updates.dataResidencyRegion);
    if (updates.allowCrossRegionProcessing !== undefined)
      addSet("allow_cross_region_processing", updates.allowCrossRegionProcessing);
    if (updates.aiProcessingRegion !== undefined) addSet("ai_processing_region", updates.aiProcessingRegion);
    if (updates.auditLogRetentionDays !== undefined)
      addSet("audit_log_retention_days", updates.auditLogRetentionDays);
    if (updates.documentVersionRetentionDays !== undefined)
      addSet("document_version_retention_days", updates.documentVersionRetentionDays);
    if (updates.deletedDocumentRetentionDays !== undefined)
      addSet("deleted_document_retention_days", updates.deletedDocumentRetentionDays);

    if (sets.length === 0) return reply.send({ ok: true });

    values.push(orgId);
    await app.db.query(
      `
        UPDATE org_settings
        SET ${sets.join(", ")}, updated_at = now()
        WHERE org_id = $${values.length}
      `,
      values
    );

    await writeAuditEvent(app.db, {
      orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "admin.settings_changed",
      resourceType: "organization",
      resourceId: orgId,
      sessionId: request.session?.id,
      success: true,
      details: { updates },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ ok: true });
  });

  app.get("/orgs/:orgId/dlp-policy", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgMember(request, reply, orgId);
    if (!member) return;

    const res = await app.db.query("SELECT policy FROM org_dlp_policies WHERE org_id = $1", [orgId]);
    if (res.rowCount !== 1) return reply.code(404).send({ error: "dlp_policy_not_found" });
    return reply.send({ policy: res.rows[0].policy });
  });

  const PutDlpPolicyBody = z.object({ policy: z.unknown() });

  app.put("/orgs/:orgId/dlp-policy", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgMember(request, reply, orgId);
    if (!member) return;
    if (!isOrgAdmin(member.role)) return reply.code(403).send({ error: "forbidden" });

    const parsed = PutDlpPolicyBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    try {
      validateDlpPolicy(parsed.data.policy);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Invalid policy";
      return reply.code(400).send({ error: "invalid_policy", message });
    }

    await app.db.query(
      `
        INSERT INTO org_dlp_policies (org_id, policy)
        VALUES ($1, $2)
        ON CONFLICT (org_id)
        DO UPDATE SET policy = EXCLUDED.policy, updated_at = now()
      `,
      [orgId, JSON.stringify(parsed.data.policy)]
    );

    await writeAuditEvent(app.db, {
      orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "admin.settings_changed",
      resourceType: "organization",
      resourceId: orgId,
      sessionId: request.session?.id,
      success: true,
      details: { updates: { dlpPolicy: true } },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ policy: parsed.data.policy });
  });
}
