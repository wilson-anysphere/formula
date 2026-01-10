import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { writeAuditEvent } from "../audit/audit";
import { isMfaEnforcedForOrg } from "../auth/mfa";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { canDocument, type DocumentRole } from "../rbac/roles";
import { signSyncToken } from "../sync/token";
import { withTransaction } from "../db/tx";
import { requireAuth } from "./auth";

async function getDocMembership(
  request: FastifyRequest,
  docId: string
): Promise<{
  orgId: string;
  role: DocumentRole | null;
}> {
  const result = await request.server.db.query(
    `
      SELECT d.org_id, dm.role
      FROM documents d
      LEFT JOIN document_members dm
        ON dm.document_id = d.id AND dm.user_id = $2
      WHERE d.id = $1
      LIMIT 1
    `,
    [docId, request.user!.id]
  );

  if (result.rowCount !== 1) {
    return { orgId: "", role: null };
  }

  const row = result.rows[0] as { org_id: string; role: DocumentRole | null };
  return { orgId: row.org_id, role: row.role };
}

async function requireDocRole(
  request: FastifyRequest,
  reply: FastifyReply,
  docId: string
): Promise<{ orgId: string; role: DocumentRole } | null> {
  const membership = await getDocMembership(request, docId);
  if (!membership.orgId) {
    reply.code(404).send({ error: "doc_not_found" });
    return null;
  }
  if (!membership.role) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return { orgId: membership.orgId, role: membership.role };
}

async function requireOrgMembership(request: FastifyRequest, orgId: string): Promise<boolean> {
  const membership = await request.server.db.query(
    "SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  return membership.rowCount === 1;
}

export function registerDocRoutes(app: FastifyInstance): void {
  const CreateDocBody = z.object({
    orgId: z.string().uuid(),
    title: z.string().min(1)
  });

  app.post("/docs", { preHandler: requireAuth }, async (request, reply) => {
    const parsed = CreateDocBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const { orgId, title } = parsed.data;
    if (!(await requireOrgMembership(request, orgId))) {
      return reply.code(403).send({ error: "forbidden" });
    }

    if ((await isMfaEnforcedForOrg(app.db, orgId)) && !request.user!.mfaTotpEnabled) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const docId = crypto.randomUUID();
    await withTransaction(app.db, async (client) => {
      await client.query(
        `
          INSERT INTO documents (id, org_id, title, created_by)
          VALUES ($1, $2, $3, $4)
        `,
        [docId, orgId, title, request.user!.id]
      );
      await client.query(
        `
          INSERT INTO document_members (document_id, user_id, role, created_by)
          VALUES ($1, $2, 'owner', $2)
        `,
        [docId, request.user!.id]
      );
    });

    await writeAuditEvent(app.db, {
      orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "document.created",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { title },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ document: { id: docId, orgId, title } });
  });

  app.get("/docs/:docId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;

    const doc = await app.db.query("SELECT id, org_id, title FROM documents WHERE id = $1", [docId]);
    return { document: doc.rows[0], role: membership.role };
  });

  const InviteBody = z.object({
    email: z.string().email(),
    role: z.enum(["owner", "admin", "editor", "commenter", "viewer"])
  });

  app.post("/docs/:docId/invite", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "share")) return reply.code(403).send({ error: "forbidden" });

    const parsed = InviteBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const email = parsed.data.email.trim().toLowerCase();
    const invitedRole = parsed.data.role as DocumentRole;

    const userRes = await app.db.query("SELECT id, email FROM users WHERE email = $1", [email]);
    if (userRes.rowCount !== 1) return reply.code(404).send({ error: "user_not_found" });

    const invitedUserId = userRes.rows[0].id as string;

    await withTransaction(app.db, async (client) => {
      // Ensure org membership for invited user.
      const orgMembership = await client.query(
        "SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2",
        [membership.orgId, invitedUserId]
      );
      if (orgMembership.rowCount !== 1) {
        await client.query(
          "INSERT INTO org_members (org_id, user_id, role) VALUES ($1, $2, 'member')",
          [membership.orgId, invitedUserId]
        );
      }

      // Upsert document membership.
      await client.query(
        `
          INSERT INTO document_members (document_id, user_id, role, created_by)
          VALUES ($1, $2, $3, $4)
          ON CONFLICT (document_id, user_id)
          DO UPDATE SET role = EXCLUDED.role
        `,
        [docId, invitedUserId, invitedRole, request.user!.id]
      );
    });

    await writeAuditEvent(app.db, {
      orgId: membership.orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "sharing.added",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { invitedEmail: email, role: invitedRole },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ ok: true });
  });

  app.get("/docs/:docId/members", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "share")) return reply.code(403).send({ error: "forbidden" });

    const members = await app.db.query(
      `
        SELECT u.id, u.email, u.name, dm.role
        FROM document_members dm
        JOIN users u ON u.id = dm.user_id
        WHERE dm.document_id = $1
        ORDER BY u.email ASC
      `,
      [docId]
    );
    return { members: members.rows };
  });

  app.post("/docs/:docId/sync-token", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;

    if (!canDocument(membership.role, "read")) return reply.code(403).send({ error: "forbidden" });
    if ((await isMfaEnforcedForOrg(app.db, membership.orgId)) && !request.user!.mfaTotpEnabled) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const { token, expiresAt } = signSyncToken({
      secret: app.config.syncTokenSecret,
      ttlSeconds: app.config.syncTokenTtlSeconds,
      claims: {
        sub: request.user!.id,
        docId,
        orgId: membership.orgId,
        role: membership.role,
        sessionId: request.session?.id
      }
    });

    await writeAuditEvent(app.db, {
      orgId: membership.orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "document.opened",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { via: "sync-token" },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ token, expiresAt: expiresAt.toISOString() });
  });

  const RangePermissionBody = z.object({
    sheetName: z.string().min(1),
    startRow: z.number().int().nonnegative(),
    startCol: z.number().int().nonnegative(),
    endRow: z.number().int().nonnegative(),
    endCol: z.number().int().nonnegative(),
    permissionType: z.enum(["read", "edit"]),
    allowedUserEmail: z.string().email()
  });

  app.post("/docs/:docId/range-permissions", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const parsed = RangePermissionBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const allowedEmail = parsed.data.allowedUserEmail.trim().toLowerCase();
    const allowedUser = await app.db.query("SELECT id FROM users WHERE email = $1", [allowedEmail]);
    if (allowedUser.rowCount !== 1) return reply.code(404).send({ error: "user_not_found" });

    const id = crypto.randomUUID();
    await app.db.query(
      `
        INSERT INTO document_range_permissions (
          id, document_id, sheet_name, start_row, start_col, end_row, end_col, permission_type, allowed_user_id, created_by
        )
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
      `,
      [
        id,
        docId,
        parsed.data.sheetName,
        parsed.data.startRow,
        parsed.data.startCol,
        parsed.data.endRow,
        parsed.data.endCol,
        parsed.data.permissionType,
        allowedUser.rows[0].id as string,
        request.user!.id
      ]
    );

    await writeAuditEvent(app.db, {
      orgId: membership.orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "sharing.modified",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: {
        type: "range-permission",
        permission: parsed.data
      },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ ok: true, id });
  });

  app.get("/docs/:docId/range-permissions", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const rows = await app.db.query(
      `
        SELECT id, sheet_name, start_row, start_col, end_row, end_col, permission_type, allowed_user_id
        FROM document_range_permissions
        WHERE document_id = $1
        ORDER BY created_at DESC
      `,
      [docId]
    );
    return { rangePermissions: rows.rows };
  });
}
