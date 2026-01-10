import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { writeAuditEvent } from "../audit/audit";
import { isMfaEnforcedForOrg } from "../auth/mfa";
import { normalizeClassification, selectorKey, validateDlpPolicy } from "../dlp/dlp";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { canDocument, type DocumentRole } from "../rbac/roles";
import { signSyncToken } from "../sync/token";
import { withTransaction } from "../db/tx";
import { requireAuth } from "./auth";

type ShareLinkRole = Exclude<DocumentRole, "owner" | "admin">;
type ShareLinkVisibility = "public" | "private";

function roleRank(role: DocumentRole): number {
  switch (role) {
    case "owner":
      return 5;
    case "admin":
      return 4;
    case "editor":
      return 3;
    case "commenter":
      return 2;
    case "viewer":
      return 1;
  }
}

function hashShareToken(token: string): string {
  return crypto.createHash("sha256").update(token).digest("hex");
}

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

  const ShareLinkBody = z
    .object({
      visibility: z.enum(["public", "private"]).default("private"),
      role: z.enum(["editor", "commenter", "viewer"]).default("viewer"),
      expiresInSeconds: z.number().int().positive().optional(),
      expiresAt: z.string().datetime().optional()
    })
    .refine((value) => !(value.expiresAt && value.expiresInSeconds), {
      message: "expiresAt and expiresInSeconds are mutually exclusive"
    });

  app.post("/docs/:docId/share-links", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "share")) return reply.code(403).send({ error: "forbidden" });

    const parsed = ShareLinkBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const settings = await app.db.query(
      "SELECT allow_public_links FROM org_settings WHERE org_id = $1",
      [membership.orgId]
    );
    const allowPublicLinks = settings.rowCount === 1 ? Boolean((settings.rows[0] as any).allow_public_links) : true;
    if (parsed.data.visibility === "public" && !allowPublicLinks) {
      return reply.code(403).send({ error: "public_links_disabled" });
    }

    const now = new Date();
    const expiresAt = parsed.data.expiresAt
      ? new Date(parsed.data.expiresAt)
      : parsed.data.expiresInSeconds
        ? new Date(now.getTime() + parsed.data.expiresInSeconds * 1000)
        : null;

    const token = crypto.randomBytes(32).toString("base64url");
    const tokenHash = hashShareToken(token);
    const id = crypto.randomUUID();

    await app.db.query(
      `
        INSERT INTO document_share_links (id, document_id, token_hash, visibility, role, created_by, expires_at)
        VALUES ($1,$2,$3,$4,$5,$6,$7)
      `,
      [id, docId, tokenHash, parsed.data.visibility, parsed.data.role, request.user!.id, expiresAt]
    );

    await writeAuditEvent(app.db, {
      orgId: membership.orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "sharing.link_created",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { visibility: parsed.data.visibility, role: parsed.data.role, expiresAt: expiresAt?.toISOString() ?? null },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({
      shareLink: {
        id,
        token,
        visibility: parsed.data.visibility as ShareLinkVisibility,
        role: parsed.data.role as ShareLinkRole,
        expiresAt: expiresAt?.toISOString() ?? null
      }
    });
  });

  app.get("/docs/:docId/share-links", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "share")) return reply.code(403).send({ error: "forbidden" });

    const links = await app.db.query(
      `
        SELECT id, visibility, role, created_at, expires_at, revoked_at
        FROM document_share_links
        WHERE document_id = $1
        ORDER BY created_at DESC
      `,
      [docId]
    );

    return reply.send({ shareLinks: links.rows });
  });

  app.post("/share-links/:token/redeem", { preHandler: requireAuth }, async (request, reply) => {
    const token = (request.params as { token: string }).token;
    if (!token) return reply.code(400).send({ error: "invalid_request" });

    const tokenHash = hashShareToken(token);
    const linkRes = await app.db.query(
      `
        SELECT l.id, l.document_id, l.visibility, l.role, l.expires_at, l.revoked_at, d.org_id
        FROM document_share_links l
        JOIN documents d ON d.id = l.document_id
        WHERE l.token_hash = $1
        LIMIT 1
      `,
      [tokenHash]
    );

    if (linkRes.rowCount !== 1) return reply.code(404).send({ error: "share_link_not_found" });

    const linkRow = linkRes.rows[0] as any;
    const expiresAt = linkRow.expires_at ? new Date(linkRow.expires_at as string) : null;
    if (linkRow.revoked_at || (expiresAt && Date.now() > expiresAt.getTime())) {
      return reply.code(404).send({ error: "share_link_not_found" });
    }

    const docId = linkRow.document_id as string;
    const orgId = linkRow.org_id as string;
    const visibility = linkRow.visibility as ShareLinkVisibility;
    const linkRole = linkRow.role as ShareLinkRole;

    const orgSettings = await app.db.query(
      "SELECT allow_external_sharing FROM org_settings WHERE org_id = $1",
      [orgId]
    );
    const allowExternalSharing =
      orgSettings.rowCount === 1 ? Boolean((orgSettings.rows[0] as any).allow_external_sharing) : true;

    const orgMembership = await app.db.query(
      "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
      [orgId, request.user!.id]
    );

    if (visibility === "private" || !allowExternalSharing) {
      if (orgMembership.rowCount !== 1) return reply.code(403).send({ error: "forbidden" });
    } else if (orgMembership.rowCount !== 1) {
      await app.db.query(
        "INSERT INTO org_members (org_id, user_id, role) VALUES ($1,$2,'member')",
        [orgId, request.user!.id]
      );
    }

    const existingDocMember = await app.db.query(
      "SELECT role FROM document_members WHERE document_id = $1 AND user_id = $2",
      [docId, request.user!.id]
    );

    const nextRole: DocumentRole =
      existingDocMember.rowCount === 1 &&
      roleRank(existingDocMember.rows[0].role as DocumentRole) > roleRank(linkRole)
        ? (existingDocMember.rows[0].role as DocumentRole)
        : linkRole;

    await withTransaction(app.db, async (client) => {
      await client.query(
        `
          INSERT INTO document_members (document_id, user_id, role, created_by)
          VALUES ($1,$2,$3,$4)
          ON CONFLICT (document_id, user_id)
          DO UPDATE SET role = EXCLUDED.role
        `,
        [docId, request.user!.id, nextRole, request.user!.id]
      );
    });

    await writeAuditEvent(app.db, {
      orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "sharing.added",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { via: "share-link", shareLinkId: linkRow.id, role: nextRole },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ ok: true, documentId: docId, role: nextRole });
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

  app.get("/docs/:docId/permissions", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;

    const rows = await app.db.query(
      `
        SELECT sheet_name, start_row, start_col, end_row, end_col, permission_type, allowed_user_id
        FROM document_range_permissions
        WHERE document_id = $1
      `,
      [docId]
    );

    type Restriction = {
      sheetName: string;
      startRow: number;
      startCol: number;
      endRow: number;
      endCol: number;
      readAllowlist: string[];
      editAllowlist: string[];
    };

    const byRange = new Map<string, Restriction>();
    const keyFor = (r: any) => `${r.sheet_name}:${r.start_row}:${r.start_col}:${r.end_row}:${r.end_col}`;

    for (const row of rows.rows as any[]) {
      const key = keyFor(row);
      let restriction = byRange.get(key);
      if (!restriction) {
        restriction = {
          sheetName: row.sheet_name as string,
          startRow: Number(row.start_row),
          startCol: Number(row.start_col),
          endRow: Number(row.end_row),
          endCol: Number(row.end_col),
          readAllowlist: [],
          editAllowlist: []
        };
        byRange.set(key, restriction);
      }
      if (row.permission_type === "read") restriction.readAllowlist.push(row.allowed_user_id as string);
      if (row.permission_type === "edit") restriction.editAllowlist.push(row.allowed_user_id as string);
    }

    return reply.send({
      permissions: {
        role: membership.role,
        rangeRestrictions: Array.from(byRange.values())
      }
    });
  });

  app.get("/docs/:docId/dlp-policy", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "read")) return reply.code(403).send({ error: "forbidden" });

    const res = await app.db.query("SELECT policy FROM document_dlp_policies WHERE document_id = $1", [docId]);
    if (res.rowCount !== 1) return reply.code(404).send({ error: "dlp_policy_not_found" });
    return reply.send({ policy: res.rows[0].policy });
  });

  const PutDocDlpPolicyBody = z.object({ policy: z.unknown() });

  app.put("/docs/:docId/dlp-policy", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const parsed = PutDocDlpPolicyBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    try {
      validateDlpPolicy(parsed.data.policy);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Invalid policy";
      return reply.code(400).send({ error: "invalid_policy", message });
    }

    await app.db.query(
      `
        INSERT INTO document_dlp_policies (document_id, policy)
        VALUES ($1, $2)
        ON CONFLICT (document_id)
        DO UPDATE SET policy = EXCLUDED.policy, updated_at = now()
      `,
      [docId, JSON.stringify(parsed.data.policy)]
    );

    await writeAuditEvent(app.db, {
      orgId: membership.orgId,
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "admin.settings_changed",
      resourceType: "document",
      resourceId: docId,
      sessionId: request.session?.id,
      success: true,
      details: { type: "dlp-policy" },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ policy: parsed.data.policy });
  });

  app.get("/docs/:docId/classifications", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "read")) return reply.code(403).send({ error: "forbidden" });

    const res = await app.db.query(
      `
        SELECT selector, classification, updated_at
        FROM document_classifications
        WHERE document_id = $1
        ORDER BY updated_at DESC
      `,
      [docId]
    );

    return {
      classifications: res.rows.map((row) => ({
        selector: row.selector,
        classification: row.classification,
        updatedAt: row.updated_at
      }))
    };
  });

  const UpsertClassificationBody = z.object({
    selector: z.unknown(),
    classification: z.unknown()
  });

  app.put("/docs/:docId/classifications", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "edit")) return reply.code(403).send({ error: "forbidden" });

    const parsed = UpsertClassificationBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    let selector;
    let classification;
    let key;

    try {
      selector = parsed.data.selector;
      if (typeof selector !== "object" || selector === null) throw new Error("Selector must be an object");
      if ((selector as any).documentId !== docId) throw new Error("Selector documentId must match route docId");

      classification = normalizeClassification(parsed.data.classification);
      key = selectorKey(selector);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Invalid classification";
      return reply.code(400).send({ error: "invalid_request", message });
    }

    const id = crypto.randomUUID();
    await app.db.query(
      `
        INSERT INTO document_classifications (id, document_id, selector_key, selector, classification)
        VALUES ($1, $2, $3, $4, $5)
        ON CONFLICT (document_id, selector_key)
        DO UPDATE SET selector = EXCLUDED.selector, classification = EXCLUDED.classification, updated_at = now()
      `,
      [id, docId, key, JSON.stringify(selector), JSON.stringify(classification)]
    );

    return reply.send({ ok: true });
  });

  app.delete("/docs/:docId/classifications/:selectorKey", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string; selectorKey: string }).docId;
    const selectorKeyParam = (request.params as { docId: string; selectorKey: string }).selectorKey;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "edit")) return reply.code(403).send({ error: "forbidden" });

    const res = await app.db.query(
      "DELETE FROM document_classifications WHERE document_id = $1 AND selector_key = $2",
      [docId, selectorKeyParam]
    );
    if (res.rowCount !== 1) return reply.code(404).send({ error: "not_found" });
    return reply.send({ ok: true });
  });
}
