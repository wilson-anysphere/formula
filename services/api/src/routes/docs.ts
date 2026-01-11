import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import {
  enforceOrgIpAllowlistForSession,
  enforceOrgIpAllowlistForSessionWithAllowlist
} from "../auth/orgIpAllowlist";
import { isMfaEnforcedForOrg } from "../auth/mfa";
import { encryptEnvelope, ENVELOPE_VERSION } from "../crypto/envelope";
import { createKeyring } from "../crypto/keyring";
import { DLP_ACTION, DLP_DECISION, normalizeClassification, selectorKey, validateDlpPolicy } from "../dlp/dlp";
import { evaluateDocumentDlpPolicy } from "../dlp/effective";
import { createDocumentVersion, getDocumentVersionData } from "../db/documentVersions";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { canDocument, type DocumentRole } from "../rbac/roles";
import { signSyncToken } from "../sync/token";
import { withTransaction } from "../db/tx";
import { requireAuth } from "./auth";

type ShareLinkRole = Exclude<DocumentRole, "owner" | "admin">;
type ShareLinkVisibility = "public" | "private";

function decodeBase64Strict(input: string): Buffer | null {
  // Buffer.from(..., "base64") is permissive and does not throw for many invalid inputs.
  // We validate the alphabet + padding and ensure a stable round-trip to reject malformed base64.
  if (typeof input !== "string") return null;
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(input)) return null;
  if (input.length % 4 !== 0) return null;

  const decoded = Buffer.from(input, "base64");
  const normalizedInput = input.replace(/=+$/, "");
  const normalizedRoundTrip = decoded.toString("base64").replace(/=+$/, "");
  if (normalizedInput !== normalizedRoundTrip) return null;
  return decoded;
}

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
  deletedAt: Date | null;
  role: DocumentRole | null;
  ipAllowlist: unknown;
}> {
  const result = await request.server.db.query(
    `
      SELECT d.org_id, d.deleted_at, dm.role, os.ip_allowlist
      FROM documents d
      LEFT JOIN org_settings os ON os.org_id = d.org_id
      LEFT JOIN document_members dm
        ON dm.document_id = d.id AND dm.user_id = $2
      WHERE d.id = $1
      LIMIT 1
    `,
    [docId, request.user!.id]
  );

  if (result.rowCount !== 1) {
    return { orgId: "", deletedAt: null, role: null, ipAllowlist: null };
  }

  const row = result.rows[0] as { org_id: string; deleted_at: Date | null; role: DocumentRole | null; ip_allowlist: unknown };
  return { orgId: row.org_id, deletedAt: row.deleted_at, role: row.role, ipAllowlist: row.ip_allowlist };
}

async function requireDocRole(
  request: FastifyRequest,
  reply: FastifyReply,
  docId: string
): Promise<{ orgId: string; deletedAt: Date | null; role: DocumentRole } | null> {
  const membership = await getDocMembership(request, docId);
  if (!membership.orgId) {
    reply.code(404).send({ error: "doc_not_found" });
    return null;
  }

  if (
    !(await enforceOrgIpAllowlistForSessionWithAllowlist(
      request,
      reply,
      membership.orgId,
      membership.ipAllowlist
    ))
  ) {
    return null;
  }

  if (!membership.role) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return { orgId: membership.orgId, deletedAt: membership.deletedAt, role: membership.role };
}

async function requireOrgMembership(request: FastifyRequest, orgId: string): Promise<boolean> {
  const membership = await request.server.db.query(
    "SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  return membership.rowCount === 1;
}

export function registerDocRoutes(app: FastifyInstance): void {
  const keyring = createKeyring(app.config);
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

    if (!(await enforceOrgIpAllowlistForSession(request, reply, orgId))) return;

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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "document.created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { title }
      })
    );

    return reply.send({ document: { id: docId, orgId, title } });
  });

  app.get("/docs/:docId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;

    const doc = await app.db.query("SELECT id, org_id, title FROM documents WHERE id = $1", [docId]);
    return { document: doc.rows[0], role: membership.role };
  });

  app.delete("/docs/:docId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    await app.db.query(
      `
        UPDATE documents
        SET deleted_at = COALESCE(deleted_at, now()), updated_at = now()
        WHERE id = $1
      `,
      [docId]
    );

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "document.deleted",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { soft: true }
      })
    );

    return reply.send({ ok: true });
  });

  const CreateDocVersionBody = z.object({
    description: z.string().max(1000).optional(),
    dataBase64: z.string()
  });

  app.post("/docs/:docId/versions", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "edit")) return reply.code(403).send({ error: "forbidden" });
    if (membership.deletedAt) return reply.code(403).send({ error: "doc_deleted" });

    const parsed = CreateDocVersionBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const data = decodeBase64Strict(parsed.data.dataBase64);
    if (!data) return reply.code(400).send({ error: "invalid_request" });

    const created = await createDocumentVersion(app.db, keyring, {
      documentId: docId,
      createdBy: request.user!.id,
      description: parsed.data.description ?? null,
      data
    });

    const createdAt = created.createdAt;
    const sizeBytes = data.length;

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "document.version_created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "document", id: docId },
          success: true,
          details: { versionId: created.id, description: parsed.data.description ?? null, sizeBytes }
        })
      );

    return reply.send({
      version: {
        id: created.id,
        createdAt,
        description: parsed.data.description ?? null,
        sizeBytes
      }
    });
  });

  app.get("/docs/:docId/versions", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "read")) return reply.code(403).send({ error: "forbidden" });

    const versions = await app.db.query(
      `
        SELECT id, created_at, created_by, description, data
        FROM document_versions
        WHERE document_id = $1
        ORDER BY created_at DESC
      `,
      [docId]
    );

    const enriched = await Promise.all(
      versions.rows.map(async (row: any) => {
        const plaintextBytes = row.data
          ? Buffer.from(row.data as any)
          : await getDocumentVersionData(app.db, keyring, String(row.id), { documentId: docId });
        const sizeBytes = plaintextBytes ? plaintextBytes.length : 0;

        return {
          id: row.id as string,
          createdAt: row.created_at as Date,
          createdBy: row.created_by as string | null,
          description: row.description as string | null,
          sizeBytes
        };
      })
    );

    return reply.send({
      versions: enriched
    });
  });

  app.get("/docs/:docId/versions/:versionId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string; versionId: string }).docId;
    const versionId = (request.params as { docId: string; versionId: string }).versionId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "read")) return reply.code(403).send({ error: "forbidden" });

    const res = await app.db.query(
      `
        SELECT id, created_at, created_by, description, data
        FROM document_versions
        WHERE document_id = $1 AND id = $2
        LIMIT 1
      `,
      [docId, versionId]
    );
    if (res.rowCount !== 1) return reply.code(404).send({ error: "version_not_found" });

    const row = res.rows[0] as any;
    const bytes =
      row.data ? Buffer.from(row.data as any) : (await getDocumentVersionData(app.db, keyring, versionId, { documentId: docId })) ?? Buffer.alloc(0);
    return reply.send({
      version: {
        id: row.id as string,
        createdAt: row.created_at as Date,
        createdBy: row.created_by as string | null,
        description: (row.description as string | null) ?? null,
        dataBase64: bytes.toString("base64")
      }
    });
  });

  app.delete("/docs/:docId/versions/:versionId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string; versionId: string }).docId;
    const versionId = (request.params as { docId: string; versionId: string }).versionId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const deleted = await app.db.query(
      "DELETE FROM document_versions WHERE document_id = $1 AND id = $2 RETURNING id",
      [docId, versionId]
    );
    if (deleted.rowCount !== 1) return reply.code(404).send({ error: "version_not_found" });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "document.version_deleted",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { versionId }
      })
    );

    return reply.send({ ok: true });
  });

  const LegalHoldBody = z.object({
    reason: z.string().max(1000).optional()
  });

  app.get("/docs/:docId/legal-hold", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;

    const hold = await app.db.query(
      `
        SELECT enabled, reason, created_at, created_by, released_at, released_by
        FROM document_legal_holds
        WHERE document_id = $1 AND org_id = $2
      `,
      [docId, membership.orgId]
    );

    return reply.send({ legalHold: hold.rowCount === 1 ? hold.rows[0] : null });
  });

  app.post("/docs/:docId/legal-hold", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const parsed = LegalHoldBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    await app.db.query(
      `
        INSERT INTO document_legal_holds (document_id, org_id, enabled, reason, created_by)
        VALUES ($1, $2, true, $3, $4)
        ON CONFLICT (document_id)
        DO UPDATE SET
          enabled = true,
          reason = EXCLUDED.reason,
          created_by = EXCLUDED.created_by,
          created_at = now(),
          released_by = NULL,
          released_at = NULL
      `,
      [docId, membership.orgId, parsed.data.reason ?? null, request.user!.id]
    );

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "retention.legal_hold_enabled",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { reason: parsed.data.reason ?? null }
      })
    );

    return reply.send({ ok: true });
  });

  app.delete("/docs/:docId/legal-hold", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const released = await app.db.query(
      `
        UPDATE document_legal_holds
        SET enabled = false,
            released_by = $2,
            released_at = now()
        WHERE document_id = $1
          AND org_id = $3
          AND enabled = true
      `,
      [docId, request.user!.id, membership.orgId]
    );

    if ((released.rowCount ?? 0) === 0) return reply.code(404).send({ error: "legal_hold_not_found" });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "retention.legal_hold_released",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: {}
      })
    );

    return reply.send({ ok: true });
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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "sharing.added",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { invitedEmail: email, role: invitedRole }
      })
    );

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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "document.opened",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { via: "sync-token" }
      })
    );

    return reply.send({ token, expiresAt: expiresAt.toISOString() });
  });

  const UpdateDocVersionBody = z.object({
    checkpointLocked: z.boolean().optional()
  });

  app.patch("/docs/:docId/versions/:versionId", { preHandler: requireAuth }, async (request, reply) => {
    const { docId, versionId } = request.params as { docId: string; versionId: string };
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "edit")) return reply.code(403).send({ error: "forbidden" });

    const parsed = UpdateDocVersionBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });
    if (parsed.data.checkpointLocked === undefined) return reply.send({ ok: true });

    const existing = await app.db.query(
      `
        SELECT
          data,
          data_ciphertext,
          data_kms_provider,
          data_kms_key_id
        FROM document_versions
        WHERE document_id = $1 AND id = $2
        LIMIT 1
      `,
      [docId, versionId]
    );
    if (existing.rowCount !== 1) return reply.code(404).send({ error: "version_not_found" });

    const existingRow = existing.rows[0] as any;
    const isEncrypted = Boolean(existingRow.data_ciphertext);
    const data = existingRow.data
      ? Buffer.from(existingRow.data as any)
      : await getDocumentVersionData(app.db, keyring, versionId, { documentId: docId });
    if (!data) return reply.code(404).send({ error: "version_not_found" });

    let envelope: any;
    try {
      envelope = JSON.parse(data.toString("utf8"));
    } catch {
      return reply.code(400).send({ error: "unsupported_version_format" });
    }
    if (!envelope || typeof envelope !== "object" || !envelope.meta || typeof envelope.meta !== "object") {
      return reply.code(400).send({ error: "unsupported_version_format" });
    }
    envelope.meta.checkpointLocked = parsed.data.checkpointLocked;
    const updatedData = Buffer.from(JSON.stringify(envelope), "utf8");

    if (!isEncrypted) {
      await app.db.query(`UPDATE document_versions SET data = $3 WHERE document_id = $1 AND id = $2`, [
        docId,
        versionId,
        updatedData
      ]);
      return reply.send({ ok: true });
    }

    const kmsProvider = String(existingRow.data_kms_provider ?? "local");
    const kmsKeyId = String(existingRow.data_kms_key_id);
    if (!kmsKeyId) return reply.code(500).send({ error: "kms_key_id_missing" });

    const kms = keyring.get(kmsProvider);
    const aad = {
      envelopeVersion: ENVELOPE_VERSION,
      blob: "document_versions.data",
      orgId: membership.orgId,
      documentId: docId,
      documentVersionId: versionId
    };

    const encrypted = await encryptEnvelope({
      plaintext: updatedData,
      kmsProvider: kms,
      orgId: membership.orgId,
      keyId: kmsKeyId,
      aadContext: aad
    });

    await app.db.query(
      `
        UPDATE document_versions
        SET data = NULL,
            data_envelope_version = $3,
            data_algorithm = $4,
            data_ciphertext = $5,
            data_iv = $6,
            data_tag = $7,
            data_encrypted_dek = $8,
            data_kms_provider = $9,
            data_kms_key_id = $10,
            data_aad = $11
        WHERE document_id = $1 AND id = $2
      `,
      [
        docId,
        versionId,
        encrypted.envelopeVersion,
        encrypted.algorithm,
        encrypted.ciphertext.toString("base64"),
        encrypted.iv.toString("base64"),
        encrypted.tag.toString("base64"),
        encrypted.encryptedDek.toString("base64"),
        encrypted.kmsProvider,
        encrypted.kmsKeyId,
        JSON.stringify(encrypted.aad)
      ]
    );

    return reply.send({ ok: true });
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

    if (parsed.data.visibility === "public") {
      const evaluation = await evaluateDocumentDlpPolicy(app.db, {
        orgId: membership.orgId,
        docId,
        action: DLP_ACTION.SHARE_EXTERNAL_LINK,
      });

      if (evaluation.decision === DLP_DECISION.BLOCK) {
        await writeAuditEvent(
          app.db,
          createAuditEvent({
            eventType: "dlp.blocked",
            actor: { type: "user", id: request.user!.id },
            context: {
              orgId: membership.orgId,
              userId: request.user!.id,
              userEmail: request.user!.email,
              sessionId: request.session?.id,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request),
            },
            resource: { type: "document", id: docId },
            success: false,
            error: { code: "dlp_blocked", message: evaluation.reasonCode },
            details: {
              action: evaluation.action,
              docId,
              classification: evaluation.classification,
              maxAllowed: evaluation.maxAllowed,
            },
          })
        );

        return reply.code(403).send({
          error: "dlp_blocked",
          reasonCode: evaluation.reasonCode,
          classification: evaluation.classification,
          maxAllowed: evaluation.maxAllowed,
        });
      }
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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "sharing.link_created",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: {
          visibility: parsed.data.visibility,
          role: parsed.data.role,
          expiresAt: expiresAt?.toISOString() ?? null
        }
      })
    );

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
        SELECT
          l.id,
          l.document_id,
          l.visibility,
          l.role,
          l.expires_at,
          l.revoked_at,
          d.org_id,
          os.allow_external_sharing,
          os.ip_allowlist
        FROM document_share_links l
        JOIN documents d ON d.id = l.document_id
        LEFT JOIN org_settings os ON os.org_id = d.org_id
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

    const ipAllowlist = (linkRow as any).ip_allowlist as unknown;
    const allowExternalSharing =
      (linkRow as any).allow_external_sharing == null ? true : Boolean((linkRow as any).allow_external_sharing);

    if (!(await enforceOrgIpAllowlistForSessionWithAllowlist(request, reply, orgId, ipAllowlist))) return;

    const orgMembership = await app.db.query(
      "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
      [orgId, request.user!.id]
    );

    if (visibility === "private" || !allowExternalSharing) {
      if (orgMembership.rowCount !== 1) return reply.code(403).send({ error: "forbidden" });
    } else if (orgMembership.rowCount !== 1) {
      const evaluation = await evaluateDocumentDlpPolicy(app.db, {
        orgId,
        docId,
        action: DLP_ACTION.SHARE_EXTERNAL_LINK,
      });

      if (evaluation.decision === DLP_DECISION.BLOCK) {
        await writeAuditEvent(
          app.db,
          createAuditEvent({
            eventType: "dlp.blocked",
            actor: { type: "user", id: request.user!.id },
            context: {
              orgId,
              userId: request.user!.id,
              userEmail: request.user!.email,
              sessionId: request.session?.id,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request),
            },
            resource: { type: "document", id: docId },
            success: false,
            error: { code: "dlp_blocked", message: evaluation.reasonCode },
            details: {
              action: evaluation.action,
              docId,
              classification: evaluation.classification,
              maxAllowed: evaluation.maxAllowed,
            },
          })
        );

        return reply.code(403).send({
          error: "dlp_blocked",
          reasonCode: evaluation.reasonCode,
          classification: evaluation.classification,
          maxAllowed: evaluation.maxAllowed,
        });
      }

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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "sharing.added",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { via: "share-link", shareLinkId: linkRow.id, role: nextRole }
      })
    );

    return reply.send({ ok: true, documentId: docId, role: nextRole });
  });

  const RangePermissionBody = z
    .object({
      sheetName: z.string().min(1),
      startRow: z.number().int().nonnegative(),
      startCol: z.number().int().nonnegative(),
      endRow: z.number().int().nonnegative(),
      endCol: z.number().int().nonnegative(),
      permissionType: z.enum(["read", "edit"]),
      allowedUserEmail: z.string().email()
    })
    .refine((value) => value.endRow >= value.startRow)
    .refine((value) => value.endCol >= value.startCol);

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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "sharing.modified",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: {
          type: "range-permission",
          permission: parsed.data
        }
      })
    );

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

  app.delete("/docs/:docId/range-permissions/:permissionId", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string; permissionId: string }).docId;
    const permissionId = (request.params as { docId: string; permissionId: string }).permissionId;
    const membership = await requireDocRole(request, reply, docId);
    if (!membership) return;
    if (!canDocument(membership.role, "admin")) return reply.code(403).send({ error: "forbidden" });

    const res = await app.db.query(
      "DELETE FROM document_range_permissions WHERE document_id = $1 AND id = $2 RETURNING id",
      [docId, permissionId]
    );
    if (res.rowCount !== 1) return reply.code(404).send({ error: "not_found" });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "sharing.modified",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { type: "range-permission", action: "deleted", id: permissionId }
      })
    );

    return reply.send({ ok: true });
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

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "admin.settings_changed",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId: membership.orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "document", id: docId },
        success: true,
        details: { type: "dlp-policy" }
      })
    );

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
