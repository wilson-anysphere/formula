import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { enforceOrgIpAllowlistForSessionWithAllowlist } from "../auth/orgIpAllowlist";
import { DLP_ACTION } from "../dlp/dlp";
import { evaluateDocumentDlpPolicy } from "../dlp/effective";
import { canDocument, type DocumentRole } from "../rbac/roles";
import { requireAuth } from "./auth";

function isValidUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

async function requireDocRead(
  request: FastifyRequest,
  reply: FastifyReply,
  docId: string
): Promise<{ orgId: string; role: DocumentRole } | null> {
  if (!isValidUuid(docId)) {
    reply.code(404).send({ error: "doc_not_found" });
    return null;
  }
  const membership = await request.server.db.query(
    `
      SELECT d.org_id, dm.role, os.ip_allowlist
      FROM documents d
      LEFT JOIN org_settings os ON os.org_id = d.org_id
      LEFT JOIN document_members dm
        ON dm.document_id = d.id AND dm.user_id = $2
      WHERE d.id = $1
      LIMIT 1
    `,
    [docId, request.user!.id]
  );

  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "doc_not_found" });
    return null;
  }

  const row = membership.rows[0] as { org_id: string; role: DocumentRole | null; ip_allowlist: unknown };
  if (request.authOrgId && request.authOrgId !== row.org_id) {
    reply.code(404).send({ error: "doc_not_found" });
    return null;
  }
  if (!(await enforceOrgIpAllowlistForSessionWithAllowlist(request, reply, row.org_id, row.ip_allowlist))) {
    return null;
  }

  if (!row.role || !canDocument(row.role, "read")) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }

  return { orgId: row.org_id, role: row.role };
}

const EvaluateDlpBody = z.object({
  action: z.enum(Object.values(DLP_ACTION) as [string, ...string[]]),
  options: z
    .object({
      includeRestrictedContent: z.boolean().optional(),
    })
    .strict()
    .optional(),
  selector: z.unknown().optional(),
});

export function registerDlpRoutes(app: FastifyInstance): void {
  app.post("/docs/:docId/dlp/evaluate", { preHandler: requireAuth }, async (request, reply) => {
    const docId = (request.params as { docId: string }).docId;
    const membership = await requireDocRead(request, reply, docId);
    if (!membership) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, membership.orgId, request.session))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const parsed = EvaluateDlpBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    let selectorValue: unknown | undefined;
    if (parsed.data.selector !== undefined) {
      try {
        const selector = parsed.data.selector;
        if (typeof selector !== "object" || selector === null) throw new Error("Selector must be an object");
        if ((selector as any).documentId !== docId) throw new Error("Selector documentId must match route docId");
        selectorValue = selector;
      } catch (error) {
        const message = error instanceof Error ? error.message : "Invalid selector";
        return reply.code(400).send({ error: "invalid_request", message });
      }
    }

    let evaluation;
    try {
      evaluation = await evaluateDocumentDlpPolicy(app.db, {
        orgId: membership.orgId,
        docId,
        action: parsed.data.action,
        options: parsed.data.options,
        selector: selectorValue,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Invalid selector";
      return reply.code(400).send({ error: "invalid_request", message });
    }

    return reply.send({
      decision: evaluation.decision,
      reasonCode: evaluation.reasonCode,
      classification: evaluation.classification,
      maxAllowed: evaluation.maxAllowed,
    });
  });
}
