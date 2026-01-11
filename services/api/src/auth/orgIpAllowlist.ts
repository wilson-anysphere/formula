import type { FastifyReply, FastifyRequest } from "fastify";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isClientIpAllowed } from "./apiKeys";

function isValidOrgId(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

export async function enforceOrgIpAllowlistForSessionWithAllowlist(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string,
  ipAllowlist: unknown
): Promise<boolean> {
  // API keys already enforce org_settings.ip_allowlist inside authenticateApiKey().
  if (request.authMethod !== "session") return true;

  const clientIp = getClientIp(request);
  if (isClientIpAllowed(clientIp, ipAllowlist)) return true;

  try {
    await writeAuditEvent(
      request.server.db,
      createAuditEvent({
        eventType: "org.ip_allowlist.blocked",
        actor: { type: "user", id: request.user?.id ?? "unknown" },
        context: {
          orgId,
          userId: request.user?.id ?? null,
          userEmail: request.user?.email ?? null,
          sessionId: request.session?.id ?? null,
          ipAddress: clientIp,
          userAgent: getUserAgent(request)
        },
        resource: { type: "organization", id: orgId },
        success: false,
        error: { code: "ip_not_allowed" },
        details: {
          method: request.method,
          path: request.url
        }
      })
    );
  } catch (err) {
    // Best-effort: allowlist enforcement must not fail open because audit plumbing failed.
    request.log?.warn?.({ err, orgId }, "org_ip_allowlist_audit_failed");
  }

  reply.code(403).send({ error: "ip_not_allowed" });
  return false;
}

export async function enforceOrgIpAllowlistForSession(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<boolean> {
  if (request.authMethod !== "session") return true;

  const orgSettings = await request.server.db.query("SELECT ip_allowlist FROM org_settings WHERE org_id = $1", [
    orgId
  ]);
  if (orgSettings.rowCount !== 1) return true;

  const ipAllowlist = (orgSettings.rows[0] as any).ip_allowlist as unknown;
  return enforceOrgIpAllowlistForSessionWithAllowlist(request, reply, orgId, ipAllowlist);
}

export async function enforceOrgIpAllowlistFromParams(
  request: FastifyRequest,
  reply: FastifyReply
): Promise<void | FastifyReply> {
  const orgId = (request.params as { orgId?: unknown } | undefined)?.orgId;
  if (typeof orgId !== "string" || orgId.length === 0) return;
  if (!isValidOrgId(orgId)) {
    return reply.code(400).send({ error: "invalid_request" });
  }
  const ok = await enforceOrgIpAllowlistForSession(request, reply, orgId);
  if (!ok) return reply;
}
