import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import type { MaybeEncryptedSecret, SiemEndpointConfig } from "../siem/types";
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

function maskSecret(_value: MaybeEncryptedSecret): "***" {
  return "***";
}

function sanitizeConfig(config: SiemEndpointConfig): SiemEndpointConfig {
  const sanitized: SiemEndpointConfig = {
    ...config
  };

  const auth = sanitized.auth;
  if (!auth) return sanitized;

  if (auth.type === "bearer") {
    sanitized.auth = { ...auth, token: maskSecret(auth.token) };
    return sanitized;
  }

  if (auth.type === "basic") {
    sanitized.auth = { ...auth, username: maskSecret(auth.username), password: maskSecret(auth.password) };
    return sanitized;
  }

  if (auth.type === "header") {
    sanitized.auth = { ...auth, value: maskSecret(auth.value) };
    return sanitized;
  }

  return sanitized;
}

function parseConfig(raw: unknown): SiemEndpointConfig | null {
  if (!raw) return null;
  if (typeof raw === "string") {
    try {
      return parseConfig(JSON.parse(raw));
    } catch {
      return null;
    }
  }
  if (typeof raw !== "object") return null;
  return raw as SiemEndpointConfig;
}

export function registerSiemRoutes(app: FastifyInstance): void {
  const Secret = z.union([
    z.string(),
    z.object({ encrypted: z.string() }),
    z.object({ ciphertext: z.string() })
  ]);

  const Auth = z.discriminatedUnion("type", [
    z.object({ type: z.literal("none") }),
    z.object({ type: z.literal("bearer"), token: Secret }),
    z.object({ type: z.literal("basic"), username: Secret, password: Secret }),
    z.object({ type: z.literal("header"), name: z.string().min(1), value: Secret })
  ]);

  const Retry = z
    .object({
      maxAttempts: z.number().int().positive().optional(),
      baseDelayMs: z.number().int().positive().optional(),
      maxDelayMs: z.number().int().positive().optional(),
      jitter: z.boolean().optional()
    })
    .optional();

  const RedactionOptions = z
    .object({
      redactionText: z.string().min(1).optional()
    })
    .optional();

  const SiemConfigBody = z.object({
    endpointUrl: z.string().url(),
    format: z.enum(["json", "cef", "leef"]).optional(),
    timeoutMs: z.number().int().positive().optional(),
    idempotencyKeyHeader: z.string().min(1).nullable().optional(),
    headers: z.record(z.string()).optional(),
    auth: Auth.optional(),
    retry: Retry,
    redactionOptions: RedactionOptions,
    batchSize: z.number().int().positive().optional()
  });

  app.get("/orgs/:orgId/siem", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const role = await requireOrgAdminRole(request, reply, orgId);
    if (!role) return;

    const res = await app.db.query("SELECT config FROM org_siem_configs WHERE org_id = $1 AND enabled = true", [
      orgId
    ]);
    if (res.rowCount !== 1) return reply.code(404).send({ error: "siem_config_not_found" });

    const config = parseConfig(res.rows[0]!.config);
    if (!config) return reply.code(500).send({ error: "siem_config_invalid" });
    return reply.send({ config: sanitizeConfig(config) });
  });

  app.put("/orgs/:orgId/siem", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const role = await requireOrgAdminRole(request, reply, orgId);
    if (!role) return;

    const parsed = SiemConfigBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const config = parsed.data;

    await app.db.query(
      `
        INSERT INTO org_siem_configs (org_id, enabled, config)
        VALUES ($1, true, $2::jsonb)
        ON CONFLICT (org_id) DO UPDATE
        SET enabled = true, config = EXCLUDED.config, updated_at = now()
      `,
      [orgId, JSON.stringify(config)]
    );

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.siem_config.updated",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "organization", id: orgId },
        success: true,
        details: { enabled: true }
      })
    );

    return reply.send({ config: sanitizeConfig(config) });
  });

  app.delete("/orgs/:orgId/siem", { preHandler: requireAuth }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const role = await requireOrgAdminRole(request, reply, orgId);
    if (!role) return;

    await app.db.query("DELETE FROM org_siem_configs WHERE org_id = $1", [orgId]);

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "org.siem_config.deleted",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "organization", id: orgId },
        success: true,
        details: {}
      })
    );

    reply.code(204).send();
  });
}
