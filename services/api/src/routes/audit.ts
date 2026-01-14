import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { PassThrough } from "node:stream";
import { z } from "zod";
import { auditLogRowToAuditEvent, redactAuditEvent, serializeBatch, type PostgresAuditLogRow } from "@formula/audit-core";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { TokenBucketRateLimiter } from "../http/rateLimit";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

const ZERO_UUID = "00000000-0000-0000-0000-000000000000";
const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function isUuid(value: string): boolean {
  return UUID_REGEX.test(value);
}

type AuditStreamCursor = {
  lastCreatedAt: Date;
  lastEventId: string;
};

function encodeStreamCursor(cursor: AuditStreamCursor): string {
  const payload = JSON.stringify({ createdAt: cursor.lastCreatedAt.toISOString(), id: cursor.lastEventId });
  return Buffer.from(payload, "utf8").toString("base64url");
}

function decodeStreamCursor(raw: string): AuditStreamCursor | null {
  if (typeof raw !== "string" || raw.trim().length === 0) return null;

  try {
    const normalized = raw.trim();
    const padded = normalized
      .replace(/-/g, "+")
      .replace(/_/g, "/")
      .padEnd(Math.ceil(normalized.length / 4) * 4, "=");
    const decoded = Buffer.from(padded, "base64").toString("utf8");
    const parsed = JSON.parse(decoded) as unknown;
    if (!parsed || typeof parsed !== "object") return null;

    const createdAtRaw = (parsed as any).createdAt;
    const idRaw = (parsed as any).id;
    if (typeof createdAtRaw !== "string" || typeof idRaw !== "string") return null;

    const createdAt = new Date(createdAtRaw);
    if (Number.isNaN(createdAt.getTime())) return null;
    // Loose UUID validation; guards against SQL injection via `::uuid` casts.
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(idRaw)) return null;

    return { lastCreatedAt: createdAt, lastEventId: idRaw };
  } catch {
    return null;
  }
}

function toDate(value: unknown): Date {
  if (value instanceof Date) return value;
  const date = new Date(String(value));
  if (Number.isNaN(date.getTime())) throw new Error(`Invalid created_at: ${String(value)}`);
  return date;
}

async function requireOrgRole(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<OrgRole | null> {
  if (request.authOrgId && request.authOrgId !== orgId) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const membership = await request.server.db.query(
    "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  return membership.rows[0].role as OrgRole;
}

async function requireOrgAdminRole(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<OrgRole | null> {
  const role = await requireOrgRole(request, reply, orgId);
  if (!role) return null;
  if (!isOrgAdmin(role)) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return role;
}

export function registerAuditRoutes(app: FastifyInstance): void {
  const ingestRateLimiter = new TokenBucketRateLimiter({ capacity: 300, refillMs: 60_000 });

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

  const AuditIngestBody = z.object({
    eventType: z.string().min(1),
    resource: z
      .object({
        type: z.string().min(1),
        id: z.string().nullable().optional(),
        name: z.string().nullable().optional()
      })
      .optional(),
    success: z.boolean().optional(),
    error: z
      .object({
        code: z.string().nullable().optional(),
        message: z.string().nullable().optional()
      })
      .optional(),
    details: z.record(z.unknown()).optional(),
    correlation: z
      .object({
        requestId: z.string().nullable().optional(),
        traceId: z.string().nullable().optional()
      })
      .optional()
  });

  const tooManyRequests = (reply: FastifyReply, retryAfterMs: number) => {
    return reply
      .header("Retry-After", String(Math.max(1, Math.ceil(retryAfterMs / 1000))))
      .code(429)
      .send({ error: "too_many_requests" });
  };

  app.post(
    "/orgs/:orgId/audit",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const role = await requireOrgAdminRole(request, reply, orgId);
      if (!role) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const ip = getClientIp(request) ?? "unknown";
      const limited = ingestRateLimiter.take(`${orgId}:${ip}`);
      if (!limited.ok) {
        app.metrics.rateLimitedTotal.inc({ route: "/orgs/:orgId/audit", reason: "org_ip" });
        return tooManyRequests(reply, limited.retryAfterMs);
      }

      const parsed = AuditIngestBody.safeParse(request.body);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      const actor = request.apiKey
        ? { type: "api_key", id: request.apiKey.id }
        : { type: "user", id: request.user!.id };

      const event = createAuditEvent({
        eventType: parsed.data.eventType,
        actor,
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: parsed.data.resource,
        success: parsed.data.success ?? true,
        error: parsed.data.error,
        details: parsed.data.details,
        correlation: parsed.data.correlation
      });

      await writeAuditEvent(app.db, event);
      app.auditStreamHub.publish(event);
      return reply.code(202).send({ id: event.id });
    }
  );

  const AuditStreamQuery = z.object({
    after: z.string().optional()
  });

  app.get(
    "/orgs/:orgId/audit/stream",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const role = await requireOrgAdminRole(request, reply, orgId);
      if (!role) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

      const query = AuditStreamQuery.safeParse(request.query);
      if (!query.success) return reply.code(400).send({ error: "invalid_request" });

      const afterParam = query.data.after;
      const lastEventIdHeader = request.headers["last-event-id"];
      const afterHeader = typeof lastEventIdHeader === "string" ? lastEventIdHeader : undefined;

      const cursorInput = afterParam ?? afterHeader;
      let decodedCursor: AuditStreamCursor | null = null;
      if (cursorInput) {
        decodedCursor = decodeStreamCursor(cursorInput);
        if (!decodedCursor) {
          const trimmed = cursorInput.trim();
          if (!isUuid(trimmed)) return reply.code(400).send({ error: "invalid_request" });

          const createdAtResult = await app.db.query<{ created_at: unknown }>(
            `
              SELECT created_at
              FROM (
                SELECT created_at
                FROM audit_log
                WHERE org_id = $1 AND id = $2
                UNION ALL
                SELECT created_at
                FROM audit_log_archive
                WHERE org_id = $1 AND id = $2
              ) AS audit_events
              LIMIT 1
            `,
            [orgId, trimmed]
          );
          if (createdAtResult.rowCount !== 1) return reply.code(400).send({ error: "invalid_request" });
          decodedCursor = { lastCreatedAt: toDate(createdAtResult.rows[0]!.created_at), lastEventId: trimmed };
        }
      }

      let cursor: AuditStreamCursor =
        decodedCursor ??
        ({
          lastCreatedAt: new Date(),
          lastEventId: ZERO_UUID
        } satisfies AuditStreamCursor);

      const stream = new PassThrough();

      reply.header("content-type", "text/event-stream");
      reply.header("cache-control", "no-cache");
      reply.header("connection", "keep-alive");
      reply.header("x-accel-buffering", "no");

      // Emit an initial comment to ensure clients treat the connection as established.
      stream.write(":ok\n\n");

      app.metrics.auditStreamClients.inc();

      let keepaliveTimer: NodeJS.Timeout | null = null;
      let drainInFlight = false;
      let drainQueued = false;
      let closed = false;
      let backpressured = false;
      let unsubscribe = () => {};

      const close = () => {
        if (closed) return;
        closed = true;
        if (keepaliveTimer) clearInterval(keepaliveTimer);
        keepaliveTimer = null;
        unsubscribe();
        app.metrics.auditStreamClients.dec();
        stream.end();
      };

      request.raw.on("close", close);
      request.raw.on("aborted", close);

      stream.on("drain", () => {
        backpressured = false;
        if (drainQueued) {
          drainQueued = false;
          void drain().catch(() => {
            // Best-effort: avoid unhandled rejections from fire-and-forget drain.
          });
        }
      });

      const drain = async () => {
        if (closed) return;
        if (backpressured) {
          drainQueued = true;
          return;
        }
        if (drainInFlight) {
          drainQueued = true;
          return;
        }

        drainInFlight = true;
        try {
          while (!closed && !backpressured) {
            const columns =
              "id, org_id, user_id, user_email, event_type, resource_type, resource_id, ip_address, user_agent, session_id, success, error_code, error_message, details, created_at";
            const values: unknown[] = [orgId, cursor.lastCreatedAt, cursor.lastEventId, 100];

            const result = await app.db.query(
              `
                SELECT ${columns}
                FROM (
                  SELECT ${columns}
                  FROM audit_log
                  WHERE org_id = $1
                  UNION ALL
                  SELECT ${columns}
                  FROM audit_log_archive
                  WHERE org_id = $1
                ) AS audit_events
                WHERE (
                  audit_events.created_at > $2::timestamptz
                  OR (
                    audit_events.created_at = $2::timestamptz
                    AND audit_events.id > $3::uuid
                  )
                )
                ORDER BY audit_events.created_at ASC, audit_events.id ASC
                LIMIT $4
              `,
              values
            );

            if (result.rows.length === 0) break;

            for (const row of result.rows) {
              if (closed) break;

              const createdAt = toDate((row as PostgresAuditLogRow).created_at);
              const id = String((row as PostgresAuditLogRow).id);
              const nextCursor = { lastCreatedAt: createdAt, lastEventId: id };

              const event = redactAuditEvent(auditLogRowToAuditEvent(row as PostgresAuditLogRow));
              const eventCursor = encodeStreamCursor(nextCursor);
              const payload = `id: ${eventCursor}\nevent: audit\ndata: ${JSON.stringify(event)}\n\n`;

              const ok = stream.write(payload);
              app.metrics.auditStreamEventsTotal.inc();
              cursor = nextCursor;

              if (!ok) {
                backpressured = true;
                drainQueued = true;
                app.metrics.auditStreamBackpressureDropsTotal.inc();
                break;
              }
            }
          }
        } catch (err) {
          stream.write("event: error\n");
          stream.write(`data: ${JSON.stringify({ error: "stream_error" })}\n\n`);
          close();
        } finally {
          drainInFlight = false;
          if (drainQueued && !closed && !backpressured) {
            drainQueued = false;
            void drain().catch(() => {
              // Best-effort: avoid unhandled rejections from fire-and-forget drain.
            });
          }
        }
      };

      unsubscribe = app.auditStreamHub.subscribe(orgId, () => {
        void drain().catch(() => {
          // Best-effort: avoid unhandled rejections from fire-and-forget drain.
        });
      });

      keepaliveTimer = setInterval(() => {
        if (closed) return;
        stream.write(":keep-alive\n\n");
        void drain().catch(() => {
          // Best-effort: avoid unhandled rejections from fire-and-forget drain.
        });
      }, 15_000);
      keepaliveTimer.unref?.();

      // Kick off an initial drain so resume cursors replay immediately.
      void drain().catch(() => {
        // Best-effort: avoid unhandled rejections from fire-and-forget drain.
      });

      return reply.send(stream);
    }
  );

  app.get(
    "/orgs/:orgId/audit",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const role = await requireOrgAdminRole(request, reply, orgId);
      if (!role) return;
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

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

      const events = result.rows.map((row) => redactAuditEvent(auditLogRowToAuditEvent(row as PostgresAuditLogRow)));
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
      if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.session))) {
        return reply.code(403).send({ error: "mfa_required" });
      }

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

      const events = result.rows.map((row) => auditLogRowToAuditEvent(row as PostgresAuditLogRow));
      const format = parsed.data.format ?? "json";
      const { contentType, body } = serializeBatch(events, { format });

      reply.header("content-type", contentType);
      reply.send(body);
    }
  );
}
