import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { authenticateScimToken } from "../auth/scimTokens";
import { withTransaction } from "../db/tx";
import { getClientIp, getUserAgent } from "../http/request-meta";

const SCIM_USER_SCHEMA = "urn:ietf:params:scim:schemas:core:2.0:User";
const SCIM_LIST_SCHEMA = "urn:ietf:params:scim:api:messages:2.0:ListResponse";
const SCIM_PATCH_SCHEMA = "urn:ietf:params:scim:api:messages:2.0:PatchOp";
const SCIM_ERROR_SCHEMA = "urn:ietf:params:scim:api:messages:2.0:Error";

type ScimErrorType =
  | "invalidFilter"
  | "invalidValue"
  | "invalidSyntax"
  | "invalidPath"
  | "uniqueness"
  | "tooMany"
  | "mutability"
  | "invalidToken"
  | "notFound";

function sendScimError(
  reply: FastifyReply,
  statusCode: number,
  options: { detail: string; scimType?: ScimErrorType }
): void {
  reply.header("content-type", "application/scim+json; charset=utf-8");
  reply.code(statusCode).send({
    schemas: [SCIM_ERROR_SCHEMA],
    status: String(statusCode),
    ...(options.scimType ? { scimType: options.scimType } : {}),
    detail: options.detail
  });
}

function extractBearerToken(request: FastifyRequest): string | null {
  const auth = request.headers.authorization;
  if (!auth || typeof auth !== "string") return null;
  const [kind, token] = auth.split(" ");
  if (kind?.toLowerCase() !== "bearer") return null;
  if (!token) return null;
  return token;
}

function getBaseUrl(request: FastifyRequest): string | null {
  const host = request.headers.host;
  if (!host || typeof host !== "string") return null;
  const proto = typeof request.protocol === "string" ? request.protocol : "https";
  return `${proto}://${host}`;
}

type DbUserRow = {
  id: string;
  email: string;
  name: string;
  created_at: Date;
  updated_at: Date;
};

function scimUserFor(options: { request: FastifyRequest; user: DbUserRow; active: boolean }): Record<string, unknown> {
  const baseUrl = getBaseUrl(options.request);
  const location = baseUrl ? `${baseUrl}/scim/v2/Users/${encodeURIComponent(options.user.id)}` : undefined;

  return {
    schemas: [SCIM_USER_SCHEMA],
    id: options.user.id,
    userName: options.user.email,
    displayName: options.user.name,
    active: options.active,
    emails: [{ value: options.user.email, primary: true }],
    meta: {
      resourceType: "User",
      created: new Date(options.user.created_at).toISOString(),
      lastModified: new Date(options.user.updated_at).toISOString(),
      ...(location ? { location } : {})
    }
  };
}

function parseUserNameEqFilter(filter: string): { ok: true; email: string } | { ok: false } {
  const match = /^\s*userName\s+eq\s+"([^"]+)"\s*$/i.exec(filter);
  if (!match) return { ok: false };
  const email = match[1]!.trim().toLowerCase();
  const emailCheck = z.string().email().safeParse(email);
  if (!emailCheck.success) return { ok: false };
  return { ok: true, email: emailCheck.data };
}

class FixedWindowRateLimiter {
  private readonly buckets = new Map<string, { count: number; resetAt: number }>();

  constructor(private readonly options: { windowMs: number; max: number }) {}

  take(key: string): { allowed: boolean; retryAfterSeconds: number; remaining: number } {
    const now = Date.now();
    const existing = this.buckets.get(key);
    if (!existing || now >= existing.resetAt) {
      const next = { count: 1, resetAt: now + this.options.windowMs };
      this.buckets.set(key, next);
      return {
        allowed: true,
        retryAfterSeconds: Math.ceil((next.resetAt - now) / 1000),
        remaining: Math.max(0, this.options.max - next.count)
      };
    }

    existing.count += 1;
    const allowed = existing.count <= this.options.max;
    return {
      allowed,
      retryAfterSeconds: Math.ceil((existing.resetAt - now) / 1000),
      remaining: allowed ? Math.max(0, this.options.max - existing.count) : 0
    };
  }
}

export function registerScimRoutes(app: FastifyInstance): void {
  // Basic abuse protection: per-org token + IP fixed-window limiter.
  const limiter = new FixedWindowRateLimiter({ windowMs: 60_000, max: 120 });

  const requireScim = async (request: FastifyRequest, reply: FastifyReply): Promise<void> => {
    const rawToken = extractBearerToken(request);
    if (!rawToken) {
      return sendScimError(reply, 401, { detail: "Missing bearer token", scimType: "invalidToken" });
    }

    const result = await authenticateScimToken(app.db, rawToken);
    if (!result.ok) {
      return sendScimError(reply, 401, { detail: "Invalid bearer token", scimType: "invalidToken" });
    }

    request.scim = { orgId: result.value.orgId };

    const ip = getClientIp(request) ?? "unknown";
    const rateKey = `${result.value.orgId}:${ip}`;
    const rate = limiter.take(rateKey);
    if (!rate.allowed) {
      reply.header("retry-after", String(rate.retryAfterSeconds));
      return sendScimError(reply, 429, { detail: "Rate limit exceeded", scimType: "tooMany" });
    }
  };

  const CreateUserBody = z.object({
    schemas: z.array(z.string()).optional(),
    userName: z.string().min(1),
    displayName: z.string().optional(),
    active: z.boolean().optional(),
    emails: z
      .array(
        z.object({
          value: z.string().min(1),
          primary: z.boolean().optional()
        })
      )
      .optional()
  });

  app.get("/scim/v2/Users", { preHandler: requireScim }, async (request, reply) => {
    const orgId = request.scim!.orgId;
    const query = request.query as { startIndex?: string; count?: string; filter?: string };

    const startIndexRaw = query.startIndex ? Number(query.startIndex) : 1;
    const countRaw = query.count ? Number(query.count) : 100;
    const startIndex = Number.isFinite(startIndexRaw) && startIndexRaw >= 1 ? Math.floor(startIndexRaw) : 1;
    const count = Number.isFinite(countRaw) && countRaw >= 0 ? Math.floor(countRaw) : 100;
    const offset = Math.max(0, startIndex - 1);

    let emailFilter: string | null = null;
    if (query.filter) {
      const parsed = parseUserNameEqFilter(query.filter);
      if (!parsed.ok) {
        return sendScimError(reply, 400, { detail: "Unsupported filter", scimType: "invalidFilter" });
      }
      emailFilter = parsed.email;
    }

    const totalRes = await app.db.query(
      `
        SELECT COUNT(*)::int AS total
        FROM org_members om
        JOIN users u ON u.id = om.user_id
        WHERE om.org_id = $1
          AND ($2::text IS NULL OR u.email = $2)
      `,
      [orgId, emailFilter]
    );
    const totalResults = Number(totalRes.rows[0]?.total ?? 0);

    const usersRes = await app.db.query(
      `
        SELECT u.id, u.email, u.name, u.created_at, u.updated_at
        FROM org_members om
        JOIN users u ON u.id = om.user_id
        WHERE om.org_id = $1
          AND ($2::text IS NULL OR u.email = $2)
        ORDER BY u.created_at ASC
        OFFSET $3
        LIMIT $4
      `,
      [orgId, emailFilter, offset, count]
    );

    const resources = usersRes.rows.map((row) =>
      scimUserFor({ request, user: row as DbUserRow, active: true })
    );

    reply.header("content-type", "application/scim+json; charset=utf-8");
    return reply.send({
      schemas: [SCIM_LIST_SCHEMA],
      totalResults,
      startIndex,
      itemsPerPage: resources.length,
      Resources: resources
    });
  });

  app.post("/scim/v2/Users", { preHandler: requireScim }, async (request, reply) => {
    const orgId = request.scim!.orgId;
    const parsed = CreateUserBody.safeParse(request.body);
    if (!parsed.success) {
      return sendScimError(reply, 400, { detail: "Invalid request body", scimType: "invalidSyntax" });
    }

    const body = parsed.data;
    const emailCandidate =
      body.emails?.find((e) => e.primary)?.value ?? body.emails?.[0]?.value ?? body.userName;
    const email = emailCandidate.trim().toLowerCase();
    const emailCheck = z.string().email().safeParse(email);
    if (!emailCheck.success) {
      return sendScimError(reply, 400, { detail: "userName must be a valid email address", scimType: "invalidValue" });
    }

    const displayName = body.displayName?.trim();
    const name = displayName && displayName.length > 0 ? displayName : email;
    const desiredActive = body.active !== false;

    const outcome = await withTransaction(app.db, async (client) => {
      const existingUser = await client.query(
        "SELECT id, email, name, created_at, updated_at FROM users WHERE email = $1 LIMIT 1",
        [email]
      );

      let user: DbUserRow;
      let created = false;

      if (existingUser.rowCount === 1) {
        user = existingUser.rows[0] as DbUserRow;
        if (displayName && displayName.length > 0 && displayName !== user.name) {
          const updated = await client.query(
            `
              UPDATE users
              SET name = $1, updated_at = now()
              WHERE id = $2
              RETURNING id, email, name, created_at, updated_at
            `,
            [displayName, user.id]
          );
          user = updated.rows[0] as DbUserRow;
        }
      } else {
        created = true;
        const userId = crypto.randomUUID();
        const inserted = await client.query(
          `
            INSERT INTO users (id, email, name)
            VALUES ($1, $2, $3)
            RETURNING id, email, name, created_at, updated_at
          `,
          [userId, email, name]
        );
        user = inserted.rows[0] as DbUserRow;
      }

      const membershipBeforeRes = await client.query(
        "SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2",
        [orgId, user.id]
      );
      const membershipBefore = membershipBeforeRes.rowCount === 1;

      let membershipAfter = membershipBefore;

      if (desiredActive) {
        const insertMembership = await client.query(
          `
            INSERT INTO org_members (org_id, user_id, role)
            VALUES ($1, $2, 'member')
            ON CONFLICT (org_id, user_id) DO NOTHING
            RETURNING 1
          `,
          [orgId, user.id]
        );
        membershipAfter = true;

        const membershipAdded = insertMembership.rowCount === 1;
        if (!created && membershipAdded && !membershipBefore) {
          await writeAuditEvent(
            client,
            createAuditEvent({
              eventType: "admin.user_reactivated",
              actor: { type: "scim", id: orgId },
              context: {
                orgId,
                ipAddress: getClientIp(request),
                userAgent: getUserAgent(request)
              },
              resource: { type: "user", id: user.id, name: user.email },
              success: true,
              details: { source: "scim" }
            })
          );
        }
      } else {
        const deleted = await client.query("DELETE FROM org_members WHERE org_id = $1 AND user_id = $2", [
          orgId,
          user.id
        ]);
        membershipAfter = false;
        if (deleted.rowCount === 1 && membershipBefore) {
          await writeAuditEvent(
            client,
            createAuditEvent({
              eventType: "admin.user_deactivated",
              actor: { type: "scim", id: orgId },
              context: {
                orgId,
                ipAddress: getClientIp(request),
                userAgent: getUserAgent(request)
              },
              resource: { type: "user", id: user.id, name: user.email },
              success: true,
              details: { source: "scim" }
            })
          );
        }
      }

      if (created) {
        await writeAuditEvent(
          client,
          createAuditEvent({
            eventType: "admin.user_created",
            actor: { type: "scim", id: orgId },
            context: {
              orgId,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request)
            },
            resource: { type: "user", id: user.id, name: user.email },
            success: true,
            details: { source: "scim", email: user.email }
          })
        );
      }

      return { user, active: membershipAfter, created };
    });

    reply.header("content-type", "application/scim+json; charset=utf-8");
    reply.code(outcome.created ? 201 : 200);
    return reply.send(scimUserFor({ request, user: outcome.user, active: outcome.active }));
  });

  app.get("/scim/v2/Users/:id", { preHandler: requireScim }, async (request, reply) => {
    const orgId = request.scim!.orgId;
    const userId = (request.params as { id: string }).id;

    const res = await app.db.query(
      `
        SELECT u.id, u.email, u.name, u.created_at, u.updated_at
        FROM org_members om
        JOIN users u ON u.id = om.user_id
        WHERE om.org_id = $1 AND u.id = $2
        LIMIT 1
      `,
      [orgId, userId]
    );

    if (res.rowCount !== 1) {
      return sendScimError(reply, 404, { detail: "User not found", scimType: "notFound" });
    }

    reply.header("content-type", "application/scim+json; charset=utf-8");
    return reply.send(scimUserFor({ request, user: res.rows[0] as DbUserRow, active: true }));
  });

  const PatchUserBody = z.object({
    schemas: z.array(z.string()).optional(),
    Operations: z.array(
      z.object({
        op: z.string().min(1),
        path: z.string().optional(),
        value: z.unknown().optional()
      })
    )
  });

  app.patch("/scim/v2/Users/:id", { preHandler: requireScim }, async (request, reply) => {
    const orgId = request.scim!.orgId;
    const userId = (request.params as { id: string }).id;

    const parsed = PatchUserBody.safeParse(request.body);
    if (!parsed.success) {
      return sendScimError(reply, 400, { detail: "Invalid request body", scimType: "invalidSyntax" });
    }

    let activeUpdate: boolean | undefined;
    let displayNameUpdate: string | undefined;

    for (const op of parsed.data.Operations) {
      if (op.op.toLowerCase() !== "replace") {
        return sendScimError(reply, 400, { detail: `Unsupported operation: ${op.op}`, scimType: "mutability" });
      }

      if (op.path) {
        const path = op.path.trim();
        if (path === "active") {
          if (typeof op.value === "boolean") activeUpdate = op.value;
          else return sendScimError(reply, 400, { detail: "active must be a boolean", scimType: "invalidValue" });
        } else if (path === "displayName") {
          if (typeof op.value === "string") displayNameUpdate = op.value;
          else return sendScimError(reply, 400, { detail: "displayName must be a string", scimType: "invalidValue" });
        }
        continue;
      }

      // No `path`: SCIM allows a value object containing the fields to update.
      if (op.value && typeof op.value === "object") {
        const value = op.value as Record<string, unknown>;
        if (value.active !== undefined) {
          if (typeof value.active === "boolean") activeUpdate = value.active;
          else return sendScimError(reply, 400, { detail: "active must be a boolean", scimType: "invalidValue" });
        }
        if (value.displayName !== undefined) {
          if (typeof value.displayName === "string") displayNameUpdate = value.displayName;
          else return sendScimError(reply, 400, { detail: "displayName must be a string", scimType: "invalidValue" });
        }
      }
    }

    const result = await withTransaction(app.db, async (client) => {
      const userRes = await client.query(
        "SELECT id, email, name, created_at, updated_at FROM users WHERE id = $1 LIMIT 1",
        [userId]
      );
      if (userRes.rowCount !== 1) return { ok: false as const };

      let user = userRes.rows[0] as DbUserRow;

      const membershipBeforeRes = await client.query(
        "SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2",
        [orgId, userId]
      );
      const membershipBefore = membershipBeforeRes.rowCount === 1;

      if (displayNameUpdate !== undefined) {
        const next = displayNameUpdate.trim();
        if (next.length > 0 && next !== user.name) {
          const updated = await client.query(
            `
              UPDATE users
              SET name = $1, updated_at = now()
              WHERE id = $2
              RETURNING id, email, name, created_at, updated_at
            `,
            [next, userId]
          );
          user = updated.rows[0] as DbUserRow;
        }
      }

      let membershipAfter = membershipBefore;

      if (activeUpdate !== undefined) {
        if (activeUpdate) {
          const insertMembership = await client.query(
            `
              INSERT INTO org_members (org_id, user_id, role)
              VALUES ($1, $2, 'member')
              ON CONFLICT (org_id, user_id) DO NOTHING
              RETURNING 1
            `,
            [orgId, userId]
          );
          membershipAfter = true;
          const membershipAdded = insertMembership.rowCount === 1;
          if (membershipAdded && !membershipBefore) {
            await writeAuditEvent(
              client,
              createAuditEvent({
                eventType: "admin.user_reactivated",
                actor: { type: "scim", id: orgId },
                context: {
                  orgId,
                  ipAddress: getClientIp(request),
                  userAgent: getUserAgent(request)
                },
                resource: { type: "user", id: user.id, name: user.email },
                success: true,
                details: { source: "scim" }
              })
            );
          }
        } else {
          const deleted = await client.query("DELETE FROM org_members WHERE org_id = $1 AND user_id = $2", [
            orgId,
            userId
          ]);
          membershipAfter = false;
          if (deleted.rowCount === 1 && membershipBefore) {
            await writeAuditEvent(
              client,
              createAuditEvent({
                eventType: "admin.user_deactivated",
                actor: { type: "scim", id: orgId },
                context: {
                  orgId,
                  ipAddress: getClientIp(request),
                  userAgent: getUserAgent(request)
                },
                resource: { type: "user", id: user.id, name: user.email },
                success: true,
                details: { source: "scim" }
              })
            );
          }
        }
      }

      return { ok: true as const, user, active: membershipAfter };
    });

    if (!result.ok) {
      return sendScimError(reply, 404, { detail: "User not found", scimType: "notFound" });
    }

    reply.header("content-type", "application/scim+json; charset=utf-8");
    return reply.send(scimUserFor({ request, user: result.user, active: result.active }));
  });

  // Minimal discovery endpoint some clients call.
  app.get("/scim/v2/ServiceProviderConfig", { preHandler: requireScim }, async (_request, reply) => {
    reply.header("content-type", "application/scim+json; charset=utf-8");
    return reply.send({
      schemas: ["urn:ietf:params:scim:schemas:core:2.0:ServiceProviderConfig"],
      patch: { supported: true },
      filter: { supported: true, maxResults: 100 },
      bulk: { supported: false },
      changePassword: { supported: false },
      sort: { supported: false },
      etag: { supported: false },
      authenticationSchemes: [
        {
          type: "oauthbearertoken",
          name: "Bearer Token",
          description: "Bearer token authentication using org SCIM token."
        }
      ]
    });
  });
}

