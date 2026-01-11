import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { parseScimToken, verifyScimTokenSecret, type ScimTokenInfo } from "../auth/scim";
import { withTransaction } from "../db/tx";
import { getClientIp, getUserAgent } from "../http/request-meta";

const SCIM_CONTENT_TYPE = "application/scim+json";

const SCIM_SCHEMA_USER = "urn:ietf:params:scim:schemas:core:2.0:User";
const SCIM_SCHEMA_LIST_RESPONSE = "urn:ietf:params:scim:api:messages:2.0:ListResponse";
const SCIM_SCHEMA_ERROR = "urn:ietf:params:scim:api:messages:2.0:Error";
const SCIM_SCHEMA_SERVICE_PROVIDER_CONFIG = "urn:ietf:params:scim:schemas:core:2.0:ServiceProviderConfig";

function isValidUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value);
}

function sendScimError(reply: FastifyReply, statusCode: number, detail: string): void {
  reply
    .code(statusCode)
    .type(SCIM_CONTENT_TYPE)
    .send({ schemas: [SCIM_SCHEMA_ERROR], status: String(statusCode), detail });
}

function extractBearerToken(request: FastifyRequest): string | null {
  const auth = request.headers.authorization;
  if (!auth || typeof auth !== "string") return null;
  const [kind, token] = auth.split(" ");
  if (kind?.toLowerCase() !== "bearer") return null;
  if (!token) return null;
  return token.trim();
}

export async function requireScimAuth(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const raw = extractBearerToken(request);
  if (!raw) return sendScimError(reply, 401, "Unauthorized");

  const parsed = parseScimToken(raw);
  if (!parsed) return sendScimError(reply, 401, "Unauthorized");

  const res = await request.server.db.query(
    `
      SELECT id, org_id, name, token_hash, created_by, created_at, last_used_at, revoked_at
      FROM org_scim_tokens
      WHERE id = $1
      LIMIT 1
    `,
    [parsed.id]
  );

  if (res.rowCount !== 1) return sendScimError(reply, 401, "Unauthorized");
  const row = res.rows[0] as {
    id: string;
    org_id: string;
    name: string;
    token_hash: string;
    created_by: string;
    created_at: Date;
    last_used_at: Date | null;
    revoked_at: Date | null;
  };

  if (row.revoked_at) return sendScimError(reply, 401, "Unauthorized");
  if (!verifyScimTokenSecret(parsed.secret, row.token_hash)) return sendScimError(reply, 401, "Unauthorized");

  await request.server.db.query("UPDATE org_scim_tokens SET last_used_at = now() WHERE id = $1", [row.id]);

  const tokenInfo: ScimTokenInfo = {
    id: row.id,
    orgId: row.org_id,
    name: row.name,
    createdBy: row.created_by,
    createdAt: new Date(row.created_at),
    lastUsedAt: row.last_used_at ? new Date(row.last_used_at) : null,
    revokedAt: row.revoked_at ? new Date(row.revoked_at) : null
  };

  request.scimToken = tokenInfo;
  request.user = undefined;
  request.session = undefined;
  request.apiKey = undefined;
  request.authMethod = "scim";
  request.authOrgId = row.org_id;
}

type DbUserRow = { id: string; email: string; name: string };

function splitName(fullName: string): { givenName: string | null; familyName: string | null } {
  const trimmed = fullName.trim();
  if (!trimmed) return { givenName: null, familyName: null };
  const firstSpace = trimmed.indexOf(" ");
  if (firstSpace === -1) return { givenName: trimmed, familyName: null };
  const givenName = trimmed.slice(0, firstSpace).trim();
  const familyName = trimmed.slice(firstSpace + 1).trim();
  return { givenName: givenName.length > 0 ? givenName : null, familyName: familyName.length > 0 ? familyName : null };
}

function toScimUser(user: DbUserRow, active: boolean): Record<string, unknown> {
  const parts = splitName(user.name);
  return {
    schemas: [SCIM_SCHEMA_USER],
    id: user.id,
    userName: user.email,
    displayName: user.name,
    name: {
      formatted: user.name,
      givenName: parts.givenName,
      familyName: parts.familyName
    },
    emails: [{ value: user.email, primary: true }],
    active,
    meta: { resourceType: "User" }
  };
}

function parseUserNameEqFilter(filterRaw: unknown): string | null {
  if (typeof filterRaw !== "string") return null;
  const filter = filterRaw.trim();
  if (filter.length === 0) return null;
  // Minimal filter support: userName eq "..."
  const match = filter.match(/^\s*userName\s+eq\s+("([^"]+)"|'([^']+)')\s*$/i);
  if (!match) return null;
  const value = (match[2] ?? match[3] ?? "").trim();
  return value.length > 0 ? value : null;
}

type PatchChanges = {
  active?: boolean;
  userName?: string;
  displayName?: string;
  nameFormatted?: string;
  nameGiven?: string;
  nameFamily?: string;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function parsePatchChanges(body: unknown): PatchChanges | null {
  if (!isRecord(body)) return null;
  const operations = body.Operations;
  if (!Array.isArray(operations)) return null;

  const changes: PatchChanges = {};

  for (const opRaw of operations) {
    if (!isRecord(opRaw)) return null;
    const op = typeof opRaw.op === "string" ? opRaw.op.toLowerCase() : null;
    if (!op || op !== "replace") return null;

    const path = typeof opRaw.path === "string" ? opRaw.path : null;
    const value = opRaw.value;

    const applyReplace = (targetPath: string | null, targetValue: unknown) => {
      const normalizedPath = targetPath?.trim();

      if (!normalizedPath) {
        // SCIM PATCH allows omitting `path` to replace multiple attributes at once.
        if (!isRecord(targetValue)) return;
        if (typeof targetValue.active === "boolean") changes.active = targetValue.active;
        if (typeof targetValue.userName === "string") changes.userName = targetValue.userName;
        if (typeof targetValue.displayName === "string") changes.displayName = targetValue.displayName;
        if (isRecord(targetValue.name)) {
          const n = targetValue.name;
          if (typeof n.formatted === "string") changes.nameFormatted = n.formatted;
          if (typeof n.givenName === "string") changes.nameGiven = n.givenName;
          if (typeof n.familyName === "string") changes.nameFamily = n.familyName;
        }
        return;
      }

      if (normalizedPath.toLowerCase() === "active" && typeof targetValue === "boolean") {
        changes.active = targetValue;
        return;
      }

      if (normalizedPath === "userName" && typeof targetValue === "string") {
        changes.userName = targetValue;
        return;
      }

      if (normalizedPath === "displayName" && typeof targetValue === "string") {
        changes.displayName = targetValue;
        return;
      }

      if (normalizedPath === "name" && isRecord(targetValue)) {
        if (typeof targetValue.formatted === "string") changes.nameFormatted = targetValue.formatted;
        if (typeof targetValue.givenName === "string") changes.nameGiven = targetValue.givenName;
        if (typeof targetValue.familyName === "string") changes.nameFamily = targetValue.familyName;
        return;
      }

      if (normalizedPath.startsWith("name.") && typeof targetValue === "string") {
        const sub = normalizedPath.slice("name.".length);
        if (sub === "formatted") changes.nameFormatted = targetValue;
        if (sub === "givenName") changes.nameGiven = targetValue;
        if (sub === "familyName") changes.nameFamily = targetValue;
      }
    };

    applyReplace(path, value);
  }

  return changes;
}

function buildDisplayName(
  email: string,
  input: {
    displayName?: string | null;
    formatted?: string | null;
    given?: string | null;
    family?: string | null;
  }
): string {
  const displayName = input.displayName?.trim();
  if (displayName) return displayName;
  const formatted = input.formatted?.trim();
  if (formatted) return formatted;
  const given = input.given?.trim();
  const family = input.family?.trim();
  const combined = [given, family].filter((v): v is string => Boolean(v && v.length > 0)).join(" ").trim();
  return combined.length > 0 ? combined : email;
}

export function registerScimRoutes(app: FastifyInstance): void {
  const CreateUserBody = z
    .object({
      userName: z.string().min(1).email(),
      displayName: z.string().min(1).optional(),
      name: z
        .object({
          formatted: z.string().min(1).optional(),
          givenName: z.string().min(1).optional(),
          familyName: z.string().min(1).optional()
        })
        .optional(),
      active: z.boolean().optional(),
      emails: z
        .array(
          z.object({
            value: z.string().min(1).email(),
            primary: z.boolean().optional()
          })
        )
        .optional()
    })
    .passthrough();

  app.register(
    async (scimApp) => {
      scimApp.addHook("onRequest", async (_request, reply) => {
        reply.type(SCIM_CONTENT_TYPE);
      });

      scimApp.addHook("preHandler", requireScimAuth);

      scimApp.get("/ServiceProviderConfig", async (_request) => {
        return {
          schemas: [SCIM_SCHEMA_SERVICE_PROVIDER_CONFIG],
          patch: { supported: true },
          bulk: { supported: false },
          filter: { supported: true, maxResults: 200 },
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
        };
      });

      scimApp.get("/Users", async (request, reply) => {
        const orgId = request.authOrgId!;

        const startIndexRaw = (request.query as any)?.startIndex;
        const countRaw = (request.query as any)?.count;
        const filterRaw = (request.query as any)?.filter;

        const startIndex = Math.max(1, Number.parseInt(String(startIndexRaw ?? "1"), 10) || 1);
        const count = Math.min(200, Math.max(0, Number.parseInt(String(countRaw ?? "100"), 10) || 100));

        const filterUserName = parseUserNameEqFilter(filterRaw);

        const values: unknown[] = [orgId];
        let where = "WHERE om.org_id = $1";
        if (filterUserName) {
          values.push(filterUserName.trim().toLowerCase());
          where += ` AND lower(u.email) = $${values.length}`;
        } else if (typeof filterRaw === "string" && filterRaw.trim().length > 0) {
          return sendScimError(reply, 400, "Unsupported filter");
        }

        const totalRes = await scimApp.db.query(
          `
            SELECT COUNT(*)::int AS total
            FROM users u
            JOIN org_members om ON om.user_id = u.id
            ${where}
          `,
          values
        );
        const totalResults = (totalRes.rows[0]?.total ?? 0) as number;

        values.push(count);
        values.push(startIndex - 1);

        const listRes = await scimApp.db.query(
          `
            SELECT u.id, u.email, u.name
            FROM users u
            JOIN org_members om ON om.user_id = u.id
            ${where}
            ORDER BY u.created_at ASC
            LIMIT $${values.length - 1} OFFSET $${values.length}
          `,
          values
        );

        const users = listRes.rows.map((row) => toScimUser(row as DbUserRow, true));

        return reply.send({
          schemas: [SCIM_SCHEMA_LIST_RESPONSE],
          totalResults,
          startIndex,
          itemsPerPage: users.length,
          Resources: users
        });
      });

      scimApp.get("/Users/:id", async (request, reply) => {
        const orgId = request.authOrgId!;
        const id = (request.params as { id: string }).id;
        if (!isValidUuid(id)) return sendScimError(reply, 404, "User not found");

        const res = await scimApp.db.query(
          `
            SELECT u.id, u.email, u.name
            FROM users u
            JOIN org_members om ON om.user_id = u.id
            WHERE om.org_id = $1 AND u.id = $2
            LIMIT 1
          `,
          [orgId, id]
        );

        if (res.rowCount !== 1) return sendScimError(reply, 404, "User not found");
        return reply.send(toScimUser(res.rows[0] as DbUserRow, true));
      });

      scimApp.post("/Users", async (request, reply) => {
        const orgId = request.authOrgId!;

        const parsed = CreateUserBody.safeParse(request.body);
        if (!parsed.success) return sendScimError(reply, 400, "Invalid request");

        const email = parsed.data.userName.trim().toLowerCase();
        const displayName = buildDisplayName(email, {
          displayName: parsed.data.displayName ?? null,
          formatted: parsed.data.name?.formatted ?? null,
          given: parsed.data.name?.givenName ?? null,
          family: parsed.data.name?.familyName ?? null
        });
        const desiredActive = parsed.data.active !== false;

        const result = await withTransaction(scimApp.db, async (client) => {
          const existing = await client.query("SELECT id FROM users WHERE email = $1 LIMIT 1", [email]);
          const userId = (existing.rows[0]?.id ?? crypto.randomUUID()) as string;
          const created = existing.rowCount !== 1;

          if (created) {
            await client.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [userId, email, displayName]);
          } else {
            await client.query("UPDATE users SET name = $2, updated_at = now() WHERE id = $1", [userId, displayName]);
          }

          let membershipInserted = false;
          let membershipDeleted = false;

          if (desiredActive) {
            const memberRes = await client.query(
              `
                INSERT INTO org_members (org_id, user_id, role)
                VALUES ($1, $2, 'member')
                ON CONFLICT (org_id, user_id) DO NOTHING
              `,
              [orgId, userId]
            );
            membershipInserted = memberRes.rowCount === 1;
          } else {
            const delRes = await client.query("DELETE FROM org_members WHERE org_id = $1 AND user_id = $2", [orgId, userId]);
            membershipDeleted = delRes.rowCount === 1;
          }

          return { userId, created, membershipInserted, membershipDeleted };
        });

        const userRow = await scimApp.db.query("SELECT id, email, name FROM users WHERE id = $1", [result.userId]);
        const user = userRow.rows[0] as DbUserRow;
        const memberAfter = await scimApp.db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
          orgId,
          user.id
        ]);
        const active = memberAfter.rowCount === 1;

        await writeAuditEvent(
          scimApp.db,
          createAuditEvent({
            eventType: "scim.user.created",
            actor: { type: "scim_token", id: request.scimToken!.id },
            context: {
              orgId,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request)
            },
            resource: { type: "user", id: user.id, name: user.email },
            success: true,
            details: {
              source: "scim",
              created: result.created,
              membershipInserted: result.membershipInserted,
              membershipDeleted: result.membershipDeleted
            }
          })
        );

        reply.code(result.created ? 201 : 200);
        return reply.send(toScimUser(user, active));
      });

      scimApp.patch("/Users/:id", async (request, reply) => {
        const orgId = request.authOrgId!;
        const id = (request.params as { id: string }).id;

        const changes = parsePatchChanges(request.body);
        if (!changes) return sendScimError(reply, 400, "Invalid request");
        if (!isValidUuid(id)) return sendScimError(reply, 404, "User not found");

        const resUser = await scimApp.db.query("SELECT id, email, name FROM users WHERE id = $1 LIMIT 1", [id]);
        if (resUser.rowCount !== 1) return sendScimError(reply, 404, "User not found");
        const user = resUser.rows[0] as DbUserRow;

        const membership = await scimApp.db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
          orgId,
          id
        ]);
        const wasMember = membership.rowCount === 1;
        if (!wasMember && changes.active !== true) return sendScimError(reply, 404, "User not found");

        let deactivatePerformed = false;
        let reactivatePerformed = false;

        try {
          await withTransaction(scimApp.db, async (client) => {
            if (changes.userName) {
              const email = changes.userName.trim().toLowerCase();
              await client.query("UPDATE users SET email = $2, updated_at = now() WHERE id = $1", [id, email]);
              user.email = email;
            }

            if (changes.displayName) {
              const nextName = changes.displayName.trim();
              if (nextName.length > 0) {
                await client.query("UPDATE users SET name = $2, updated_at = now() WHERE id = $1", [id, nextName]);
                user.name = nextName;
              }
            } else if (changes.nameFormatted || changes.nameGiven || changes.nameFamily) {
              const nextName = buildDisplayName(user.email, {
                formatted: changes.nameFormatted ?? null,
                given: changes.nameGiven ?? null,
                family: changes.nameFamily ?? null
              });
              await client.query("UPDATE users SET name = $2, updated_at = now() WHERE id = $1", [id, nextName]);
              user.name = nextName;
            }

            if (changes.active === false) {
              const del = await client.query("DELETE FROM org_members WHERE org_id = $1 AND user_id = $2", [orgId, id]);
              deactivatePerformed = del.rowCount === 1;
            } else if (changes.active === true) {
              const ins = await client.query(
                `
                  INSERT INTO org_members (org_id, user_id, role)
                  VALUES ($1, $2, 'member')
                  ON CONFLICT (org_id, user_id) DO NOTHING
                `,
                [orgId, id]
              );
              reactivatePerformed = ins.rowCount === 1;
            }
          });
        } catch (err: any) {
          // Handle unique email constraint in a SCIM-friendly way.
          if (typeof err?.code === "string" && err.code === "23505") {
            return sendScimError(reply, 409, "Email already in use");
          }
          throw err;
        }

        const memberAfter = await scimApp.db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
          orgId,
          id
        ]);
        const active = memberAfter.rowCount === 1;

        if (deactivatePerformed) {
          await writeAuditEvent(
            scimApp.db,
            createAuditEvent({
              eventType: "scim.user.deactivated",
              actor: { type: "scim_token", id: request.scimToken!.id },
              context: {
                orgId,
                ipAddress: getClientIp(request),
                userAgent: getUserAgent(request)
              },
              resource: { type: "user", id, name: user.email },
              success: true,
              details: { source: "scim" }
            })
          );
        }

        if (reactivatePerformed) {
          await writeAuditEvent(
            scimApp.db,
            createAuditEvent({
              eventType: "scim.user.reactivated",
              actor: { type: "scim_token", id: request.scimToken!.id },
              context: {
                orgId,
                ipAddress: getClientIp(request),
                userAgent: getUserAgent(request)
              },
              resource: { type: "user", id, name: user.email },
              success: true,
              details: { source: "scim" }
            })
          );
        }

        return reply.send(toScimUser(user, active));
      });

      scimApp.delete("/Users/:id", async (request, reply) => {
        const orgId = request.authOrgId!;
        const id = (request.params as { id: string }).id;
        if (!isValidUuid(id)) return sendScimError(reply, 404, "User not found");

        const res = await scimApp.db.query("DELETE FROM org_members WHERE org_id = $1 AND user_id = $2", [orgId, id]);
        if (res.rowCount !== 1) return sendScimError(reply, 404, "User not found");

        await writeAuditEvent(
          scimApp.db,
          createAuditEvent({
            eventType: "scim.user.removed_from_org",
            actor: { type: "scim_token", id: request.scimToken!.id },
            context: {
              orgId,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request)
            },
            resource: { type: "user", id },
            success: true,
            details: { source: "scim" }
          })
        );

        return reply.code(204).send();
      });
    },
    { prefix: "/scim/v2" }
  );
}
