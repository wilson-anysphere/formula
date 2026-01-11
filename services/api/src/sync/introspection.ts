import type { Pool } from "pg";
import { z } from "zod";

import { isClientIpAllowed } from "../auth/apiKeys";
import type { DocumentRole } from "../rbac/roles";

type JwtModule = typeof import("jsonwebtoken");

function loadJwt(): JwtModule {
  // Use runtime require so Vitest/Vite SSR doesn't try to transform jsonwebtoken
  // (a CJS package that relies on relative requires like `./decode`).
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  return require("jsonwebtoken") as JwtModule;
}

const RoleSchema = z.enum(["owner", "admin", "editor", "commenter", "viewer"]);

const SyncJwtClaimsSchema = z
  .object({
    sub: z.string().uuid(),
    docId: z.string().uuid(),
    orgId: z.string().uuid(),
    role: RoleSchema,
    sessionId: z.string().uuid().optional()
  })
  // Allow standard claims (iat, exp, aud, jti, etc).
  .passthrough();

export type SyncIntrospectionResult =
  | {
      active: true;
      userId: string;
      orgId: string;
      role: DocumentRole;
      sessionId?: string;
    }
  | {
      active: false;
      reason: string;
      userId?: string;
      orgId?: string;
      role?: DocumentRole;
      sessionId?: string;
    };

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

function requireObject(payload: unknown): Record<string, unknown> | null {
  if (!payload || typeof payload !== "object") return null;
  return payload as Record<string, unknown>;
}

export async function introspectSyncToken(
  db: Pool,
  params: {
    secret: string;
    token: string;
    docId: string;
    clientIp?: string | null;
    userAgent?: string | null;
  }
): Promise<SyncIntrospectionResult> {
  const jwt = loadJwt();

  let verified: unknown;
  try {
    verified = jwt.verify(params.token, params.secret, {
      algorithms: ["HS256"],
      audience: "formula-sync"
    });
  } catch (err) {
    const reason =
      err instanceof Error && err.name === "TokenExpiredError" ? "token_expired" : "invalid_token";
    return { active: false, reason };
  }

  const payload = requireObject(verified);
  if (!payload) return { active: false, reason: "invalid_claims" };

  const parsedClaims = SyncJwtClaimsSchema.safeParse(payload);
  if (!parsedClaims.success) return { active: false, reason: "invalid_claims" };

  const claims = parsedClaims.data;
  const userId = claims.sub;
  const orgId = claims.orgId;
  const role = claims.role as DocumentRole;

  if (claims.docId !== params.docId) {
    return { active: false, reason: "doc_mismatch", userId, orgId, role };
  }

  if (claims.sessionId) {
    const sessionRes = await db.query(
      `
        SELECT user_id, expires_at, revoked_at
        FROM sessions
        WHERE id = $1
        LIMIT 1
      `,
      [claims.sessionId]
    );

    if (sessionRes.rowCount !== 1) {
      return { active: false, reason: "session_not_found", userId, orgId, role };
    }

    const session = sessionRes.rows[0] as { user_id: string; expires_at: Date; revoked_at: Date | null };

    if (session.user_id !== userId) {
      return { active: false, reason: "session_user_mismatch", userId, orgId, role };
    }
    if (session.revoked_at) {
      return { active: false, reason: "session_revoked", userId, orgId, role };
    }

    const expiresAtMs = new Date(session.expires_at).getTime();
    if (!Number.isFinite(expiresAtMs) || expiresAtMs <= Date.now()) {
      return { active: false, reason: "session_expired", userId, orgId, role };
    }
  }

  const docRes = await db.query(
    `
      SELECT d.org_id, d.deleted_at, dm.role AS member_role, os.ip_allowlist
      FROM documents d
      JOIN org_settings os ON os.org_id = d.org_id
      LEFT JOIN document_members dm
        ON dm.document_id = d.id AND dm.user_id = $2
      WHERE d.id = $1
      LIMIT 1
    `,
    [claims.docId, userId]
  );

  if (docRes.rowCount !== 1) {
    return { active: false, reason: "doc_not_found", userId, orgId, role };
  }

  const row = docRes.rows[0] as {
    org_id: string;
    deleted_at: Date | null;
    member_role: unknown;
    ip_allowlist: unknown;
  };

  if (row.org_id !== orgId) {
    return { active: false, reason: "org_mismatch", userId, orgId, role };
  }

  if (row.deleted_at) {
    return { active: false, reason: "doc_deleted", userId, orgId, role };
  }

  const memberRoleParsed = RoleSchema.safeParse(row.member_role);
  const memberRole = memberRoleParsed.success ? (memberRoleParsed.data as DocumentRole) : null;
  if (!memberRole) {
    return { active: false, reason: "not_member", userId, orgId, role };
  }

  // Permissions may change after a sync token is minted (e.g. owner demotes an
  // editor to viewer). Treat the token role as an upper bound and clamp it to
  // the current DB role to prevent privilege escalation without requiring the
  // client to refresh its sync token.
  const effectiveRole = roleRank(role) > roleRank(memberRole) ? memberRole : role;

  if (!isClientIpAllowed(params.clientIp ?? null, row.ip_allowlist)) {
    return { active: false, reason: "ip_not_allowed", userId, orgId, role, sessionId: claims.sessionId };
  }

  return {
    active: true,
    userId,
    orgId,
    role: effectiveRole,
    sessionId: claims.sessionId
  };
}
