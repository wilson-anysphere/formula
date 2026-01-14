import crypto from "node:crypto";
import type { Pool, PoolClient } from "pg";

export interface SessionInfo {
  id: string;
  userId: string;
  mfaSatisfied: boolean;
  expiresAt: Date;
  revokedAt: Date | null;
}

export interface AuthenticatedUser {
  id: string;
  email: string;
  name: string;
  mfaTotpEnabled: boolean;
}

export interface SessionLookupResult {
  session: SessionInfo;
  user: AuthenticatedUser;
}

export interface CreateSessionInput {
  userId: string;
  expiresAt: Date;
  ipAddress?: string | null;
  userAgent?: string | null;
  mfaSatisfied?: boolean;
}

export function generateSessionToken(): string {
  return crypto.randomBytes(32).toString("hex");
}

export function hashSessionToken(token: string): string {
  return crypto.createHash("sha256").update(token).digest("hex");
}

export async function createSession(
  client: PoolClient,
  input: CreateSessionInput
): Promise<{ sessionId: string; token: string; expiresAt: Date }> {
  const sessionId = crypto.randomUUID();
  const token = generateSessionToken();
  const tokenHash = hashSessionToken(token);

  await client.query(
    `
      INSERT INTO sessions (id, user_id, token_hash, expires_at, ip_address, user_agent, mfa_satisfied)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
    `,
    [
      sessionId,
      input.userId,
      tokenHash,
      input.expiresAt,
      input.ipAddress ?? null,
      input.userAgent ?? null,
      input.mfaSatisfied ?? false
    ]
  );

  return { sessionId, token, expiresAt: input.expiresAt };
}

export async function lookupSessionByToken(pool: Pool, token: string): Promise<SessionLookupResult | null> {
  const tokenHash = hashSessionToken(token);
  const result = await pool.query(
    `
      SELECT
        s.id AS session_id,
        s.user_id AS session_user_id,
        s.mfa_satisfied AS session_mfa_satisfied,
        s.expires_at AS session_expires_at,
        s.revoked_at AS session_revoked_at,
        u.id AS user_id,
        u.email AS user_email,
        u.name AS user_name,
        u.mfa_totp_enabled AS user_mfa_totp_enabled
      FROM sessions s
      JOIN users u ON u.id = s.user_id
      WHERE s.token_hash = $1
        AND s.revoked_at IS NULL
        AND s.expires_at > now()
      LIMIT 1
    `,
    [tokenHash]
  );

  if (result.rowCount !== 1) return null;
  const row = result.rows[0] as {
    session_id: string;
    session_user_id: string;
    session_mfa_satisfied: boolean;
    session_expires_at: Date;
    session_revoked_at: Date | null;
    user_id: string;
    user_email: string;
    user_name: string;
    user_mfa_totp_enabled: boolean;
  };

  // Best-effort "touch" for session activity.
  void pool.query("UPDATE sessions SET last_used_at = now() WHERE id = $1", [row.session_id]).catch(() => {
    // Best-effort; ignore failures updating the timestamp.
  });

  return {
    session: {
      id: row.session_id,
      userId: row.session_user_id,
      mfaSatisfied: Boolean(row.session_mfa_satisfied),
      expiresAt: new Date(row.session_expires_at),
      revokedAt: row.session_revoked_at ? new Date(row.session_revoked_at) : null
    },
    user: {
      id: row.user_id,
      email: row.user_email,
      name: row.user_name,
      mfaTotpEnabled: row.user_mfa_totp_enabled
    }
  };
}

export async function revokeSession(pool: Pool, sessionId: string): Promise<void> {
  await pool.query("UPDATE sessions SET revoked_at = now() WHERE id = $1", [sessionId]);
}

type Queryable = Pick<Pool, "query"> | Pick<PoolClient, "query">;

export async function setSessionMfaSatisfied(
  db: Queryable,
  options: { sessionId: string; userId: string; mfaSatisfied: boolean }
): Promise<void> {
  await db.query("UPDATE sessions SET mfa_satisfied = $1 WHERE id = $2 AND user_id = $3", [
    options.mfaSatisfied,
    options.sessionId,
    options.userId
  ]);
}
