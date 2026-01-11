import crypto from "node:crypto";
import type { DocumentRole } from "../rbac/roles";

type JwtModule = typeof import("jsonwebtoken");

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function loadJwt(): JwtModule {
  // Use runtime require so Vitest/Vite SSR doesn't try to transform jsonwebtoken
  // (a CJS package that relies on relative requires like `./decode`).
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  return require("jsonwebtoken") as JwtModule;
}

export interface SyncTokenClaims {
  sub: string;
  docId: string;
  orgId: string;
  role: DocumentRole;
  /**
   * Optional issuing session id.
   *
   * Revoking a session (`sessions.revoked_at`) implicitly revokes all derived sync
   * tokens because sync-server can revalidate this session via the internal
   * `/internal/sync/introspect` endpoint.
   */
  sessionId?: string;
}

export function signSyncToken(params: {
  secret: string;
  ttlSeconds: number;
  claims: SyncTokenClaims;
}): { token: string; expiresAt: Date } {
  const jwt = loadJwt();
  const expiresAt = new Date(Date.now() + params.ttlSeconds * 1000);
  const token = jwt.sign(params.claims, params.secret, {
    algorithm: "HS256",
    expiresIn: params.ttlSeconds,
    audience: "formula-sync",
    // Add a unique identifier so we can support explicit sync token revocation in
    // the future if needed (in addition to session-based revocation).
    jwtid: crypto.randomUUID()
  });
  return { token, expiresAt };
}

function isDocumentRole(value: unknown): value is DocumentRole {
  return value === "owner" || value === "admin" || value === "editor" || value === "commenter" || value === "viewer";
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" && UUID_REGEX.test(value);
}

export function verifySyncToken(params: { token: string; secret: string }): SyncTokenClaims {
  const jwt = loadJwt();
  const verified = jwt.verify(params.token, params.secret, {
    algorithms: ["HS256"],
    audience: "formula-sync"
  });

  if (!verified || typeof verified !== "object") {
    throw new Error("invalid_sync_token");
  }

  const payload = verified as Record<string, unknown>;
  const sub = payload.sub;
  const docId = payload.docId;
  const orgId = payload.orgId;
  const role = payload.role;
  const sessionId = payload.sessionId;

  if (!isUuid(sub)) throw new Error("invalid_sync_token");
  if (!isUuid(docId)) throw new Error("invalid_sync_token");
  if (!isUuid(orgId)) throw new Error("invalid_sync_token");
  if (!isDocumentRole(role)) throw new Error("invalid_sync_token");
  if (sessionId !== undefined && !isUuid(sessionId)) {
    throw new Error("invalid_sync_token");
  }

  const claims: SyncTokenClaims = {
    sub,
    docId,
    orgId,
    role
  };

  if (sessionId !== undefined) {
    claims.sessionId = sessionId;
  }

  return claims;
}
