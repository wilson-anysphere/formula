import type { IncomingMessage } from "node:http";
import jwt from "jsonwebtoken";

import { normalizeRestriction } from "../../../packages/collab/permissions/index.js";

import type { AuthMode } from "./config.js";

export type SyncRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

export type AuthContext = {
  userId: string;
  tokenType: "opaque" | "jwt";
  docId: string;
  orgId: string | null;
  role: SyncRole;
  sessionId?: string | null;
  rangeRestrictions?: unknown[];
};

export class AuthError extends Error {
  constructor(
    message: string,
    public readonly statusCode: 401 | 403 = 401
  ) {
    super(message);
    this.name = "AuthError";
  }
}

function parseBearerAuthorizationHeader(value: string | undefined): string | null {
  if (!value) return null;
  const match = value.match(/^Bearer\s+(.+)$/i);
  return match ? match[1] : null;
}

export function extractToken(req: IncomingMessage, url: URL): string | null {
  const fromQuery = url.searchParams.get("token");
  if (fromQuery) return fromQuery;

  const header = req.headers["authorization"];
  if (typeof header === "string") return parseBearerAuthorizationHeader(header);
  return null;
}

function isStringArray(value: unknown): value is string[] {
  return (
    Array.isArray(value) &&
    value.every((v) => typeof v === "string" && v.length > 0)
  );
}

function requireJwtPayloadObject(payload: unknown): Record<string, unknown> {
  if (!payload || typeof payload !== "object") {
    throw new AuthError("Invalid JWT payload", 403);
  }
  return payload as Record<string, unknown>;
}

function authorizeDocAccessFromJwtPayload(
  payload: Record<string, unknown>,
  docName: string
): string {
  const docId = payload.docId;
  if (docId !== undefined) {
    if (typeof docId !== "string" || docId.length === 0) {
      throw new AuthError('JWT "docId" claim must be a non-empty string', 403);
    }

    if (docId !== docName) {
      throw new AuthError("Token is not authorized for this document", 403);
    }

    return docId;
  }

  const docs = payload.docs;
  const doc = payload.doc;

  const allowedDocs = isStringArray(docs)
    ? docs
    : typeof doc === "string" && doc.length > 0
      ? [doc]
      : null;

  if (!allowedDocs) {
    throw new AuthError(
      'JWT is missing a "docId" (string), "docs" (string[]) or "doc" (string) claim',
      403
    );
  }

  if (allowedDocs.includes("*")) return docName;
  if (!allowedDocs.includes(docName)) {
    throw new AuthError("Token is not authorized for this document", 403);
  }

  return docName;
}

function parseRoleFromJwtPayload(payload: Record<string, unknown>): SyncRole {
  const role = payload.role;
  if (role === undefined) return "editor";

  if (
    role === "owner" ||
    role === "admin" ||
    role === "editor" ||
    role === "commenter" ||
    role === "viewer"
  ) {
    return role;
  }

  throw new AuthError('JWT "role" claim is invalid', 403);
}

function parseOptionalStringClaim(
  payload: Record<string, unknown>,
  claim: string
): string | null | undefined {
  const value = payload[claim];
  if (value === undefined) return undefined;
  if (value === null) return null;
  if (typeof value === "string" && value.length > 0) return value;
  throw new AuthError(`JWT "${claim}" claim must be a non-empty string`, 403);
}

function parseOptionalRangeRestrictionsClaim(
  payload: Record<string, unknown>
): unknown[] | undefined {
  const value = payload.rangeRestrictions;
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) {
    throw new AuthError('JWT "rangeRestrictions" claim must be an array', 403);
  }

  for (const [index, entry] of value.entries()) {
    try {
      normalizeRestriction(entry);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      throw new AuthError(
        `JWT "rangeRestrictions"[${index}] is invalid: ${message}`,
        403
      );
    }
  }

  return value as unknown[];
}

export function authenticateRequest(
  auth: AuthMode,
  token: string | null,
  docName: string
): AuthContext {
  if (!token) throw new AuthError("Missing token", 401);

  if (auth.mode === "opaque") {
    if (token !== auth.token) throw new AuthError("Invalid token", 401);
    return {
      userId: "opaque",
      tokenType: "opaque",
      docId: docName,
      orgId: null,
      role: "owner",
    };
  }

  let verifiedPayload: unknown;
  try {
    verifiedPayload = jwt.verify(token, auth.secret, {
      algorithms: ["HS256"],
      issuer: auth.issuer,
      audience: auth.audience,
    });
  } catch {
    throw new AuthError("Invalid token", 401);
  }

  const payload = requireJwtPayloadObject(verifiedPayload);
  const resolvedDocId = authorizeDocAccessFromJwtPayload(payload, docName);

  const sub = payload.sub;
  const userId = typeof sub === "string" && sub.length > 0 ? sub : "jwt";
  const role = parseRoleFromJwtPayload(payload);

  const orgIdValue = parseOptionalStringClaim(payload, "orgId");
  const orgId = orgIdValue === undefined ? null : orgIdValue;

  const sessionId = parseOptionalStringClaim(payload, "sessionId");
  const rangeRestrictions = parseOptionalRangeRestrictionsClaim(payload);

  const ctx: AuthContext = {
    userId,
    tokenType: "jwt",
    docId: resolvedDocId,
    orgId,
    role,
  };

  if (sessionId !== undefined) {
    ctx.sessionId = sessionId;
  }

  if (rangeRestrictions !== undefined) {
    ctx.rangeRestrictions = rangeRestrictions;
  }

  return ctx;
}
