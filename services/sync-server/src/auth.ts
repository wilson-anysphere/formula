import type { IncomingMessage } from "node:http";
import jwt from "jsonwebtoken";

import type { AuthMode } from "./config.js";

export type AuthContext = {
  userId: string;
  tokenType: "opaque" | "jwt";
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

function authorizeDocAccessFromJwtPayload(payload: unknown, docName: string) {
  if (!payload || typeof payload !== "object") {
    throw new AuthError("Invalid JWT payload", 403);
  }

  const docs = (payload as { docs?: unknown }).docs;
  const doc = (payload as { doc?: unknown }).doc;

  const allowedDocs = isStringArray(docs)
    ? docs
    : typeof doc === "string" && doc.length > 0
      ? [doc]
      : null;

  if (!allowedDocs) {
    throw new AuthError(
      'JWT is missing a "docs" (string[]) or "doc" (string) claim',
      403
    );
  }

  if (allowedDocs.includes("*")) return;
  if (!allowedDocs.includes(docName)) {
    throw new AuthError("Token is not authorized for this document", 403);
  }
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
    };
  }

  const payload = jwt.verify(token, auth.secret, {
    algorithms: ["HS256"],
    issuer: auth.issuer,
    audience: auth.audience,
  });

  authorizeDocAccessFromJwtPayload(payload, docName);

  const sub =
    payload && typeof payload === "object" && "sub" in payload
      ? (payload as { sub?: unknown }).sub
      : undefined;

  return {
    userId: typeof sub === "string" && sub.length > 0 ? sub : "jwt",
    tokenType: "jwt",
  };
}

