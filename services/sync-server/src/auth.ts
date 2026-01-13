import type { IncomingMessage } from "node:http";
import crypto from "node:crypto";
import jwt from "jsonwebtoken";

import { normalizeRestriction } from "../../../packages/collab/permissions/index.js";

import type { AuthMode } from "./config.js";
import { statusCodeForIntrospectionReason } from "./introspection-reasons.js";
import type {
  IntrospectionRequestResult,
  SyncServerMetrics,
} from "./metrics.js";

export type SyncRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

export type AuthContext = {
  userId: string;
  tokenType: "opaque" | "jwt" | "introspect";
  docId: string;
  orgId: string | null;
  role: SyncRole;
  sessionId?: string | null;
  rangeRestrictions?: unknown[];
};

export class AuthError extends Error {
  constructor(
    message: string,
    public readonly statusCode: 401 | 403 | 414 | 503 = 401
  ) {
    super(message);
    this.name = "AuthError";
  }
}

function timingSafeEqualStrings(a: string, b: string): boolean {
  const aBuf = Buffer.from(a);
  const bBuf = Buffer.from(b);
  if (aBuf.length !== bBuf.length) return false;
  return crypto.timingSafeEqual(aBuf, bBuf);
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

export type IntrospectCache = Map<
  string,
  {
    expiresAtMs: number;
    ctx: AuthContext;
  }
>;

const introspectCacheLastSweepAtMs = new WeakMap<IntrospectCache, number>();

function maybeSweepIntrospectCache(cache: IntrospectCache, now: number, cacheMs: number): void {
  const sweepIntervalMs = Math.max(cacheMs, 30_000);
  const maxEntriesBeforeSweep = 10_000;
  const lastSweepAtMs = introspectCacheLastSweepAtMs.get(cache) ?? 0;

  if (cache.size <= maxEntriesBeforeSweep && now - lastSweepAtMs < sweepIntervalMs) {
    return;
  }

  for (const [key, entry] of cache) {
    if (entry.expiresAtMs <= now) {
      cache.delete(key);
    }
  }
  introspectCacheLastSweepAtMs.set(cache, now);
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

function sha256Hex(value: string): string {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function parseJwtExpMs(token: string): number | null {
  try {
    const decoded = jwt.decode(token);
    if (!decoded || typeof decoded !== "object") return null;
    const exp = (decoded as any).exp;
    return typeof exp === "number" && Number.isFinite(exp) ? exp * 1000 : null;
  } catch {
    return null;
  }
}

async function introspectTokenWithRetry(
  auth: Extract<AuthMode, { mode: "introspect" }>,
  token: string,
  docId: string,
  options: { clientIp?: string | null; userAgent?: string | null } = {}
): Promise<{
  userId: string;
  orgId: string;
  role: SyncRole;
  sessionId?: string | null;
  rangeRestrictions?: unknown[];
}> {
  const url = new URL("/internal/sync/introspect", auth.url).toString();

  const timeoutMs = 5_000;
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-internal-admin-token": auth.token,
        },
        body: JSON.stringify({
          token,
          docId,
          ...(options.clientIp ? { clientIp: options.clientIp } : {}),
          ...(options.userAgent ? { userAgent: options.userAgent } : {}),
        }),
        signal: AbortSignal.timeout(timeoutMs),
      });

      if (res.status === 401 || res.status === 403) {
        throw new AuthError("Invalid token", res.status === 401 ? 401 : 403);
      }

      if (!res.ok) {
        // Retry on transient errors (gateway issues, timeouts, 5xx).
        if (res.status >= 500 && res.status <= 599 && attempt < maxAttempts) {
          const backoffMs =
            Math.min(250 * 2 ** (attempt - 1), 2_000) +
            Math.floor(Math.random() * 100);
          await new Promise((resolve) => setTimeout(resolve, backoffMs));
          continue;
        }
        throw new AuthError("Authentication service unavailable", 503);
      }

      let body: any;
      try {
        body = (await res.json()) as any;
      } catch {
        throw new AuthError("Authentication service unavailable", 503);
      }

      if (!body || typeof body !== "object") {
        throw new AuthError("Authentication service unavailable", 503);
      }

      const active =
        typeof body.active === "boolean"
          ? body.active
          : typeof body.ok === "boolean"
            ? body.ok
            : null;
      if (active === null) {
        throw new AuthError("Authentication service unavailable", 503);
      }

      if (!active) {
        const reason =
          typeof body.reason === "string" && body.reason.length > 0
            ? body.reason
            : typeof body.error === "string" && body.error.length > 0
              ? body.error
              : undefined;
        throw new AuthError(
          reason ? `Token inactive: ${reason}` : "Invalid token",
          statusCodeForIntrospectionReason(reason)
        );
      }

      const userId = body.userId;
      const orgId = body.orgId;
      const role = body.role;
      const sessionId = body.sessionId;
      const rangeRestrictions = body.rangeRestrictions;

      if (typeof userId !== "string" || userId.length === 0) {
        throw new AuthError("Authentication service unavailable", 503);
      }
      if (typeof orgId !== "string" || orgId.length === 0) {
        throw new AuthError("Authentication service unavailable", 503);
      }
      if (
        role !== "owner" &&
        role !== "admin" &&
        role !== "editor" &&
        role !== "commenter" &&
        role !== "viewer"
      ) {
        throw new AuthError("Authentication service unavailable", 503);
      }

      if (sessionId !== undefined && sessionId !== null && typeof sessionId !== "string") {
        throw new AuthError("Authentication service unavailable", 503);
      }

      if (rangeRestrictions !== undefined) {
        if (!Array.isArray(rangeRestrictions)) {
          throw new AuthError("Authentication service unavailable", 503);
        }
        for (const entry of rangeRestrictions) {
          try {
            normalizeRestriction(entry);
          } catch {
            throw new AuthError("Authentication service unavailable", 503);
          }
        }
      }

      return {
        userId,
        orgId,
        role,
        sessionId: sessionId === undefined ? undefined : sessionId,
        rangeRestrictions:
          rangeRestrictions === undefined ? undefined : (rangeRestrictions as unknown[]),
      };
    } catch (err) {
      if (err instanceof AuthError) {
        throw err;
      }

      if (attempt < maxAttempts) {
        const backoffMs =
          Math.min(250 * 2 ** (attempt - 1), 2_000) +
          Math.floor(Math.random() * 100);
        await new Promise((resolve) => setTimeout(resolve, backoffMs));
        continue;
      }

      throw new AuthError("Authentication service unavailable", 503);
    }
  }

  throw new AuthError("Authentication service unavailable", 503);
}

export async function authenticateRequest(
  auth: AuthMode,
  token: string | null,
  docName: string,
  options: {
    introspectCache?: IntrospectCache | null;
    clientIp?: string | null;
    userAgent?: string | null;
    metrics?: Pick<SyncServerMetrics, "introspectionRequestDurationMs">;
    /**
     * Optional token length limit (in bytes, utf-8). When exceeded, the token is
     * rejected before any hashing/verification/introspection work occurs.
     */
    maxTokenBytes?: number;
  } = {}
): Promise<AuthContext> {
  if (!token) throw new AuthError("Missing token", 401);

  const maxTokenBytes = options.maxTokenBytes ?? 0;
  if (maxTokenBytes > 0 && Buffer.byteLength(token, "utf8") > maxTokenBytes) {
    throw new AuthError("Token too long", 414);
  }

  if (auth.mode === "opaque") {
    if (!timingSafeEqualStrings(token, auth.token)) throw new AuthError("Invalid token", 401);
    return {
      userId: "opaque",
      tokenType: "opaque",
      docId: docName,
      orgId: null,
      role: "owner",
    };
  }

  if (auth.mode === "introspect") {
    const cache = options.introspectCache ?? null;
    // Cache entries are scoped per (token, doc, clientIp) so a single user/session
    // token can be safely used across multiple documents and cannot be replayed
    // from a different client IP during the cache window.
    const tokenHash = cache
      ? sha256Hex(`${token}\n${docName}\n${options.clientIp ?? ""}`)
      : null;
    const now = Date.now();
    if (cache && tokenHash) {
      maybeSweepIntrospectCache(cache, now, auth.cacheMs);
      const cached = cache.get(tokenHash);
      if (cached) {
        if (cached.expiresAtMs > now) {
          return cached.ctx;
        }
        cache.delete(tokenHash);
      }
    }

    try {
      let metricResult: IntrospectionRequestResult = "error";
      const startHr = process.hrtime.bigint();
      const result = await (async () => {
        try {
          const value = await introspectTokenWithRetry(auth, token, docName, {
            clientIp: options.clientIp ?? null,
            userAgent: options.userAgent ?? null,
          });
          metricResult = "ok";
          return value;
        } catch (err) {
          const inactive =
            err instanceof AuthError && (err.statusCode === 401 || err.statusCode === 403);
          metricResult = inactive ? "inactive" : "error";
          throw err;
        } finally {
          const durationMs = Number(process.hrtime.bigint() - startHr) / 1e6;
          try {
            options.metrics?.introspectionRequestDurationMs.set(
              { path: "auth_mode", result: metricResult },
              durationMs
            );
          } catch {
            // ignore
          }
        }
      })();
      const ctx: AuthContext = {
        userId: result.userId,
        tokenType: "introspect",
        docId: docName,
        orgId: result.orgId,
        role: result.role,
      };

      if (result.sessionId !== undefined) {
        ctx.sessionId = result.sessionId;
      }

      if (result.rangeRestrictions !== undefined) {
        ctx.rangeRestrictions = result.rangeRestrictions;
      }

      if (cache && tokenHash) {
        const expMs = parseJwtExpMs(token);
        const cacheUntil = now + auth.cacheMs;
        const expiresAtMs =
          typeof expMs === "number" && Number.isFinite(expMs)
            ? Math.min(expMs, cacheUntil)
            : cacheUntil;
        cache.set(tokenHash, { expiresAtMs, ctx });
      }

      return ctx;
    } catch (err) {
      const invalidToken =
        err instanceof AuthError && (err.statusCode === 401 || err.statusCode === 403);
      if (auth.failOpen && !invalidToken) {
        return {
          userId: "introspect-fail-open",
          tokenType: "introspect",
          docId: docName,
          orgId: null,
          role: "owner",
        };
      }
      throw err;
    }
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

  // Optional hardening: require a JWT `exp` claim to prevent long-lived tokens.
  //
  // We treat missing `exp` as an authentication failure (401), similar to an
  // invalid/expired token.
  if (auth.requireExp) {
    const exp = payload.exp;
    if (exp === undefined) {
      throw new AuthError('JWT "exp" claim is required', 401);
    }
  }

  const resolvedDocId = authorizeDocAccessFromJwtPayload(payload, docName);

  const sub = payload.sub;
  if (auth.requireSub) {
    if (typeof sub !== "string" || sub.length === 0) {
      throw new AuthError('JWT "sub" claim is required', 403);
    }
  }
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
