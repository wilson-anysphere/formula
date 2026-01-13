import crypto from "node:crypto";

import type { SyncRole } from "./auth.js";
import type { IntrospectionRequestResult, SyncServerMetrics } from "./metrics.js";

export type SyncTokenIntrospectionResult = {
  active: boolean;
  reason?: string;
  userId?: string;
  orgId?: string;
  role?: SyncRole;
};

function parseRole(value: unknown): SyncRole | undefined {
  switch (value) {
    case "owner":
    case "admin":
    case "editor":
    case "commenter":
    case "viewer":
      return value;
    default:
      return undefined;
  }
}

function sha256Hex(value: string): string {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function parseIntrospectionResult(value: unknown): SyncTokenIntrospectionResult {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid introspection response (expected JSON object)");
  }

  const obj = value as Record<string, unknown>;
  const active =
    typeof obj.active === "boolean"
      ? obj.active
      : typeof obj.ok === "boolean"
        ? obj.ok
        : null;
  if (active === null) {
    throw new Error('Invalid introspection response (missing boolean "active"/"ok")');
  }

  const reason =
    typeof obj.reason === "string" && obj.reason.length > 0
      ? obj.reason
      : typeof obj.error === "string" && obj.error.length > 0
        ? obj.error
        : undefined;
  const userId = typeof obj.userId === "string" && obj.userId.length > 0 ? obj.userId : undefined;
  const orgId = typeof obj.orgId === "string" && obj.orgId.length > 0 ? obj.orgId : undefined;
  const role = parseRole(obj.role);

  return {
    active,
    reason,
    userId,
    orgId,
    role,
  };
}

export type SyncTokenIntrospectionClient = {
  introspect: (params: {
    token: string;
    docId: string;
    clientIp?: string;
    userAgent?: string;
  }) => Promise<SyncTokenIntrospectionResult>;
};

export class SyncTokenIntrospectionOverCapacityError extends Error {
  constructor(message: string = "Introspection over capacity") {
    super(message);
    this.name = "SyncTokenIntrospectionOverCapacityError";
  }
}

export function createSyncTokenIntrospectionClient(config: {
  url: string;
  token: string;
  cacheTtlMs: number;
  maxConcurrent?: number;
  metrics?: Pick<
    SyncServerMetrics,
    "introspectionOverCapacityTotal" | "introspectionRequestsTotal" | "introspectionRequestDurationMs"
  >;
}): SyncTokenIntrospectionClient {
  const cache = new Map<
    string,
    {
      expiresAtMs: number;
      value?: SyncTokenIntrospectionResult;
      inFlight?: Promise<SyncTokenIntrospectionResult>;
    }
  >();

  const cacheTtlMs = Math.max(0, Math.floor(config.cacheTtlMs));
  const cacheSweepIntervalMs = Math.max(cacheTtlMs, 30_000);
  const maxEntriesBeforeSweep = 10_000;
  let lastCacheSweepAtMs = Date.now();

  const maxConcurrent = Math.max(0, Math.floor(config.maxConcurrent ?? 50));
  const metrics = config.metrics;
  let activeRequests = 0;

  const cacheKey = (params: { token: string; docId: string; clientIp?: string }) =>
    // Hash the key so we don't retain raw JWTs/opaque tokens (which could be large
    // and sensitive) in the cache map keys.
    sha256Hex(`${params.token}\n${params.docId}\n${params.clientIp ?? ""}`);

  const introspectInner: SyncTokenIntrospectionClient["introspect"] = async (params) => {
    const key = cacheKey(params);
    const now = Date.now();

    // Opportunistically sweep expired entries to prevent unbounded growth when
    // tokens are mostly one-off (e.g. short-lived JWTs).
    if (cacheTtlMs > 0 && cache.size > 0) {
      const shouldSweep =
        cache.size > maxEntriesBeforeSweep || now - lastCacheSweepAtMs > cacheSweepIntervalMs;
      if (shouldSweep) {
        for (const [cachedKey, entry] of cache) {
          if (entry.expiresAtMs <= now && !entry.inFlight) {
            cache.delete(cachedKey);
          }
        }
        lastCacheSweepAtMs = now;
      }
    }

    const cached = cache.get(key);

    if (cacheTtlMs > 0 && cached && cached.value && cached.expiresAtMs > now) {
      return cached.value;
    }

    if (cacheTtlMs > 0 && cached?.inFlight) {
      return await cached.inFlight;
    }

    const task = (async () => {
      let acquired = false;
      let recorded = false;
      const record = (result: "ok" | "inactive" | "error") => {
        if (recorded) return;
        recorded = true;
        metrics?.introspectionRequestsTotal.inc({ result });
      };

      if (maxConcurrent > 0) {
        if (activeRequests >= maxConcurrent) {
          metrics?.introspectionOverCapacityTotal.inc();
          record("error");
          throw new SyncTokenIntrospectionOverCapacityError();
        }
        activeRequests += 1;
        acquired = true;
      }

      try {
        const res = await fetch(config.url, {
          method: "POST",
          headers: {
            "content-type": "application/json",
            // API internal endpoints use the same header name as other internal ops.
            "x-internal-admin-token": config.token,
          },
          body: JSON.stringify({
            token: params.token,
            docId: params.docId,
            clientIp: params.clientIp,
            userAgent: params.userAgent,
          }),
          signal: AbortSignal.timeout(5_000),
        });

        const json = (await res.json().catch(() => null)) as unknown;

        if (res.status === 401 || res.status === 403) {
          // Token is inactive/invalid (or the introspection endpoint rejected our request).
          // Never treat a 401/403 response as an active token, even if the body is malformed or claims otherwise.
          const reason =
            json && typeof json === "object"
              ? typeof (json as any).reason === "string" && (json as any).reason.length > 0
                ? ((json as any).reason as string)
                : typeof (json as any).error === "string" && (json as any).error.length > 0
                  ? ((json as any).error as string)
                  : "forbidden"
              : "forbidden";

          record("inactive");
          return { active: false, reason };
        }

        if (!res.ok) {
          record("error");
          throw new Error(`Introspection request failed (${res.status})`);
        }

        const parsed = parseIntrospectionResult(json);
        record(parsed.active ? "ok" : "inactive");
        return parsed;
      } catch (err) {
        record("error");
        throw err;
      } finally {
        if (acquired) {
          activeRequests = Math.max(0, activeRequests - 1);
        }
      }
    })();

    if (cacheTtlMs > 0) {
      cache.set(key, {
        expiresAtMs: now + cacheTtlMs,
        inFlight: task
      });
    }

    try {
      const value = await task;
      if (cacheTtlMs > 0) {
        cache.set(key, { expiresAtMs: now + cacheTtlMs, value });
      }
      return value;
    } catch (err) {
      cache.delete(key);
      throw err;
    }
  };

  const introspect: SyncTokenIntrospectionClient["introspect"] = async (params) => {
    const startHr = process.hrtime.bigint();
    let metricResult: IntrospectionRequestResult = "error";
    try {
      const value = await introspectInner(params);
      metricResult = value.active ? "ok" : "inactive";
      return value;
    } catch (err) {
      metricResult = "error";
      throw err;
    } finally {
      const durationMs = Number(process.hrtime.bigint() - startHr) / 1e6;
      try {
        metrics?.introspectionRequestDurationMs.set(
          { path: "jwt_revalidation", result: metricResult },
          durationMs
        );
      } catch {
        // ignore
      }
    }
  };

  return { introspect };
}
