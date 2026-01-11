import type { SyncRole } from "./auth.js";

export type SyncTokenIntrospectionResult = {
  ok: boolean;
  userId?: string;
  orgId?: string;
  role?: SyncRole;
  sessionId?: string | null;
  error?: string;
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

function parseIntrospectionResult(value: unknown): SyncTokenIntrospectionResult {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid introspection response (expected JSON object)");
  }

  const obj = value as Record<string, unknown>;
  if (typeof obj.ok !== "boolean") {
    throw new Error('Invalid introspection response (missing boolean "ok")');
  }

  const userId = typeof obj.userId === "string" && obj.userId.length > 0 ? obj.userId : undefined;
  const orgId = typeof obj.orgId === "string" && obj.orgId.length > 0 ? obj.orgId : undefined;
  const role = parseRole(obj.role);
  const sessionId =
    obj.sessionId === undefined || obj.sessionId === null || (typeof obj.sessionId === "string" && obj.sessionId.length > 0)
      ? (obj.sessionId as string | null | undefined)
      : undefined;
  const error = typeof obj.error === "string" && obj.error.length > 0 ? obj.error : undefined;

  return {
    ok: obj.ok,
    userId,
    orgId,
    role,
    sessionId,
    error
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

export function createSyncTokenIntrospectionClient(config: {
  url: string;
  token: string;
  cacheTtlMs: number;
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

  const cacheKey = (params: { token: string; docId: string; clientIp?: string }) =>
    `${params.token}\n${params.docId}\n${params.clientIp ?? ""}`;

  const introspect: SyncTokenIntrospectionClient["introspect"] = async (params) => {
    const key = cacheKey(params);
    const now = Date.now();
    const cached = cache.get(key);

    if (cacheTtlMs > 0 && cached && cached.value && cached.expiresAtMs > now) {
      return cached.value;
    }

    if (cacheTtlMs > 0 && cached?.inFlight) {
      return await cached.inFlight;
    }

    const task = (async () => {
      const res = await fetch(config.url, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          // API internal endpoints use the same header name as other internal ops.
          "x-internal-admin-token": config.token
        },
        body: JSON.stringify({
          token: params.token,
          docId: params.docId,
          clientIp: params.clientIp,
          userAgent: params.userAgent
        }),
        signal: AbortSignal.timeout(5_000)
      });

      const json = (await res.json().catch(() => null)) as unknown;

      if (res.status === 403) {
        // Token is inactive/invalid. API returns `{ ok: false, error: "forbidden" }`.
        if (!json) return { ok: false, error: "forbidden" };
        return parseIntrospectionResult(json);
      }

      if (!res.ok) {
        throw new Error(`Introspection request failed (${res.status})`);
      }

      return parseIntrospectionResult(json);
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

  return { introspect };
}
