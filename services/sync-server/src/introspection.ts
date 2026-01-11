import type { SyncRole } from "./auth.js";

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

      if (res.status === 401 || res.status === 403) {
        // Token is inactive/invalid (or the introspection endpoint rejected our request).
        const fallbackReason =
          json && typeof json === "object"
            ? typeof (json as any).error === "string" && (json as any).error.length > 0
              ? ((json as any).error as string)
              : typeof (json as any).reason === "string" && (json as any).reason.length > 0
                ? ((json as any).reason as string)
                : "forbidden"
            : "forbidden";

        if (!json) return { active: false, reason: fallbackReason };
        try {
          return parseIntrospectionResult(json);
        } catch {
          return { active: false, reason: fallbackReason };
        }
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
