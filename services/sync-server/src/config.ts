import { readFileSync } from "node:fs";
import path from "node:path";

import { KeyRing } from "../../../packages/security/crypto/keyring.js";

export type AuthMode =
  | {
      mode: "opaque";
      token: string;
    }
  | {
      mode: "jwt-hs256";
      secret: string;
      issuer?: string;
      audience?: string;
    }
  | {
      mode: "introspect";
      /**
       * Base URL for the API internal endpoint (e.g. https://api.internal.example.com).
       */
      url: string;
      /**
       * Shared secret matching API `INTERNAL_ADMIN_TOKEN` (sent as `x-internal-admin-token`).
       */
      token: string;
      cacheMs: number;
      /**
       * If true, allow connections when the introspection endpoint is unavailable.
       *
       * This is only honored in non-production environments.
       */
      failOpen: boolean;
    };

export type SyncServerConfig = {
  host: string;
  port: number;
  trustProxy: boolean;
  gc: boolean;
  /**
   * Optional allowlist of accepted WebSocket `Origin` header values.
   *
   * When set, browser-originated websocket upgrades (requests that include an
   * `Origin` header) must match one of these values. Requests without an Origin
   * header are always allowed to support non-browser clients.
   *
   * When unset/empty, all origins are allowed (current behavior).
   */
  allowedOrigins: string[] | null;
  tls:
    | {
        certPath: string;
        keyPath: string;
      }
    | null;

  metrics: {
    /**
     * Whether to expose `/metrics` publicly. `/internal/metrics` remains protected
     * by `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`.
     */
    public: boolean;
  };

  dataDir: string;
  disableDataDirLock: boolean;
  persistence: {
    backend: "leveldb" | "file";
    compactAfterUpdates: number;
    leveldbDocNameHashing: boolean;
    encryption:
      | {
          mode: "off";
        }
      | {
          mode: "keyring";
          keyRing: KeyRing;
          strict: boolean;
        };
  };

  auth: AuthMode;
  /**
   * When enabled, the sync-server will enforce JWT range restrictions (fail-closed)
   * for incoming Yjs updates that touch spreadsheet cells.
   */
  enforceRangeRestrictions: boolean;

  /**
   * Optional API-side token introspection for sync JWTs.
   *
   * When enabled, sync-server will call the API internal endpoint on each websocket
   * connection to ensure the issuing session is still active and document
   * permissions haven't been revoked.
   */
  introspection: {
    url: string;
    token: string;
    cacheTtlMs: number;
    /**
     * Maximum number of concurrent in-flight HTTP calls to the introspection
     * endpoint. `0` disables the limit.
     */
    maxConcurrent: number;
  } | null;

  /**
   * Optional shared secret for internal admin endpoints (purge, retention ops, etc).
   *
   * Disabled by default. To enable, set `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`.
   */
  internalAdminToken: string | null;
  retention: {
    ttlMs: number;
    sweepIntervalMs: number;
    tombstoneTtlMs: number;
  };

  limits: {
    maxConnections: number;
    maxConnectionsPerIp: number;
    maxConnectionsPerDoc: number;
    maxConnAttemptsPerWindow: number;
    connAttemptWindowMs: number;
    maxMessageBytes: number;
    maxMessagesPerWindow: number;
    messageWindowMs: number;
    maxMessagesPerIpWindow: number;
    ipMessageWindowMs: number;
    maxAwarenessStateBytes: number;
    maxAwarenessEntries: number;
    maxMessagesPerDocWindow: number;
    docMessageWindowMs: number;
    maxBranchingCommitsPerDoc: number;
    maxVersionsPerDoc: number;
  };

  logLevel: string;
};

function envBool(value: string | undefined, defaultValue: boolean): boolean {
  if (value === undefined) return defaultValue;
  return value === "1" || value.toLowerCase() === "true";
}

function envInt(value: string | undefined, defaultValue: number): number {
  if (value === undefined || value === "") return defaultValue;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : defaultValue;
}

function loadKeyRingFromEnv(): KeyRing {
  const keyRingJsonEnv = process.env.SYNC_SERVER_ENCRYPTION_KEYRING_JSON;
  const keyRingPath = process.env.SYNC_SERVER_ENCRYPTION_KEYRING_PATH;
  const keyBase64Env = process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64;

  let json: string | null = null;
  if (keyRingJsonEnv && keyRingJsonEnv.trim().length > 0) {
    json = keyRingJsonEnv;
  } else if (keyRingPath && keyRingPath.trim().length > 0) {
    try {
      json = readFileSync(keyRingPath, "utf8");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      throw new Error(
        `Failed to read KeyRing JSON from SYNC_SERVER_ENCRYPTION_KEYRING_PATH (${keyRingPath}): ${message}`
      );
    }
  } else if (keyBase64Env && keyBase64Env.trim().length > 0) {
    const keyBase64 = keyBase64Env.trim();
    const decoded = Buffer.from(keyBase64, "base64");
    if (decoded.byteLength !== 32) {
      throw new Error(
        "SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64 must be a base64-encoded 32-byte (256-bit) key."
      );
    }
    return new KeyRing({
      currentVersion: 1,
      keysByVersion: new Map([[1, decoded]]),
    });
  }

  if (!json) {
    throw new Error(
      "SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring requires SYNC_SERVER_ENCRYPTION_KEYRING_JSON, SYNC_SERVER_ENCRYPTION_KEYRING_PATH, or SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64."
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch (err) {
    throw new Error(
      `Invalid KeyRing JSON (SYNC_SERVER_ENCRYPTION_KEYRING_JSON/SYNC_SERVER_ENCRYPTION_KEYRING_PATH): ${String(err)}`
    );
  }

  return KeyRing.fromJSON(parsed);
}

export function loadConfigFromEnv(): SyncServerConfig {
  const nodeEnv = process.env.NODE_ENV ?? "development";

  const host = process.env.SYNC_SERVER_HOST ?? "127.0.0.1";
  const port = envInt(process.env.SYNC_SERVER_PORT, 1234);
  if (port < 0 || port > 65_535) {
    throw new Error(
      `SYNC_SERVER_PORT must be an integer between 0 and 65535 (got ${port}).`
    );
  }
  const trustProxy = envBool(process.env.SYNC_SERVER_TRUST_PROXY, false);
  const gc = envBool(process.env.SYNC_SERVER_GC, true);

  const disablePublicMetrics = envBool(process.env.SYNC_SERVER_DISABLE_PUBLIC_METRICS, false);

  const allowedOriginsEnv = process.env.SYNC_SERVER_ALLOWED_ORIGINS;
  const allowedOrigins = (() => {
    const raw = allowedOriginsEnv?.trim() ?? "";
    if (raw.length === 0) return null;
    const parts = raw
      .split(",")
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
    if (parts.length === 0) return null;
    return [...new Set(parts)];
  })();

  const enforceRangeRestrictionsDefault = nodeEnv === "production";
  const enforceRangeRestrictions = envBool(
    process.env.SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS,
    enforceRangeRestrictionsDefault
  );
  const tlsCertPath = process.env.SYNC_SERVER_TLS_CERT_PATH?.trim() ?? "";
  const tlsKeyPath = process.env.SYNC_SERVER_TLS_KEY_PATH?.trim() ?? "";
  const tls =
    tlsCertPath.length > 0 || tlsKeyPath.length > 0
      ? (() => {
          if (tlsCertPath.length === 0 || tlsKeyPath.length === 0) {
            throw new Error(
              "SYNC_SERVER_TLS_CERT_PATH and SYNC_SERVER_TLS_KEY_PATH must both be set to enable TLS."
            );
          }
           return { certPath: tlsCertPath, keyPath: tlsKeyPath };
         })()
       : null;

  const dataDir =
    process.env.SYNC_SERVER_DATA_DIR ??
    path.resolve(process.cwd(), ".sync-server-data");

  const disableDataDirLock = envBool(
    process.env.SYNC_SERVER_DISABLE_DATA_DIR_LOCK,
    false
  );

  const backendEnv = process.env.SYNC_SERVER_PERSISTENCE_BACKEND ?? "leveldb";
  const backend =
    backendEnv === "file" || backendEnv === "leveldb" ? backendEnv : "leveldb";

  const compactAfterUpdates = envInt(
    process.env.SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES,
    200
  );

  const encryptionEnv = process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION ?? "off";
  const encryptionKeyBase64 = process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64;
  const encryptionMode =
    encryptionEnv === "keyring" || encryptionKeyBase64
      ? "keyring"
      : encryptionEnv === "off"
        ? "off"
        : "off";
  const encryptionStrictDefault = nodeEnv === "production";
  const encryptionStrict = envBool(
    process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT,
    encryptionStrictDefault
  );
  const encryption =
    encryptionMode === "keyring"
      ? {
          mode: "keyring" as const,
          keyRing: loadKeyRingFromEnv(),
          strict: encryptionStrict,
        }
      : { mode: "off" as const };

  const leveldbDocNameHashing = envBool(
    process.env.SYNC_SERVER_LEVELDB_DOCNAME_HASHING,
    false
  );
  const retentionTtlMs = envInt(process.env.SYNC_SERVER_RETENTION_TTL_MS, 0);
  const retentionSweepIntervalMs = envInt(
    process.env.SYNC_SERVER_RETENTION_SWEEP_INTERVAL_MS,
    0
  );

  const authModeEnv = process.env.SYNC_SERVER_AUTH_MODE?.trim();
  const opaqueToken = process.env.SYNC_SERVER_AUTH_TOKEN;
  const jwtSecret = process.env.SYNC_SERVER_JWT_SECRET;
  const introspectUrl = process.env.SYNC_SERVER_INTROSPECT_URL;
  const introspectToken = process.env.SYNC_SERVER_INTROSPECT_TOKEN;
  const introspectCacheMs = envInt(process.env.SYNC_SERVER_INTROSPECT_CACHE_MS, 30_000);
  const introspectFailOpenEnv = envBool(process.env.SYNC_SERVER_INTROSPECT_FAIL_OPEN, false);
  const introspectFailOpen = nodeEnv === "production" ? false : introspectFailOpenEnv;

  const internalAdminTokenEnv = process.env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN;
  const internalAdminToken =
    internalAdminTokenEnv && internalAdminTokenEnv.length > 0
      ? internalAdminTokenEnv
      : null;

  const defaultTombstoneTtlMs = 7 * 24 * 60 * 60 * 1000;
  const tombstoneTtlMs =
    process.env.SYNC_SERVER_TOMBSTONE_TTL_MS !== undefined &&
    process.env.SYNC_SERVER_TOMBSTONE_TTL_MS !== ""
      ? envInt(process.env.SYNC_SERVER_TOMBSTONE_TTL_MS, defaultTombstoneTtlMs)
      : retentionTtlMs > 0
        ? retentionTtlMs
        : defaultTombstoneTtlMs;

  let auth: AuthMode;
  if (authModeEnv === "introspect") {
    if (!introspectUrl || introspectUrl.trim().length === 0) {
      throw new Error("SYNC_SERVER_AUTH_MODE=introspect requires SYNC_SERVER_INTROSPECT_URL");
    }
    if (!introspectToken || introspectToken.trim().length === 0) {
      throw new Error("SYNC_SERVER_AUTH_MODE=introspect requires SYNC_SERVER_INTROSPECT_TOKEN");
    }
    auth = {
      mode: "introspect",
      url: introspectUrl,
      token: introspectToken,
      cacheMs: introspectCacheMs,
      failOpen: introspectFailOpen,
    };
  } else if (authModeEnv === "opaque") {
    if (!opaqueToken) {
      throw new Error("SYNC_SERVER_AUTH_MODE=opaque requires SYNC_SERVER_AUTH_TOKEN");
    }
    auth = { mode: "opaque", token: opaqueToken };
  } else if (authModeEnv === "jwt-hs256" || authModeEnv === "jwt") {
    if (!jwtSecret) {
      throw new Error("SYNC_SERVER_AUTH_MODE=jwt-hs256 requires SYNC_SERVER_JWT_SECRET");
    }
    auth = {
      mode: "jwt-hs256",
      secret: jwtSecret,
      issuer: process.env.SYNC_SERVER_JWT_ISSUER,
      audience: process.env.SYNC_SERVER_JWT_AUDIENCE ?? "formula-sync",
    };
  } else if (opaqueToken) {
    auth = { mode: "opaque", token: opaqueToken };
  } else if (jwtSecret) {
    auth = {
      mode: "jwt-hs256",
      secret: jwtSecret,
      issuer: process.env.SYNC_SERVER_JWT_ISSUER,
      audience: process.env.SYNC_SERVER_JWT_AUDIENCE ?? "formula-sync",
    };
  } else {
    if (nodeEnv === "production") {
      throw new Error(
        "Auth is required in production. Set SYNC_SERVER_AUTH_TOKEN, SYNC_SERVER_JWT_SECRET, or SYNC_SERVER_AUTH_MODE=introspect."
      );
    }
    // Dev default: force auth (still), but with a known token.
    auth = { mode: "opaque", token: "dev-token" };
  }

  const introspectionUrlRaw = process.env.SYNC_SERVER_INTROSPECTION_URL?.trim();
  const introspectionTokenRaw = process.env.SYNC_SERVER_INTROSPECTION_TOKEN?.trim();
  const introspectionCacheTtlMs = envInt(process.env.SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS, 15_000);
  const introspectionMaxConcurrent = Math.max(
    0,
    envInt(process.env.SYNC_SERVER_INTROSPECTION_MAX_CONCURRENT, 50)
  );

  let introspection: SyncServerConfig["introspection"] = null;
  if (introspectionUrlRaw) {
    if (!introspectionTokenRaw) {
      throw new Error(
        "SYNC_SERVER_INTROSPECTION_URL requires SYNC_SERVER_INTROSPECTION_TOKEN (shared secret for API internal endpoints)."
      );
    }

    // Allow operators to pass either a base API URL ("https://api.internal") or
    // the full endpoint URL ("https://api.internal/internal/sync/introspect").
    const url = introspectionUrlRaw.includes("/internal/sync/introspect")
      ? introspectionUrlRaw
      : new URL("/internal/sync/introspect", introspectionUrlRaw).toString();

    introspection = {
      url,
      token: introspectionTokenRaw,
      cacheTtlMs: introspectionCacheTtlMs,
      maxConcurrent: introspectionMaxConcurrent,
    };
  }

  // Reserved history quota defaults:
  //
  // The sync-server can optionally store versioning/branching metadata directly in the Y.Doc
  // (roots like `versions` and `branching:commits`). When operators disable the reserved-root
  // guard (see `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED` in `src/server.ts`), clients are allowed
  // to write to those roots.
  //
  // In development/test, default these limits to 0 (disabled) for convenience.
  // In production, default to conservative non-zero limits to prevent unbounded history growth
  // if the reserved-root guard is disabled.
  //
  // Tune (or disable by setting to 0) via:
  // - SYNC_SERVER_MAX_VERSIONS_PER_DOC
  // - SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC
  const defaultMaxVersionsPerDoc = nodeEnv === "production" ? 500 : 0;
  const defaultMaxBranchingCommitsPerDoc = nodeEnv === "production" ? 5_000 : 0;

  return {
    host,
    port,
    trustProxy,
    gc,
    allowedOrigins,
    tls,
    metrics: {
      public: !disablePublicMetrics,
    },
    dataDir,
    disableDataDirLock,
    persistence: {
      backend,
      compactAfterUpdates,
      leveldbDocNameHashing,
      encryption,
    },
    auth,
    enforceRangeRestrictions,
    introspection,
    internalAdminToken,
    retention: {
      ttlMs: retentionTtlMs,
      sweepIntervalMs: retentionSweepIntervalMs,
      tombstoneTtlMs,
    },
    limits: {
      maxConnections: envInt(process.env.SYNC_SERVER_MAX_CONNECTIONS, 1000),
      maxConnectionsPerIp: envInt(
        process.env.SYNC_SERVER_MAX_CONNECTIONS_PER_IP,
        25
      ),
      maxConnectionsPerDoc: envInt(process.env.SYNC_SERVER_MAX_CONNECTIONS_PER_DOC, 0),
      maxConnAttemptsPerWindow: envInt(
        process.env.SYNC_SERVER_MAX_CONN_ATTEMPTS_PER_WINDOW,
        60
      ),
      connAttemptWindowMs: envInt(
        process.env.SYNC_SERVER_CONN_ATTEMPT_WINDOW_MS,
        60_000
      ),
      maxMessagesPerWindow: envInt(
        process.env.SYNC_SERVER_MAX_MESSAGES_PER_WINDOW,
        2_000
      ),
      messageWindowMs: envInt(
        process.env.SYNC_SERVER_MESSAGE_WINDOW_MS,
        10_000
      ),
      maxMessagesPerIpWindow: envInt(
        process.env.SYNC_SERVER_MAX_MESSAGES_PER_IP_WINDOW,
        0
      ),
      ipMessageWindowMs: envInt(
        process.env.SYNC_SERVER_IP_MESSAGE_WINDOW_MS,
        0
      ),
      maxMessageBytes: envInt(
        process.env.SYNC_SERVER_MAX_MESSAGE_BYTES,
        2 * 1024 * 1024
      ),
      maxAwarenessStateBytes: envInt(
        process.env.SYNC_SERVER_MAX_AWARENESS_STATE_BYTES,
        64 * 1024
      ),
      maxAwarenessEntries: envInt(process.env.SYNC_SERVER_MAX_AWARENESS_ENTRIES, 10),
      maxMessagesPerDocWindow: envInt(
        process.env.SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW,
        10_000
      ),
      docMessageWindowMs: envInt(
        process.env.SYNC_SERVER_DOC_MESSAGE_WINDOW_MS,
        10_000
      ),
      maxBranchingCommitsPerDoc: Math.max(
        0,
        envInt(
          process.env.SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC,
          defaultMaxBranchingCommitsPerDoc
        )
      ),
      maxVersionsPerDoc: Math.max(
        0,
        envInt(process.env.SYNC_SERVER_MAX_VERSIONS_PER_DOC, defaultMaxVersionsPerDoc)
      ),
    },
    logLevel: process.env.LOG_LEVEL ?? "info",
  };
}
