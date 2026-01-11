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
    };

export type SyncServerConfig = {
  host: string;
  port: number;
  trustProxy: boolean;
  gc: boolean;

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
        };
    /**
     * Optional encryption for y-leveldb persistence values.
     *
     * Note: this only encrypts LevelDB *values*. Keys (including doc ids) remain
     * plaintext unless `SYNC_SERVER_LEVELDB_DOCNAME_HASHING` is enabled.
     */
    leveldbEncryption?: {
      key: Buffer;
      strict: boolean;
    };
  };

  auth: AuthMode;

  /**
   * Optional shared secret for internal admin endpoints (purge, retention ops, etc).
   *
   * Disabled by default. To enable, set `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`.
   * For convenience in multi-service deployments, `INTERNAL_ADMIN_TOKEN` is also
   * accepted as a fallback unless `SYNC_SERVER_INTERNAL_ADMIN_TOKEN` is set
   * (even to an empty string).
   */
  internalAdminToken: string | null;
  retention: {
    ttlMs: number;
    sweepIntervalMs: number;
  };

  limits: {
    maxConnections: number;
    maxConnectionsPerIp: number;
    maxConnAttemptsPerWindow: number;
    connAttemptWindowMs: number;
    maxMessagesPerWindow: number;
    messageWindowMs: number;
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

  let json: string | null = null;
  if (keyRingJsonEnv && keyRingJsonEnv.trim().length > 0) {
    json = keyRingJsonEnv;
  } else if (keyRingPath && keyRingPath.trim().length > 0) {
    json = readFileSync(keyRingPath, "utf8");
  }

  if (!json) {
    throw new Error(
      "SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring requires SYNC_SERVER_ENCRYPTION_KEYRING_JSON or SYNC_SERVER_ENCRYPTION_KEYRING_PATH."
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
  const trustProxy = envBool(process.env.SYNC_SERVER_TRUST_PROXY, false);
  const gc = envBool(process.env.SYNC_SERVER_GC, true);

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
  const encryptionMode =
    encryptionEnv === "keyring" || encryptionEnv === "off" ? encryptionEnv : "off";
  const encryption =
    encryptionMode === "keyring"
      ? {
          mode: "keyring" as const,
          keyRing: loadKeyRingFromEnv(),
        }
      : { mode: "off" as const };

  const leveldbEncryptionKeyB64 =
    process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64;
  const leveldbEncryptionStrictDefault = nodeEnv === "production";
  const leveldbEncryptionStrict = envBool(
    process.env.SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT,
    leveldbEncryptionStrictDefault
  );
  const leveldbEncryptionKey =
    leveldbEncryptionKeyB64 === undefined
      ? null
      : Buffer.from(leveldbEncryptionKeyB64, "base64");
  if (leveldbEncryptionKey !== null && leveldbEncryptionKey.byteLength !== 32) {
    throw new Error(
      "SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64 must be a base64-encoded 32-byte (256-bit) key."
    );
  }
  const leveldbEncryption =
    leveldbEncryptionKey === null
      ? undefined
      : {
          key: leveldbEncryptionKey,
          strict: leveldbEncryptionStrict,
        };

  const leveldbDocNameHashing = envBool(
    process.env.SYNC_SERVER_LEVELDB_DOCNAME_HASHING,
    false
  );

  const internalAdminToken =
    process.env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN !== undefined
      ? process.env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN || null
      : process.env.INTERNAL_ADMIN_TOKEN || null;

  const retentionTtlMs = envInt(process.env.SYNC_SERVER_RETENTION_TTL_MS, 0);
  const retentionSweepIntervalMs = envInt(
    process.env.SYNC_SERVER_RETENTION_SWEEP_INTERVAL_MS,
    0
  );

  const opaqueToken = process.env.SYNC_SERVER_AUTH_TOKEN;
  const jwtSecret = process.env.SYNC_SERVER_JWT_SECRET;

  let auth: AuthMode;
  if (opaqueToken) {
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
        "Auth is required in production. Set SYNC_SERVER_AUTH_TOKEN or SYNC_SERVER_JWT_SECRET."
      );
    }
    // Dev default: force auth (still), but with a known token.
    auth = { mode: "opaque", token: "dev-token" };
  }

  return {
    host,
    port,
    trustProxy,
    gc,
    dataDir,
    disableDataDirLock,
    persistence: {
      backend,
      compactAfterUpdates,
      leveldbDocNameHashing,
      encryption,
      leveldbEncryption,
    },
    auth,
    internalAdminToken,
    retention: {
      ttlMs: retentionTtlMs,
      sweepIntervalMs: retentionSweepIntervalMs,
    },
    limits: {
      maxConnections: envInt(process.env.SYNC_SERVER_MAX_CONNECTIONS, 1000),
      maxConnectionsPerIp: envInt(
        process.env.SYNC_SERVER_MAX_CONNECTIONS_PER_IP,
        25
      ),
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
    },
    logLevel: process.env.LOG_LEVEL ?? "info",
  };
}
