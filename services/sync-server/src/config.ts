import path from "node:path";

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
  persistence: {
    backend: "leveldb" | "file";
    compactAfterUpdates: number;
  };

  auth: AuthMode;

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

export function loadConfigFromEnv(): SyncServerConfig {
  const nodeEnv = process.env.NODE_ENV ?? "development";

  const host = process.env.SYNC_SERVER_HOST ?? "127.0.0.1";
  const port = envInt(process.env.SYNC_SERVER_PORT, 1234);
  const trustProxy = envBool(process.env.SYNC_SERVER_TRUST_PROXY, false);
  const gc = envBool(process.env.SYNC_SERVER_GC, true);

  const dataDir =
    process.env.SYNC_SERVER_DATA_DIR ??
    path.resolve(process.cwd(), ".sync-server-data");

  const backendEnv = process.env.SYNC_SERVER_PERSISTENCE_BACKEND ?? "leveldb";
  const backend =
    backendEnv === "file" || backendEnv === "leveldb" ? backendEnv : "leveldb";

  const compactAfterUpdates = envInt(
    process.env.SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES,
    200
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
    persistence: { backend, compactAfterUpdates },
    auth,
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
