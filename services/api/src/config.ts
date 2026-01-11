import { deriveSecretStoreKey, type SecretStoreKeyring } from "./secrets/secretStore";

export interface AppConfig {
  port: number;
  databaseUrl: string;
  sessionCookieName: string;
  sessionTtlSeconds: number;
  cookieSecure: boolean;
  /**
   * Whether to trust reverse proxy headers (e.g. `X-Forwarded-For`) when deriving
   * `request.ip`. Keep this disabled unless the API is behind a trusted proxy
   * that strips spoofed forwarding headers.
   */
  trustProxy?: boolean;
  /**
   * Canonical external base URL (origin) for the API (e.g. `https://api.example.com`).
   *
   * This is used for security-sensitive flows that need to generate absolute URLs,
   * such as OIDC `redirect_uri`. Production deployments must set this explicitly
   * to avoid host-header / forwarded-header spoofing.
   */
  publicBaseUrl?: string;
  /**
   * Comma-separated allowlist of allowed CORS origins.
   *
   * - In production, defaults to no allowed origins unless explicitly configured.
   * - In dev/test, defaults to common localhost origins for the web UI.
   */
  corsAllowedOrigins: string[];
  syncTokenSecret: string;
  syncTokenTtlSeconds: number;
  /**
   * Keyring used to encrypt values stored in the database-backed secret store
   * (`secrets` table).
   *
   * Production deployments should use `SECRET_STORE_KEYS_JSON` so keys can be
   * rotated without downtime. `SECRET_STORE_KEY` remains supported as a legacy
   * single-key mode.
   */
  secretStoreKeys: SecretStoreKeyring;
  /**
   * Internal base URL for sync-server, used to purge persisted CRDT state when
   * documents are hard-deleted by retention policy.
   *
   * If unset (or if `syncServerInternalAdminToken` is unset), sync purge
   * integration is disabled.
   */
  syncServerInternalUrl?: string;
  /**
   * Shared secret for sync-server internal endpoints.
   *
   * If unset (or if `syncServerInternalUrl` is unset), sync purge integration is
   * disabled.
   */
  syncServerInternalAdminToken?: string;
  /**
   * Master secret for the LocalKmsProvider (dev/test).
   *
   * In production, set this to a high-entropy value and/or use a real KMS
   * provider (aws/gcp/azure).
   */
  localKmsMasterKey: string;
  /**
   * Enable AWS KMS provider support (requires @aws-sdk/client-kms).
   */
  awsKmsEnabled: boolean;
  /**
   * AWS region to use for KMS operations when awsKmsEnabled=true.
   */
  awsRegion?: string;
  /**
   * If null, retention sweeps are disabled.
   */
  retentionSweepIntervalMs: number | null;
  /**
   * Optional shared secret for internal endpoints (retention sweeps, etc).
   * If unset, internal endpoints are disabled.
   */
  internalAdminToken?: string;
}

function parseIntEnv(value: string | undefined, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return parsed;
}

const DEV_SYNC_TOKEN_SECRET = "dev-sync-token-secret-change-me";
const DEV_SECRET_STORE_KEY = "dev-secret-store-key-change-me";
const DEV_LOCAL_KMS_MASTER_KEY = "dev-local-kms-master-key-change-me";

function readStringEnv(value: string | undefined, fallback: string): string {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

function parseCorsAllowedOrigins(value: string | undefined, nodeEnv: string): string[] {
  const raw = typeof value === "string" ? value.trim() : "";
  const entries = raw.length > 0 ? raw.split(",") : [];

  const parsed = entries
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
    .map((entry) => {
      const url = new URL(entry);
      if (url.origin === "null") throw new Error("CORS origin must not be null");
      return url.origin;
    });

  // Deduplicate after normalization.
  if (parsed.length > 0) return Array.from(new Set(parsed));

  if (nodeEnv === "production") return [];

  return [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "http://localhost:3000",
    "http://127.0.0.1:3000"
  ];
}

function loadSecretStoreKeys(env: NodeJS.ProcessEnv, legacySecret: string): SecretStoreKeyring {
  const rawJson = typeof env.SECRET_STORE_KEYS_JSON === "string" ? env.SECRET_STORE_KEYS_JSON.trim() : "";
  if (rawJson.length > 0) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(rawJson) as unknown;
    } catch (err) {
      throw new Error(
        `SECRET_STORE_KEYS_JSON must be valid JSON (${err instanceof Error ? err.message : "parse failed"})`
      );
    }

    if (!parsed || typeof parsed !== "object") {
      throw new Error("SECRET_STORE_KEYS_JSON must be a JSON object");
    }

    const record = parsed as Record<string, unknown>;
    const currentKeyId = record.currentKeyId;
    const keysValue = record.keys;

    if (typeof currentKeyId !== "string" || currentKeyId.trim().length === 0) {
      throw new Error("SECRET_STORE_KEYS_JSON.currentKeyId must be a non-empty string");
    }
    if (!keysValue || typeof keysValue !== "object" || Array.isArray(keysValue)) {
      throw new Error("SECRET_STORE_KEYS_JSON.keys must be an object mapping keyId -> base64 key");
    }

    const keys: Record<string, Buffer> = {};
    for (const [keyId, value] of Object.entries(keysValue as Record<string, unknown>)) {
      if (typeof value !== "string" || value.trim().length === 0) {
        throw new Error(`SECRET_STORE_KEYS_JSON.keys.${keyId} must be a base64 string`);
      }
      const raw = Buffer.from(value, "base64");
      if (raw.byteLength !== 32) {
        throw new Error(`SECRET_STORE_KEYS_JSON.keys.${keyId} must decode to 32 bytes (got ${raw.byteLength})`);
      }
      keys[keyId] = raw;
    }

    if (!keys[currentKeyId]) {
      throw new Error(`SECRET_STORE_KEYS_JSON currentKeyId=${currentKeyId} missing from keys`);
    }

    return { currentKeyId, keys };
  }

  const legacyKey = deriveSecretStoreKey(legacySecret);
  return { currentKeyId: "legacy", keys: { legacy: legacyKey } };
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const nodeEnv = env.NODE_ENV ?? "development";
  const port = parseIntEnv(env.PORT, 3000);
  const databaseUrl = readStringEnv(env.DATABASE_URL, "postgres://postgres:postgres@localhost:5432/formula");
  const sessionCookieName = readStringEnv(env.SESSION_COOKIE_NAME, "formula_session");
  const sessionTtlSeconds = parseIntEnv(env.SESSION_TTL_SECONDS, 60 * 60 * 24);
  const cookieSecure = env.COOKIE_SECURE === "true";
  const trustProxy = env.TRUST_PROXY === "true";
  const publicBaseUrlEnv = readStringEnv(env.PUBLIC_BASE_URL, "");
  const publicBaseUrl =
    publicBaseUrlEnv.length > 0 ? (() => {
      const url = new URL(publicBaseUrlEnv);
      if (url.origin === "null") throw new Error("PUBLIC_BASE_URL must not be null");
      return url.origin;
    })() : undefined;
  const corsAllowedOrigins = parseCorsAllowedOrigins(env.CORS_ALLOWED_ORIGINS, nodeEnv);
  const syncTokenSecret = readStringEnv(env.SYNC_TOKEN_SECRET, DEV_SYNC_TOKEN_SECRET);
  const syncTokenTtlSeconds = parseIntEnv(env.SYNC_TOKEN_TTL_SECONDS, 60 * 5);

  const secretStoreKey = readStringEnv(env.SECRET_STORE_KEY, DEV_SECRET_STORE_KEY);
  const secretStoreKeys = loadSecretStoreKeys(env, secretStoreKey);

  const syncServerInternalUrl = readStringEnv(env.SYNC_SERVER_INTERNAL_URL, "");
  const syncServerInternalAdminToken = readStringEnv(env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN, "");
  const localKmsMasterKey = readStringEnv(env.LOCAL_KMS_MASTER_KEY, DEV_LOCAL_KMS_MASTER_KEY);
  const awsKmsEnabled = env.AWS_KMS_ENABLED === "true";
  const awsRegion = readStringEnv(env.AWS_REGION, "");
  const retentionSweepIntervalMs =
    env.RETENTION_SWEEP_INTERVAL_MS === "0"
      ? null
      : parseIntEnv(env.RETENTION_SWEEP_INTERVAL_MS, 60 * 60 * 1000);
  const internalAdminToken = readStringEnv(env.INTERNAL_ADMIN_TOKEN, "");

  const config: AppConfig = {
    port,
    databaseUrl,
    sessionCookieName,
    sessionTtlSeconds,
    cookieSecure,
    trustProxy,
    publicBaseUrl,
    corsAllowedOrigins,
    syncTokenSecret,
    syncTokenTtlSeconds,
    secretStoreKeys,
    syncServerInternalUrl: syncServerInternalUrl.length > 0 ? syncServerInternalUrl : undefined,
    syncServerInternalAdminToken:
      syncServerInternalAdminToken.length > 0 ? syncServerInternalAdminToken : undefined,
    localKmsMasterKey,
    awsKmsEnabled,
    awsRegion: awsRegion.length > 0 ? awsRegion : undefined,
    retentionSweepIntervalMs,
    internalAdminToken: internalAdminToken.length > 0 ? internalAdminToken : undefined
  };

  if (nodeEnv === "production") {
    if (!config.publicBaseUrl) {
      throw new Error("Refusing to start in production without PUBLIC_BASE_URL");
    }
    if (!config.publicBaseUrl.startsWith("https://")) {
      throw new Error("Refusing to start in production with PUBLIC_BASE_URL that is not https");
    }

    const invalidSecrets: string[] = [];
    if (config.syncTokenSecret === DEV_SYNC_TOKEN_SECRET) invalidSecrets.push("SYNC_TOKEN_SECRET");
    if (rawJsonIsEmpty(env) && secretStoreKey === DEV_SECRET_STORE_KEY) invalidSecrets.push("SECRET_STORE_KEY");
    if (config.localKmsMasterKey === DEV_LOCAL_KMS_MASTER_KEY) invalidSecrets.push("LOCAL_KMS_MASTER_KEY");
    if (invalidSecrets.length > 0) {
      throw new Error(
        `Refusing to start with default development secrets in production: ${invalidSecrets.join(", ")}`
      );
    }

    if (!config.cookieSecure) {
      throw new Error("Refusing to start in production with COOKIE_SECURE!=true");
    }
  }

  return config;
}

function rawJsonIsEmpty(env: NodeJS.ProcessEnv): boolean {
  const rawJson = typeof env.SECRET_STORE_KEYS_JSON === "string" ? env.SECRET_STORE_KEYS_JSON.trim() : "";
  return rawJson.length === 0;
}

