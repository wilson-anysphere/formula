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
   * Publicly reachable base URL for this API service (scheme + host [+ optional path prefix]).
   *
   * This is used for security-sensitive flows that need to generate absolute URLs,
   * such as OIDC `redirect_uri`. Production deployments must set this explicitly
   * to avoid host-header / forwarded-header spoofing.
   *
   * Environment variables:
   * - `PUBLIC_BASE_URL` (preferred)
   * - `EXTERNAL_BASE_URL` (legacy alias)
   */
  publicBaseUrl?: string;
  /**
   * When `publicBaseUrl` is not configured (dev/test), the server may fall back
   * to deriving the external base URL from request headers, but only when:
   * - Fastify has `trustProxy=true`, and
   * - the derived host is in this allowlist.
   *
   * Env: `PUBLIC_BASE_URL_HOST_ALLOWLIST` (comma-separated).
   */
  publicBaseUrlHostAllowlist: string[];
  /**
   * Comma-separated allowlist of allowed CORS origins.
   *
   * - In production, defaults to no allowed origins unless explicitly configured.
   * - In dev/test, defaults to common localhost origins for the web UI.
   */
  corsAllowedOrigins: string[];
  /**
   * Whether to send `Access-Control-Allow-Credentials: true` for allowed origins.
   *
   * Defaults to true because the API supports cookie-based sessions.
   *
   * Env: `CORS_ALLOW_CREDENTIALS` (true/false)
   */
  corsAllowCredentials?: boolean;
  syncTokenSecret: string;
  syncTokenTtlSeconds: number;
  /**
   * Keyring used to encrypt values stored in the database-backed secret store
   * (`secrets` table).
   *
   * Supported environment variables (highest priority first):
   *
   * - `SECRET_STORE_KEYS_JSON`: JSON object containing `{ currentKeyId, keys }`,
   *   where `keys` is a map of keyId -> base64-encoded 32-byte AES key.
   * - `SECRET_STORE_KEYS`: comma-separated list of `<keyId>:<base64>` entries.
   *   The **last** key id is treated as current for encryption; all keys are
   *   valid for decryption.
   * - `SECRET_STORE_KEY`: legacy single secret (hashed with SHA-256 to derive a
   *   32-byte AES-256 key). Still supported for smooth upgrades.
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
   * Legacy secret used to decrypt historical (envelope schema v1) rows that were
   * encrypted using the previous HKDF-based local KMS model.
   *
   * The canonical local KMS provider persists versioned KEKs in Postgres
   * (`org_kms_local_state`) and does not require this secret for new writes.
   *
   * Note: This remains configurable so existing deployments can migrate without
   * data loss.
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
   * Interval for deleting stale auth state rows (`oidc_auth_states`, `saml_auth_states`)
   * plus SAML request cache entries (`saml_request_cache`).
   *
   * If null, cleanup is disabled.
   *
   * Env: `OIDC_AUTH_STATE_CLEANUP_INTERVAL_MS` (set to `0` to disable)
   */
  oidcAuthStateCleanupIntervalMs: number | null;
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

function readBooleanEnv(value: string | undefined, fallback: boolean): boolean {
  if (typeof value !== "string") return fallback;
  const trimmed = value.trim().toLowerCase();
  if (trimmed === "true") return true;
  if (trimmed === "false") return false;
  return fallback;
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

function parseHostAllowlist(value: string | undefined): string[] {
  const raw = typeof value === "string" ? value.trim() : "";
  if (!raw) return ["localhost", "127.0.0.1", "[::1]"];
  const parsed = raw
    .split(",")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
  return parsed.length > 0 ? parsed : ["localhost", "127.0.0.1", "[::1]"];
}

function normalizePublicBaseUrl(value: string): string {
  const raw = value.trim();
  if (raw.length === 0) throw new Error("PUBLIC_BASE_URL must be non-empty");

  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error("PUBLIC_BASE_URL must be a valid URL");
  }

  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("PUBLIC_BASE_URL must start with http:// or https://");
  }
  if (url.username || url.password) {
    throw new Error("PUBLIC_BASE_URL must not include credentials");
  }

  // Avoid surprising redirects based on query parameters / fragments.
  url.search = "";
  url.hash = "";
  // Strip trailing slashes so URL joining is stable.
  url.pathname = url.pathname.replace(/\/+$/, "");

  const pathname = url.pathname === "/" ? "" : url.pathname;
  return `${url.origin}${pathname}`;
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
    if (currentKeyId.includes(":")) {
      throw new Error("SECRET_STORE_KEYS_JSON.currentKeyId must not contain ':'");
    }
    if (!keysValue || typeof keysValue !== "object" || Array.isArray(keysValue)) {
      throw new Error("SECRET_STORE_KEYS_JSON.keys must be an object mapping keyId -> base64 key");
    }

    const keys: Record<string, Buffer> = {};
    for (const [keyId, value] of Object.entries(keysValue as Record<string, unknown>)) {
      if (keyId.length === 0) {
        throw new Error("SECRET_STORE_KEYS_JSON.keys must not contain empty key ids");
      }
      if (keyId.includes(":")) {
        throw new Error(`SECRET_STORE_KEYS_JSON.keys.${keyId} must not contain ':'`);
      }
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

  const rawKeys = typeof env.SECRET_STORE_KEYS === "string" ? env.SECRET_STORE_KEYS.trim() : "";
  if (rawKeys.length > 0) {
    const parts = rawKeys
      .split(",")
      .map((part) => part.trim())
      .filter((part) => part.length > 0);

    if (parts.length === 0) {
      throw new Error("SECRET_STORE_KEYS must contain at least one <keyId>:<base64> entry");
    }

    const keys: Record<string, Buffer> = {};
    let currentKeyId: string | null = null;

    for (const entry of parts) {
      const idx = entry.indexOf(":");
      if (idx <= 0) {
        throw new Error(`SECRET_STORE_KEYS entries must be in the form <keyId>:<base64> (got "${entry}")`);
      }
      const keyId = entry.slice(0, idx).trim();
      const value = entry.slice(idx + 1).trim();
      if (!keyId) {
        throw new Error(`SECRET_STORE_KEYS entry is missing keyId (got "${entry}")`);
      }
      if (keyId.includes(":")) {
        throw new Error(`SECRET_STORE_KEYS keyId must not contain ':' (got "${keyId}")`);
      }
      if (!value) {
        throw new Error(`SECRET_STORE_KEYS entry is missing base64 key for keyId=${keyId}`);
      }
      if (keys[keyId]) {
        throw new Error(`SECRET_STORE_KEYS contains duplicate keyId=${keyId}`);
      }

      const raw = Buffer.from(value, "base64");
      if (raw.byteLength !== 32) {
        throw new Error(`SECRET_STORE_KEYS keyId=${keyId} must decode to 32 bytes (got ${raw.byteLength})`);
      }
      keys[keyId] = raw;
      currentKeyId = keyId; // last entry is current
    }

    if (!currentKeyId) {
      throw new Error("SECRET_STORE_KEYS must contain at least one key");
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

  const publicBaseUrlEnv = readStringEnv(env.PUBLIC_BASE_URL ?? env.EXTERNAL_BASE_URL, "");
  const publicBaseUrl =
    publicBaseUrlEnv.length > 0
      ? normalizePublicBaseUrl(publicBaseUrlEnv)
      : undefined;
  const publicBaseUrlHostAllowlist = parseHostAllowlist(env.PUBLIC_BASE_URL_HOST_ALLOWLIST);

  const corsAllowedOrigins = parseCorsAllowedOrigins(env.CORS_ALLOWED_ORIGINS, nodeEnv);
  const corsAllowCredentials = readBooleanEnv(env.CORS_ALLOW_CREDENTIALS, true);
  const syncTokenSecret = readStringEnv(env.SYNC_TOKEN_SECRET, DEV_SYNC_TOKEN_SECRET);
  const syncTokenTtlSeconds = parseIntEnv(env.SYNC_TOKEN_TTL_SECONDS, 60 * 5);

  const secretStoreKey = readStringEnv(env.SECRET_STORE_KEY, DEV_SECRET_STORE_KEY);
  const secretStoreKeys = loadSecretStoreKeys(env, secretStoreKey);

  const syncServerInternalUrl = readStringEnv(env.SYNC_SERVER_INTERNAL_URL, "");
  const syncServerInternalAdminToken = readStringEnv(env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN, "");
  // Legacy-only: required to decrypt/migrate historical envelope schema v1 rows.
  // New deployments using the canonical DB-backed local KMS do not need it.
  const localKmsMasterKey = readStringEnv(env.LOCAL_KMS_MASTER_KEY, "");
  const awsKmsEnabled = env.AWS_KMS_ENABLED === "true";
  const awsRegion = readStringEnv(env.AWS_REGION, "");
  const retentionSweepIntervalMs =
    env.RETENTION_SWEEP_INTERVAL_MS === "0"
      ? null
      : parseIntEnv(env.RETENTION_SWEEP_INTERVAL_MS, 60 * 60 * 1000);
  const oidcAuthStateCleanupIntervalMs =
    env.OIDC_AUTH_STATE_CLEANUP_INTERVAL_MS === "0"
      ? null
      : parseIntEnv(env.OIDC_AUTH_STATE_CLEANUP_INTERVAL_MS, 60 * 1000);
  const internalAdminToken = readStringEnv(env.INTERNAL_ADMIN_TOKEN, "");

  const config: AppConfig = {
    port,
    databaseUrl,
    sessionCookieName,
    sessionTtlSeconds,
    cookieSecure,
    trustProxy,
    publicBaseUrl,
    publicBaseUrlHostAllowlist,
    corsAllowedOrigins,
    corsAllowCredentials,
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
    oidcAuthStateCleanupIntervalMs,
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
    const devSecretStoreKey = deriveSecretStoreKey(DEV_SECRET_STORE_KEY);
    const usingDevSecretStoreKey = Object.values(config.secretStoreKeys.keys).some(
      (key) => Buffer.compare(key, devSecretStoreKey) === 0
    );
    if (usingDevSecretStoreKey) {
      invalidSecrets.push(secretStoreKeySource(env));
    }
    const rawLocalKmsMasterKey =
      typeof env.LOCAL_KMS_MASTER_KEY === "string" ? env.LOCAL_KMS_MASTER_KEY.trim() : "";
    if (rawLocalKmsMasterKey === DEV_LOCAL_KMS_MASTER_KEY) invalidSecrets.push("LOCAL_KMS_MASTER_KEY");
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

function secretStoreKeySource(env: NodeJS.ProcessEnv): string {
  const rawJson = typeof env.SECRET_STORE_KEYS_JSON === "string" ? env.SECRET_STORE_KEYS_JSON.trim() : "";
  if (rawJson.length > 0) return "SECRET_STORE_KEYS_JSON";
  const rawKeys = typeof env.SECRET_STORE_KEYS === "string" ? env.SECRET_STORE_KEYS.trim() : "";
  if (rawKeys.length > 0) return "SECRET_STORE_KEYS";
  return "SECRET_STORE_KEY";
}
