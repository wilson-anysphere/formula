export interface AppConfig {
  port: number;
  databaseUrl: string;
  sessionCookieName: string;
  sessionTtlSeconds: number;
  cookieSecure: boolean;
  syncTokenSecret: string;
  syncTokenTtlSeconds: number;
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

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const port = parseIntEnv(env.PORT, 3000);
  const databaseUrl = env.DATABASE_URL ?? "postgres://postgres:postgres@localhost:5432/formula";
  const sessionCookieName = env.SESSION_COOKIE_NAME ?? "formula_session";
  const sessionTtlSeconds = parseIntEnv(env.SESSION_TTL_SECONDS, 60 * 60 * 24);
  const cookieSecure = env.COOKIE_SECURE === "true";
  const syncTokenSecret = env.SYNC_TOKEN_SECRET ?? "dev-sync-token-secret-change-me";
  const syncTokenTtlSeconds = parseIntEnv(env.SYNC_TOKEN_TTL_SECONDS, 60 * 5);
  const syncServerInternalUrl = env.SYNC_SERVER_INTERNAL_URL;
  const syncServerInternalAdminToken = env.SYNC_SERVER_INTERNAL_ADMIN_TOKEN;
  const retentionSweepIntervalMs =
    env.RETENTION_SWEEP_INTERVAL_MS === "0"
      ? null
      : parseIntEnv(env.RETENTION_SWEEP_INTERVAL_MS, 60 * 60 * 1000);
  const internalAdminToken = env.INTERNAL_ADMIN_TOKEN;

  return {
    port,
    databaseUrl,
    sessionCookieName,
    sessionTtlSeconds,
    cookieSecure,
    syncTokenSecret,
    syncTokenTtlSeconds,
    syncServerInternalUrl,
    syncServerInternalAdminToken,
    retentionSweepIntervalMs,
    internalAdminToken
  };
}
