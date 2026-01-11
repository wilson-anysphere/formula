import { fileScopeExact, httpScope, sqlScope } from "./scopes.js";

/**
 * @typedef {import("./store.js").CredentialStore} CredentialStore
 * @typedef {import("./scopes.js").CredentialScope} CredentialScope
 */

/**
 * @typedef {{
 *   credentialId: string;
 *   id: string;
 *   kind: string;
 *   getSecret: () => Promise<unknown>;
 * }} CredentialHandle
 */

/**
 * @typedef {(args: {
 *   connectorId: string;
 *   scope: CredentialScope;
 *   request: unknown;
 * }) => Promise<unknown | null | undefined>} CredentialPrompt
 */

function parseSqlConnectionScope(connection) {
  if (typeof connection === "string") {
    try {
      const url = new URL(connection);
      const server = url.host;
      const database = url.pathname ? url.pathname.replace(/^\//, "") : null;
      const user = url.username || null;
      return sqlScope({ server, database, user });
    } catch {
      // ignore
    }
  }

  if (connection && typeof connection === "object" && !Array.isArray(connection)) {
    // Some hosts pass a structured connection descriptor with an embedded URL.
    // Prefer parsing it to derive a stable {server, database, user} scope.
    // @ts-ignore - runtime access
    const urlValue = connection.url;
    if (typeof urlValue === "string" && urlValue.length > 0) {
      try {
        const url = new URL(urlValue);
        const server = url.host;
        const database = url.pathname ? url.pathname.replace(/^\//, "") : null;
        const user = url.username || null;
        return sqlScope({ server, database, user });
      } catch {
        // ignore
      }
    }

    // @ts-ignore - runtime access
    const server = connection.server ?? connection.host ?? connection.hostname ?? null;
    if (typeof server === "string" && server.length > 0) {
      // @ts-ignore - runtime access
      const port = connection.port ?? null;
      // @ts-ignore - runtime access
      const database = connection.database ?? connection.db ?? null;
      // @ts-ignore - runtime access
      const user = connection.user ?? connection.username ?? null;
      let serverWithPort = server;
      if (typeof port === "number" && Number.isFinite(port) && port > 0) {
        // Best-effort: append port if provided separately (common for Postgres).
        // Avoid trying to interpret whether `server` already contains a port.
        const needsBrackets = server.includes(":") && !(server.startsWith("[") && server.endsWith("]"));
        const host = needsBrackets ? `[${server}]` : server;
        serverWithPort = `${host}:${port}`;
      }
      return sqlScope({
        server: serverWithPort,
        database: typeof database === "string" ? database : null,
        user: typeof user === "string" ? user : null,
      });
    }
  }

  return null;
}

/**
 * First-class credential manager for Power Query connectors.
 *
 * This is intended to be used as the host application's implementation of
 * `QueryEngine({ onCredentialRequest })`.
 *
 * Credential secrets are connector-specific. For example, the built-in HTTP
 * connector understands:
 * - `{ headers: Record<string,string> }` for legacy header injection
 * - `{ oauth2: { providerId: string, scopes?: string[] } }` for OAuth2 bearer
 *   token handling when an `OAuth2Manager` is configured on `HttpConnector`.
 */
export class CredentialManager {
  /**
   * @param {{
   *   store: CredentialStore;
   *   prompt?: CredentialPrompt;
   * }} opts
   */
  constructor(opts) {
    if (!opts?.store) throw new TypeError("store is required");
    this.store = opts.store;
    this.prompt = opts.prompt ?? null;
  }

  /**
   * Resolve a connector request into a credential scope.
   *
   * @param {string} connectorId
   * @param {any} request
   * @returns {CredentialScope | null}
   */
  resolveScope(connectorId, request) {
    if (!request || typeof request !== "object") return null;
    if (connectorId === "http") {
      if (typeof request.url !== "string") return null;
      try {
        return httpScope({ url: request.url, realm: request.realm ?? null });
      } catch {
        return null;
      }
    }
    if (connectorId === "file") {
      if (typeof request.path !== "string") return null;
      return fileScopeExact({ path: request.path });
    }
    if (connectorId === "sql") {
      return parseSqlConnectionScope(request.connection);
    }
    return null;
  }

  /**
   * Implementation for `QueryEngine({ onCredentialRequest })`.
   *
   * The returned value is a handle that is safe to embed in cache keys (via its
   * stable `credentialId`). Connectors can call `getSecret()` to retrieve the
   * underlying secret when they are about to execute.
   *
   * @param {string} connectorId
   * @param {{ request: unknown }} details
   * @returns {Promise<CredentialHandle | undefined>}
   */
  async onCredentialRequest(connectorId, details) {
    const request = details?.request;
    const scope = this.resolveScope(connectorId, request);
    if (!scope) return undefined;

    const existing = await this.store.get(scope);
    if (existing) {
      return {
        credentialId: existing.id,
        id: existing.id,
        kind: connectorId,
        getSecret: async () => {
          // Avoid capturing the secret in a closure to keep the handle lightweight.
          const latest = await this.store.get(scope);
          return latest?.secret;
        },
      };
    }

    if (!this.prompt) return undefined;
    const provided = await this.prompt({ connectorId, scope, request });
    if (provided == null) return undefined;

    const created = await this.store.set(scope, provided);
    return {
      credentialId: created.id,
      id: created.id,
      kind: connectorId,
      getSecret: async () => {
        const latest = await this.store.get(scope);
        return latest?.secret;
      },
    };
  }
}
