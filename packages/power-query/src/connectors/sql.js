import { DataTable } from "../table.js";

/**
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
 */

/**
 * @typedef {Object} SqlConnectorRequest
 * @property {unknown} connection
 * @property {string} sql
 * @property {unknown[] | undefined} [params]
 */

/**
 * @typedef {Object} SqlConnectorOptions
 * @property {((connection: unknown, sql: string, options?: { params?: unknown[]; signal?: AbortSignal; credentials?: unknown }) => Promise<DataTable>) | undefined} [querySql]
 */

export class SqlConnector {
  /**
   * @param {SqlConnectorOptions} [options]
   */
  constructor(options = {}) {
    this.id = "sql";
    this.permissionKind = "database:query";
    this.querySql = options.querySql ?? null;
  }

  /**
   * @param {SqlConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    return {
      connector: "sql",
      connection: request.connection,
      sql: request.sql,
      params: request.params ?? null,
    };
  }

  /**
   * @param {SqlConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<ConnectorResult>}
   */
  async execute(request, options = {}) {
    if (!this.querySql) {
      throw new Error("Database source requires a SqlConnector querySql adapter");
    }

    const now = options.now ?? (() => Date.now());
    let credentials = options.credentials;
    if (
      credentials &&
      typeof credentials === "object" &&
      !Array.isArray(credentials) &&
      // @ts-ignore - runtime access
      typeof credentials.getSecret === "function"
    ) {
      // Credential handle convention: hosts can return an object with a stable
      // credentialId plus a `getSecret()` method. This keeps secret retrieval
      // inside the connector execution path (and out of cache key hashing).
      // @ts-ignore - runtime call
      credentials = await credentials.getSecret();
    }
    const table = await this.querySql(request.connection, request.sql, {
      params: request.params,
      signal: options.signal,
      credentials,
    });

    return {
      table,
      meta: {
        refreshedAt: new Date(now()),
        schema: { columns: table.columns, inferred: true },
        rowCount: table.rows.length,
        rowCountEstimate: table.rows.length,
        provenance: { kind: "sql", sql: request.sql },
      },
    };
  }
}
