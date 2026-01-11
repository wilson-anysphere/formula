import { DataTable } from "../table.js";

/**
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
 */

/**
 * @typedef {Object} SqlConnectorRequest
 * @property {unknown} connection
 * @property {string} sql
 */

/**
 * @typedef {Object} SqlConnectorOptions
 * @property {((connection: unknown, sql: string, options?: { signal?: AbortSignal; credentials?: unknown }) => Promise<DataTable>) | undefined} [querySql]
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
    const table = await this.querySql(request.connection, request.sql, { signal: options.signal, credentials: options.credentials });

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

