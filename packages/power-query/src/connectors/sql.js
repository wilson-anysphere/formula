import { DataTable } from "../table.js";
import { hashValue } from "../cache/key.js";

/**
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
 */

/**
 * @typedef {Object} SqlConnectorRequest
 * @property {string | undefined} [connectionId]
 * @property {unknown} connection
 * @property {string} sql
 * @property {unknown[] | undefined} [params]
 */

/**
 * @typedef {{
 *   columns: string[];
 *   types?: Record<string, import("../model.js").DataType>;
 * }} SqlConnectorSchema
 */

/**
 * @typedef {Object} SqlConnectorOptions
 * @property {((connection: unknown, sql: string, options?: { params?: unknown[]; signal?: AbortSignal; credentials?: unknown }) => Promise<DataTable>) | undefined} [querySql]
 * @property {((connection: unknown) => unknown) | undefined} [getConnectionIdentity]
 *   Return a JSON-serializable identity for the provided connection descriptor.
 *   The engine will hash this identity for stable cache keys + folding.
 * @property {((request: SqlConnectorRequest, options?: { signal?: AbortSignal; credentials?: unknown }) => Promise<SqlConnectorSchema>) | undefined} [getSchema]
 *   Optional hook for schema discovery (columns/types) used to enable SQL folding
 *   of operations like `renameColumn` / `changeType` when the source column list
 *   is not pre-specified.
 */

export class SqlConnector {
  /**
   * @param {SqlConnectorOptions} [options]
   */
  constructor(options = {}) {
    this.id = "sql";
    this.permissionKind = "database:query";
    this.querySql = options.querySql ?? null;
    this.getConnectionIdentity = options.getConnectionIdentity ?? defaultGetConnectionIdentity;
    this.getSchema = options.getSchema ?? null;
  }

  /**
   * @param {SqlConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    const connectionId = resolveConnectionId(request, this.getConnectionIdentity);
    return {
      connector: "sql",
      ...(connectionId ? { connectionId } : { missingConnectionId: true }),
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

/**
 * @param {unknown} connection
 * @returns {unknown}
 */
function defaultGetConnectionIdentity(connection) {
  if (connection == null) return null;
  const type = typeof connection;
  if (type === "string" || type === "number" || type === "boolean") {
    return connection;
  }

  if (type === "object" && !Array.isArray(connection)) {
    // Common case: host apps pass connection descriptors like `{ id: "db1", ... }`.
    // Prefer a stable string identifier instead of hashing the entire object.
    // @ts-ignore - runtime inspection
    if (typeof connection.id === "string" && connection.id) return connection.id;
  }

  return null;
}

/**
 * @param {SqlConnectorRequest} request
 * @param {(connection: unknown) => unknown} getConnectionIdentity
 * @returns {string | null}
 */
function resolveConnectionId(request, getConnectionIdentity) {
  if (typeof request.connectionId === "string" && request.connectionId) {
    return request.connectionId;
  }

  const identity = getConnectionIdentity(request.connection);
  if (identity == null) return null;
  if (typeof identity === "string") return identity;
  return hashValue(identity);
}
