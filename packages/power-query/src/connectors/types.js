/**
 * Connector contracts for `packages/power-query`.
 *
 * This package intentionally stays JS + JSDoc so it can run without a TS build.
 * Host applications can provide their own connectors (or wrappers around native
 * capabilities) while relying on a stable request/response shape.
 */
 
/**
 * @typedef {import("../model.js").DataType} DataType
 * @typedef {import("../table.js").DataTable} DataTable
 */
 
/**
 * Schema information returned by connectors.
 *
 * This is intentionally "Power Query-ish": connectors can provide type
 * inference, but callers should treat this as best-effort.
 *
 * @typedef {{
 *   columns: Array<{ name: string, type: DataType }>;
 *   inferred?: boolean;
 * }} SchemaInfo
 */
 
/**
 * Common metadata returned by all connectors.
 *
 * @typedef {Object} ConnectorMeta
 * @property {Date} refreshedAt When the connector fetched/loaded the data.
 * @property {Date | undefined} [sourceTimestamp]
 *   Optional timestamp for when the underlying source last changed, if the
 *   connector can determine it (e.g. file mtime, HTTP Last-Modified).
 * @property {SchemaInfo} schema Best-effort schema output.
 * @property {number} rowCount Exact row count for the returned table.
 * @property {number | undefined} [rowCountEstimate] Best-effort estimate for the total row count on the source.
 * @property {Record<string, unknown>} provenance Provenance information (URL/path/connection identifiers).
 */
 
/**
 * @typedef {Object} ConnectorResult
 * @property {DataTable} table
 * @property {ConnectorMeta} meta
 */
 
 /**
  * @typedef {Object} ConnectorExecuteOptions
  * @property {AbortSignal | undefined} [signal]
  * @property {unknown} [credentials]
  * @property {(() => number) | undefined} [now]
  */

/**
 * OAuth2 configuration for HTTP requests.
 *
 * Used by:
 * - `HttpConnectorRequest.auth`
 * - `onCredentialRequest("http", { request })` responses (via `credentials.oauth2`)
 *
 * @typedef {Object} HttpConnectorOAuth2Config
 * @property {string} providerId Stable provider/config identifier.
 * @property {string[] | undefined} [scopes] Optional scope list.
 */

/**
 * Credentials shape understood by the built-in `HttpConnector`.
 *
 * Host applications can return this object from `onCredentialRequest("http", ...)`:
 * - `headers` is merged into the request headers (legacy / escape hatch).
 * - `oauth2` delegates bearer token handling to an `OAuth2Manager` configured
 *   on the connector.
 *
 * Precedence: if `request.auth` is provided, it overrides `credentials.oauth2`.
 *
 * @typedef {Object} HttpConnectorCredentials
 * @property {Record<string, string> | undefined} [headers]
 * @property {HttpConnectorOAuth2Config | undefined} [oauth2]
 */

 /**
  * Base connector interface.
  *
  * @template Request
 * @typedef {{
 *   id: string;
 *   /**
 *    * A permission kind to pass to `onPermissionRequest`. Host apps decide what
 *    * to do with it; the library only provides the hook.
 *    */
 *   permissionKind: string;
 *   /**
 *    * Return a stable, JSON-serializable representation of the request to be
 *    * used as an input for cache-key generation.
 *    */
 *   getCacheKey: (request: Request) => unknown;
 *   execute: (request: Request, options?: ConnectorExecuteOptions) => Promise<ConnectorResult>;
 * }} Connector
 */
 
export {};
