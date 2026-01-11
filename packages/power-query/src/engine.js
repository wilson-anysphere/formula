import { applyOperation } from "./steps.js";
import { ArrowTableAdapter } from "./arrowTable.js";
import { DataTable } from "./table.js";

import { hashValue } from "./cache/key.js";
import { deserializeAnyTable, deserializeTable, serializeAnyTable } from "./cache/serialize.js";
import { FileConnector } from "./connectors/file.js";
import { HttpConnector } from "./connectors/http.js";
import { SqlConnector } from "./connectors/sql.js";
import { QueryFoldingEngine } from "./folding/sql.js";
import { normalizePostgresPlaceholders } from "./folding/placeholders.js";
import { computeParquetProjectionColumns, computeParquetRowLimit } from "./parquetProjection.js";
import { collectSourcePrivacy, distinctPrivacyLevels, shouldBlockCombination } from "./privacy/firewall.js";
import { getSourceIdForProvenance, getSourceIdForQuerySource } from "./privacy/sourceId.js";
import { getPrivacyLevel } from "./privacy/levels.js";

/**
 * Lazy-load Arrow/parquet helpers from `@formula/data-io`.
 *
 * Power Query's core engine can run without Arrow support (e.g. for CSV/API/SQL
 * sources). Keeping this import lazy avoids hard-failing in environments where
 * optional Arrow dependencies are not present.
 *
 * @returns {Promise<typeof import("../../data-io/src/index.js")>}
 */
let dataIoModulePromise = null;
async function loadDataIoModule() {
  if (!dataIoModulePromise) {
    dataIoModulePromise = import("../../data-io/src/index.js");
  }
  return dataIoModulePromise;
}

/**
 * @typedef {import("./model.js").Query} Query
 * @typedef {import("./model.js").QuerySource} QuerySource
 * @typedef {import("./model.js").QueryStep} QueryStep
 * @typedef {import("./model.js").QueryOperation} QueryOperation
 * @typedef {import("./table.js").ITable} ITable
 * @typedef {import("./connectors/types.js").ConnectorMeta} ConnectorMeta
 * @typedef {import("./connectors/types.js").SchemaInfo} SchemaInfo
 */

/**
 * @typedef {{
 *   tables?: Record<string, ITable>;
 *   queries?: Record<string, Query>;
 *   // Optional pre-computed query results that can be reused when resolving query
 *   // references (source.type === "query") and merge/append dependencies. This is
 *   // primarily used by dependency-aware refresh orchestration ("Refresh All") to
 *   // avoid re-executing shared upstream queries.
 *   queryResults?: Record<string, QueryExecutionResult>;
 *   // Optional host-provided version/signature per table name. When provided, table
 *   // sources incorporate the signature into the query cache key so cached results
 *   // reflect workbook edits.
 *   tableSignatures?: Record<string, unknown>;
 *   // Optional callback used to resolve a signature/version for a table name. If both
 *   // `getTableSignature` and `tableSignatures` are supplied, the callback wins.
 *   getTableSignature?: (tableName: string) => unknown;
 *   privacy?: { levelsBySourceId: Record<string, import("./privacy/levels.js").PrivacyLevel> };
 * }} QueryExecutionContext
 */

/**
 * Shared state for a group of query executions.
 *
 * A session allows the engine to reuse credential/permission prompts and other
 * deterministic values (like the "current time") across multiple query
 * executions.
 *
 * @typedef {{
 *   credentialCache: Map<string, Promise<unknown>>;
 *   permissionCache: Map<string, Promise<boolean>>;
 *   now?: () => number;
 * }} QueryExecutionSession
 */

/**
 * @typedef {{
 *   type: "cache:hit" | "cache:miss" | "cache:set";
 *   queryId: string;
 *   cacheKey: string;
 * } | {
 *   type: "source:start" | "source:complete";
 *   queryId: string;
 *   sourceType: QuerySource["type"];
 * } | {
 *   type: "step:start" | "step:complete";
 *   queryId: string;
 *   stepIndex: number;
 *   stepId: string;
 *   operation: QueryOperation["type"];
 * } | {
 *   type: "privacy:firewall";
 *   queryId: string;
 *   phase: "folding" | "combine";
 *   mode: "enforce" | "warn";
 *   action: "prevent-folding" | "warn" | "block";
 *   operation: "merge" | "append";
 *   stepIndex?: number;
 *   stepId?: string;
 *   sources: Array<{ sourceId: string; level: import("./privacy/levels.js").PrivacyLevel }>;
 *   message: string;
 * }} EngineProgressEvent
 */

/**
 * @typedef {{
 *   limit?: number;
 *   // Execute up to and including this step index.
 *   maxStepIndex?: number;
 *   signal?: AbortSignal;
 *   onProgress?: (event: EngineProgressEvent) => void;
 *   cache?: { mode?: "use" | "refresh" | "bypass"; ttlMs?: number; validation?: "none" | "source-state" };
 * }} ExecuteOptions
 */

/**
 * @typedef {Object} QueryExecutionMeta
 * @property {string} queryId
 * @property {Date} startedAt
 * @property {Date} completedAt
 * @property {Date} refreshedAt
 *   When the underlying data was last refreshed (i.e. the cache entry was
 *   created, or the refresh just completed).
 * @property {ConnectorMeta[]} sources Metadata for every source the query touched (including referenced queries).
 * @property {SchemaInfo} outputSchema
 * @property {number} outputRowCount
 * @property {{ key: string; hit: boolean } | undefined} [cache]
 * @property {{
 *   dialect?: import("./folding/dialect.js").SqlDialectName;
 *   planType: "local" | "sql" | "hybrid";
 *   sql: string;
 *   params: unknown[];
 *   steps: import("./folding/sql.js").FoldingExplainStep[];
 *   // Index within the executed step list where local execution begins.
 *   // Only present for hybrid plans.
 *   localStepOffset?: number;
 * } | undefined} [folding]
 */

/**
 * @typedef {{
 *   table: ITable;
 *   meta: QueryExecutionMeta;
 * }} QueryExecutionResult
 */

/**
 * Host-provided hooks used during query execution.
 *
 * `onCredentialRequest(connectorId, { request })` may return:
 * - `undefined` / `null`: no credentials
 * - an object understood by the target connector (e.g. `HttpConnector` supports
 *   `{ headers }` and `{ oauth2: { providerId, scopes? } }`)
 * - a credential handle with a `getSecret()` method (see `CredentialManager`)
 *
 * The engine memoizes the returned credentials per request within a single
 * execution so repeated source calls don't repeatedly prompt the user.
 *
 * @typedef {{
 *   onPermissionRequest?: (kind: string, details: unknown) => boolean | Promise<boolean>;
 *   onCredentialRequest?: (connectorId: string, details: unknown) => unknown | Promise<unknown>;
 * }} QueryEngineHooks
 */

/**
 * @typedef {Object} QueryEngineOptions
 * @property {{ querySql: (connection: unknown, sql: string, options?: any) => Promise<ITable> } | undefined} [databaseAdapter]
 *   Backwards-compatible adapter from the prototype. Prefer supplying a `SqlConnector`.
 * @property {{ fetchTable: (url: string, options: { method: string; headers?: Record<string, string> }) => Promise<ITable> } | undefined} [apiAdapter]
 *   Backwards-compatible adapter from the prototype. Prefer supplying a `HttpConnector`.
 * @property {{
 *   readText?: (path: string) => Promise<string>;
 *   readBinary?: (path: string) => Promise<Uint8Array>;
 *   readParquetTable?: (path: string, options?: { signal?: AbortSignal }) => Promise<DataTable>;
 *   stat?: (path: string) => Promise<{ mtimeMs: number }>;
 * } | undefined} [fileAdapter]
 *   Backwards-compatible adapter from the prototype. Prefer supplying a `FileConnector`.
 * @property {Partial<{ file: FileConnector; http: HttpConnector; sql: SqlConnector } & Record<string, any>> | undefined} [connectors]
 * @property {import("./cache/cache.js").CacheManager | undefined} [cache]
 * @property {number | undefined} [defaultCacheTtlMs]
 * @property {{ enabled?: boolean; dialect?: import("./folding/dialect.js").SqlDialectName | import("./folding/dialect.js").SqlDialect } | undefined} [sqlFolding]
 *   When enabled and a dialect is known (either via `source.dialect` or this
 *   default dialect), the engine will execute a foldable prefix of operations
 *   in the source database via `QueryFoldingEngine`.
 * @property {"ignore" | "enforce" | "warn" | undefined} [privacyMode]
 *   Privacy enforcement mode for Power Query-style "privacy levels" / formula firewall.
 *   Defaults to `"ignore"` for backwards compatibility.
 * @property {QueryEngineHooks["onPermissionRequest"] | undefined} [onPermissionRequest]
 * @property {QueryEngineHooks["onCredentialRequest"] | undefined} [onCredentialRequest]
 */

/**
 * @param {AbortSignal | undefined} signal
 */
function throwIfAborted(signal) {
  if (!signal?.aborted) return;
  const err = new Error("Aborted");
  err.name = "AbortError";
  throw err;
}

/**
 * @param {{ id: string, getCacheKey: (request: any) => unknown }} connector
 * @param {any} request
 */
function buildConnectorSourceKey(connector, request) {
  return `${connector.id}:${hashValue(connector.getCacheKey(request))}`;
}

/**
 * @param {ConnectorMeta} meta
 * @returns {{ refreshedAtMs: number, sourceTimestampMs?: number, etag?: string, sourceKey?: string, schema: any, rowCount: number, rowCountEstimate?: number, provenance: any }}
 */
function serializeConnectorMeta(meta) {
  return {
    refreshedAtMs: meta.refreshedAt.getTime(),
    sourceTimestampMs: meta.sourceTimestamp ? meta.sourceTimestamp.getTime() : undefined,
    etag: meta.etag,
    sourceKey: meta.sourceKey,
    schema: meta.schema,
    rowCount: meta.rowCount,
    rowCountEstimate: meta.rowCountEstimate,
    provenance: meta.provenance,
  };
}

/**
 * @param {any} data
 * @returns {ConnectorMeta}
 */
function deserializeConnectorMeta(data) {
  return {
    refreshedAt: new Date(data.refreshedAtMs),
    sourceTimestamp: data.sourceTimestampMs != null ? new Date(data.sourceTimestampMs) : undefined,
    etag: typeof data.etag === "string" ? data.etag : undefined,
    sourceKey: typeof data.sourceKey === "string" ? data.sourceKey : undefined,
    schema: data.schema,
    rowCount: data.rowCount,
    rowCountEstimate: data.rowCountEstimate,
    provenance: data.provenance,
  };
}

/**
 * @param {QueryExecutionMeta} meta
 * @returns {any}
 */
function serializeQueryMeta(meta) {
  return {
    queryId: meta.queryId,
    refreshedAtMs: meta.refreshedAt.getTime(),
    sources: meta.sources.map(serializeConnectorMeta),
    outputSchema: meta.outputSchema,
    outputRowCount: meta.outputRowCount,
    folding: meta.folding,
  };
}

/**
 * @param {any} data
 * @param {Date} startedAt
 * @param {Date} completedAt
 * @param {{ key: string; hit: boolean } | undefined} cache
 * @returns {QueryExecutionMeta}
 */
function deserializeQueryMeta(data, startedAt, completedAt, cache) {
  return {
    queryId: data.queryId,
    startedAt,
    completedAt,
    refreshedAt: new Date(data.refreshedAtMs),
    sources: Array.isArray(data.sources) ? data.sources.map(deserializeConnectorMeta) : [],
    outputSchema: data.outputSchema,
    outputRowCount: data.outputRowCount,
    cache,
    folding: data.folding,
  };
}

/**
 * Extract a stable credential identifier from a credentials object.
 *
 * Host applications should return credential handles that include a stable
 * `credentialId` (or `id`) so cache keys can vary by credential without
 * embedding secret material.
 *
 * @param {unknown} credentials
 * @returns {string | null}
 */
function extractCredentialId(credentials) {
  if (!credentials) return null;
  if (typeof credentials !== "object" || Array.isArray(credentials)) return null;
  // @ts-ignore - runtime access
  const id = credentials.credentialId ?? credentials.id ?? null;
  return typeof id === "string" && id.length > 0 ? id : null;
}

export class QueryEngine {
  /**
   * @param {QueryEngineOptions} [options]
   */
  constructor(options = {}) {
    this.onPermissionRequest = options.onPermissionRequest ?? null;
    this.onCredentialRequest = options.onCredentialRequest ?? null;
    this.fileAdapter = options.fileAdapter ?? null;
    this.privacyMode = options.privacyMode ?? "ignore";

    /** @type {WeakMap<object, Set<string>>} */
    this._tableSourceIds = new WeakMap();

    /** @type {WeakMap<object, string>} */
    this._ephemeralObjectIds = new WeakMap();
    this._ephemeralObjectIdCounter = 0;

    /** @type {Map<string, any>} */
    this.connectors = new Map();

    const fileConnector =
      options.connectors?.file ??
      new FileConnector({
        readText: options.fileAdapter?.readText,
        readParquetTable: options.fileAdapter?.readParquetTable,
        stat: options.fileAdapter?.stat,
      });
    const httpConnector = options.connectors?.http ?? new HttpConnector({ fetchTable: options.apiAdapter?.fetchTable });
    const sqlConnector = options.connectors?.sql ?? new SqlConnector({ querySql: options.databaseAdapter?.querySql });

    this.connectors.set(fileConnector.id, fileConnector);
    this.connectors.set(httpConnector.id, httpConnector);
    this.connectors.set(sqlConnector.id, sqlConnector);

    if (options.connectors) {
      for (const connector of Object.values(options.connectors)) {
        if (!connector || typeof connector !== "object") continue;
        if (typeof connector.id === "string") {
          this.connectors.set(connector.id, connector);
        }
      }
    }

    this.cache = options.cache ?? null;
    this.defaultCacheTtlMs = options.defaultCacheTtlMs ?? null;

    this.sqlFoldingEnabled = options.sqlFolding?.enabled ?? true;
    this.sqlFoldingDialect = options.sqlFolding?.dialect ?? null;
    this.foldingEngine = new QueryFoldingEngine();

    /** @type {Map<string, Promise<{ columns: string[], types?: Record<string, import("./model.js").DataType> }>>} */
    this.databaseSchemaCache = new Map();
  }

  /**
   * Generate a stable, per-engine identifier for an object reference.
   *
   * This is used as a fallback for permission/credential prompt caching when we
   * don't have a stable, user-provided identity (e.g. for opaque DB connection
   * handles). The ID is only stable for the lifetime of this engine instance.
   *
   * @private
   * @param {unknown} value
   * @returns {string | null}
   */
  getEphemeralObjectId(value) {
    if (!value) return null;
    const type = typeof value;
    if (type !== "object" && type !== "function") return null;
    const obj = /** @type {object} */ (value);
    const existing = this._ephemeralObjectIds.get(obj);
    if (existing) return existing;
    const next = `obj:${++this._ephemeralObjectIdCounter}`;
    this._ephemeralObjectIds.set(obj, next);
    return next;
  }

  /**
   * Build a stable cache-key input for a connector request.
   *
   * Prefer the connector's `getCacheKey(request)` (which should be JSON-safe and
   * avoid opaque handles). For SQL requests without a stable connection identity,
   * include an ephemeral per-object identifier to avoid collisions between
   * different connection handles in the same session.
   *
   * @private
   * @param {string} connectorId
   * @param {any} request
   * @returns {unknown}
   */
  buildConnectorRequestCacheKey(connectorId, request) {
    const connector = this.connectors.get(connectorId);
    /** @type {any} */
    let keyInput = request;
    if (connector && typeof connector.getCacheKey === "function") {
      try {
        keyInput = connector.getCacheKey(request);
      } catch {
        keyInput = request;
      }
    }

    if (
      connectorId === "sql" &&
      keyInput &&
      typeof keyInput === "object" &&
      !Array.isArray(keyInput) &&
      // @ts-ignore - runtime indexing
      keyInput.missingConnectionId === true
    ) {
      const refId = this.getEphemeralObjectId(request?.connection);
      if (refId) {
        keyInput = { ...keyInput, connectionRefId: refId };
      }
    }

    return keyInput;
  }

  /**
   * @private
   * @param {ITable} table
   * @returns {Set<string>}
   */
  getTableSourceIds(table) {
    const ids = this._tableSourceIds.get(/** @type {any} */ (table));
    return ids ?? new Set();
  }

  /**
   * @private
   * @param {ITable} table
   * @param {Iterable<string>} sourceIds
   */
  setTableSourceIds(table, sourceIds) {
    this._tableSourceIds.set(/** @type {any} */ (table), new Set(sourceIds));
  }

  /**
   * @private
   * @param {ConnectorMeta[]} metas
   * @returns {Set<string>}
   */
  collectSourceIdsFromMetas(metas) {
    /** @type {Set<string>} */
    const ids = new Set();
    for (const meta of metas) {
      const sourceId = getSourceIdForProvenance(meta.provenance);
      if (sourceId) ids.add(sourceId);
    }
    return ids;
  }

  /**
   * Enforce the Power Query formula firewall for *local* data combination steps
   * (`merge` / `append`).
   *
   * @private
   * @param {{
   *   queryId: string;
   *   operation: "merge" | "append";
   *   sourceIds: Set<string>;
   *   context: QueryExecutionContext;
   *   options: ExecuteOptions;
   *   stepIndex?: number;
   *   stepId?: string;
   * }} args
   */
  enforceFirewallForCombination(args) {
    if (this.privacyMode === "ignore") return;
    if (this.privacyMode !== "warn" && this.privacyMode !== "enforce") return;

    const infos = collectSourcePrivacy(args.sourceIds, args.context.privacy?.levelsBySourceId);
    const levels = distinctPrivacyLevels(infos);
    if (levels.size <= 1) return;

    const message = `Formula firewall detected ${args.operation} across privacy levels (${Array.from(levels).join(", ")})`;
    const mode = this.privacyMode;

    if (mode === "enforce" && shouldBlockCombination(infos)) {
      args.options.onProgress?.({
        type: "privacy:firewall",
        queryId: args.queryId,
        phase: "combine",
        mode,
        action: "block",
        operation: args.operation,
        stepIndex: args.stepIndex,
        stepId: args.stepId,
        sources: infos,
        message,
      });
      const err = new Error(
        `Formula.Firewall: Query '${args.queryId}' blocked combining sources with incompatible privacy levels (${Array.from(levels).join(", ")})`,
      );
      err.name = "PrivacyError";
      throw err;
    }

    args.options.onProgress?.({
      type: "privacy:firewall",
      queryId: args.queryId,
      phase: "combine",
      mode,
      action: "warn",
      operation: args.operation,
      stepIndex: args.stepIndex,
      stepId: args.stepId,
      sources: infos,
      message,
    });
  }

  /**
   * Execute a full query.
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions} [options]
   * @returns {Promise<ITable>}
   */
  async executeQuery(query, context = {}, options = {}) {
    const { table } = await this.executeQueryWithMeta(query, context, options);
    return table;
  }

  /**
   * Execute a query and return refresh/caching metadata.
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions} [options]
   * @returns {Promise<QueryExecutionResult>}
   */
  async executeQueryWithMeta(query, context = {}, options = {}) {
    const session = this.createSession();
    return this.executeQueryWithMetaInSession(query, context, options, session);
  }

  /**
   * Create a shared execution session for running multiple queries.
   *
   * @param {{ now?: () => number }} [options]
   * @returns {QueryExecutionSession}
   */
  createSession(options = {}) {
    return {
      credentialCache: new Map(),
      permissionCache: new Map(),
      now: options.now,
    };
  }

  /**
   * Execute a query with a shared session (credential/permission caches).
   *
   * This is the preferred entry point for dependency-aware "Refresh All"
   * orchestration where multiple queries should share prompts and other
   * deterministic state.
   *
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions} [options]
   * @param {QueryExecutionSession} session
   * @returns {Promise<QueryExecutionResult>}
   */
  async executeQueryWithMetaInSession(query, context = {}, options = {}, session) {
    const now = session.now ?? (() => Date.now());
    return this.executeQueryInternal(
      query,
      context,
      options,
      { credentialCache: session.credentialCache, permissionCache: session.permissionCache, now },
      new Set([query.id]),
    );
  }

  /**
   * Compute a deterministic cache key for a query execution.
   *
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @returns {Promise<string | null>}
   */
  async getCacheKey(query, context = {}, options = {}) {
    if (!this.cache) return null;
    /** @type {Map<string, Promise<unknown>>} */
    const credentialCache = new Map();
    /** @type {Map<string, Promise<boolean>>} */
    const permissionCache = new Map();
    const now = () => Date.now();
    const state = { credentialCache, permissionCache, now };
    return this.computeCacheKey(query, context, options, state, new Set([query.id]));
  }

  /**
   * Manual invalidation helper.
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions} [options]
   */
  async invalidateQueryCache(query, context = {}, options = {}) {
    if (!this.cache) return;
    const key = await this.getCacheKey(query, context, options);
    if (key) await this.cache.delete(key);
  }

  /**
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<QueryExecutionResult>}
   */
  async executeQueryInternal(query, context, options, state, callStack) {
    throwIfAborted(options.signal);

    const startedAt = new Date(state.now());
    const cacheMode = options.cache?.mode ?? "use";
    const cacheValidation = options.cache?.validation ?? "source-state";
    const cacheTtlMs = options.cache?.ttlMs ?? this.defaultCacheTtlMs ?? undefined;

    /** @type {string | null} */
    let cacheKey = null;
    if (this.cache && cacheMode !== "bypass") {
      cacheKey = await this.computeCacheKey(query, context, options, state, callStack);
      if (cacheKey && cacheMode === "use") {
        /** @type {import("./cache/cache.js").CacheEntry | null} */
        let cached = null;
        try {
          cached = await this.cache.getEntry(cacheKey);
        } catch {
          cached = null;
        }
        if (cached) {
          const payload = /** @type {any} */ (cached.value);
          let cacheHitValid = cacheValidation === "none";
          if (!cacheHitValid) {
            try {
              cacheHitValid = await this.validateCacheEntry(query, context, options, state, callStack, payload?.meta);
            } catch {
              cacheHitValid = false;
            }
          }
          if (cacheHitValid) {
            const completedAt = new Date(state.now());
            try {
              const table =
                payload?.version === 2
                  ? deserializeAnyTable(payload.table)
                  : payload?.version === 1
                    ? deserializeTable(payload.table)
                    : payload?.table?.kind
                      ? deserializeAnyTable(payload.table)
                       : deserializeTable(payload.table);
              const meta = deserializeQueryMeta(payload.meta, startedAt, completedAt, { key: cacheKey, hit: true });
              this.setTableSourceIds(table, this.collectSourceIdsFromMetas(meta.sources));
              options.onProgress?.({ type: "cache:hit", queryId: query.id, cacheKey });
              return { table, meta };
            } catch {
              // Treat cache corruption as a miss so we can recover on the next refresh.
              try {
                await this.cache.delete(cacheKey);
              } catch {
                // ignore
              }
            }
          }
          options.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey });
        } else {
          options.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey });
        }
      }
    }

    /** @type {ConnectorMeta[]} */
    const sources = [];

    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);

    /** @type {import("./folding/sql.js").FoldingExplainResult | null} */
    let foldingExplain = null;
    /** @type {import("./folding/sql.js").CompiledQueryPlan | null} */
    let foldedPlan = null;
    /** @type {import("./folding/dialect.js").SqlDialectName | import("./folding/dialect.js").SqlDialect | null} */
    let foldedDialect = null;
    if (this.sqlFoldingEnabled && query.source.type === "database") {
      const dialect = query.source.dialect ?? this.sqlFoldingDialect;
      foldedDialect = dialect ?? null;

      const sqlConnector = this.connectors.get("sql");
      const getConnectionIdentity =
        sqlConnector && typeof sqlConnector.getConnectionIdentity === "function"
          ? (connection) => sqlConnector.getConnectionIdentity(connection)
          : undefined;

      let sourceForFolding = query.source;
      if (dialect && query.source.columns == null && sqlConnector && typeof sqlConnector.getSchema === "function") {
        const connectionId = resolveDatabaseConnectionId(query.source, sqlConnector);
        const request = {
          connectionId: connectionId ?? undefined,
          connection: query.source.connection,
          sql: query.source.query,
        };

        /** @type {string | null} */
        let schemaCacheKey = null;
        /** @type {Promise<{ columns: string[], types?: Record<string, import("./model.js").DataType> }> | null} */
        let schemaPromise = null;
        try {
          throwIfAborted(options.signal);
          await this.assertPermission(sqlConnector.permissionKind, { source: query.source, request }, state);
          const credentials = await this.getCredentials("sql", request, state);
          const credentialId = extractCredentialId(credentials);
          const schemaCacheable = credentials == null || credentialId != null;

          if (connectionId && schemaCacheable) {
            schemaCacheKey = `pq:schema:v2:${hashValue({
              connectionId,
              sql: query.source.query,
              credentialsHash: credentialId ? hashValue(credentialId) : null,
            })}`;
            schemaPromise = this.databaseSchemaCache.get(schemaCacheKey) ?? null;
          }

          if (!schemaPromise) {
            schemaPromise = Promise.resolve(sqlConnector.getSchema(request, { signal: options.signal, credentials }));
            if (schemaCacheKey) {
              schemaPromise = schemaPromise.catch((err) => {
                this.databaseSchemaCache.delete(schemaCacheKey);
                throw err;
              });
              this.databaseSchemaCache.set(schemaCacheKey, schemaPromise);
            }
          }
        } catch {
          schemaPromise = null;
        }

        if (schemaPromise) {
          try {
            const schema = await schemaPromise;
            if (schema && Array.isArray(schema.columns) && schema.columns.length > 0) {
              sourceForFolding = { ...query.source, columns: schema.columns.slice() };
            }
          } catch {
            // Schema discovery is best-effort; ignore failures and let folding fall back to hybrid/local execution.
          }
        }
      }

      foldingExplain = this.foldingEngine.explain(
        { ...query, source: sourceForFolding, steps },
        {
          dialect: dialect ?? undefined,
          queries: context.queries ?? undefined,
          getConnectionIdentity,
          privacyMode: this.privacyMode,
          privacyLevelsBySourceId: context.privacy?.levelsBySourceId,
        },
      );
      foldedPlan = foldingExplain.plan;
    }

    /** @type {string | null} */
    let executedSql = null;
    /** @type {unknown[] | null} */
    let executedParams = null;
    /** @type {number | undefined} */
    let localStepOffset = undefined;

    if (query.source.type === "database") {
      executedSql = query.source.query;
      executedParams = [];
    }

    if (foldedPlan && Array.isArray(foldedPlan.diagnostics) && (this.privacyMode === "warn" || this.privacyMode === "enforce")) {
      for (const diag of foldedPlan.diagnostics) {
        options.onProgress?.({
          type: "privacy:firewall",
          queryId: query.id,
          phase: "folding",
          mode: this.privacyMode,
          action: "prevent-folding",
          operation: diag.operation,
          sources: diag.sources,
          message: diag.message,
        });
      }
    }

    /** @type {ITable} */
    let table;

    if (
      foldedPlan &&
      (foldedPlan.type === "sql" || foldedPlan.type === "hybrid") &&
      query.source.type === "database" &&
      foldedDialect
    ) {
      const sqlToRun =
        foldedPlan.type === "sql" && options.limit != null
          ? `SELECT * FROM (${foldedPlan.sql}) AS t LIMIT ?`
          : foldedPlan.sql;
      const paramsToRun =
        foldedPlan.type === "sql" && options.limit != null ? [...foldedPlan.params, options.limit] : foldedPlan.params;

      const dialectName = typeof foldedDialect === "string" ? foldedDialect : foldedDialect.name;
      executedSql = dialectName === "postgres" ? normalizePostgresPlaceholders(sqlToRun, paramsToRun.length) : sqlToRun;
      executedParams = paramsToRun;
      const sourceResult = await this.loadDatabaseQueryWithMeta(
        query.source,
        sqlToRun,
        paramsToRun,
        foldedDialect,
        callStack,
        options,
        state,
      );
      sources.push(...sourceResult.sources);
      table = sourceResult.table;

      if (foldedPlan.type === "hybrid" && foldedPlan.localSteps.length > 0) {
        const offset = steps.indexOf(foldedPlan.localSteps[0]);
        localStepOffset = offset >= 0 ? offset : 0;
        table = await this.executeSteps(
          table,
          foldedPlan.localSteps,
          context,
          options,
          state,
          callStack,
          sources,
          localStepOffset,
        );
      }
    } else {
      let source = query.source;
      if (source.type === "parquet" && this.fileAdapter?.readBinary) {
        const projection = computeParquetProjectionColumns(steps);
        const rowLimit = computeParquetRowLimit(steps, options.limit);
        const nextOptions = projection || rowLimit != null ? { ...(source.options ?? {}) } : null;

        if (projection && projection.length > 0 && nextOptions) {
          const existing = Array.isArray(source.options?.columns) ? source.options.columns : [];
          nextOptions.columns = Array.from(new Set([...existing, ...projection]));
        }

        if (rowLimit != null && nextOptions) {
          const existing = typeof source.options?.limit === "number" ? source.options.limit : null;
          nextOptions.limit = existing == null ? rowLimit : Math.min(existing, rowLimit);
        }

        if (nextOptions) {
          source = { ...source, options: nextOptions };
        }
      }

      const sourceResult = await this.loadSourceWithMeta(source, context, callStack, options, state);
      sources.push(...sourceResult.sources);
      table = sourceResult.table;
      table = await this.executeSteps(table, steps, context, options, state, callStack, sources);
    }

    if (options.limit != null) {
      table = table.head(options.limit);
    }

    const completedAt = new Date(state.now());

    /** @type {SchemaInfo} */
    const outputSchema = { columns: table.columns, inferred: true };

    /** @type {QueryExecutionMeta} */
    const meta = {
      queryId: query.id,
      startedAt,
      completedAt,
      refreshedAt: completedAt,
      sources,
      outputSchema,
      outputRowCount: table.rowCount,
      cache: cacheKey ? { key: cacheKey, hit: false } : undefined,
      folding:
      foldingExplain && query.source.type === "database" && executedSql && executedParams
          ? {
              dialect: foldedDialect ? (typeof foldedDialect === "string" ? foldedDialect : foldedDialect.name) : undefined,
              planType: foldingExplain.plan.type,
              sql: executedSql,
              params: executedParams,
              steps: foldingExplain.steps,
              localStepOffset: foldingExplain.plan.type === "hybrid" ? localStepOffset : undefined,
            }
          : undefined,
    };

    // Ensure the final materialized table is tagged with the full source set so
    // downstream merge/append operations can enforce privacy levels correctly.
    //
    // Note: Some sources (e.g. SQL connections without a stable identity) may
    // only have an ephemeral per-engine id stored on the table, and will not be
    // recoverable from connector provenance alone. Union the existing tags with
    // the provenance-derived tags rather than overwriting them.
    const existingSourceIds = this.getTableSourceIds(table);
    const metaSourceIds = this.collectSourceIdsFromMetas(sources);
    this.setTableSourceIds(table, new Set([...existingSourceIds, ...metaSourceIds]));

    if (this.cache && cacheKey && cacheMode !== "bypass" && (table instanceof DataTable || table instanceof ArrowTableAdapter)) {
      try {
        await this.cache.set(
          cacheKey,
          { version: 2, table: serializeAnyTable(table), meta: serializeQueryMeta(meta) },
          { ttlMs: cacheTtlMs },
        );
        options.onProgress?.({ type: "cache:set", queryId: query.id, cacheKey });
      } catch {
        // Best-effort: cache failures should not fail the query execution.
      }
    }

    return { table, meta };
  }

  /**
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<string | null>}
   */
  async computeCacheKey(query, context, options, state, callStack) {
    if (!this.cache) return null;

    const signature = await this.buildQuerySignature(query, context, options, state, callStack);
    if (signature && typeof signature === "object" && signature.$cacheable === false) return null;
    return `pq:v1:${hashValue(signature)}`;
  }

  /**
   * Validate a cached query result against the current state of its sources.
   *
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @param {any} cachedMeta
   * @returns {Promise<boolean>}
   */
  async validateCacheEntry(query, context, options, state, callStack, cachedMeta) {
    // No metadata -> can't validate; force refresh.
    if (!cachedMeta || !Array.isArray(cachedMeta.sources)) return false;

    /** @type {Map<string, { sourceTimestampMs?: number, etag?: string }>} */
    const cachedStates = new Map();
    for (const source of cachedMeta.sources) {
      if (!source || typeof source !== "object") continue;
      // @ts-ignore - runtime indexing
      const key = source.sourceKey;
      if (typeof key !== "string" || key === "") continue;
      // @ts-ignore - runtime indexing
      const ts = source.sourceTimestampMs;
      // @ts-ignore - runtime indexing
      const etag = source.etag;
      cachedStates.set(key, {
        sourceTimestampMs: typeof ts === "number" ? ts : undefined,
        etag: typeof etag === "string" ? etag : undefined,
      });
    }

    /** @type {Map<string, { connector: any, request: any, credentials: unknown }>} */
    const targets = new Map();
    await this.collectSourceStateTargets(query, context, options, state, callStack, targets);

    if (targets.size === 0) return true;
    if (cachedStates.size === 0) return false;

    for (const [sourceKey, target] of targets.entries()) {
      const cached = cachedStates.get(sourceKey);
      if (!cached) return false;

      const probe = target.connector?.getSourceState;
      if (typeof probe !== "function") continue;

      /** @type {import("./connectors/types.js").SourceState} */
      let currentState = {};
      try {
        currentState = await probe.call(target.connector, target.request, {
          signal: options.signal,
          credentials: target.credentials,
          now: state.now,
        });
      } catch {
        // If the probe fails (offline / server doesn't support HEAD), fall back to the cached result.
        continue;
      }

      const currentTimestamp = currentState?.sourceTimestamp;
      const currentEtag = currentState?.etag;
      const currentHasState =
        (currentTimestamp instanceof Date && !Number.isNaN(currentTimestamp.getTime())) ||
        (typeof currentEtag === "string" && currentEtag !== "");
      const cachedHasState =
        (typeof cached.sourceTimestampMs === "number" && Number.isFinite(cached.sourceTimestampMs)) ||
        (typeof cached.etag === "string" && cached.etag !== "");

      // If we can see state now but it wasn't captured in the cached entry, force a refresh so future hits can validate.
      if (!cachedHasState && currentHasState) return false;

      if (typeof cached.etag === "string" && typeof currentEtag === "string" && cached.etag !== currentEtag) return false;
      if (typeof cached.sourceTimestampMs === "number" && currentTimestamp instanceof Date) {
        const currentMs = currentTimestamp.getTime();
        if (!Number.isNaN(currentMs) && cached.sourceTimestampMs !== currentMs) return false;
      }
    }

    return true;
  }

  /**
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @param {Map<string, { connector: any, request: any, credentials: unknown }>} out
   */
  async collectSourceStateTargets(query, context, options, state, callStack, out) {
    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);

    await this.collectSourceStateTargetsFromSource(query.source, context, options, state, callStack, out);

    for (const step of steps) {
      if (step.operation.type === "merge") {
        const dep = context.queries?.[step.operation.rightQuery];
        if (!dep) continue;
        if (callStack.has(dep.id)) continue;
        const nextStack = new Set(callStack);
        nextStack.add(dep.id);
        const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
        await this.collectSourceStateTargets(dep, context, depOptions, state, nextStack, out);
      } else if (step.operation.type === "append") {
        for (const id of step.operation.queries) {
          const dep = context.queries?.[id];
          if (!dep) continue;
          if (callStack.has(dep.id)) continue;
          const nextStack = new Set(callStack);
          nextStack.add(dep.id);
          const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
          await this.collectSourceStateTargets(dep, context, depOptions, state, nextStack, out);
        }
      }
    }
  }

  /**
   * @private
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @param {Map<string, { connector: any, request: any, credentials: unknown }>} out
   */
  async collectSourceStateTargetsFromSource(source, context, options, state, callStack, out) {
    if (source.type === "query") {
      const target = context.queries?.[source.queryId];
      if (!target) return;
      if (callStack.has(target.id)) return;
      const nextStack = new Set(callStack);
      nextStack.add(target.id);
      const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
      await this.collectSourceStateTargets(target, context, depOptions, state, nextStack, out);
      return;
    }

    if (source.type === "csv" || source.type === "json" || source.type === "parquet") {
      const connector = this.connectors.get("file");
      if (!connector || typeof connector.getSourceState !== "function") return;
      const request =
        source.type === "csv"
          ? { format: "csv", path: source.path, csv: source.options ?? {} }
          : source.type === "json"
            ? { format: "json", path: source.path, json: { jsonPath: source.jsonPath ?? "" } }
            : { format: "parquet", path: source.path };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("file", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);
      if (!out.has(sourceKey)) out.set(sourceKey, { connector, request, credentials });
      return;
    }

    if (source.type === "api") {
      const connector = this.connectors.get("http");
      if (!connector || typeof connector.getSourceState !== "function") return;
      const request = { url: source.url, method: source.method, headers: source.headers ?? {}, auth: source.auth, responseType: "auto" };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("http", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);
      if (!out.has(sourceKey)) out.set(sourceKey, { connector, request, credentials });
      return;
    }

    if (source.type === "database") {
      const connector = this.connectors.get("sql");
      if (!connector || typeof connector.getSourceState !== "function") return;
      const connectionId = resolveDatabaseConnectionId(source, connector);
      const request = { connectionId: connectionId ?? undefined, connection: source.connection, sql: source.query };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("sql", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);
      if (!out.has(sourceKey)) out.set(sourceKey, { connector, request, credentials });
      return;
    }
  }

  /**
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<unknown | null>}
   */
  async buildQuerySignature(query, context, options, state, callStack) {
    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);

    let cacheable = true;
    const sourceSignature = await this.buildSourceSignature(query.source, context, state, callStack);
    if (sourceSignature && typeof sourceSignature === "object" && sourceSignature.$cacheable === false) {
      cacheable = false;
    }

    /** @type {Record<string, unknown>} */
    const signature = {
      source: sourceSignature,
      steps: steps.map((s) => s.operation),
      options: { limit: options.limit ?? null, maxStepIndex: options.maxStepIndex ?? null },
      privacy: { mode: this.privacyMode },
    };

    // Merge/append steps refer to other queries; include their signatures so the cache key changes when dependencies change.
    for (const step of steps) {
      if (step.operation.type === "merge") {
        const dep = context.queries?.[step.operation.rightQuery];
        if (dep) {
          if (callStack.has(dep.id)) {
            signature[`merge:${step.operation.rightQuery}`] = { queryId: dep.id, cycle: true };
            continue;
          }
          const nextStack = new Set(callStack);
          nextStack.add(dep.id);
          const depSignature = await this.buildQuerySignature(dep, context, {}, state, nextStack);
          if (depSignature && typeof depSignature === "object" && depSignature.$cacheable === false) {
            cacheable = false;
          }
          signature[`merge:${step.operation.rightQuery}`] = depSignature;
        }
      } else if (step.operation.type === "append") {
        for (const id of step.operation.queries) {
          const dep = context.queries?.[id];
          if (dep) {
            if (callStack.has(dep.id)) {
              signature[`append:${id}`] = { queryId: dep.id, cycle: true };
              continue;
            }
            const nextStack = new Set(callStack);
            nextStack.add(dep.id);
            const depSignature = await this.buildQuerySignature(dep, context, {}, state, nextStack);
            if (depSignature && typeof depSignature === "object" && depSignature.$cacheable === false) {
              cacheable = false;
            }
            signature[`append:${id}`] = depSignature;
          }
        }
      }
    }

    signature.$cacheable = cacheable;
    return signature;
  }

  /**
   * @private
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<unknown | null>}
   */
  async buildSourceSignature(source, context, state, callStack) {
    if (source.type === "query") {
      const target = context.queries?.[source.queryId];
      if (!target) return { type: "query", queryId: source.queryId, missing: true };
      if (callStack.has(target.id)) {
        return { type: "query", queryId: source.queryId, cycle: true };
      }
      const nextStack = new Set(callStack);
      nextStack.add(target.id);
      const query = await this.buildQuerySignature(target, context, {}, state, nextStack);
      return {
        type: "query",
        queryId: source.queryId,
        query,
        $cacheable: query && typeof query === "object" ? query.$cacheable !== false : true,
      };
    }

    if (source.type === "range") {
      const sourceId = getSourceIdForQuerySource(source);
      return {
        type: "range",
        sourceId,
        privacyLevel: getPrivacyLevel(context.privacy?.levelsBySourceId, sourceId),
        hasHeaders: source.range.hasHeaders ?? true,
        values: source.range.values,
        $cacheable: true,
      };
    }
    if (source.type === "table") {
      const sourceId = getSourceIdForQuerySource(source);
      const signature =
        typeof context.getTableSignature === "function"
          ? context.getTableSignature(source.table)
          : context.tableSignatures
            ? context.tableSignatures[source.table]
            : undefined;
      const privacyLevel = getPrivacyLevel(context.privacy?.levelsBySourceId, sourceId);
      if (signature === undefined) {
        return { type: "table", sourceId, privacyLevel, table: source.table, missingSignature: true, $cacheable: false };
      }
      return {
        type: "table",
        sourceId,
        table: source.table,
        privacyLevel,
        signature,
        $cacheable: true,
      };
    }

    if (source.type === "csv" || source.type === "json" || source.type === "parquet") {
      const connector = this.connectors.get("file");
      if (!connector) return { type: source.type, missingConnector: "file" };
      const sourceId = getSourceIdForQuerySource(source);
      const request =
        source.type === "csv"
          ? { format: "csv", path: source.path, csv: source.options ?? {} }
          : source.type === "json"
            ? { format: "json", path: source.path, json: { jsonPath: source.jsonPath ?? "" } }
            : { format: "parquet", path: source.path };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("file", request, state);
      const credentialId = extractCredentialId(credentials);
      const cacheable = credentials == null || credentialId != null;
      return {
        type: source.type,
        sourceId,
        privacyLevel: getPrivacyLevel(context.privacy?.levelsBySourceId, sourceId),
        request: connector.getCacheKey(request),
        credentialsHash: credentialId ? hashValue(credentialId) : null,
        $cacheable: cacheable,
      };
    }

    if (source.type === "api") {
      const connector = this.connectors.get("http");
      if (!connector) return { type: "api", missingConnector: "http" };
      const sourceId = getSourceIdForQuerySource(source);
      const request = {
        url: source.url,
        method: source.method,
        headers: source.headers ?? {},
        auth: source.auth,
        responseType: "auto",
      };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("http", request, state);
      const credentialId = extractCredentialId(credentials);
      const cacheable = credentials == null || credentialId != null;
      return {
        type: "api",
        sourceId,
        privacyLevel: getPrivacyLevel(context.privacy?.levelsBySourceId, sourceId),
        request: connector.getCacheKey(request),
        credentialsHash: credentialId ? hashValue(credentialId) : null,
        $cacheable: cacheable,
      };
    }

    if (source.type === "database") {
      const connector = this.connectors.get("sql");
      if (!connector) return { type: "database", missingConnector: "sql" };
      const connectionId = resolveDatabaseConnectionId(source, connector);
      const connectionRefId = this.getEphemeralObjectId(source.connection);
      const sourceId = connectionId ? `sql:${connectionId}` : connectionRefId ? `sql:${connectionRefId}` : getSourceIdForQuerySource(source);
      const privacyLevel = getPrivacyLevel(context.privacy?.levelsBySourceId, sourceId);
      if (!connectionId) {
        return {
          type: "database",
          sourceId,
          privacyLevel,
          dialect: source.dialect ?? null,
          request: connector.getCacheKey({ connection: source.connection, sql: source.query }),
          credentialsHash: null,
          missingConnectionId: true,
          $cacheable: false,
        };
      }

      const request = { connectionId, connection: source.connection, sql: source.query };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("sql", request, state);
      const credentialId = extractCredentialId(credentials);
      const cacheable = credentials == null || credentialId != null;
      return {
        type: "database",
        sourceId,
        privacyLevel,
        connectionId,
        dialect: source.dialect ?? null,
        request: connector.getCacheKey(request),
        credentialsHash: credentialId ? hashValue(credentialId) : null,
        $cacheable: cacheable,
      };
    }

    /** @type {never} */
    const exhausted = source;
    throw new Error(`Unsupported source type '${exhausted.type}'`);
  }

  /**
   * Execute a list of steps starting from an already-materialized table.
   * @param {ITable} table
   * @param {QueryStep[]} steps
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} [options]
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} [state]
   * @param {Set<string>} [callStack]
   * @param {ConnectorMeta[]} [sources]
   * @param {number} [stepIndexOffset]
   * @returns {Promise<ITable>}
   */
  async executeSteps(table, steps, context, options = {}, state, callStack, sources, stepIndexOffset = 0) {
    let current = table;
    const queryId = callStack ? Array.from(callStack).at(-1) ?? "<unknown>" : "<unknown>";
    for (let i = 0; i < steps.length; i++) {
      throwIfAborted(options.signal);
      const step = steps[i];
      const stepIndex = i + stepIndexOffset;
      options.onProgress?.({
        type: "step:start",
        queryId,
        stepIndex,
        stepId: step.id,
        operation: step.operation.type,
      });
      current = await this.applyStep(current, step.operation, context, options, state, callStack, sources, { stepIndex, stepId: step.id });
      options.onProgress?.({
        type: "step:complete",
        queryId,
        stepIndex,
        stepId: step.id,
        operation: step.operation.type,
      });
    }
    return current;
  }

  /**
   * @param {ITable} table
   * @param {QueryOperation} operation
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} [options]
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} [state]
   * @param {Set<string>} [callStack]
   * @param {ConnectorMeta[]} [sources]
   * @param {{ stepIndex: number; stepId: string } | undefined} [stepContext]
   * @returns {Promise<ITable>}
   */
  async applyStep(table, operation, context, options = {}, state, callStack, sources, stepContext) {
    throwIfAborted(options.signal);
    switch (operation.type) {
      case "merge":
        return this.mergeTables(table, operation, context, options, state, callStack, sources, stepContext);
      case "append":
        return this.appendTables(table, operation, context, options, state, callStack, sources, stepContext);
      default:
        // Pure local transforms preserve the source set.
        {
          const next = applyOperation(table, operation);
          this.setTableSourceIds(next, this.getTableSourceIds(table));
          return next;
        }
    }
  }

  /**
   * Load a query source into a materialized table.
   *
   * This is exposed for advanced callers, but most hosts should use
   * `executeQuery` / `executeQueryWithMeta`.
   *
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {Set<string>} callStack
   * @param {ExecuteOptions} [options]
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} [state]
   * @returns {Promise<ITable>}
   */
  async loadSource(
    source,
    context,
    callStack,
    options = {},
    state = { credentialCache: new Map(), permissionCache: new Map(), now: () => Date.now() },
  ) {
    const result = await this.loadSourceWithMeta(source, context, callStack, options, state);
    return result.table;
  }

  /**
   * @private
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {Set<string>} callStack
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @returns {Promise<{ table: ITable, meta: ConnectorMeta, sources: ConnectorMeta[] }>}
   */
  async loadSourceWithMeta(source, context, callStack, options, state) {
    throwIfAborted(options.signal);

    options.onProgress?.({ type: "source:start", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });

    if (source.type === "range") {
      const hasHeaders = source.range.hasHeaders ?? true;
      const table = DataTable.fromGrid(source.range.values, { hasHeaders, inferTypes: true });
      this.setTableSourceIds(table, [getSourceIdForQuerySource(source) ?? "workbook:range"]);
      const meta = {
        refreshedAt: new Date(state.now()),
        schema: { columns: table.columns, inferred: true },
        rowCount: table.rowCount,
        rowCountEstimate: table.rowCount,
        provenance: { kind: "range" },
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table, meta, sources: [meta] };
    }

    if (source.type === "table") {
      const table = context.tables?.[source.table];
      if (!table) {
        throw new Error(`Unknown table '${source.table}'`);
      }
      this.setTableSourceIds(table, [getSourceIdForQuerySource(source) ?? `workbook:table:${source.table}`]);
      const meta = {
        refreshedAt: new Date(state.now()),
        schema: { columns: table.columns, inferred: true },
        rowCount: table.rowCount,
        rowCountEstimate: table.rowCount,
        provenance: { kind: "table", table: source.table },
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table, meta, sources: [meta] };
    }

    if (source.type === "query") {
      if (callStack.has(source.queryId)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${source.queryId}`);
      }

      const existing = context.queryResults?.[source.queryId];
      if (existing) {
        const { table, meta: queryMeta } = existing;
        // When query results are provided by the host (e.g. refresh orchestration),
        // the table may not have been produced by this engine instance. Ensure it
        // still carries a correct source set for privacy firewall enforcement.
        if (queryMeta?.sources) {
          this.setTableSourceIds(table, this.collectSourceIdsFromMetas(queryMeta.sources));
        }
        const meta = {
          refreshedAt: queryMeta.refreshedAt,
          schema: queryMeta.outputSchema,
          rowCount: queryMeta.outputRowCount,
          rowCountEstimate: queryMeta.outputRowCount,
          provenance: { kind: "query", queryId: source.queryId },
        };
        options.onProgress?.({
          type: "source:complete",
          queryId: Array.from(callStack).at(-1) ?? "<unknown>",
          sourceType: source.type,
        });
        return { table, meta, sources: [meta, ...queryMeta.sources] };
      }

      const target = context.queries?.[source.queryId];
      if (!target) throw new Error(`Unknown query '${source.queryId}'`);
      if (callStack.has(target.id)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${target.id}`);
      }
      const nextStack = new Set(callStack);
      nextStack.add(target.id);
      const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
      const { table, meta: queryMeta } = await this.executeQueryInternal(target, context, depOptions, state, nextStack);
      this.setTableSourceIds(table, this.collectSourceIdsFromMetas(queryMeta.sources));
      const meta = {
        refreshedAt: queryMeta.refreshedAt,
        schema: queryMeta.outputSchema,
        rowCount: queryMeta.outputRowCount,
        rowCountEstimate: queryMeta.outputRowCount,
        provenance: { kind: "query", queryId: source.queryId },
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table, meta, sources: [meta, ...queryMeta.sources] };
    }

    if (source.type === "csv" || source.type === "json" || source.type === "parquet") {
      const connector = this.connectors.get("file");
      if (!connector) throw new Error("File source requires a FileConnector");
      const request =
        source.type === "csv"
          ? { format: "csv", path: source.path, csv: source.options ?? {} }
          : source.type === "json"
            ? { format: "json", path: source.path, json: { jsonPath: source.jsonPath ?? "" } }
            : { format: "parquet", path: source.path };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("file", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);

      const cacheMode = options.cache?.mode ?? "use";
      const cacheValidation = options.cache?.validation ?? "source-state";
      /** @type {import("./connectors/types.js").SourceState} */
      let sourceState = {};
      if (this.cache && cacheMode !== "bypass" && cacheValidation === "source-state" && typeof connector.getSourceState === "function") {
        try {
          sourceState = await connector.getSourceState(request, { signal: options.signal, credentials, now: state.now });
        } catch {
          sourceState = {};
        }
      }

      // Prefer the Arrow-backed Parquet path when the host can provide raw bytes.
      // This avoids materializing row arrays and lets downstream steps stay columnar.
      if (source.type === "parquet" && this.fileAdapter?.readBinary) {
        const bytes = await this.fileAdapter.readBinary(source.path);
        throwIfAborted(options.signal);
        const { parquetToArrowTable } = await loadDataIoModule();
        const arrowTable = await parquetToArrowTable(bytes, source.options);
        const table = new ArrowTableAdapter(arrowTable);
        this.setTableSourceIds(table, [getSourceIdForQuerySource(source) ?? source.path]);
        const meta = {
          refreshedAt: new Date(state.now()),
          sourceTimestamp: sourceState.sourceTimestamp,
          etag: sourceState.etag,
          sourceKey,
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rowCount,
          rowCountEstimate: table.rowCount,
          provenance: { kind: "file", path: source.path, format: "parquet" },
        };
        options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
        return { table, meta, sources: [meta] };
      }

      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      this.setTableSourceIds(result.table, [getSourceIdForQuerySource(source) ?? source.path]);
      const meta = {
        ...result.meta,
        sourceTimestamp: sourceState.sourceTimestamp ?? result.meta.sourceTimestamp,
        etag: sourceState.etag ?? result.meta.etag,
        sourceKey,
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table: result.table, meta, sources: [meta] };
    }

    if (source.type === "api") {
      const connector = this.connectors.get("http");
      if (!connector) throw new Error("API source requires an HttpConnector");
      const request = { url: source.url, method: source.method, headers: source.headers ?? {}, auth: source.auth, responseType: "auto" };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("http", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);

      const cacheMode = options.cache?.mode ?? "use";
      const cacheValidation = options.cache?.validation ?? "source-state";
      /** @type {import("./connectors/types.js").SourceState} */
      let sourceState = {};
      if (this.cache && cacheMode !== "bypass" && cacheValidation === "source-state" && typeof connector.getSourceState === "function") {
        try {
          sourceState = await connector.getSourceState(request, { signal: options.signal, credentials, now: state.now });
        } catch {
          sourceState = {};
        }
      }

      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      this.setTableSourceIds(result.table, [getSourceIdForQuerySource(source) ?? source.url]);
      const meta = {
        ...result.meta,
        sourceTimestamp: sourceState.sourceTimestamp ?? result.meta.sourceTimestamp,
        etag: sourceState.etag ?? result.meta.etag,
        sourceKey,
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table: result.table, meta, sources: [meta] };
    }

    if (source.type === "database") {
      const connector = this.connectors.get("sql");
      if (!connector) throw new Error("Database source requires a SqlConnector");
      const connectionId = resolveDatabaseConnectionId(source, connector);
      const request = { connectionId: connectionId ?? undefined, connection: source.connection, sql: source.query };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("sql", request, state);
      const sourceKey = buildConnectorSourceKey(connector, request);

      const cacheMode = options.cache?.mode ?? "use";
      const cacheValidation = options.cache?.validation ?? "source-state";
      /** @type {import("./connectors/types.js").SourceState} */
      let sourceState = {};
      if (this.cache && cacheMode !== "bypass" && cacheValidation === "source-state" && typeof connector.getSourceState === "function") {
        try {
          sourceState = await connector.getSourceState(request, { signal: options.signal, credentials, now: state.now });
        } catch {
          sourceState = {};
        }
      }

      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      const connectionRefId = this.getEphemeralObjectId(source.connection);
      const sqlSourceId =
        connectionId ? `sql:${connectionId}` : connectionRefId ? `sql:${connectionRefId}` : getSourceIdForQuerySource(source) ?? "<unknown-sql>";
      this.setTableSourceIds(result.table, [sqlSourceId]);
      const meta = {
        ...result.meta,
        sourceTimestamp: sourceState.sourceTimestamp ?? result.meta.sourceTimestamp,
        etag: sourceState.etag ?? result.meta.etag,
        sourceKey,
      };
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { table: result.table, meta, sources: [meta] };
    }

    /** @type {never} */
    const exhausted = source;
    throw new Error(`Unsupported source type '${exhausted.type}'`);
  }

  /**
   * Execute a database query through the SQL connector while preserving the
   * normal source metadata/progress reporting.
   *
   * This is used by SQL folding execution to run a folded SQL statement with
   * parameters.
   *
   * @private
   * @param {import("./model.js").DatabaseQuerySource} source
   * @param {string} sql
   * @param {unknown[]} params
   * @param {import("./folding/dialect.js").SqlDialectName | import("./folding/dialect.js").SqlDialect} dialect
   * @param {Set<string>} callStack
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @returns {Promise<{ table: DataTable, meta: ConnectorMeta, sources: ConnectorMeta[] }>}
   */
  async loadDatabaseQueryWithMeta(source, sql, params, dialect, callStack, options, state) {
    throwIfAborted(options.signal);
    options.onProgress?.({ type: "source:start", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: "database" });

    const connector = this.connectors.get("sql");
    if (!connector) throw new Error("Database source requires a SqlConnector");

    const dialectName = typeof dialect === "string" ? dialect : dialect.name;
    let normalizedSql = sql;
    if (dialectName === "postgres") {
      normalizedSql = normalizePostgresPlaceholders(sql, params.length);
    }
    const connectionId = resolveDatabaseConnectionId(source, connector);
    const request = { connectionId: connectionId ?? undefined, connection: source.connection, sql: normalizedSql, params };
    const signatureRequest = { connectionId: connectionId ?? undefined, connection: source.connection, sql: source.query };
    // Important: permission/credential prompts should be consistent regardless of
    // whether SQL folding runs. Use the source-signature request (connection +
    // base SQL) instead of the derived folded SQL statement.
    await this.assertPermission(connector.permissionKind, { source, request: signatureRequest }, state);
    const credentials = await this.getCredentials("sql", signatureRequest, state);

    const sourceKey = buildConnectorSourceKey(connector, signatureRequest);

    const cacheMode = options.cache?.mode ?? "use";
    const cacheValidation = options.cache?.validation ?? "source-state";
    /** @type {import("./connectors/types.js").SourceState} */
    let sourceState = {};
    if (this.cache && cacheMode !== "bypass" && cacheValidation === "source-state" && typeof connector.getSourceState === "function") {
      try {
        sourceState = await connector.getSourceState(signatureRequest, { signal: options.signal, credentials, now: state.now });
      } catch {
        sourceState = {};
      }
    }

    const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
    const connectionRefId = this.getEphemeralObjectId(source.connection);
    const sqlSourceId =
      connectionId ? `sql:${connectionId}` : connectionRefId ? `sql:${connectionRefId}` : getSourceIdForQuerySource(source) ?? "<unknown-sql>";
    this.setTableSourceIds(result.table, [sqlSourceId]);
    const meta = {
      ...result.meta,
      sourceTimestamp: sourceState.sourceTimestamp ?? result.meta.sourceTimestamp,
      etag: sourceState.etag ?? result.meta.etag,
      sourceKey,
    };

    options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: "database" });
    return { table: result.table, meta, sources: [meta] };
  }

  /**
   * @private
   * @param {string} kind
   * @param {unknown} details
   * @param {{ permissionCache?: Map<string, Promise<boolean>> }} [state]
   */
  async assertPermission(kind, details, state) {
    if (!this.onPermissionRequest) return;
    const cache = state?.permissionCache;
    /** @type {any} */
    const req =
      details && typeof details === "object" && !Array.isArray(details)
        ? // @ts-ignore - runtime access
          details.request
        : undefined;
    const sourceType =
      details && typeof details === "object" && !Array.isArray(details)
        ? // @ts-ignore - runtime access
          details.source?.type
        : null;
    const connectorId =
      sourceType === "database"
        ? "sql"
        : sourceType === "api"
          ? "http"
          : sourceType === "csv" || sourceType === "json" || sourceType === "parquet"
            ? "file"
            : null;
    const keyInput = connectorId ? this.buildConnectorRequestCacheKey(connectorId, req) : req ?? details;
    const key = cache ? `${kind}:${hashValue(keyInput)}` : null;
    const allowedPromise = key
      ? cache.get(key) ?? Promise.resolve(this.onPermissionRequest(kind, details))
      : Promise.resolve(this.onPermissionRequest(kind, details));
    if (key && cache && !cache.has(key)) cache.set(key, allowedPromise);
    const allowed = await allowedPromise;
    if (allowed === false) {
      throw new Error(`Permission denied: ${kind}`);
    }
  }

  /**
   * @private
   * @param {string} connectorId
   * @param {unknown} request
   * @param {{ credentialCache: Map<string, Promise<unknown>> }} state
   * @returns {Promise<unknown>}
   */
  async getCredentials(connectorId, request, state) {
    if (!this.onCredentialRequest) return undefined;
    const keyInput = this.buildConnectorRequestCacheKey(connectorId, request);
    const key = `${connectorId}:${hashValue(keyInput)}`;
    const existing = state.credentialCache.get(key);
    if (existing) return existing;
    const promise = Promise.resolve(this.onCredentialRequest(connectorId, { request }));
    state.credentialCache.set(key, promise);
    return promise;
  }

  /**
   * @param {ITable} left
   * @param {import("./model.js").MergeOp} op
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async mergeTables(left, op, context, options = {}, state, callStack, sources = [], stepContext) {
    const queryId = callStack ? Array.from(callStack).at(-1) ?? "<unknown>" : "<unknown>";
    if (callStack?.has(op.rightQuery)) {
      throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${op.rightQuery}`);
    }

    /** @type {ITable} */
    let right;
    /** @type {QueryExecutionMeta | null} */
    let rightMeta = null;

    const existing = context.queryResults?.[op.rightQuery];
    if (existing) {
      right = existing.table;
      rightMeta = existing.meta;
      if (rightMeta?.sources) {
        this.setTableSourceIds(right, this.collectSourceIdsFromMetas(rightMeta.sources));
      }
    } else {
      const query = context.queries?.[op.rightQuery];
      if (!query) throw new Error(`Unknown query '${op.rightQuery}'`);
      if (callStack?.has(query.id)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${query.id}`);
      }
      const nextStack = callStack ? new Set(callStack) : new Set();
      nextStack.add(query.id);

      const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
      const depState = state ?? { credentialCache: new Map(), permissionCache: new Map(), now: () => Date.now() };
      const result = await this.executeQueryInternal(query, context, depOptions, depState, nextStack);
      right = result.table;
      rightMeta = result.meta;
    }

    if (rightMeta) sources.push(...rightMeta.sources);

    const leftSourceIds = this.getTableSourceIds(left);
    const rightSourceIds = this.getTableSourceIds(right);
    const combinedSourceIds = new Set([...leftSourceIds, ...rightSourceIds]);
    this.enforceFirewallForCombination({
      queryId,
      operation: "merge",
      sourceIds: combinedSourceIds,
      context,
      options,
      stepIndex: stepContext?.stepIndex,
      stepId: stepContext?.stepId,
    });

    const leftKeyIdx = left.getColumnIndex(op.leftKey);
    const rightKeyIdx = right.getColumnIndex(op.rightKey);

    /** @type {Map<unknown, number[]>} */
    const rightIndex = new Map();
    for (let rowIndex = 0; rowIndex < right.rowCount; rowIndex++) {
      const key = right.getCell(rowIndex, rightKeyIdx);
      const bucket = rightIndex.get(key);
      if (bucket) bucket.push(rowIndex);
      else rightIndex.set(key, [rowIndex]);
    }

    const rightColumnsToInclude = right.columns
      .map((col, idx) => ({ col, idx }))
      .filter(({ col }) => col.name !== op.rightKey || op.rightKey !== op.leftKey);

    const leftNames = new Set(left.columns.map((c) => c.name));
    const rightColumns = rightColumnsToInclude.map(({ col }) => {
      if (!leftNames.has(col.name)) return col;
      return { ...col, name: `${col.name}.right` };
    });

    const outColumns = [...left.columns, ...rightColumns];

    /** @type {unknown[][]} */
    const outRows = [];

    const emit = (leftRowIndex, rightRowIndex) => {
      const row = new Array(outColumns.length);
      let offset = 0;

      if (leftRowIndex == null) {
        for (let i = 0; i < left.columns.length; i++) row[offset++] = null;
      } else {
        for (let i = 0; i < left.columns.length; i++) row[offset++] = left.getCell(leftRowIndex, i);
      }

      if (rightRowIndex == null) {
        for (let i = 0; i < rightColumnsToInclude.length; i++) row[offset++] = null;
      } else {
        for (const { idx } of rightColumnsToInclude) {
          row[offset++] = right.getCell(rightRowIndex, idx);
        }
      }

      outRows.push(row);
    };

    if (op.joinType === "inner" || op.joinType === "left" || op.joinType === "full") {
      /** @type {Set<number>} */
      const matchedRight = new Set();

      for (let leftRowIndex = 0; leftRowIndex < left.rowCount; leftRowIndex++) {
        const matchIndices = rightIndex.get(left.getCell(leftRowIndex, leftKeyIdx)) ?? [];
        if (matchIndices.length === 0) {
          if (op.joinType !== "inner") emit(leftRowIndex, null);
          continue;
        }

        for (const rightIdx of matchIndices) {
          matchedRight.add(rightIdx);
          emit(leftRowIndex, rightIdx);
        }
      }

      if (op.joinType === "full") {
        for (let rightRowIndex = 0; rightRowIndex < right.rowCount; rightRowIndex++) {
          if (!matchedRight.has(rightRowIndex)) {
            emit(null, rightRowIndex);
          }
        }
      }

      const out = new DataTable(outColumns, outRows);
      this.setTableSourceIds(out, combinedSourceIds);
      return out;
    }

    if (op.joinType === "right") {
      /** @type {Map<unknown, number[]>} */
      const leftIndex = new Map();
      for (let rowIndex = 0; rowIndex < left.rowCount; rowIndex++) {
        const key = left.getCell(rowIndex, leftKeyIdx);
        const bucket = leftIndex.get(key);
        if (bucket) bucket.push(rowIndex);
        else leftIndex.set(key, [rowIndex]);
      }

      for (let rightRowIndex = 0; rightRowIndex < right.rowCount; rightRowIndex++) {
        const matchIndices = leftIndex.get(right.getCell(rightRowIndex, rightKeyIdx)) ?? [];
        if (matchIndices.length === 0) {
          emit(null, rightRowIndex);
          continue;
        }
        for (const leftIdx of matchIndices) {
          emit(leftIdx, rightRowIndex);
        }
      }

      const out = new DataTable(outColumns, outRows);
      this.setTableSourceIds(out, combinedSourceIds);
      return out;
    }

    throw new Error(`Unsupported joinType '${op.joinType}'`);
  }

  /**
   * @param {ITable} current
   * @param {import("./model.js").AppendOp} op
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async appendTables(current, op, context, options = {}, state, callStack, sources = [], stepContext) {
    const queryId = callStack ? Array.from(callStack).at(-1) ?? "<unknown>" : "<unknown>";
    const tables = [current];
    const combinedSourceIds = new Set(this.getTableSourceIds(current));
    for (const id of op.queries) {
      if (callStack?.has(id)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${id}`);
      }

      const existing = context.queryResults?.[id];
      if (existing) {
        sources.push(...existing.meta.sources);
        if (existing.meta?.sources) {
          this.setTableSourceIds(existing.table, this.collectSourceIdsFromMetas(existing.meta.sources));
        }
        for (const sourceId of this.getTableSourceIds(existing.table)) combinedSourceIds.add(sourceId);
        tables.push(existing.table);
        continue;
      }

      const query = context.queries?.[id];
      if (!query) throw new Error(`Unknown query '${id}'`);
      if (callStack?.has(query.id)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${query.id}`);
      }
      const nextStack = callStack ? new Set(callStack) : new Set();
      nextStack.add(query.id);
      const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
      const depState = state ?? { credentialCache: new Map(), permissionCache: new Map(), now: () => Date.now() };
      const { table, meta } = await this.executeQueryInternal(query, context, depOptions, depState, nextStack);
      sources.push(...meta.sources);
      for (const sourceId of this.getTableSourceIds(table)) combinedSourceIds.add(sourceId);
      tables.push(table);
    }

    this.enforceFirewallForCombination({
      queryId,
      operation: "append",
      sourceIds: combinedSourceIds,
      context,
      options,
      stepIndex: stepContext?.stepIndex,
      stepId: stepContext?.stepId,
    });

    /** @type {string[]} */
    const columns = [];
    /** @type {Map<string, { type: string }>} */
    const columnMeta = new Map();

    for (const table of tables) {
      for (const col of table.columns) {
        if (columnMeta.has(col.name)) continue;
        columnMeta.set(col.name, { type: col.type });
        columns.push(col.name);
      }
    }

    const outColumns = columns.map((name) => ({ name, type: columnMeta.get(name)?.type ?? "any" }));

    const outRows = [];
    for (const table of tables) {
      const index = new Map(table.columns.map((c, idx) => [c.name, idx]));
      for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
        outRows.push(columns.map((name) => table.getCell(rowIndex, index.get(name) ?? -1)));
      }
    }

    const out = new DataTable(outColumns, outRows);
    this.setTableSourceIds(out, combinedSourceIds);
    return out;
  }

  /**
   * Execute a query and stream the resulting grid batches to `onBatch`.
   *
   * This is intended for progressively populating a spreadsheet-like UI without needing to
   * materialize a full `table.toGrid()` result in memory.
   *
   * @param {Query} query
   * @param {QueryExecutionContext} [context]
   * @param {ExecuteOptions & {
   *   batchSize?: number;
   *   includeHeader?: boolean;
   *   onBatch: (batch: { rowOffset: number; values: unknown[][] }) => Promise<void> | void;
   * }} options
   */
  async executeQueryStreaming(query, context = {}, options) {
    const batchSize = options.batchSize ?? 1024;
    const includeHeader = options.includeHeader ?? true;
    const onBatch = options.onBatch;

    const { batchSize: _batchSize, includeHeader: _includeHeader, onBatch: _onBatch, ...executeOptions } = options;
    const table = await this.executeQuery(query, context, executeOptions);

    for await (const batch of tableToGridBatches(table, { batchSize, includeHeader })) {
      await onBatch(batch);
    }

    return table;
  }
}

/**
 * Resolve a stable identity for a database connection descriptor.
 *
 * This is used for:
 * - deterministic cache keys for database sources
 * - schema discovery caching (per engine instance)
 *
 * @param {import("./model.js").DatabaseQuerySource} source
 * @param {any} connector
 * @returns {string | null}
 */
function resolveDatabaseConnectionId(source, connector) {
  if (typeof source.connectionId === "string" && source.connectionId) {
    return source.connectionId;
  }

  const connection = source.connection;

  // Prefer connector-provided identity hook when available. This allows hosts to
  // define identities that incorporate additional fields (e.g. server + database)
  // even when a generic `connection.id` property is present.
  if (connector && typeof connector.getConnectionIdentity === "function") {
    try {
      const identity = connector.getConnectionIdentity(connection);
      if (identity != null) {
        if (typeof identity === "string") return identity;
        return hashValue(identity);
      }
    } catch {
      // Fall through to conservative fallback below.
    }
  }

  // Conservative fallback: treat `{ id: string }` as a stable identity when no
  // connector hook is available (or it returns null/throws).
  if (connection && typeof connection === "object" && !Array.isArray(connection)) {
    // @ts-ignore - runtime inspection
    if (typeof connection.id === "string" && connection.id) return connection.id;
  }

  return null;
}

/**
 * @param {ITable} table
 * @param {{ batchSize: number; includeHeader: boolean }} options
 */
async function* tableToGridBatches(table, options) {
  if (table instanceof ArrowTableAdapter) {
    const batchSize = options.batchSize;
    const includeHeader = options.includeHeader;
    const baseOffset = includeHeader ? 1 : 0;

    if (includeHeader) {
      yield { rowOffset: 0, values: [table.columns.map((c) => c.name)] };
    }

    const { arrowTableToGridBatches } = await loadDataIoModule();
    for await (const batch of arrowTableToGridBatches(table.table, { batchSize, includeHeader: false })) {
      yield { rowOffset: baseOffset + batch.rowOffset, values: batch.values };
    }
    return;
  }

  const batchSize = options.batchSize;
  const includeHeader = options.includeHeader;

  if (includeHeader) {
    yield { rowOffset: 0, values: [table.columns.map((c) => c.name)] };
  }

  const baseOffset = includeHeader ? 1 : 0;
  for (let rowStart = 0; rowStart < table.rowCount; rowStart += batchSize) {
    const slice = [];
    const end = Math.min(table.rowCount, rowStart + batchSize);
    for (let rowIndex = rowStart; rowIndex < end; rowIndex++) {
      slice.push(table.getRow(rowIndex));
    }
    yield { rowOffset: baseOffset + rowStart, values: slice };
  }
}

// Backwards-compatible exports for consumers that relied on the original engine helpers.
export { parseCsv, parseCsvCell } from "./connectors/file.js";
