import { applyOperation } from "./steps.js";
import { arrowTableToGridBatches, parquetToArrowTable } from "../../data-io/src/index.js";
import { ArrowTableAdapter } from "./arrowTable.js";
import { DataTable } from "./table.js";

import { hashValue } from "./cache/key.js";
import { deserializeTable, serializeTable } from "./cache/serialize.js";
import { FileConnector } from "./connectors/file.js";
import { HttpConnector } from "./connectors/http.js";
import { SqlConnector } from "./connectors/sql.js";
import { QueryFoldingEngine } from "./folding/sql.js";
import { computeParquetProjectionColumns, computeParquetRowLimit } from "./parquetProjection.js";

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
 * }} QueryExecutionContext
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
 * }} EngineProgressEvent
 */

/**
 * @typedef {{
 *   limit?: number;
 *   // Execute up to and including this step index.
 *   maxStepIndex?: number;
 *   signal?: AbortSignal;
 *   onProgress?: (event: EngineProgressEvent) => void;
 *   cache?: { mode?: "use" | "refresh" | "bypass"; ttlMs?: number };
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
 */

/**
 * @typedef {{
 *   table: ITable;
 *   meta: QueryExecutionMeta;
 * }} QueryExecutionResult
 */

/**
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
 * } | undefined} [fileAdapter]
 *   Backwards-compatible adapter from the prototype. Prefer supplying a `FileConnector`.
 * @property {Partial<{ file: FileConnector; http: HttpConnector; sql: SqlConnector } & Record<string, any>> | undefined} [connectors]
 * @property {import("./cache/cache.js").CacheManager | undefined} [cache]
 * @property {number | undefined} [defaultCacheTtlMs]
 * @property {{ enabled?: boolean; dialect?: import("./folding/dialect.js").SqlDialectName | import("./folding/dialect.js").SqlDialect } | undefined} [sqlFolding]
 *   When enabled and a dialect is known (either via `source.dialect` or this
 *   default dialect), the engine will execute a foldable prefix of operations
 *   in the source database via `QueryFoldingEngine`.
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
 * @param {ConnectorMeta} meta
 * @returns {{ refreshedAtMs: number, sourceTimestampMs?: number, schema: any, rowCount: number, rowCountEstimate?: number, provenance: any }}
 */
function serializeConnectorMeta(meta) {
  return {
    refreshedAtMs: meta.refreshedAt.getTime(),
    sourceTimestampMs: meta.sourceTimestamp ? meta.sourceTimestamp.getTime() : undefined,
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
  };
}

export class QueryEngine {
  /**
   * @param {QueryEngineOptions} [options]
   */
  constructor(options = {}) {
    this.onPermissionRequest = options.onPermissionRequest ?? null;
    this.onCredentialRequest = options.onCredentialRequest ?? null;
    this.fileAdapter = options.fileAdapter ?? null;

    /** @type {Map<string, any>} */
    this.connectors = new Map();

    const fileConnector =
      options.connectors?.file ??
      new FileConnector({ readText: options.fileAdapter?.readText, readParquetTable: options.fileAdapter?.readParquetTable });
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
    /** @type {Map<string, Promise<unknown>>} */
    const credentialCache = new Map();
    /** @type {Map<string, Promise<boolean>>} */
    const permissionCache = new Map();

    const now = () => Date.now();
    const result = await this.executeQueryInternal(
      query,
      context,
      options,
      { credentialCache, permissionCache, now },
      new Set([query.id]),
    );

    return result;
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
    const cacheTtlMs = options.cache?.ttlMs ?? this.defaultCacheTtlMs ?? undefined;

    /** @type {string | null} */
    let cacheKey = null;
    if (this.cache && cacheMode !== "bypass") {
      cacheKey = await this.computeCacheKey(query, context, options, state, callStack);
      if (cacheKey && cacheMode === "use") {
        const cached = await this.cache.getEntry(cacheKey);
        if (cached) {
          options.onProgress?.({ type: "cache:hit", queryId: query.id, cacheKey });
          const completedAt = new Date(state.now());
          const payload = /** @type {any} */ (cached.value);
          const table = deserializeTable(payload.table);
          const meta = deserializeQueryMeta(
            payload.meta,
            startedAt,
            completedAt,
            { key: cacheKey, hit: true },
          );
          return { table, meta };
        }
        options.onProgress?.({ type: "cache:miss", queryId: query.id, cacheKey });
      }
    }

    /** @type {ConnectorMeta[]} */
    const sources = [];

    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);

    /** @type {import("./folding/sql.js").CompiledQueryPlan | null} */
    let foldedPlan = null;
    /** @type {import("./folding/dialect.js").SqlDialectName | import("./folding/dialect.js").SqlDialect | null} */
    let foldedDialect = null;
    if (this.sqlFoldingEnabled && query.source.type === "database") {
      const dialect = query.source.dialect ?? this.sqlFoldingDialect;
      if (dialect) {
        foldedPlan = this.foldingEngine.compile({ ...query, steps }, { dialect, queries: context.queries ?? undefined });
        foldedDialect = dialect;
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
      const sourceResult = await this.loadDatabaseQueryWithMeta(
        query.source,
        foldedPlan.sql,
        foldedPlan.params,
        foldedDialect,
        callStack,
        options,
        state,
      );
      sources.push(...sourceResult.sources);
      table = sourceResult.table;

      if (foldedPlan.type === "hybrid" && foldedPlan.localSteps.length > 0) {
        const offset = steps.indexOf(foldedPlan.localSteps[0]);
        table = await this.executeSteps(
          table,
          foldedPlan.localSteps,
          context,
          options,
          state,
          callStack,
          sources,
          offset >= 0 ? offset : 0,
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
    };

    if (this.cache && cacheKey && cacheMode !== "bypass" && table instanceof DataTable) {
      await this.cache.set(
        cacheKey,
        { version: 1, table: serializeTable(table), meta: serializeQueryMeta(meta) },
        { ttlMs: cacheTtlMs },
      );
      options.onProgress?.({ type: "cache:set", queryId: query.id, cacheKey });
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
    return `pq:v1:${hashValue(signature)}`;
  }

  /**
   * @private
   * @param {Query} query
   * @param {QueryExecutionContext} context
   * @param {ExecuteOptions} options
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<unknown>}
   */
  async buildQuerySignature(query, context, options, state, callStack) {
    const maxStepIndex = options.maxStepIndex ?? query.steps.length - 1;
    const steps = query.steps.slice(0, maxStepIndex + 1);

    /** @type {Record<string, unknown>} */
    const signature = {
      source: await this.buildSourceSignature(query.source, context, state, callStack),
      steps: steps.map((s) => s.operation),
      options: { limit: options.limit ?? null, maxStepIndex: options.maxStepIndex ?? null },
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
          signature[`merge:${step.operation.rightQuery}`] = await this.buildQuerySignature(dep, context, {}, state, nextStack);
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
            signature[`append:${id}`] = await this.buildQuerySignature(dep, context, {}, state, nextStack);
          }
        }
      }
    }

    return signature;
  }

  /**
   * @private
   * @param {QuerySource} source
   * @param {QueryExecutionContext} context
   * @param {{ credentialCache: Map<string, Promise<unknown>>, permissionCache: Map<string, Promise<boolean>>, now: () => number }} state
   * @param {Set<string>} callStack
   * @returns {Promise<unknown>}
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
      return { type: "query", queryId: source.queryId, query: await this.buildQuerySignature(target, context, {}, state, nextStack) };
    }

    if (source.type === "range") {
      return { type: "range", hasHeaders: source.range.hasHeaders ?? true, values: source.range.values };
    }
    if (source.type === "table") {
      return { type: "table", table: source.table };
    }

    if (source.type === "csv" || source.type === "json" || source.type === "parquet") {
      const connector = this.connectors.get("file");
      if (!connector) return { type: source.type, missingConnector: "file" };
      const request =
        source.type === "csv"
          ? { format: "csv", path: source.path, csv: source.options ?? {} }
          : source.type === "json"
            ? { format: "json", path: source.path, json: { jsonPath: source.jsonPath ?? "" } }
            : { format: "parquet", path: source.path };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("file", request, state);
      return { type: source.type, request: connector.getCacheKey(request), credentialsHash: hashValue(credentials ?? null) };
    }

    if (source.type === "api") {
      const connector = this.connectors.get("http");
      if (!connector) return { type: "api", missingConnector: "http" };
      const request = {
        url: source.url,
        method: source.method,
        headers: source.headers ?? {},
        responseType: "auto",
      };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("http", request, state);
      return { type: "api", request: connector.getCacheKey(request), credentialsHash: hashValue(credentials ?? null) };
    }

    if (source.type === "database") {
      const connector = this.connectors.get("sql");
      if (!connector) return { type: "database", missingConnector: "sql" };
      const request = { connection: source.connection, sql: source.query };
      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("sql", request, state);
      return {
        type: "database",
        dialect: source.dialect ?? null,
        request: connector.getCacheKey(request),
        credentialsHash: hashValue(credentials ?? null),
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
      options.onProgress?.({
        type: "step:start",
        queryId,
        stepIndex: i + stepIndexOffset,
        stepId: step.id,
        operation: step.operation.type,
      });
      current = await this.applyStep(current, step.operation, context, options, state, callStack, sources);
      options.onProgress?.({
        type: "step:complete",
        queryId,
        stepIndex: i + stepIndexOffset,
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
   * @returns {Promise<ITable>}
   */
  async applyStep(table, operation, context, options = {}, state, callStack, sources) {
    throwIfAborted(options.signal);
    switch (operation.type) {
      case "merge":
        return this.mergeTables(table, operation, context, options, state, callStack, sources);
      case "append":
        return this.appendTables(table, operation, context, options, state, callStack, sources);
      default:
        return applyOperation(table, operation);
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
      const target = context.queries?.[source.queryId];
      if (!target) throw new Error(`Unknown query '${source.queryId}'`);
      if (callStack.has(target.id)) {
        throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${target.id}`);
      }
      const nextStack = new Set(callStack);
      nextStack.add(target.id);
      const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
      const { table, meta: queryMeta } = await this.executeQueryInternal(target, context, depOptions, state, nextStack);
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

      // Prefer the Arrow-backed Parquet path when the host can provide raw bytes.
      // This avoids materializing row arrays and lets downstream steps stay columnar.
      if (source.type === "parquet" && this.fileAdapter?.readBinary) {
        const bytes = await this.fileAdapter.readBinary(source.path);
        throwIfAborted(options.signal);
        const arrowTable = await parquetToArrowTable(bytes, source.options);
        const table = new ArrowTableAdapter(arrowTable);
        const meta = {
          refreshedAt: new Date(state.now()),
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rowCount,
          rowCountEstimate: table.rowCount,
          provenance: { kind: "file", path: source.path, format: "parquet" },
        };
        options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
        return { table, meta, sources: [meta] };
      }

      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { ...result, sources: [result.meta] };
    }

    if (source.type === "api") {
      const connector = this.connectors.get("http");
      if (!connector) throw new Error("API source requires an HttpConnector");
      const request = { url: source.url, method: source.method, headers: source.headers ?? {}, responseType: "auto" };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("http", request, state);
      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { ...result, sources: [result.meta] };
    }

    if (source.type === "database") {
      const connector = this.connectors.get("sql");
      if (!connector) throw new Error("Database source requires a SqlConnector");
      const request = { connection: source.connection, sql: source.query };

      await this.assertPermission(connector.permissionKind, { source, request }, state);
      const credentials = await this.getCredentials("sql", request, state);
      const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });
      options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: source.type });
      return { ...result, sources: [result.meta] };
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
      let idx = 0;
      normalizedSql = sql.replaceAll("?", () => `$${++idx}`);
    }
    const request = { connection: source.connection, sql: normalizedSql, params };
    await this.assertPermission(connector.permissionKind, { source, request }, state);
    const credentials = await this.getCredentials("sql", request, state);
    const result = await connector.execute(request, { signal: options.signal, credentials, now: state.now });

    options.onProgress?.({ type: "source:complete", queryId: Array.from(callStack).at(-1) ?? "<unknown>", sourceType: "database" });
    return { ...result, sources: [result.meta] };
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
    const key = cache ? `${kind}:${hashValue(details)}` : null;
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
    const key = `${connectorId}:${hashValue(request)}`;
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
  async mergeTables(left, op, context, options = {}, state, callStack, sources = []) {
    const query = context.queries?.[op.rightQuery];
    if (!query) throw new Error(`Unknown query '${op.rightQuery}'`);
    if (callStack?.has(query.id)) {
      throw new Error(`Query reference cycle detected: ${Array.from(callStack).join(" -> ")} -> ${query.id}`);
    }
    const nextStack = callStack ? new Set(callStack) : new Set();
    nextStack.add(query.id);

    const depOptions = { ...options, limit: undefined, maxStepIndex: undefined };
    const depState = state ?? { credentialCache: new Map(), permissionCache: new Map(), now: () => Date.now() };
    const { table: right, meta: rightMeta } = await this.executeQueryInternal(query, context, depOptions, depState, nextStack);
    sources.push(...rightMeta.sources);

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

      return new DataTable(outColumns, outRows);
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

      return new DataTable(outColumns, outRows);
    }

    throw new Error(`Unsupported joinType '${op.joinType}'`);
  }

  /**
   * @param {ITable} current
   * @param {import("./model.js").AppendOp} op
   * @param {QueryExecutionContext} context
   * @returns {Promise<DataTable>}
   */
  async appendTables(current, op, context, options = {}, state, callStack, sources = []) {
    const tables = [current];
    for (const id of op.queries) {
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
      tables.push(table);
    }

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

    return new DataTable(outColumns, outRows);
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
   *   onBatch: (batch: { rowOffset: number; values: unknown[][] }) => Promise<void> | void;
   * }} options
   */
  async executeQueryStreaming(query, context = {}, options) {
    const batchSize = options.batchSize ?? 1024;
    const onBatch = options.onBatch;

    const { batchSize: _batchSize, onBatch: _onBatch, ...executeOptions } = options;
    const table = await this.executeQuery(query, context, executeOptions);

    for await (const batch of tableToGridBatches(table, { batchSize, includeHeader: true })) {
      await onBatch(batch);
    }

    return table;
  }
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
