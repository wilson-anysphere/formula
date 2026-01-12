/**
 * Public TypeScript surface for `@formula/power-query`.
 *
 * This package is implemented in JS + JSDoc for runtime portability. These
 * declarations define the stable API consumed by TypeScript packages (e.g. the
 * desktop app) without requiring deep imports into implementation internals.
 */

// -----------------------------------------------------------------------------
// Query model (see `src/model.js`)
// -----------------------------------------------------------------------------

export type DataType = "any" | "string" | "number" | "boolean" | "date";

export type RangeQuerySource = {
  type: "range";
  range: { values: unknown[][]; hasHeaders?: boolean };
};

export type TableQuerySource = {
  type: "table";
  table: string;
};

export type CSVQuerySource = {
  type: "csv";
  path: string;
  options?: { delimiter?: string; hasHeaders?: boolean };
};

export type JSONQuerySource = {
  type: "json";
  path: string;
  jsonPath?: string;
};

export type ParquetQuerySource = {
  type: "parquet";
  path: string;
  options?: Record<string, unknown>;
};

export type DatabaseQuerySource = {
  type: "database";
  connectionId?: string;
  connection: unknown;
  query: string;
  dialect?: "postgres" | "mysql" | "sqlite";
  columns?: string[];
};

export type APIQuerySource = {
  type: "api";
  url: string;
  method: string;
  headers?: Record<string, string>;
  auth?: { type: "oauth2"; providerId: string; scopes?: string[] | string };
};

export type QueryRefSource = {
  type: "query";
  queryId: string;
};

export type QuerySource =
  | RangeQuerySource
  | TableQuerySource
  | CSVQuerySource
  | JSONQuerySource
  | ParquetQuerySource
  | DatabaseQuerySource
  | APIQuerySource
  | QueryRefSource;

export type ComparisonPredicate = {
  type: "comparison";
  column: string;
  operator:
    | "equals"
    | "notEquals"
    | "greaterThan"
    | "greaterThanOrEqual"
    | "lessThan"
    | "lessThanOrEqual"
    | "contains"
    | "startsWith"
    | "endsWith"
    | "isNull"
    | "isNotNull";
  value?: unknown;
  caseSensitive?: boolean;
};

export type AndPredicate = { type: "and"; predicates: FilterPredicate[] };
export type OrPredicate = { type: "or"; predicates: FilterPredicate[] };
export type NotPredicate = { type: "not"; predicate: FilterPredicate };

export type FilterPredicate = ComparisonPredicate | AndPredicate | OrPredicate | NotPredicate;

export type SortSpec = {
  column: string;
  direction?: "ascending" | "descending";
  nulls?: "first" | "last";
};

export type Aggregation = {
  column: string;
  op: "sum" | "count" | "average" | "min" | "max" | "countDistinct";
  as?: string;
};

export type SelectColumnsOp = { type: "selectColumns"; columns: string[] };
export type RemoveColumnsOp = { type: "removeColumns"; columns: string[] };
export type FilterRowsOp = { type: "filterRows"; predicate: FilterPredicate };
export type SortRowsOp = { type: "sortRows"; sortBy: SortSpec[] };
export type GroupByOp = { type: "groupBy"; groupColumns: string[]; aggregations: Aggregation[] };
export type AddColumnOp = { type: "addColumn"; name: string; formula: string };
export type RenameColumnOp = { type: "renameColumn"; oldName: string; newName: string };
export type ChangeTypeOp = { type: "changeType"; column: string; newType: DataType };
export type TakeOp = { type: "take"; count: number };
export type PivotOp = { type: "pivot"; rowColumn: string; valueColumn: string; aggregation: Aggregation["op"] };
export type UnpivotOp = { type: "unpivot"; columns: string[]; nameColumn: string; valueColumn: string };
export type MergeOp = {
  type: "merge";
  rightQuery: string;
  joinType: "inner" | "left" | "right" | "full";
  leftKey: string;
  rightKey: string;
};
export type AppendOp = { type: "append"; queries: string[] };
export type DistinctRowsOp = { type: "distinctRows"; columns: string[] | null };
export type RemoveRowsWithErrorsOp = { type: "removeRowsWithErrors"; columns: string[] | null };
export type TransformColumnSpec = { column: string; formula: string; newType: DataType | null };
export type TransformColumnsOp = { type: "transformColumns"; transforms: TransformColumnSpec[] };
export type FillDownOp = { type: "fillDown"; columns: string[] };
export type ReplaceValuesOp = { type: "replaceValues"; column: string; find: unknown; replace: unknown };
export type SplitColumnOp = { type: "splitColumn"; column: string; delimiter: string };

export type QueryOperation =
  | SelectColumnsOp
  | RemoveColumnsOp
  | FilterRowsOp
  | SortRowsOp
  | GroupByOp
  | AddColumnOp
  | RenameColumnOp
  | ChangeTypeOp
  | TakeOp
  | PivotOp
  | UnpivotOp
  | MergeOp
  | AppendOp
  | DistinctRowsOp
  | RemoveRowsWithErrorsOp
  | TransformColumnsOp
  | FillDownOp
  | ReplaceValuesOp
  | SplitColumnOp;

export type QueryStep = {
  id: string;
  name: string;
  operation: QueryOperation;
  resultSchema?: unknown;
};

export type RefreshPolicy =
  | { type: "manual" }
  | { type: "interval"; intervalMs: number }
  | { type: "on-open" }
  | { type: "cron"; cron: string };

// -----------------------------------------------------------------------------
// Cron helpers (see `src/cron.js`)
// -----------------------------------------------------------------------------

export type CronSchedule = {
  source: string;
  minutes: number[];
  minutesSet: boolean[];
  hours: number[];
  hoursSet: boolean[];
  daysOfMonth: number[];
  daysOfMonthSet: boolean[];
  daysOfMonthAny: boolean;
  months: number[];
  monthsSet: boolean[];
  daysOfWeek: number[];
  daysOfWeekSet: boolean[];
  daysOfWeekAny: boolean;
};

export function parseCronExpression(expression: string): CronSchedule;

export type Query = {
  id: string;
  name: string;
  source: QuerySource;
  steps: QueryStep[];
  destination?: unknown;
  refreshPolicy?: RefreshPolicy;
};

// -----------------------------------------------------------------------------
// Table primitives (see `src/table.js`)
// -----------------------------------------------------------------------------

export type Column = { name: string; type: DataType };

export type ColumnVector = {
  length: number;
  get: (index: number) => unknown;
};

export interface ITable {
  columns: Column[];
  readonly rowCount: number;
  readonly columnCount: number;
  getColumnIndex(name: string): number;
  getColumnVector(index: number): ColumnVector;
  getCell(rowIndex: number, colIndex: number): unknown;
  getRow(rowIndex: number): unknown[];
  iterRows(): IterableIterator<unknown[]>;
  toGrid(options?: { includeHeader?: boolean }): unknown[][];
  head(limit: number): ITable;
}

export class DataTable implements ITable {
  columns: Column[];
  rows: unknown[][];

  constructor(columns: Column[], rows: unknown[][]);

  static fromGrid(grid: unknown[][], options?: { hasHeaders?: boolean; inferTypes?: boolean }): DataTable;

  get rowCount(): number;
  get columnCount(): number;

  getColumnIndex(name: string): number;
  getColumnVector(index: number): ColumnVector;
  getCell(rowIndex: number, colIndex: number): unknown;
  getRow(rowIndex: number): unknown[];
  iterRows(): IterableIterator<unknown[]>;
  toGrid(options?: { includeHeader?: boolean }): unknown[][];
  head(limit: number): DataTable;
}

/**
 * Columnar table adapter backed by an Arrow JS table.
 *
 * Note: this type intentionally avoids referencing `apache-arrow` in the public
 * surface, because Arrow support is optional and provided via `@formula/data-io`.
 */
export class ArrowTableAdapter implements ITable {
  table: any;
  columns: Column[];

  constructor(table: any, columns?: Column[]);

  get rowCount(): number;
  get columnCount(): number;

  getColumnIndex(name: string): number;
  getColumnVector(index: number): ColumnVector;
  getCell(rowIndex: number, colIndex: number): unknown;
  getRow(rowIndex: number): unknown[];
  iterRows(): IterableIterator<unknown[]>;
  toGrid(options?: { includeHeader?: boolean }): unknown[][];
  head(limit: number): ArrowTableAdapter;
}

// -----------------------------------------------------------------------------
// Cache primitives (see `src/cache/*`)
// -----------------------------------------------------------------------------

export type CacheEntry = {
  value: unknown;
  createdAtMs: number;
  expiresAtMs: number | null;
};

export type CacheLimits = {
  maxEntries?: number;
  maxBytes?: number;
};

export type CacheStore = {
  get: (key: string) => Promise<CacheEntry | null>;
  set: (key: string, entry: CacheEntry) => Promise<void>;
  delete: (key: string) => Promise<void>;
  clear?: () => Promise<void>;
  pruneExpired?: (nowMs?: number) => Promise<void>;
  prune?: (options: { nowMs: number; maxEntries?: number; maxBytes?: number }) => Promise<void>;
};

export type CacheManagerOptions = {
  store: CacheStore;
  now?: () => number;
  limits?: CacheLimits;
};

export class CacheManager {
  readonly store: CacheStore;
  readonly now: () => number;
  readonly limits: CacheLimits | null;

  constructor(options: CacheManagerOptions);

  getEntry(key: string): Promise<CacheEntry | null>;
  get(key: string): Promise<unknown | null>;
  set(key: string, value: unknown, options?: { ttlMs?: number }): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
  pruneExpired(nowMs?: number): Promise<void>;
  prune(limits?: CacheLimits): Promise<void>;
}

export class MemoryCacheStore implements CacheStore {
  get(key: string): Promise<CacheEntry | null>;
  set(key: string, entry: CacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
}

export class FileSystemCacheStore implements CacheStore {
  constructor(options: { directory: string });
  get(key: string): Promise<CacheEntry | null>;
  set(key: string, entry: CacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
}

export class IndexedDBCacheStore implements CacheStore {
  constructor(options?: { dbName?: string; storeName?: string });
  get(key: string): Promise<CacheEntry | null>;
  set(key: string, entry: CacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
}

export type CacheCryptoProvider = {
  keyVersion: number;
  encryptBytes: (
    plaintext: Uint8Array,
    aad?: Uint8Array,
  ) => Promise<{ keyVersion: number; iv: Uint8Array; tag: Uint8Array; ciphertext: Uint8Array }>;
  decryptBytes: (
    payload: { keyVersion: number; iv: Uint8Array; tag: Uint8Array; ciphertext: Uint8Array },
    aad?: Uint8Array,
  ) => Promise<Uint8Array>;
};

export class EncryptedCacheStore implements CacheStore {
  readonly store: CacheStore;
  readonly crypto: CacheCryptoProvider;
  readonly storeId: string | undefined;

  constructor(options: { store: CacheStore; crypto: CacheCryptoProvider; storeId?: string });

  get(key: string): Promise<CacheEntry | null>;
  set(key: string, entry: CacheEntry): Promise<void>;
  delete(key: string): Promise<void>;
  clear(): Promise<void>;
  pruneExpired(nowMs?: number): Promise<void>;
  prune(options: { nowMs: number; maxEntries?: number; maxBytes?: number }): Promise<void>;
}

export function createWebCryptoCacheProvider(options: { keyVersion: number; keyBytes: Uint8Array }): Promise<CacheCryptoProvider>;

// Deterministic key helpers.
export function stableStringify(value: unknown): string;
export function fnv1a64(input: string): string;
export function hashValue(value: unknown): string;

// -----------------------------------------------------------------------------
// Connectors (see `src/connectors/*`)
// -----------------------------------------------------------------------------

export type SchemaInfo = {
  columns: Array<{ name: string; type: DataType }>;
  inferred?: boolean;
};

export type ConnectorMeta = {
  refreshedAt: Date;
  sourceTimestamp?: Date;
  etag?: string;
  sourceKey?: string;
  schema: SchemaInfo;
  rowCount: number;
  rowCountEstimate?: number;
  provenance: Record<string, unknown>;
};

export type ConnectorExecuteOptions = {
  signal?: AbortSignal;
  credentials?: unknown;
  now?: () => number;
};

export type SourceState = {
  sourceTimestamp?: Date;
  etag?: string;
};

export type ConnectorResult = {
  table: ITable;
  meta: ConnectorMeta;
};

export type Connector<Request = any> = {
  id: string;
  permissionKind: string;
  getCacheKey: (request: Request) => unknown;
  execute: (request: Request, options?: ConnectorExecuteOptions) => Promise<ConnectorResult>;
  getSourceState?: (request: Request, options?: ConnectorExecuteOptions) => Promise<SourceState>;
};

export type HttpConnectorOAuth2Config = {
  providerId: string;
  scopes?: string[];
};

export type HttpConnectorCredentials = {
  headers?: Record<string, string>;
  oauth2?: HttpConnectorOAuth2Config;
};

export type HttpConnectorRequest = {
  url: string;
  method?: string;
  headers?: Record<string, string>;
  auth?: { type: "oauth2"; providerId: string; scopes?: string[] | string };
  responseType?: "auto" | "json" | "csv" | "text";
  jsonPath?: string;
};

export type HttpConnectorOptions = {
  fetch?: typeof fetch;
  fetchTable?: (
    url: string,
    options: { method: string; headers?: Record<string, string>; signal?: AbortSignal; credentials?: unknown },
  ) => Promise<DataTable>;
  oauth2Manager?: OAuth2Manager;
  oauth2RetryStatusCodes?: number[];
};

export class HttpConnector implements Connector<HttpConnectorRequest> {
  readonly id: string;
  readonly permissionKind: string;

  constructor(options?: HttpConnectorOptions);

  getCacheKey(request: HttpConnectorRequest): unknown;
  execute(request: HttpConnectorRequest, options?: ConnectorExecuteOptions): Promise<ConnectorResult>;
  getSourceState(request: HttpConnectorRequest, options?: ConnectorExecuteOptions): Promise<SourceState>;
}

export type ODataConnectorRequest = {
  url: string;
  headers?: Record<string, string>;
  query?: string;
  rowsPath?: string;
  jsonPath?: string;
  limit?: number;
  auth?: { type: "oauth2"; providerId: string; scopes?: string[] | string };
};

export type ODataConnectorOptions = {
  fetch?: typeof fetch;
  oauth2Manager?: OAuth2Manager;
  oauth2RetryStatusCodes?: number[];
};

export class ODataConnector implements Connector<ODataConnectorRequest> {
  readonly id: string;
  readonly permissionKind: string;

  constructor(options?: ODataConnectorOptions);

  getCacheKey(request: ODataConnectorRequest): unknown;
  execute(request: ODataConnectorRequest, options?: ConnectorExecuteOptions): Promise<ConnectorResult>;
  getSourceState(request: ODataConnectorRequest, options?: ConnectorExecuteOptions): Promise<SourceState>;
}

export type SharePointConnectorRequest = {
  siteUrl: string;
  mode: "contents" | "files";
  options?: { auth?: { type: "oauth2"; providerId: string; scopes?: string[] | string } | null; includeContent?: boolean; recursive?: boolean };
  url?: string;
  method?: string;
};

export type SharePointConnectorOptions = {
  fetch?: typeof fetch;
  oauth2Manager?: OAuth2Manager;
  oauth2RetryStatusCodes?: number[];
};

export class SharePointConnector implements Connector<SharePointConnectorRequest> {
  readonly id: string;
  readonly permissionKind: string;

  constructor(options?: SharePointConnectorOptions);

  getCacheKey(request: SharePointConnectorRequest): unknown;
  execute(request: SharePointConnectorRequest, options?: ConnectorExecuteOptions): Promise<ConnectorResult>;
}

export type FileConnectorRequest = {
  format: "csv" | "json" | "parquet";
  path: string;
  csv?: { delimiter?: string; hasHeaders?: boolean };
  json?: { jsonPath?: string };
};

export type FileConnectorOptions = {
  readText?: (path: string) => Promise<string>;
  readParquetTable?: (path: string, options?: { signal?: AbortSignal }) => Promise<DataTable>;
  stat?: (path: string) => Promise<{ mtimeMs: number }>;
};

export class FileConnector implements Connector<FileConnectorRequest> {
  readonly id: string;
  readonly permissionKind: string;

  constructor(options?: FileConnectorOptions);

  getCacheKey(request: FileConnectorRequest): unknown;
  execute(request: FileConnectorRequest, options?: ConnectorExecuteOptions): Promise<ConnectorResult>;
  getSourceState(request: FileConnectorRequest, options?: ConnectorExecuteOptions): Promise<SourceState>;
}

export type SqlConnectorRequest = {
  connectionId?: string;
  connection: unknown;
  sql: string;
  params?: unknown[];
};

export type SqlConnectorSchema = {
  columns: string[];
  types?: Record<string, DataType>;
};

export type SqlConnectorOptions = {
  querySql?: (
    connection: unknown,
    sql: string,
    options?: { params?: unknown[]; signal?: AbortSignal; credentials?: unknown },
  ) => Promise<DataTable>;
  getConnectionIdentity?: (connection: unknown) => unknown;
  getSchema?: (request: SqlConnectorRequest, options?: { signal?: AbortSignal; credentials?: unknown }) => Promise<SqlConnectorSchema>;
  getSourceState?: (request: SqlConnectorRequest, options?: ConnectorExecuteOptions) => Promise<SourceState>;
};

export class SqlConnector implements Connector<SqlConnectorRequest> {
  readonly id: string;
  readonly permissionKind: string;
  readonly getConnectionIdentity: (connection: unknown) => unknown;

  constructor(options?: SqlConnectorOptions);

  getCacheKey(request: SqlConnectorRequest): unknown;
  execute(request: SqlConnectorRequest, options?: ConnectorExecuteOptions): Promise<ConnectorResult>;
}

export function parseCsv(text: string, options?: { delimiter?: string }): string[][];
export function parseCsvCell(value: string): unknown;
export function parseCsvStream(chunks: AsyncIterable<string>, options?: { delimiter?: string }): AsyncGenerator<string[]>;
export function parseCsvStreamBatches(chunks: AsyncIterable<string>, options?: { delimiter?: string; batchSize?: number }): AsyncGenerator<string[][]>;

export function decodeBinaryTextStream(
  chunks: AsyncIterable<Uint8Array>,
  options?: { signal?: AbortSignal; encoding?: string },
): AsyncGenerator<string>;
export function decodeBinaryText(bytes: Uint8Array, options?: { encoding?: string }): string;

// -----------------------------------------------------------------------------
// Engine (see `src/engine.js`)
// -----------------------------------------------------------------------------

export type EngineProgressEvent =
  | { type: "cache:hit" | "cache:miss" | "cache:set"; queryId: string; cacheKey: string }
  | { type: "source:start" | "source:complete"; queryId: string; sourceType: QuerySource["type"] }
  | { type: "step:start" | "step:complete"; queryId: string; stepIndex: number; stepId: string; operation: QueryOperation["type"] }
  | {
      type: "privacy:firewall";
      queryId: string;
      phase: "folding" | "combine";
      mode: "enforce" | "warn";
      action: "prevent-folding" | "warn" | "block";
      operation: "merge" | "append";
      stepIndex?: number;
      stepId?: string;
      sources: Array<{ sourceId: string; level: PrivacyLevel }>;
      message: string;
    };

export type ExecuteOptions = {
  limit?: number;
  maxStepIndex?: number;
  signal?: AbortSignal;
  onProgress?: (event: EngineProgressEvent) => void;
  cache?: { mode?: "use" | "refresh" | "bypass"; ttlMs?: number; validation?: "none" | "source-state" };
};

export type QueryExecutionMeta = {
  queryId: string;
  startedAt: Date;
  completedAt: Date;
  refreshedAt: Date;
  sources: ConnectorMeta[];
  outputSchema: SchemaInfo;
  outputRowCount: number;
  cache?: { key: string; hit: boolean };
  folding?: unknown;
};

export type QueryExecutionResult = {
  table: ITable;
  meta: QueryExecutionMeta;
};

export type QueryExecutionContext = {
  tables?: Record<string, ITable>;
  queries?: Record<string, Query>;
  queryResults?: Record<string, QueryExecutionResult>;
  tableSignatures?: Record<string, unknown>;
  getTableSignature?: (tableName: string) => unknown;
  privacy?: { levelsBySourceId: Record<string, PrivacyLevel> };
};

export type QueryEngineHooks = {
  onPermissionRequest?: (kind: string, details: unknown) => boolean | Promise<boolean>;
  onCredentialRequest?: (connectorId: string, details: unknown) => unknown | Promise<unknown>;
};

export type QueryEngineOptions = {
  databaseAdapter?: { querySql: (connection: unknown, sql: string, options?: any) => Promise<ITable> };
  apiAdapter?: { fetchTable: (url: string, options: { method: string; headers?: Record<string, string> }) => Promise<ITable> };
  fileAdapter?: {
    readText?: (path: string) => Promise<string>;
    readBinary?: (path: string) => Promise<Uint8Array>;
    readParquetTable?: (path: string, options?: { signal?: AbortSignal }) => Promise<DataTable>;
    stat?: (path: string) => Promise<{ mtimeMs: number }>;
  };
  connectors?: Partial<{ file: FileConnector; http: HttpConnector; sql: SqlConnector } & Record<string, any>>;
  cache?: CacheManager;
  defaultCacheTtlMs?: number;
  sqlFolding?: { enabled?: boolean; dialect?: any };
  privacyMode?: "ignore" | "enforce" | "warn";
} & QueryEngineHooks;

export type QueryExecutionSession = {
  credentialCache: Map<string, Promise<unknown>>;
  permissionCache: Map<string, Promise<boolean>>;
  now?: () => number;
};

export class QueryEngine {
  readonly connectors: Map<string, any>;
  readonly cache: CacheManager | null;
  readonly fileAdapter: QueryEngineOptions["fileAdapter"] | null;

  constructor(options?: QueryEngineOptions);

  executeQuery(query: Query, context?: QueryExecutionContext, options?: ExecuteOptions): Promise<ITable>;
  executeQueryWithMeta(query: Query, context?: QueryExecutionContext, options?: ExecuteOptions): Promise<QueryExecutionResult>;

  createSession(options?: { now?: () => number }): QueryExecutionSession;

  executeQueryWithMetaInSession(
    query: Query,
    context: QueryExecutionContext | undefined,
    options: ExecuteOptions | undefined,
    session: QueryExecutionSession,
  ): Promise<QueryExecutionResult>;

  executeQueryStreaming(
    query: Query,
    context: QueryExecutionContext | undefined,
    options: ExecuteOptions & {
      batchSize?: number;
      includeHeader?: boolean;
      onBatch: (batch: { rowOffset: number; values: unknown[][] }) => Promise<void> | void;
    },
  ): Promise<ITable>;

  getCacheKey(query: Query, context?: QueryExecutionContext, options?: ExecuteOptions): Promise<string | null>;
  invalidateQueryCache(query: Query, context?: QueryExecutionContext, options?: ExecuteOptions): Promise<void>;
}

// -----------------------------------------------------------------------------
// Refresh orchestration (see `src/refresh.js` and `src/refreshGraph.js`)
// -----------------------------------------------------------------------------

export type RefreshReason = "manual" | "interval" | "on-open" | "cron";

export type RefreshJobInfo = {
  id: string;
  queryId: string;
  reason: RefreshReason;
  queuedAt: Date;
  startedAt?: Date;
  completedAt?: Date;
};

export type RefreshEvent =
  | { type: "queued"; job: RefreshJobInfo }
  | { type: "started"; job: RefreshJobInfo }
  | { type: "progress"; job: RefreshJobInfo; event: EngineProgressEvent }
  | { type: "completed"; job: RefreshJobInfo; result: QueryExecutionResult }
  | { type: "error"; job: RefreshJobInfo; error: unknown }
  | { type: "cancelled"; job: RefreshJobInfo };

export type RefreshHandle = {
  id: string;
  queryId: string;
  promise: Promise<QueryExecutionResult>;
  cancel: () => void;
};

export type RefreshState = { [queryId: string]: { policy: RefreshPolicy; lastRunAtMs?: number } };
export type RefreshStateStore = { load(): Promise<RefreshState>; save(state: RefreshState): Promise<void> };

export type RefreshManagerOptions = {
  engine: QueryEngine;
  getContext?: () => QueryExecutionContext;
  concurrency?: number;
  timers?: { setTimeout: typeof setTimeout; clearTimeout: typeof clearTimeout };
  now?: () => number;
  timezone?: "local" | "utc";
  stateStore?: RefreshStateStore;
};

export class RefreshManager {
  readonly ready: Promise<void>;

  constructor(options: RefreshManagerOptions);

  onEvent(handler: (event: RefreshEvent) => void): () => void;
  registerQuery(query: Query, policy?: RefreshPolicy): void;
  unregisterQuery(queryId: string): void;
  triggerOnOpen(queryId?: string): void;
  refresh(queryId: string, reason?: RefreshReason): RefreshHandle;
  dispose(): void;
}

// Backwards compatible shim (see `src/refresh.js`).
export class QueryScheduler {
  constructor(options: { engine: QueryEngine; getContext?: () => QueryExecutionContext; concurrency?: number });
  schedule(query: Query, onResult: (table: DataTable, meta: any) => void): void;
  unschedule(queryId: string): void;
  refreshNow(query: Query): Promise<DataTable>;
}

export function computeQueryDependencies(query: Query): string[];

export type RefreshPhase = "dependency" | "target";

export type RefreshGraphEvent =
  | { type: "queued"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo }
  | { type: "started"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo }
  | { type: "progress"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo; event: EngineProgressEvent }
  | { type: "completed"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo; result: QueryExecutionResult }
  | { type: "error"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo; error: unknown }
  | { type: "cancelled"; sessionId: string; phase: RefreshPhase; job: RefreshJobInfo };

export type RefreshAllHandle = {
  sessionId: string;
  queryIds: string[];
  promise: Promise<Record<string, QueryExecutionResult>>;
  cancel: () => void;
};

export class RefreshOrchestrator {
  constructor(options: { engine: QueryEngine; getContext?: () => QueryExecutionContext; concurrency?: number; now?: () => number });

  onEvent(handler: (event: RefreshGraphEvent) => void): () => void;
  registerQuery(query: Query): void;
  unregisterQuery(queryId: string): void;
  refreshAll(queryIds?: string[] | undefined, reason?: RefreshReason): RefreshAllHandle;
}

export class InMemoryRefreshStateStore implements RefreshStateStore {
  constructor(initialState?: RefreshState);
  load(): Promise<RefreshState>;
  save(state: RefreshState): Promise<void>;
}

// -----------------------------------------------------------------------------
// Query folding (see `src/folding/*`)
// -----------------------------------------------------------------------------

export class QueryFoldingEngine {
  constructor(options?: any);
  compile(query: Query, options?: any): any;
}

// -----------------------------------------------------------------------------
// Sheet helpers (see `src/sheet.js`)
// -----------------------------------------------------------------------------

export class InMemorySheet {
  getCell(row: number, col: number): unknown;
  setCell(row: number, col: number, value: unknown): void;
}

export function writeTableToSheet(
  table: ITable,
  sheet: InMemorySheet,
  options?: { startRow?: number; startCol?: number },
): void;

// -----------------------------------------------------------------------------
// M language helpers (see `src/m/*`)
// -----------------------------------------------------------------------------

export function parseM(text: string): any;
export function compileMToQuery(ast: any): Query;
export function prettyPrintQueryToM(query: Query): string;

// -----------------------------------------------------------------------------
// Operation helpers (see `src/steps.js`)
// -----------------------------------------------------------------------------

export function applyOperation(table: ITable, operation: QueryOperation): ITable;
export function compileRowFormula(table: DataTable, formula: string): (values: unknown[]) => unknown;

// -----------------------------------------------------------------------------
// OAuth2 helpers (see `src/oauth2/*`)
// -----------------------------------------------------------------------------

export type OAuth2ProviderConfig = {
  id: string;
  clientId: string;
  clientSecret?: string;
  tokenEndpoint: string;
  authorizationEndpoint?: string;
  redirectUri?: string;
  deviceAuthorizationEndpoint?: string;
  defaultScopes?: string[];
  authorizationParams?: Record<string, string>;
};

export type OAuth2Broker = {
  openAuthUrl: (url: string) => void | Promise<void>;
  waitForRedirect?: (redirectUri: string) => Promise<string>;
  deviceCodePrompt?: (code: string, verificationUri: string) => void | Promise<void>;
};

export type GetAccessTokenOptions = {
  providerId: string;
  scopes?: string[];
  signal?: AbortSignal;
  now?: () => number;
  forceRefresh?: boolean;
};

export type OAuth2AccessTokenResult = {
  accessToken: string;
  expiresAtMs: number | null;
  refreshToken: string | null;
};

export class OAuth2Manager {
  readonly tokenStore: any;

  constructor(options?: {
    tokenStore?: any;
    fetch?: typeof fetch;
    now?: () => number;
    clockSkewMs?: number;
    persistAccessToken?: boolean;
  });

  registerProvider(config: OAuth2ProviderConfig): void;
  getProvider(providerId: string): OAuth2ProviderConfig;
  makeStoreKey(providerId: string, scopes: string[] | undefined): { providerId: string; scopesHash: string };
  static keyString(key: { providerId: string; scopesHash: string }): string;
  getAccessToken(options: GetAccessTokenOptions): Promise<OAuth2AccessTokenResult>;
  exchangeAuthorizationCode(options: {
    providerId: string;
    code: string;
    redirectUri?: string;
    codeVerifier?: string;
    scopes?: string[] | string;
    signal?: AbortSignal;
    now?: () => number;
  }): Promise<OAuth2AccessTokenResult>;
  authorizeWithPkce(options: {
    providerId: string;
    scopes?: string[] | string;
    broker: OAuth2Broker;
    signal?: AbortSignal;
    now?: () => number;
  }): Promise<OAuth2AccessTokenResult>;
  authorizeWithDeviceCode(options: {
    providerId: string;
    scopes?: string[] | string;
    broker: OAuth2Broker;
    signal?: AbortSignal;
    now?: () => number;
  }): Promise<OAuth2AccessTokenResult>;
  persistTokens(key: any, entry: any): Promise<void>;
  clearTokens(options: { providerId: string; scopes?: string[] | string }): Promise<void>;
}

export class OAuth2TokenClient {
  constructor(options?: { fetch?: typeof fetch; now?: () => number });
}

export class OAuth2TokenError extends Error {}

export class InMemoryOAuthTokenStore {
  constructor(snapshot?: Record<string, any>);

  static keyString(key: any): string;

  snapshot(): Record<string, any>;
  get(key: any): Promise<any>;
  set(key: any, entry: any): Promise<void>;
  delete(key: any): Promise<void>;
}

export class CredentialStoreOAuthTokenStore {
  constructor(store: CredentialStore);
}

export function createCodeVerifier(): string;
export function createCodeChallenge(codeVerifier: string): string;
export function normalizeScopes(scopes?: string[] | string | null | undefined): { scopes: string[]; scopesHash: string };

// -----------------------------------------------------------------------------
// Privacy helpers (see `src/privacy/*`)
// -----------------------------------------------------------------------------

export type PrivacyLevel = "public" | "organizational" | "private" | "unknown";

export function getPrivacyLevel(levelsBySourceId: Record<string, PrivacyLevel> | undefined, sourceId: string | null | undefined): PrivacyLevel;
export function privacyRank(level: PrivacyLevel): number;

export function normalizeFilePath(path: string): string;
export function getFileSourceId(path: string): string;
export function getHttpSourceId(url: string): string;
export function getSqlSourceId(connectionId: string): string;
export function getSourceIdForProvenance(provenance: any): string | null;
export function getSourceIdForQuerySource(source: QuerySource): string | null;

// -----------------------------------------------------------------------------
// Credentials helpers (see `src/credentials/*`)
// -----------------------------------------------------------------------------

export type CredentialEntry = { id: string; secret: unknown };

export type CredentialStore = {
  get: (scope: any) => Promise<CredentialEntry | null>;
  set: (scope: any, secret: unknown) => Promise<CredentialEntry>;
  delete: (scope: any) => Promise<void>;
};

export class CredentialManager {
  constructor(options: { store: CredentialStore; prompt?: (args: any) => Promise<unknown | null | undefined> });
  onCredentialRequest(connectorId: string, details: unknown): Promise<any>;
}

export class InMemoryCredentialStore implements CredentialStore {
  get(scope: any): Promise<CredentialEntry | null>;
  set(scope: any, secret: unknown): Promise<CredentialEntry>;
  delete(scope: any): Promise<void>;
}

export class KeychainCredentialStore implements CredentialStore {
  constructor(options: { keychainProvider: any; service?: string; accountPrefix?: string });
  get(scope: any): Promise<CredentialEntry | null>;
  set(scope: any, secret: unknown): Promise<CredentialEntry>;
  delete(scope: any): Promise<void>;
}

export type HttpCredentialScope = { type: "http"; origin: string; realm?: string | null };
export type FileCredentialScope = { type: "file"; match: "exact" | "prefix"; path: string };
export type SqlCredentialScope = { type: "sql"; server: string; database?: string | null; user?: string | null };
export type OAuth2CredentialScope = { type: "oauth2"; providerId: string; scopesHash: string };
export type CredentialScope = HttpCredentialScope | FileCredentialScope | SqlCredentialScope | OAuth2CredentialScope;

export function httpScope(args: { url: string; realm?: string | null }): HttpCredentialScope;
export function fileScopeExact(args: { path: string }): FileCredentialScope;
export function fileScopePrefix(args: { pathPrefix: string }): FileCredentialScope;
export function sqlScope(args: { server: string; database?: string | null; user?: string | null }): SqlCredentialScope;
export function oauth2Scope(args: { providerId: string; scopesHash: string }): OAuth2CredentialScope;

export const scopes: {
  httpScope: typeof httpScope;
  fileScopeExact: typeof fileScopeExact;
  fileScopePrefix: typeof fileScopePrefix;
  sqlScope: typeof sqlScope;
  oauth2Scope: typeof oauth2Scope;
};

export function credentialScopeKey(scope: any): string;
export function randomId(bytes?: number): string;

// -----------------------------------------------------------------------------
// Value types (see `src/values.js`)
// -----------------------------------------------------------------------------

export const MS_PER_DAY: number;

export class PqDecimal {
  value: string;
  constructor(value: string | number | bigint);
  toString(): string;
  valueOf(): number;
}

export class PqTime {
  milliseconds: number;
  constructor(milliseconds: number);
  static from(input: string): PqTime | null;
  toString(): string;
  valueOf(): number;
}

export class PqDuration {
  milliseconds: number;
  constructor(milliseconds: number);
  static from(input: string): PqDuration | null;
  toString(): string;
  valueOf(): number;
}

export class PqDateTimeZone {
  date: Date;
  offsetMinutes: number;
  constructor(dateUtc: Date, offsetMinutes: number);
  static from(input: string): PqDateTimeZone | null;
  toDate(): Date;
  toString(): string;
  valueOf(): number;
}
