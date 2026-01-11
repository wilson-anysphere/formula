/**
 * Query model based on `docs/07-power-features.md`.
 *
 * This file is intentionally JS + JSDoc so it can run in Node without a TS build.
 */

/**
 * @typedef {"any" | "string" | "number" | "boolean" | "date"} DataType
 */

/**
 * @typedef {{
 *   type: "range";
 *   // A range is represented as a 2D grid (rows x columns). The first row is
 *   // treated as headers by default.
 *   range: { values: unknown[][]; hasHeaders?: boolean };
 * }} RangeQuerySource
 */

/**
 * @typedef {{
 *   type: "table";
 *   table: string;
 * }} TableQuerySource
 */

/**
 * @typedef {{
 *   type: "csv";
 *   path: string;
 *   options?: { delimiter?: string; hasHeaders?: boolean };
 * }} CSVQuerySource
 */

/**
 * @typedef {{
 *   type: "json";
 *   path: string;
 *   jsonPath?: string;
 * }} JSONQuerySource
 */

/**
 * @typedef {{
 *   type: "parquet";
 *   path: string;
 *   // Passed through to `@formula/data-io`'s `parquetToArrowTable` reader options.
 *   options?: Record<string, unknown>;
 * }} ParquetQuerySource
 */

/**
 * @typedef {{
 *   type: "database";
 *   /**
 *    * Optional stable identifier for the database connection.
 *    *
 *    * This must be stable across refreshes (and JSON-serializable) so it can be
 *    * used for deterministic cache keys and to decide whether folding can safely
 *    * `merge`/`append` across queries.
 *    *
 *    * If omitted, the engine will try to derive an identity from the connection
 *    * descriptor:
 *    * - Prefer `SqlConnector.getConnectionIdentity(connection)` when available.
 *    * - Fall back to `connection.id` when present.
 *    *
 *    * When no stable identity is available, database sources are treated as
 *    * non-cacheable to avoid incorrect reuse across connections.
 *    */
 *   connectionId?: string;
 *   connection: unknown;
 *   query: string;
 *   /**
 *    * SQL dialect name used for query folding / compilation.
 *    *
 *    * This is optional because some host apps might not know the exact backend
 *    * dialect, but pushdown folding should only be enabled when the dialect is
 *    * known to avoid generating incompatible SQL.
 *    */
 *   dialect?: "postgres" | "mysql" | "sqlite";
 *   /**
 *    * Optional column names for the query result.
 *    *
 *    * Some folding operations (e.g. `renameColumn`, `changeType`) require an
 *    * explicit projection list to avoid duplicate output columns. When this
 *    * metadata is available, the folding engine can push those operations down.
 *    *
 *    * When omitted and SQL folding is enabled, the engine may attempt to discover
 *    * the schema via an optional `SqlConnector.getSchema` hook.
 *    */
 *   columns?: string[];
 * }} DatabaseQuerySource
 */

/**
 * @typedef {{
 *   type: "api";
 *   url: string;
 *   method: string;
 *   headers?: Record<string, string>;
 *   // Optional per-request auth configuration. This is forwarded to
 *   // `HttpConnectorRequest.auth`.
 *   auth?: { type: "oauth2"; providerId: string; scopes?: string[] };
 * }} APIQuerySource
 */

/**
 * @typedef {{
 *   type: "query";
 *   queryId: string;
 * }} QueryRefSource
 */

/**
 * @typedef {RangeQuerySource | TableQuerySource | CSVQuerySource | JSONQuerySource | ParquetQuerySource | DatabaseQuerySource | APIQuerySource | QueryRefSource} QuerySource
 */

/**
 * @typedef {{
 *   type: "comparison";
 *   column: string;
 *   operator:
 *     | "equals"
 *     | "notEquals"
 *     | "greaterThan"
 *     | "greaterThanOrEqual"
 *     | "lessThan"
 *     | "lessThanOrEqual"
 *     | "contains"
 *     | "startsWith"
 *     | "endsWith"
 *     | "isNull"
 *     | "isNotNull";
 *   value?: unknown;
 *   caseSensitive?: boolean;
 * }} ComparisonPredicate
 */

/**
 * @typedef {{ type: "and"; predicates: FilterPredicate[] }} AndPredicate
 */

/**
 * @typedef {{ type: "or"; predicates: FilterPredicate[] }} OrPredicate
 */

/**
 * @typedef {{ type: "not"; predicate: FilterPredicate }} NotPredicate
 */

/**
 * @typedef {ComparisonPredicate | AndPredicate | OrPredicate | NotPredicate} FilterPredicate
 */

/**
 * @typedef {{
 *   column: string;
 *   direction?: "ascending" | "descending";
 *   nulls?: "first" | "last";
 * }} SortSpec
 */

/**
 * @typedef {{
 *   column: string;
 *   op: "sum" | "count" | "average" | "min" | "max" | "countDistinct";
 *   as?: string;
 * }} Aggregation
 */

/**
 * @typedef {{ type: "selectColumns"; columns: string[] }} SelectColumnsOp
 * @typedef {{ type: "removeColumns"; columns: string[] }} RemoveColumnsOp
 * @typedef {{ type: "filterRows"; predicate: FilterPredicate }} FilterRowsOp
 * @typedef {{ type: "sortRows"; sortBy: SortSpec[] }} SortRowsOp
 * @typedef {{ type: "groupBy"; groupColumns: string[]; aggregations: Aggregation[] }} GroupByOp
 * @typedef {{ type: "addColumn"; name: string; formula: string }} AddColumnOp
 * @typedef {{ type: "renameColumn"; oldName: string; newName: string }} RenameColumnOp
 * @typedef {{ type: "changeType"; column: string; newType: DataType }} ChangeTypeOp
 * @typedef {{ type: "take"; count: number }} TakeOp
 * @typedef {{ type: "pivot"; rowColumn: string; valueColumn: string; aggregation: Aggregation["op"] }} PivotOp
 * @typedef {{ type: "unpivot"; columns: string[]; nameColumn: string; valueColumn: string }} UnpivotOp
 * @typedef {{ type: "merge"; rightQuery: string; joinType: "inner" | "left" | "right" | "full"; leftKey: string; rightKey: string }} MergeOp
 * @typedef {{ type: "append"; queries: string[] }} AppendOp
 * @typedef {{ type: "distinctRows"; columns: string[] | null }} DistinctRowsOp
 * @typedef {{ type: "removeRowsWithErrors"; columns: string[] | null }} RemoveRowsWithErrorsOp
 * @typedef {{ column: string; formula: string; newType: DataType | null }} TransformColumnSpec
 * @typedef {{ type: "transformColumns"; transforms: TransformColumnSpec[] }} TransformColumnsOp
 * @typedef {{ type: "fillDown"; columns: string[] }} FillDownOp
 * @typedef {{ type: "replaceValues"; column: string; find: unknown; replace: unknown }} ReplaceValuesOp
 * @typedef {{ type: "splitColumn"; column: string; delimiter: string }} SplitColumnOp
 *
 * @typedef {SelectColumnsOp | RemoveColumnsOp | FilterRowsOp | SortRowsOp | GroupByOp | AddColumnOp | RenameColumnOp | ChangeTypeOp | TakeOp | PivotOp | UnpivotOp | MergeOp | AppendOp | DistinctRowsOp | RemoveRowsWithErrorsOp | TransformColumnsOp | FillDownOp | ReplaceValuesOp | SplitColumnOp} QueryOperation
 */

/**
 * @typedef {{
 *   id: string;
 *   name: string;
 *   operation: QueryOperation;
 *   resultSchema?: unknown;
 * }} QueryStep
 */

/**
 * @typedef {{
 *   type: "manual"
 * } | {
 *   type: "interval",
 *   intervalMs: number
 * } | {
 *   type: "on-open"
 * } | {
 *   type: "cron",
 *   /**
 *    * 5-field cron schedule: `minute hour day-of-month month day-of-week`
 *    *
 *    * Example: `*\/15 9-17 * * 1-5` (every 15 minutes during business hours, Mon-Fri)
 *    */
 *   cron: string
 * }} RefreshPolicy
 */

/**
 * @typedef {{
 *   id: string;
 *   name: string;
 *   source: QuerySource;
 *   steps: QueryStep[];
 *   destination?: unknown;
 *   refreshPolicy?: RefreshPolicy;
 * }} Query
 */

export {};
