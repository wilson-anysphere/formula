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
 *   type: "database";
 *   connection: unknown;
 *   query: string;
 *   /**
 *    * Optional column names for the query result.
 *    *
 *    * Some folding operations (e.g. `renameColumn`, `changeType`) require an
 *    * explicit projection list to avoid duplicate output columns. When this
 *    * metadata is available, the folding engine can push those operations down.
 *    */
 *   columns?: string[];
 }} DatabaseQuerySource
 */

/**
 * @typedef {{
 *   type: "api";
 *   url: string;
 *   method: string;
 *   headers?: Record<string, string>;
 * }} APIQuerySource
 */

/**
 * @typedef {{
 *   type: "query";
 *   queryId: string;
 * }} QueryRefSource
 */

/**
 * @typedef {RangeQuerySource | TableQuerySource | CSVQuerySource | JSONQuerySource | DatabaseQuerySource | APIQuerySource | QueryRefSource} QuerySource
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
 * @typedef {{ type: "fillDown"; columns: string[] }} FillDownOp
 * @typedef {{ type: "replaceValues"; column: string; find: unknown; replace: unknown }} ReplaceValuesOp
 * @typedef {{ type: "splitColumn"; column: string; delimiter: string }} SplitColumnOp
 *
 * @typedef {SelectColumnsOp | RemoveColumnsOp | FilterRowsOp | SortRowsOp | GroupByOp | AddColumnOp | RenameColumnOp | ChangeTypeOp | TakeOp | PivotOp | UnpivotOp | MergeOp | AppendOp | FillDownOp | ReplaceValuesOp | SplitColumnOp} QueryOperation
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
 *   type: "cron",
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
