/**
 * Scalar cell input/value transported over the worker RPC boundary.
 *
 * `null` represents an empty cell. When used as an *input* (e.g. `setCell("A1", null)`),
 * the engine should clear the stored cell entry (sparse semantics) rather than storing
 * an explicit blank.
 *
 * Workbook JSON exports should omit empty cells entirely instead of emitting `"A1": null`
 * entries to keep payloads compact.
 */
export type CellScalar = number | string | boolean | null;

/**
 * Workbook-wide calculation settings (Excel `calcPr`).
 *
 * This mirrors the JS-friendly DTO exposed by `crates/formula-wasm` via
 * `WasmWorkbook.getCalcSettings()` / `WasmWorkbook.setCalcSettings()`.
 */
export type CalculationMode = "automatic" | "automaticNoTable" | "manual";

export interface IterativeCalcSettings {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
}

export interface CalcSettings {
  calculationMode: CalculationMode;
  calculateBeforeSave: boolean;
  fullPrecision: boolean;
  fullCalcOnLoad: boolean;
  iterative: IterativeCalcSettings;
}

/**
 * JSON-friendly rich cell value transported over the worker RPC boundary.
 *
 * This mirrors `formula_model::CellValue`'s `{type,value}` tagged schema.
 *
 * Note: This type is intentionally minimal (best-effort) to avoid coupling the TS
 * public API too tightly to the Rust model while rich values are still evolving.
 */
export type CellRef = { row: number; col: number };

export type RichTextRunStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  color?: unknown;
  font?: string;
  size_100pt?: number;
};

export type RichTextRun = {
  start: number;
  end: number;
  style: RichTextRunStyle;
};

export type RichTextValue = {
  text: string;
  // Use a readonly array so callers can conveniently construct rich text values with
  // `as const` without fighting TS's readonly tuple inference.
  runs: ReadonlyArray<RichTextRun>;
};

export type EntityValue = {
  /** Optional discriminator (e.g. "stock", "geography"). */
  entityType?: string;
  /** Optional entity id (e.g. "AAPL"). */
  entityId?: string;
  /**
   * User-visible string representation (what Excel renders in the grid).
   *
   * Note: older payloads may use `display`.
   */
  displayValue?: string;
  /** Legacy alias for `displayValue` (accepted by formula-model). */
  display?: string;
  properties?: Record<string, CellValueRich>;
};

export type RecordValue = {
  fields?: Record<string, CellValueRich>;
  displayField?: string;
  /** Optional precomputed display string (legacy / fallback). */
  displayValue?: string;
  /** Legacy alias for `displayValue` (accepted by formula-model). */
  display?: string;
};

export type ImageValue = {
  imageId: string;
  altText?: string;
  width?: number;
  height?: number;
};

export type ArrayValue = {
  data: CellValueRich[][];
};

export type SpillValue = {
  origin: CellRef;
};

export type CellValueRich =
  | { type: "empty" }
  | { type: "number"; value: number }
  | { type: "string"; value: string }
  | { type: "boolean"; value: boolean }
  | { type: "error"; value: string }
  | { type: "rich_text"; value: RichTextValue }
  | { type: "entity"; value: EntityValue }
  | { type: "record"; value: RecordValue }
  | { type: "image"; value: ImageValue }
  | { type: "array"; value: ArrayValue }
  | { type: "spill"; value: SpillValue };

export interface CellData {
  sheet: string;
  address: string;
  input: CellScalar;
  value: CellScalar;
}

/**
 * Compact range cell payload used by the `getRangeCompact` RPC.
 *
 * Shape: `[input, value]`
 */
export type CellDataCompact = [input: CellScalar, value: CellScalar];

export interface CellDataRich {
  sheet: string;
  address: string;
  input: CellValueRich;
  value: CellValueRich;
}

export interface CellChange {
  sheet: string;
  address: string;
  value: CellScalar;
}

/**
 * Best-effort used range metadata for a worksheet.
 *
 * Semantics:
 * - Coordinates are 0-based (engine coordinates).
 * - Range is inclusive (`start_*` and `end_*` are both included).
 */
export interface SheetUsedRangeDto {
  start_row: number;
  end_row: number;
  start_col: number;
  end_col: number;
}

/**
 * Lightweight workbook metadata DTO returned by the worker's `getWorkbookInfo` RPC.
 *
 * This is compatible with `@formula/workbook-backend`'s `WorkbookInfo` shape, with additive
 * optional metadata (dimensions + best-effort used range).
 */
export interface WorkbookSheetInfoDto {
  id: string;
  name: string;
  visibility?: "visible" | "hidden" | "veryHidden";
  tabColor?: {
    rgb?: string;
    theme?: number;
    indexed?: number;
    tint?: number;
    auto?: boolean;
  };
  /**
   * Logical row count for the worksheet grid.
   *
   * Only present when the sheet differs from Excel defaults.
   */
  rowCount?: number;
  /**
   * Logical column count for the worksheet grid.
   *
   * Only present when the sheet differs from Excel defaults.
   */
  colCount?: number;
  /**
   * Best-effort used range computed from stored sparse inputs (may be omitted or null when empty).
   */
  usedRange?: SheetUsedRangeDto | null;
}

export interface WorkbookInfoDto {
  path: string | null;
  origin_path: string | null;
  sheets: WorkbookSheetInfoDto[];
}

/**
 * Host-provided system/environment metadata surfaced via Excel-compatible `INFO()` keys.
 *
 * All fields are optional; `null`/empty string clears the value.
 *
 * Note: `memavail` and `totmem` must be finite numbers.
 */
export interface EngineInfoDto {
  system?: string | null;
  directory?: string | null;
  osversion?: string | null;
  release?: string | null;
  version?: string | null;
  memavail?: number | null;
  totmem?: number | null;
}

// Pivots

export type PivotValue =
  | { type: "blank" }
  | { type: "number"; value: number }
  | { type: "date"; value: string }
  | { type: "text"; value: string }
  | { type: "bool"; value: boolean };

export type PivotFieldType = "blank" | "number" | "date" | "text" | "bool" | "mixed";

export interface PivotSchemaField {
  name: string;
  fieldType: PivotFieldType;
  sampleValues: PivotValue[];
}

export interface PivotSchema {
  fields: PivotSchemaField[];
  recordCount: number;
}

/**
 * Canonical reference to a pivot field in `formula_model::pivots`.
 *
 * Notes:
 * - For worksheet/range-backed pivots, `PivotFieldRef` is serialized as a plain string (the header
 *   name). We keep that backward compatible representation here via the `string` variant.
 * - For Data Model pivots, refs are structured objects to remove ambiguity:
 *   - `{ table, column }` for columns
 *   - `{ measure }` for measures
 *
 * The Rust deserializer also accepts `{ name }` as an alternate measure shape for backward
 * compatibility with earlier payloads.
 */
export type PivotFieldRef =
  | string
  | { table: string; column: string }
  | { measure: string }
  | { name: string };

export type PivotSortOrder = "ascending" | "descending" | "manual";

export interface PivotField {
  sourceField: PivotFieldRef;
  sortOrder?: PivotSortOrder;
  // Note: manual sort values are represented in Rust as `PivotKeyPart` (including
  // numeric variants encoded as raw f64 bit-pattern u64s). That is not currently
  // representable losslessly in JS, so keep this shape permissive.
  manualSort?: unknown[];
}

export type PivotAggregationType =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNumbers"
  | "stdDev"
  | "stdDevP"
  | "var"
  | "varP";

export type PivotShowAsType =
  | "normal"
  | "percentOfGrandTotal"
  | "percentOfRowTotal"
  | "percentOfColumnTotal"
  | "percentOf"
  | "percentDifferenceFrom"
  | "runningTotal"
  | "rankAscending"
  | "rankDescending";

export interface PivotValueField {
  sourceField: PivotFieldRef;
  name: string;
  aggregation: PivotAggregationType;
  numberFormat?: string;
  showAs?: PivotShowAsType;
  baseField?: PivotFieldRef;
  baseItem?: string;
}

export interface PivotFilterField {
  sourceField: PivotFieldRef;
  // Allowed values are represented in Rust as `HashSet<PivotKeyPart>` which is not
  // currently representable losslessly in JS (it serializes numeric keys as raw
  // u64 bit patterns). For now treat as an opaque payload.
  allowed?: unknown;
}

export interface PivotCalculatedField {
  name: string;
  formula: string;
}

export interface PivotCalculatedItem {
  field: string;
  name: string;
  formula: string;
}

export type PivotLayout = "compact" | "outline" | "tabular";
export type PivotSubtotalPosition = "top" | "bottom" | "none";

export interface PivotGrandTotals {
  rows: boolean;
  columns: boolean;
}

/**
 * Canonical pivot config schema expected by `crates/formula-wasm` (`formula_model::pivots::PivotConfig`).
 */
export interface PivotConfig {
  rowFields: PivotField[];
  columnFields: PivotField[];
  valueFields: PivotValueField[];
  filterFields: PivotFilterField[];
  calculatedFields?: PivotCalculatedField[];
  calculatedItems?: PivotCalculatedItem[];
  layout: PivotLayout;
  subtotals: PivotSubtotalPosition;
  grandTotals: PivotGrandTotals;
}

export interface PivotCellWrite {
  sheet: string;
  address: string;
  value: CellScalar;
  /** Optional Excel-compatible number format code (e.g. "m/d/yyyy", "$#,##0.00"). */
  numberFormat?: string;
}

export interface PivotCalculationResult {
  writes: PivotCellWrite[];
}

export type PivotFieldItems = PivotValue[];

// What-If / Goal Seek
export type GoalSeekRecalcMode = "singleThreaded" | "multiThreaded";

export interface GoalSeekRequest {
  /** A1 address within `sheet` (no `Sheet!` prefix). */
  targetCell: string;
  targetValue: number;
  /** A1 address within `sheet` (no `Sheet!` prefix). */
  changingCell: string;
  sheet?: string;
  tolerance?: number;
  maxIterations?: number;
  /**
   * Optional finite differencing step size. If omitted, the engine chooses a value based on the
   * current input.
   */
  derivativeStep?: number;
  /** Minimum absolute derivative before falling back to bisection. */
  minDerivative?: number;
  /** Maximum number of bracket expansion attempts when falling back to bisection. */
  maxBracketExpansions?: number;
  /**
   * Legacy option kept for API compatibility.
   *
   * Note: `crates/formula-wasm` currently ignores this field (it is safe to pass but has no effect).
   */
  recalcMode?: GoalSeekRecalcMode;
}

export type GoalSeekStatus =
  | "Converged"
  | "MaxIterationsReached"
  | "NoBracketFound"
  | "NumericalFailure"
  | (string & {});

export interface GoalSeekResult {
  /**
   * Goal seek solver status.
   *
   * Known values match the Rust `GoalSeekStatus` enum. Treat this as best-effort and be prepared
   * for additional statuses in future engine versions.
   */
  status: GoalSeekStatus;
  solution: number;
  iterations: number;
  finalOutput: number;
  finalError: number;
}

export interface GoalSeekResponse {
  result: GoalSeekResult;
  changes: CellChange[];
}

/**
 * Structural edit operation applied with Excel-like semantics.
 *
 * Notes:
 * - `row` / `col` are 0-indexed (engine coordinates).
 * - `address` / `range` use A1 notation (e.g. `A1`, `A1:B2`).
 */
export type EditOp =
  | { type: "InsertRows"; sheet: string; row: number; count: number }
  | { type: "DeleteRows"; sheet: string; row: number; count: number }
  | { type: "InsertCols"; sheet: string; col: number; count: number }
  | { type: "DeleteCols"; sheet: string; col: number; count: number }
  | { type: "InsertCellsShiftRight"; sheet: string; range: string }
  | { type: "InsertCellsShiftDown"; sheet: string; range: string }
  | { type: "DeleteCellsShiftLeft"; sheet: string; range: string }
  | { type: "DeleteCellsShiftUp"; sheet: string; range: string }
  | { type: "MoveRange"; sheet: string; src: string; dstTopLeft: string }
  | { type: "CopyRange"; sheet: string; src: string; dstTopLeft: string }
  | { type: "Fill"; sheet: string; src: string; dst: string };

export interface EditCellSnapshot {
  value: CellScalar;
  formula?: string;
}

export interface EditCellChange {
  sheet: string;
  address: string;
  before?: EditCellSnapshot;
  after?: EditCellSnapshot;
}

export interface EditMovedRange {
  sheet: string;
  from: string;
  to: string;
}

export interface EditFormulaRewrite {
  sheet: string;
  address: string;
  before: string;
  after: string;
}

export interface EditResult {
  changedCells: EditCellChange[];
  movedRanges: EditMovedRange[];
  formulaRewrites: EditFormulaRewrite[];
}

/**
 * Request item for `rewriteFormulasForCopyDelta`.
 *
 * This mirrors the Rust `rewrite_formula_for_copy_delta` helper used by structural edits
 * (copy/fill) so UI code can shift formulas using engine semantics without applying a
 * full workbook edit.
 */
export interface RewriteFormulaForCopyDeltaRequest {
  formula: string;
  /** Row delta in 0-based engine coordinates. */
  deltaRow: number;
  /** Column delta in 0-based engine coordinates. */
  deltaCol: number;
}

/**
 * Span in a formula string.
 *
 * Offsets are expressed as **UTF-16 code unit** indices (the same indexing used
 * by JS `string.slice()` / `string.length`).
 *
 * Spans use `[start, end)` semantics (start inclusive, end exclusive).
 */
export interface FormulaSpan {
  start: number;
  end: number;
}

/**
 * Token returned by `lexFormula`.
 *
 * This type intentionally mirrors the Rust wasm DTO shape produced by
 * `crates/formula-wasm` (`LexTokenDto`).
 */
export type FormulaCoord =
  | { kind: "A1"; index: number; abs: boolean }
  | { kind: "Offset"; delta: number };

export type FormulaToken =
  | { kind: "Number"; span: FormulaSpan; value: string }
  | { kind: "String"; span: FormulaSpan; value: string }
  | { kind: "Boolean"; span: FormulaSpan; value: boolean }
  | { kind: "Error"; span: FormulaSpan; value: string }
  | { kind: "Cell"; span: FormulaSpan; row: number; col: number; row_abs: boolean; col_abs: boolean }
  | { kind: "R1C1Cell"; span: FormulaSpan; row: FormulaCoord; col: FormulaCoord }
  | { kind: "R1C1Row"; span: FormulaSpan; row: FormulaCoord }
  | { kind: "R1C1Col"; span: FormulaSpan; col: FormulaCoord }
  | { kind: "Ident"; span: FormulaSpan; value: string }
  | { kind: "QuotedIdent"; span: FormulaSpan; value: string }
  | { kind: "Whitespace"; span: FormulaSpan; value: string }
  | { kind: "Intersect"; span: FormulaSpan; value: string }
  | { kind: "LParen"; span: FormulaSpan }
  | { kind: "RParen"; span: FormulaSpan }
  | { kind: "LBrace"; span: FormulaSpan }
  | { kind: "RBrace"; span: FormulaSpan }
  | { kind: "LBracket"; span: FormulaSpan }
  | { kind: "RBracket"; span: FormulaSpan }
  | { kind: "Bang"; span: FormulaSpan }
  | { kind: "Colon"; span: FormulaSpan }
  | { kind: "Dot"; span: FormulaSpan }
  | { kind: "ArgSep"; span: FormulaSpan }
  | { kind: "Union"; span: FormulaSpan }
  | { kind: "ArrayRowSep"; span: FormulaSpan }
  | { kind: "ArrayColSep"; span: FormulaSpan }
  | { kind: "Plus"; span: FormulaSpan }
  | { kind: "Minus"; span: FormulaSpan }
  | { kind: "Star"; span: FormulaSpan }
  | { kind: "Slash"; span: FormulaSpan }
  | { kind: "Caret"; span: FormulaSpan }
  | { kind: "Amp"; span: FormulaSpan }
  | { kind: "Percent"; span: FormulaSpan }
  | { kind: "Hash"; span: FormulaSpan }
  | { kind: "Eq"; span: FormulaSpan }
  | { kind: "Ne"; span: FormulaSpan }
  | { kind: "Lt"; span: FormulaSpan }
  | { kind: "Gt"; span: FormulaSpan }
  | { kind: "Le"; span: FormulaSpan }
  | { kind: "Ge"; span: FormulaSpan }
  | { kind: "At"; span: FormulaSpan }
  | { kind: "Eof"; span: FormulaSpan };

export interface FormulaParseError {
  message: string;
  span: FormulaSpan;
}

export interface FunctionContext {
  name: string;
  /** 0-indexed argument index. */
  argIndex: number;
}

export interface FormulaPartialParseResult {
  /**
   * Partial AST representation (Rust DTO; currently treated as opaque by the TS API).
   */
  ast: unknown;
  error: FormulaParseError | null;
  context: {
    function: FunctionContext | null;
  };
}

/**
 * Parse options accepted by the WASM editor tooling helpers (`lexFormula` / `parseFormulaPartial`).
 *
 * This intentionally mirrors the JS-friendly DTO supported by `crates/formula-wasm`.
 */
export interface FormulaParseOptions {
  /**
   * Formula locale id (e.g. `"en-US"`, `"de-DE"`).
   *
   * Note: supported locales depend on the engine build.
   */
  localeId?: string;
  referenceStyle?: "A1" | "R1C1";
}

/**
 * Locale metadata used by formula parsing/rendering.
 *
 * This mirrors the JS-friendly DTO returned by `crates/formula-wasm` via `getLocaleInfo(localeId)`.
 */
export interface FormulaLocaleInfo {
  localeId: string;
  decimalSeparator: string;
  argSeparator: string;
  arrayRowSeparator: string;
  arrayColSeparator: string;
  thousandsSeparator?: string;
  isRtl: boolean;
  booleanTrue: string;
  booleanFalse: string;
}

/**
 * Result returned by `lexFormulaPartial`.
 */
export interface FormulaPartialLexResult {
  tokens: FormulaToken[];
  error: FormulaParseError | null;
}

export interface RpcOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}

/**
 * Compressed range-run formatting segment applied to a single column.
 *
 * Mirrors `crates/formula-engine`'s `FormatRun` and `DocumentController`'s `formatRunsByCol`
 * representation.
 *
 * Runs use half-open row intervals `[startRow, endRowExclusive)` in 0-based engine coordinates.
 */
export type FormatRun = { startRow: number; endRowExclusive: number; styleId: number };

/**
 * RPC method names supported by the `packages/engine` Web Worker protocol.
 *
 * Keep this in sync with:
 * - `packages/engine/src/worker/EngineWorker.ts` (client)
 * - `packages/engine/src/engine.worker.ts` (worker dispatch)
 */
export type RpcMethod =
  | "ping"
  // Editor tooling (module-level; independent of workbook state)
  | "supportedLocaleIds"
  | "getLocaleInfo"
  | "lexFormula"
  | "lexFormulaPartial"
  | "parseFormulaPartial"
  | "rewriteFormulasForCopyDelta"
  | "canonicalizeFormula"
  | "localizeFormula"
  // Workbook lifecycle
  | "newWorkbook"
  | "loadFromJson"
  | "loadFromXlsxBytes"
  | "loadFromEncryptedXlsxBytes"
  // Workbook reads/writes
  | "getWorkbookInfo"
  | "toJson"
  | "getCell"
  | "getCellRich"
  | "getRange"
  | "getRangeCompact"
  | "setCells"
  | "setCellRich"
  | "setRange"
  | "setLocale"
  | "getCalcSettings"
  | "setCalcSettings"
  | "setEngineInfo"
  | "setInfoOrigin"
  | "setInfoOriginForSheet"
  | "recalculate"
  | "applyOperation"
  | "goalSeek"
  | "getPivotSchema"
  | "getPivotFieldItems"
  | "getPivotFieldItemsPaged"
  | "calculatePivot"
  | "setSheetDimensions"
  | "getSheetDimensions"
  | "renameSheet"
  | "setSheetOrigin"
  | "setSheetDisplayName"
  // Workbook metadata
  | "setWorkbookFileMetadata"
  | "setCellStyleId"
  | "setRowStyleId"
  | "setColStyleId"
  | "setFormatRunsByCol"
  | "setSheetDefaultStyleId"
  | "setColWidth"
  | "setColWidthChars"
  | "setColHidden"
  | "setColFormatRuns"
  | "internStyle";

/**
 * Minimal JSON-friendly cell style payload used by `internStyle`.
 *
 * This type is intentionally kept flexible: different frontends may have
 * different style encodings (e.g. snake_case formula-model styles vs
 * camelCase OOXML-ish styles). The WASM boundary is responsible for validating
 * and normalizing.
 */
export type WorkbookStyleDto = Record<string, unknown>;

export interface InitMessage {
  type: "init";
  port: MessagePort;
  /**
   * URL to the WASM module (typically the wasm-bindgen JS glue entrypoint).
   *
   * Pass an empty string to disable WASM loading; RPC requests will fail with
   * `worker not initialized`.
   */
  wasmModuleUrl: string;
  /**
   * Optional URL to the `.wasm` binary.
   *
   * If provided, the worker will pass it to the wasm-bindgen init function.
   * This can be useful when bundlers fingerprint assets such that the wrapper
   * can't reliably derive the `.wasm` URL from its own `import.meta.url`.
   */
  wasmBinaryUrl?: string;
}

export interface ReadyMessage {
  type: "ready";
}

export interface RpcRequest {
  type: "request";
  id: number;
  method: RpcMethod;
  /**
   * Params are structured-cloned across the MessagePort.
   *
   * Some calls may include large binary payloads (e.g. `Uint8Array` workbook bytes)
   * that are posted with an explicit transfer list so their underlying
   * `ArrayBuffer` moves ownership to the worker without an extra copy.
   */
  params: unknown;
}

export interface RpcCancel {
  type: "cancel";
  id: number;
}

export interface RpcResponseOk {
  type: "response";
  id: number;
  ok: true;
  result: unknown;
}

export interface RpcResponseErr {
  type: "response";
  id: number;
  ok: false;
  error: string;
}

export type WorkerInboundMessage = RpcRequest | RpcCancel;
export type WorkerOutboundMessage = ReadyMessage | RpcResponseOk | RpcResponseErr;
