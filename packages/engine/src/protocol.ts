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
 * JSON-friendly rich cell value transported over the worker RPC boundary.
 *
 * This mirrors `formula_model::CellValue`'s `{type,value}` tagged schema.
 *
 * Note: This type is intentionally minimal (best-effort) to avoid coupling the TS
 * public API too tightly to the Rust model while rich values are still evolving.
 */
export type CellValueRich =
  | { type: "empty" }
  | { type: "number"; value: number }
  | { type: "string"; value: string }
  | { type: "boolean"; value: boolean }
  | { type: "error"; value: string }
  | { type: "rich_text"; value: unknown }
  | { type: "entity"; value: unknown }
  | { type: "record"; value: unknown }
  | { type: "image"; value: unknown }
  | { type: "array"; value: unknown }
  | { type: "spill"; value: unknown };

export interface CellData {
  sheet: string;
  address: string;
  input: CellScalar;
  value: CellScalar;
}

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
 * Span in a formula string.
 *
 * Offsets are expressed as **UTF-16 code unit** indices (the same indexing used
 * by JS `string.slice()` / `string.length`).
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

export interface RpcOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}

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
  method: string;
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
