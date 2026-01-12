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

export interface CellData {
  sheet: string;
  address: string;
  input: CellScalar;
  value: CellScalar;
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
 * Note: This type intentionally mirrors the Rust wasm DTO shape. Treat unknown
 * fields as implementation details; consumers should rely on `kind` + `span`
 * for stable behavior.
 */
export interface FormulaToken {
  kind: string;
  span: FormulaSpan;
  /**
   * Optional token payload (depends on `kind`).
   *
   * For example:
   * - literals may include a raw string/number representation
   * - reference tokens may include structured metadata (row/col, absolutes)
   */
  value?: unknown;
}

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
