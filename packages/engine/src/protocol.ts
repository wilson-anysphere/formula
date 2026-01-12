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
