import type {
  CellChange,
  CellData,
  CellDataRich,
  CellScalar,
  CellValueRich,
  EditOp,
  EditResult,
  GoalSeekRequest,
  GoalSeekResponse,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaParseOptions,
  FormulaToken,
  PivotCalculationResult,
  PivotConfig,
  PivotSchema,
  RewriteFormulaForCopyDeltaRequest,
  RpcOptions,
} from "./protocol.ts";
import { defaultWasmBinaryUrl, defaultWasmModuleUrl } from "./wasm.ts";
import { EngineWorker } from "./worker/EngineWorker.ts";

export interface EngineClient {
  /**
   * Force initialization of the underlying worker/WASM engine.
   *
   * Note: the engine may still lazy-load WASM on first request.
   */
  init(): Promise<void>;
  newWorkbook(): Promise<void>;
  loadWorkbookFromJson(json: string): Promise<void>;
  /**
   * Load a workbook from raw `.xlsx` bytes.
   *
   * The underlying `ArrayBuffer` may be transferred to a Worker thread to avoid
   * an extra structured-clone copy. When the buffer is transferred it becomes
   * detached on the calling side.
   *
   * If `bytes` is a view into a larger buffer (e.g. a `subarray()`), the engine
   * may first copy it to a compact buffer so only the view range is transferred.
   */
  loadWorkbookFromXlsxBytes(bytes: Uint8Array, options?: RpcOptions): Promise<void>;
  toJson(): Promise<string>;
  getCell(address: string, sheet?: string, options?: RpcOptions): Promise<CellData>;
  /**
   * Read a cell's rich `{type,value}` input/value.
   *
   * This is an additive API: rich values are not representable in the legacy
   * scalar workbook JSON schema returned by `toJson()`.
   */
  getCellRich?(address: string, sheet?: string, options?: RpcOptions): Promise<CellDataRich>;
  getRange(range: string, sheet?: string, options?: RpcOptions): Promise<CellData[][]>;
  /**
   * Set a single cell, batched across the current microtask to minimize RPC
   * overhead.
   */
  setCell(address: string, value: CellScalar, sheet?: string): Promise<void>;
  /**
   * Set a cell's rich `{type,value}` input.
   *
   * For scalar inputs (number/string/bool/error), callers should prefer
   * `setCell` for compatibility; `setCellRich` exists for entity/record/image
   * rich values.
   */
  setCellRich?(address: string, value: CellValueRich | null, sheet?: string, options?: RpcOptions): Promise<void>;
  /**
   * Set multiple cells in a single RPC call.
   *
   * Useful when applying large delta batches (paste, imports) without creating
   * per-cell promises.
   */
  setCells(
    updates: Array<{ address: string; value: CellScalar; sheet?: string }>,
    options?: RpcOptions
  ): Promise<void>;
  setRange(range: string, values: CellScalar[][], sheet?: string, options?: RpcOptions): Promise<void>;
  /**
   * Set the locale used by the WASM engine when interpreting user-entered formulas and when
   * parsing locale-sensitive strings at runtime (criteria, VALUE/DATE parsing, etc).
   *
   * Returns `false` when the locale id is not supported by the engine build.
   */
  setLocale(localeId: string, options?: RpcOptions): Promise<boolean>;
  /**
   * Recalculate the workbook and return value-change deltas.
   *
   * Note: the `sheet` argument is accepted for API symmetry with other surfaces,
   * but the web/WASM engine intentionally returns *all* value changes across the
   * workbook even when a sheet is provided. Callers rely on this to keep
   * cross-sheet caches coherent.
   */
  recalculate(sheet?: string, options?: RpcOptions): Promise<CellChange[]>;

  /**
   * Inspect a worksheet range (headers + records) and return a pivot schema for UI prompting.
   *
   * Additive API: older WASM builds may not export `getPivotSchema`.
   */
  getPivotSchema?(sheet: string, sourceRangeA1: string, sampleSize?: number, options?: RpcOptions): Promise<PivotSchema>;

  /**
   * Calculate a pivot table from a source range and return the list of worksheet writes needed to
   * render the pivot at `destinationTopLeftA1`.
   *
   * Does **not** mutate workbook state; callers can apply the returned writes as desired.
   */
  calculatePivot?(
    sheet: string,
    sourceRangeA1: string,
    destinationTopLeftA1: string,
    config: PivotConfig,
    options?: RpcOptions
  ): Promise<PivotCalculationResult>;

  /**
   * Run Goal Seek (what-if analysis) over the current workbook.
   *
   * This is an additive API; older WASM builds may not export `goalSeek`.
   */
  goalSeek?(request: GoalSeekRequest, options?: RpcOptions): Promise<GoalSeekResponse>;
  /**
   * Configure the logical worksheet dimensions (row/column count) for a sheet.
   *
   * This affects whole-row/whole-column references like `1:1` / `A:A`.
   */
  setSheetDimensions(sheet: string, rows: number, cols: number, options?: RpcOptions): Promise<void>;
  getSheetDimensions(sheet: string, options?: RpcOptions): Promise<{ rows: number; cols: number }>;

  /**
   * Apply an Excel-like structural edit operation (insert/delete rows/cols, move/copy/fill).
   *
   * Note: does **not** implicitly recalculate. Call `recalculate()` explicitly to
   * update cached formula results.
   */
  applyOperation(op: EditOp, options?: RpcOptions): Promise<EditResult>;

  /**
   * Rewrite formulas as if they were copied by `(deltaRow, deltaCol)`.
   *
   * This is intended for UI layers (clipboard paste, fill handle) that need the
   * engine's shifting semantics without mutating workbook state.
   */
  rewriteFormulasForCopyDelta(requests: RewriteFormulaForCopyDeltaRequest[], options?: RpcOptions): Promise<string[]>;

  /**
   * Tokenize a formula string for editor tooling (syntax highlighting, etc).
   *
   * This call is independent of any loaded workbook.
   */
  lexFormula(formula: string, options?: FormulaParseOptions, rpcOptions?: RpcOptions): Promise<FormulaToken[]>;
  /**
   * Tokenize a formula string for editor tooling (syntax highlighting, etc).
   *
   * Convenience overload for `lexFormula(formula, undefined, rpcOptions)`.
   */
  lexFormula(formula: string, rpcOptions?: RpcOptions): Promise<FormulaToken[]>;

  /**
   * Best-effort lexer for editor syntax highlighting (never throws).
   *
   * This call is independent of any loaded workbook.
   */
  lexFormulaPartial(
    formula: string,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialLexResult>;
  /**
   * Best-effort lexer for editor syntax highlighting (never throws).
   *
   * Convenience overload for `lexFormulaPartial(formula, undefined, rpcOptions)`.
   */
  lexFormulaPartial(formula: string, rpcOptions?: RpcOptions): Promise<FormulaPartialLexResult>;

  /**
   * Best-effort partial parse for editor/autocomplete scenarios.
   *
   * `cursor` (when provided) is expressed as a **UTF-16 code unit** offset (JS
   * string indexing). This matches the span units returned by `lexFormula`.
   *
   * This call is independent of any loaded workbook.
   */
  parseFormulaPartial(
    formula: string,
    cursor?: number,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialParseResult>;
  /**
   * Best-effort partial parse at the end of the formula string (cursor defaults to `formula.length`).
   *
   * This overload is a convenience so callers can pass parse options without having to provide an
   * explicit cursor.
   */
  parseFormulaPartial(formula: string, options?: FormulaParseOptions, rpcOptions?: RpcOptions): Promise<FormulaPartialParseResult>;
  terminate(): void;
}

export function createEngineClient(options?: { wasmModuleUrl?: string; wasmBinaryUrl?: string }): EngineClient {
  if (typeof Worker === "undefined") {
    throw new Error("createEngineClient() requires a Worker-capable environment");
  }

  const wasmModuleUrl = options?.wasmModuleUrl ?? defaultWasmModuleUrl();
  const wasmBinaryUrl = options?.wasmBinaryUrl ?? defaultWasmBinaryUrl();

  // Vite supports Worker construction via `new URL(..., import.meta.url)` and will
  // bundle the Worker entrypoint correctly for both dev and production builds.
  //
  // In React 18 StrictMode (dev-only), effects intentionally run twice
  // (setup → cleanup → setup). Callers may `terminate()` between setups, so we
  // support tearing down the worker and reconnecting on-demand.
  let worker: Worker | null = null;
  let engine: EngineWorker | null = null;
  let enginePromise: Promise<EngineWorker> | null = null;
  let generation = 0;

  const ensureWorker = () => {
    if (worker) return worker;
    worker = new Worker(new URL("./engine.worker.ts", import.meta.url), {
      type: "module"
    });
    return worker;
  };

  const connect = () => {
    if (enginePromise) {
      return enginePromise;
    }

    const connectGeneration = ++generation;
    const activeWorker = ensureWorker();

    enginePromise = EngineWorker.connect({
      worker: activeWorker,
      wasmModuleUrl,
      wasmBinaryUrl
    });

    void enginePromise
      .then((connected) => {
        // If the caller terminated/restarted while we were connecting, immediately
        // dispose the stale connection.
        if (connectGeneration !== generation) {
          connected.terminate();
          return;
        }
        engine = connected;
      })
      .catch(() => {
        // Allow retries on the next call.
        if (connectGeneration !== generation) {
          return;
        }
        enginePromise = null;
        engine = null;
        worker?.terminate();
        worker = null;
      });

    return enginePromise;
  };

  const withEngine = async <T>(fn: (engine: EngineWorker) => Promise<T>): Promise<T> => {
    const connected = await connect();
    return await fn(connected);
  };

  return {
    init: async () => {
      await connect();
    },
    newWorkbook: async () => await withEngine((connected) => connected.newWorkbook()),
    loadWorkbookFromJson: async (json) => await withEngine((connected) => connected.loadWorkbookFromJson(json)),
    loadWorkbookFromXlsxBytes: async (bytes, rpcOptions) =>
      await withEngine((connected) => connected.loadWorkbookFromXlsxBytes(bytes, rpcOptions)),
    toJson: async () => await withEngine((connected) => connected.toJson()),
    getCell: async (address, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getCell(address, sheet, rpcOptions)),
    getCellRich: async (address, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getCellRich(address, sheet, rpcOptions)),
    getRange: async (range, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getRange(range, sheet, rpcOptions)),
    setCell: async (address, value, sheet) => await withEngine((connected) => connected.setCell(address, value, sheet)),
    setCellRich: async (address, value, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setCellRich(address, value, sheet, rpcOptions)),
    setCells: async (updates, rpcOptions) => await withEngine((connected) => connected.setCells(updates, rpcOptions)),
    setRange: async (range, values, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setRange(range, values, sheet, rpcOptions)),
    setLocale: async (localeId, rpcOptions) => await withEngine((connected) => connected.setLocale(localeId, rpcOptions)),
    recalculate: async (sheet, rpcOptions) => await withEngine((connected) => connected.recalculate(sheet, rpcOptions)),
    getPivotSchema: async (sheet, sourceRangeA1, sampleSize, rpcOptions) =>
      await withEngine((connected) => connected.getPivotSchema(sheet, sourceRangeA1, sampleSize, rpcOptions)),
    calculatePivot: async (sheet, sourceRangeA1, destinationTopLeftA1, config, rpcOptions) =>
      await withEngine((connected) =>
        connected.calculatePivot(sheet, sourceRangeA1, destinationTopLeftA1, config, rpcOptions)
      ),
    goalSeek: async (request, rpcOptions) => await withEngine((connected) => connected.goalSeek(request, rpcOptions)),
    setSheetDimensions: async (sheet, rows, cols, rpcOptions) =>
      await withEngine((connected) => connected.setSheetDimensions(sheet, rows, cols, rpcOptions)),
    getSheetDimensions: async (sheet, rpcOptions) =>
      await withEngine((connected) => connected.getSheetDimensions(sheet, rpcOptions)),
    applyOperation: async (op, rpcOptions) => await withEngine((connected) => connected.applyOperation(op, rpcOptions)),
    rewriteFormulasForCopyDelta: async (requests, rpcOptions) =>
      await withEngine((connected) => connected.rewriteFormulasForCopyDelta(requests, rpcOptions)),
    lexFormula: async (formula: string, optionsOrRpcOptions?: FormulaParseOptions | RpcOptions, rpcOptions?: RpcOptions) =>
      await withEngine((connected) => (connected.lexFormula as any)(formula, optionsOrRpcOptions, rpcOptions)),
    lexFormulaPartial: async (
      formula: string,
      optionsOrRpcOptions?: FormulaParseOptions | RpcOptions,
      rpcOptions?: RpcOptions
    ) => await withEngine((connected) => (connected.lexFormulaPartial as any)(formula, optionsOrRpcOptions, rpcOptions)),
    parseFormulaPartial: async (
      formula: string,
      cursorOrOptions?: number | FormulaParseOptions,
      optionsOrRpcOptions?: FormulaParseOptions | RpcOptions,
      rpcOptions?: RpcOptions
    ) =>
      await withEngine((connected) =>
        // EngineWorker implements overloads/argument parsing; forward through.
        (connected.parseFormulaPartial as any)(formula, cursorOrOptions, optionsOrRpcOptions, rpcOptions)
      ),
    terminate: () => {
      generation++;
      enginePromise = null;
      engine?.terminate();
      engine = null;
      worker?.terminate();
      worker = null;
    }
  };
}

declare global {
  // Set by Playwright E2E tests via `page.addInitScript` before the app loads.
  // When enabled, we expose `createEngineClient` on `globalThis` so E2E tests can
  // create a fresh engine instance from the production bundle without relying on
  // Vite dev-server-only module resolution.
  // eslint-disable-next-line no-var
  var __FORMULA_E2E__: boolean | undefined;
  // eslint-disable-next-line no-var
  var __FORMULA_ENGINE_E2E__: { createEngineClient: typeof createEngineClient } | undefined;
}

if (typeof globalThis !== "undefined" && (globalThis as any).__FORMULA_E2E__) {
  (globalThis as any).__FORMULA_ENGINE_E2E__ = { createEngineClient };
}
