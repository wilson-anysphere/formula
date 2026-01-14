import type {
  CalcSettings,
  CellChange,
  CellData,
  CellDataCompact,
  CellDataRich,
  CellScalar,
  CellValueRich,
  EngineInfoDto,
  FormatRun,
  WorkbookStyleDto,
  EditOp,
  EditResult,
  GoalSeekRequest,
  GoalSeekResponse,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaParseOptions,
  FormulaToken,
  FormulaLocaleInfo,
  PivotCalculationResult,
  PivotConfig,
  PivotFieldItems,
  PivotSchema,
  WorkbookInfoDto,
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
   * Load a workbook from raw `.xlsx`/`.xlsm` bytes.
   *
   * The underlying `ArrayBuffer` may be transferred to a Worker thread to avoid
   * an extra structured-clone copy. When the buffer is transferred it becomes
   * detached on the calling side.
   *
   * If `bytes` is a view into a larger buffer (e.g. a `subarray()`), the engine
   * may first copy it to a compact buffer so only the view range is transferred.
   */
  loadWorkbookFromXlsxBytes(bytes: Uint8Array, options?: RpcOptions): Promise<void>;
  /**
   * Load a workbook from Office-encrypted `.xlsx`/`.xlsm`/`.xlsb` bytes and decrypt it using `password`.
   *
   * Additive API: older WASM builds / worker bundles may not support this call.
   */
  loadWorkbookFromEncryptedXlsxBytes?(bytes: Uint8Array, password: string, options?: RpcOptions): Promise<void>;
  toJson(): Promise<string>;
  /**
   * Return lightweight workbook metadata (sheet list + best-effort used ranges).
   *
   * Additive API: older workers / WASM builds may not support this call.
   */
  getWorkbookInfo?(options?: RpcOptions): Promise<WorkbookInfoDto>;
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
   * Fetch a range using a compact payload shape (`[input, value]` per cell) to avoid
   * allocating redundant `{sheet,address}` data for every cell.
   *
   * Additive API: older WASM builds may not export this method.
   */
  getRangeCompact?(range: string, sheet?: string, options?: RpcOptions): Promise<CellDataCompact[][]>;
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
   * Set workbook-level file metadata used by Excel-compatible functions like `CELL("filename")`.
   */
  setWorkbookFileMetadata(directory: string | null, filename: string | null, options?: RpcOptions): Promise<void>;
  /**
   * Set a cell's style id (formatting metadata).
   */
  setCellStyleId(address: string, styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  /**
   * Set (or clear) a column width override.
   *
   * `col` is 0-indexed (engine coordinates). `width=null` clears the override.
   *
   * `width` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
   * Prefer `setColWidthChars` for an explicit unit name.
   */
  setColWidth(col: number, width: number | null, sheet?: string, options?: RpcOptions): Promise<void>;
  /**
   * Set a column's hidden flag.
   *
   * `col` is 0-indexed (engine coordinates).
   */
  setColHidden(col: number, hidden: boolean, sheet?: string, options?: RpcOptions): Promise<void>;
  /**
   * Intern (deduplicate) a style into the workbook's shared style table, returning its id.
   *
   * Style id `0` is always the default style. Passing `null` is treated as the default style.
   */
  internStyle(style: WorkbookStyleDto | null, options?: RpcOptions): Promise<number>;
  /**
   * Set the locale used by the WASM engine when interpreting user-entered formulas and when
   * parsing locale-sensitive strings at runtime (criteria, VALUE/DATE parsing, etc).
   *
   * Returns `false` when the locale id is not supported by the engine build.
   */
  setLocale(localeId: string, options?: RpcOptions): Promise<boolean>;

  /**
   * Read the workbook calculation settings (`calcPr`).
   */
  getCalcSettings(options?: RpcOptions): Promise<CalcSettings>;

  /**
   * Replace the workbook calculation settings (`calcPr`).
   */
  setCalcSettings(settings: CalcSettings, options?: RpcOptions): Promise<void>;
  /**
   * Set host-provided system/environment metadata surfaced via Excel `INFO()` keys.
   */
  setEngineInfo(info: EngineInfoDto, options?: RpcOptions): Promise<void>;
  /**
   * Set (or clear) the workbook-level default for `INFO("origin")`.
   */
  setInfoOrigin(origin: string | null, options?: RpcOptions): Promise<void>;
  /**
   * Set (or clear) the per-sheet override for `INFO("origin")`.
   */
  setInfoOriginForSheet(sheet: string, origin: string | null, options?: RpcOptions): Promise<void>;
  /**
   * Replace the range-run formatting runs for a column.
   *
   * Runs are expressed as half-open row intervals `[startRow, endRowExclusive)`.
   */
  setColFormatRuns(
    sheet: string,
    col: number,
    runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
    options?: RpcOptions,
  ): Promise<void>;
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
   * Return the unique pivot item values for a given field (worksheet/range-backed pivots).
   *
   * Additive API: older WASM builds may not export `getPivotFieldItems`.
   */
  getPivotFieldItems?(
    sheet: string,
    sourceRangeA1: string,
    field: string,
    options?: RpcOptions
  ): Promise<PivotFieldItems>;

  /**
   * Paged variant of `getPivotFieldItems` for large cardinality fields.
   *
   * Additive API: older WASM builds may not export `getPivotFieldItemsPaged`.
   */
  getPivotFieldItemsPaged?(
    sheet: string,
    sourceRangeA1: string,
    field: string,
    offset: number,
    limit: number,
    options?: RpcOptions
  ): Promise<PivotFieldItems>;

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
   * Set the top-left visible cell ("origin") for a worksheet view.
   *
   * This is host-provided UI metadata surfaced to formulas via `INFO("origin")`.
   */
  setSheetOrigin(sheet: string, origin: string | null, options?: RpcOptions): Promise<void>;

  /**
   * Update a sheet's user-visible display (tab) name without changing its stable id/key.
   *
   * This influences functions like `CELL("address")` and runtime sheet-name resolution in
   * functions like `INDIRECT`.
   */
  setSheetDisplayName?(sheetId: string, name: string, options?: RpcOptions): Promise<void>;

  /**
   * Rename a worksheet and rewrite formulas that reference it (Excel-like).
   *
   * Returns `false` when `oldName` does not exist or `newName` conflicts with another sheet.
   */
  renameSheet(oldName: string, newName: string, options?: RpcOptions): Promise<boolean>;

  /**
   * Set (or clear) a per-column width override.
   *
   * `widthChars` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
   */
  setColWidthChars(sheet: string, col: number, widthChars: number | null, options?: RpcOptions): Promise<void>;
  /**
   * Set a row-level formatting style id (layered formatting).
   *
   * Preferred signature: `(sheet, row, styleId)` where `null` clears the row style.
   * Legacy signature: `(row, styleId, sheet?)` where `styleId=0` clears the row style.
   */
  setRowStyleId?: {
    (sheet: string, row: number, styleId: number | null, options?: RpcOptions): Promise<void>;
    (row: number, styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  };
  /**
   * Set a column-level formatting style id (layered formatting).
   *
   * Preferred signature: `(sheet, col, styleId)` where `null` clears the column style.
   * Legacy signature: `(col, styleId, sheet?)` where `styleId=0` clears the column style.
   */
  setColStyleId?: {
    (sheet: string, col: number, styleId: number | null, options?: RpcOptions): Promise<void>;
    (col: number, styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  };
  /**
   * Set the sheet default style id (layered formatting base).
   *
   * Preferred signature: `(sheet, styleId)` where `null` resets to the default style.
   * Legacy signature: `(styleId, sheet?)` where `styleId=0` clears the override.
   */
  setSheetDefaultStyleId?: {
    (sheet: string, styleId: number | null, options?: RpcOptions): Promise<void>;
    (styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  };
  /**
   * Replace the compressed range-run formatting layer for a column (DocumentController `formatRunsByCol`).
   *
   * Preferred signature: `(sheet, col, runs)`.
   * Legacy signature: `(col, runs, sheet?)`.
   */
  setFormatRunsByCol?: {
    (
      sheet: string,
      col: number,
      runs: FormatRun[],
      options?: RpcOptions
    ): Promise<void>;
    (
      col: number,
      runs: FormatRun[],
      sheet?: string,
      options?: RpcOptions
    ): Promise<void>;
  };

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
   * Canonicalize a locale-specific formula string into the engine's persisted form.
   *
   * This call is independent of any loaded workbook.
   */
  canonicalizeFormula?(
    formula: string,
    localeId: string,
    referenceStyle?: "A1" | "R1C1",
    options?: RpcOptions
  ): Promise<string>;

  /**
   * Localize a canonical (English) formula string for display in `localeId`.
   *
   * This call is independent of any loaded workbook.
   */
  localizeFormula?(
    formula: string,
    localeId: string,
    referenceStyle?: "A1" | "R1C1",
    options?: RpcOptions
  ): Promise<string>;

  /**
   * Return the list of formula locale ids supported by the underlying engine build.
   *
   * This call is independent of any loaded workbook.
   */
  supportedLocaleIds?(options?: RpcOptions): Promise<string[]>;

  /**
   * Return locale metadata used by formula parsing/rendering (separators, boolean literals, RTL flag, etc).
   *
   * This call is independent of any loaded workbook.
   */
  getLocaleInfo?(localeId: string, options?: RpcOptions): Promise<FormulaLocaleInfo>;

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

export function createEngineClient(options?: {
  wasmModuleUrl?: string;
  wasmBinaryUrl?: string;
  /**
   * Maximum time to wait for the worker "ready" handshake before failing the connection.
   *
   * The engine worker posts "ready" immediately after receiving the init message (before WASM is
   * loaded), so this should only trip when the worker fails to start (bundle load errors, CSP
   * issues, etc). Failing fast avoids hanging `init()` forever.
   */
  connectTimeoutMs?: number;
}): EngineClient {
  if (typeof Worker === "undefined") {
    throw new Error("createEngineClient() requires a Worker-capable environment");
  }

  const wasmModuleUrl = options?.wasmModuleUrl ?? defaultWasmModuleUrl();
  const wasmBinaryUrl = options?.wasmBinaryUrl ?? defaultWasmBinaryUrl();
  const connectTimeoutMs = options?.connectTimeoutMs ?? 10_000;

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
  let connectAbort: AbortController | null = null;

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

    // If the caller terminates while we're waiting for the worker "ready" handshake,
    // the underlying EngineWorker.connect() promise can hang forever (it has no
    // built-in timeout/rejection path). Wrap it so pending `init()`/calls can fail
    // fast on teardown (important for StrictMode cleanup, tests, and multi-document
    // flows).
    connectAbort?.abort();
    const abortController = new AbortController();
    connectAbort = abortController;

    const rawConnectPromise = EngineWorker.connect({
      worker: activeWorker,
      wasmModuleUrl,
      wasmBinaryUrl,
      signal: abortController.signal,
      timeoutMs: connectTimeoutMs
    });

    enginePromise = new Promise<EngineWorker>((resolve, reject) => {
      const signal = abortController.signal;
      const onAbort = () => reject(new Error("engine terminated while connecting"));

      if (signal.aborted) {
        reject(new Error("engine terminated while connecting"));
        return;
      }

      signal.addEventListener("abort", onAbort, { once: true });

      rawConnectPromise.then(
        (connected) => {
          signal.removeEventListener("abort", onAbort);
          resolve(connected);
        },
        (err) => {
          signal.removeEventListener("abort", onAbort);
          reject(err);
        }
      );
    });
    // `createEngineClient` is often used in React effects where `init()` may be
    // fired-and-forgotten (or cancelled during StrictMode cleanup). Attach a no-op
    // rejection handler so a failed/aborted connect attempt never surfaces as an
    // unhandled rejection when callers don't await.
    void enginePromise.catch(() => {});

    void rawConnectPromise
      .then((connected) => {
        // If the caller terminated/restarted while we were connecting, immediately
        // dispose the stale connection.
        if (connectGeneration !== generation) {
          connected.terminate();
          return;
        }
        engine = connected;
        if (connectAbort === abortController) {
          connectAbort = null;
        }
      })
      .catch(() => {
        // Allow retries on the next call.
        if (connectGeneration !== generation) {
          return;
        }
        enginePromise = null;
        engine = null;
        if (connectAbort === abortController) {
          connectAbort = null;
        }
        try {
          worker?.terminate();
        } catch {
          // ignore
        }
        worker = null;
      });

    return enginePromise;
  };

  const withEngine = async <T>(fn: (engine: EngineWorker) => Promise<T>): Promise<T> => {
    const connected = await connect();
    return await fn(connected);
  };

  return {
    init: () => {
      const promise = connect().then(() => undefined) as Promise<void>;
      // Callers sometimes fire-and-forget initialization (e.g. app startup effects). Attach a no-op
      // rejection handler so connect failures/timeouts don't surface as unhandled rejections when
      // the returned promise isn't awaited.
      void promise.catch(() => {});
      return promise;
    },
    newWorkbook: async () => await withEngine((connected) => connected.newWorkbook()),
    loadWorkbookFromJson: async (json) => await withEngine((connected) => connected.loadWorkbookFromJson(json)),
    loadWorkbookFromXlsxBytes: async (bytes, rpcOptions) =>
      await withEngine((connected) => connected.loadWorkbookFromXlsxBytes(bytes, rpcOptions)),
    loadWorkbookFromEncryptedXlsxBytes: async (bytes, password, rpcOptions) =>
      await withEngine((connected) => connected.loadWorkbookFromEncryptedXlsxBytes(bytes, password, rpcOptions)),
    toJson: async () => await withEngine((connected) => connected.toJson()),
    getWorkbookInfo: async (rpcOptions) => await withEngine((connected) => connected.getWorkbookInfo(rpcOptions)),
    getCell: async (address, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getCell(address, sheet, rpcOptions)),
    getCellRich: async (address, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getCellRich(address, sheet, rpcOptions)),
    getRange: async (range, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getRange(range, sheet, rpcOptions)),
    getRangeCompact: async (range, sheet, rpcOptions) =>
      await withEngine((connected) => connected.getRangeCompact(range, sheet, rpcOptions)),
    setCell: async (address, value, sheet) => await withEngine((connected) => connected.setCell(address, value, sheet)),
    setCellRich: async (address, value, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setCellRich(address, value, sheet, rpcOptions)),
    setCells: async (updates, rpcOptions) => await withEngine((connected) => connected.setCells(updates, rpcOptions)),
    setRange: async (range, values, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setRange(range, values, sheet, rpcOptions)),
    setWorkbookFileMetadata: async (directory, filename, rpcOptions) =>
      await withEngine((connected) => connected.setWorkbookFileMetadata(directory, filename, rpcOptions)),
    setCellStyleId: async (address, styleId, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setCellStyleId(address, styleId, sheet, rpcOptions)),
    // Style-layer RPC methods support multiple call signatures (legacy sheet-last + new sheet-first).
    // Forward through to EngineWorker which normalizes arguments.
    setRowStyleId: async (...args: any[]) =>
      await withEngine((connected) => (connected.setRowStyleId as any).call(connected, ...args)),
    setColStyleId: async (...args: any[]) =>
      await withEngine((connected) => (connected.setColStyleId as any).call(connected, ...args)),
    setSheetDefaultStyleId: async (...args: any[]) =>
      await withEngine((connected) => (connected.setSheetDefaultStyleId as any).call(connected, ...args)),
    setFormatRunsByCol: async (...args: any[]) =>
      await withEngine((connected) => (connected.setFormatRunsByCol as any).call(connected, ...args)),
    setColWidth: async (col, width, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setColWidth(col, width, sheet, rpcOptions)),
    setColHidden: async (col, hidden, sheet, rpcOptions) =>
      await withEngine((connected) => connected.setColHidden(col, hidden, sheet, rpcOptions)),
    internStyle: async (style, rpcOptions) => await withEngine((connected) => connected.internStyle(style, rpcOptions)),
    setLocale: async (localeId, rpcOptions) => await withEngine((connected) => connected.setLocale(localeId, rpcOptions)),
    getCalcSettings: async (rpcOptions) => await withEngine((connected) => connected.getCalcSettings(rpcOptions)),
    setCalcSettings: async (settings, rpcOptions) =>
      await withEngine((connected) => connected.setCalcSettings(settings, rpcOptions)),
    setEngineInfo: async (info, rpcOptions) => await withEngine((connected) => connected.setEngineInfo(info, rpcOptions)),
    setInfoOrigin: async (origin, rpcOptions) => await withEngine((connected) => connected.setInfoOrigin(origin, rpcOptions)),
    setInfoOriginForSheet: async (sheet, origin, rpcOptions) =>
      await withEngine((connected) => connected.setInfoOriginForSheet(sheet, origin, rpcOptions)),
    setColFormatRuns: async (sheet, col, runs, rpcOptions) =>
      await withEngine((connected) => connected.setColFormatRuns(sheet, col, runs, rpcOptions)),
    recalculate: async (sheet, rpcOptions) => await withEngine((connected) => connected.recalculate(sheet, rpcOptions)),
    getPivotSchema: async (sheet, sourceRangeA1, sampleSize, rpcOptions) =>
      await withEngine((connected) => connected.getPivotSchema(sheet, sourceRangeA1, sampleSize, rpcOptions)),
    getPivotFieldItems: async (sheet, sourceRangeA1, field, rpcOptions) =>
      await withEngine((connected) => connected.getPivotFieldItems(sheet, sourceRangeA1, field, rpcOptions)),
    getPivotFieldItemsPaged: async (sheet, sourceRangeA1, field, offset, limit, rpcOptions) =>
      await withEngine((connected) =>
        connected.getPivotFieldItemsPaged(sheet, sourceRangeA1, field, offset, limit, rpcOptions)
      ),
    calculatePivot: async (sheet, sourceRangeA1, destinationTopLeftA1, config, rpcOptions) =>
      await withEngine((connected) =>
        connected.calculatePivot(sheet, sourceRangeA1, destinationTopLeftA1, config, rpcOptions)
      ),
    goalSeek: async (request, rpcOptions) => await withEngine((connected) => connected.goalSeek(request, rpcOptions)),
    setSheetDimensions: async (sheet, rows, cols, rpcOptions) =>
      await withEngine((connected) => connected.setSheetDimensions(sheet, rows, cols, rpcOptions)),
    getSheetDimensions: async (sheet, rpcOptions) =>
      await withEngine((connected) => connected.getSheetDimensions(sheet, rpcOptions)),
    renameSheet: async (oldName, newName, rpcOptions) =>
      await withEngine((connected) => connected.renameSheet(oldName, newName, rpcOptions)),
    setSheetOrigin: (sheet, origin, rpcOptions) => {
      const promise = withEngine((connected) => connected.setSheetOrigin(sheet, origin, rpcOptions));
      // Sheet origin updates are often called fire-and-forget (scroll path). Attach a no-op rejection
      // handler so teardown/connect races don't surface as unhandled rejections.
      void promise.catch(() => {});
      return promise;
    },
    setSheetDisplayName: async (sheetId, name, rpcOptions) =>
      await withEngine((connected) => connected.setSheetDisplayName(sheetId, name, rpcOptions)),
    setColWidthChars: async (sheet, col, widthChars, rpcOptions) =>
      await withEngine((connected) => connected.setColWidthChars(sheet, col, widthChars, rpcOptions)),
    applyOperation: async (op, rpcOptions) => await withEngine((connected) => connected.applyOperation(op, rpcOptions)),
    rewriteFormulasForCopyDelta: async (requests, rpcOptions) =>
      await withEngine((connected) => connected.rewriteFormulasForCopyDelta(requests, rpcOptions)),
    canonicalizeFormula: async (formula, localeId, referenceStyle, rpcOptions) =>
      await withEngine((connected) => connected.canonicalizeFormula(formula, localeId, referenceStyle, rpcOptions)),
    localizeFormula: async (formula, localeId, referenceStyle, rpcOptions) =>
      await withEngine((connected) => connected.localizeFormula(formula, localeId, referenceStyle, rpcOptions)),
    supportedLocaleIds: async (rpcOptions) => await withEngine((connected) => connected.supportedLocaleIds(rpcOptions)),
    getLocaleInfo: async (localeId, rpcOptions) => await withEngine((connected) => connected.getLocaleInfo(localeId, rpcOptions)),
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
      try {
        connectAbort?.abort();
      } catch {
        // ignore
      }
      connectAbort = null;
      enginePromise = null;
      engine?.terminate();
      engine = null;
      try {
        worker?.terminate();
      } catch {
        // ignore
      }
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
