import type {
  CalcSettings,
  CellChange,
  CellData,
  CellDataCompact,
  CellDataRich,
  CellScalar,
  CellValueRich,
  EngineInfoDto,
  EditOp,
  EditResult,
  FormatRun,
  GoalSeekRequest,
  GoalSeekResponse,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaLocaleInfo,
  FormulaParseOptions,
  FormulaToken,
  InitMessage,
  PivotCalculationResult,
  PivotConfig,
  PivotFieldItems,
  PivotSchema,
  WorkbookInfoDto,
  WorkbookStyleDto,
  RewriteFormulaForCopyDeltaRequest,
  RpcCancel,
  RpcMethod,
  RpcOptions,
  RpcRequest,
  RpcResponseErr,
  RpcResponseOk,
  WorkerOutboundMessage
} from "../protocol.ts";
import { isFormulaInput, normalizeFormulaText } from "../backend/formula.ts";

export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): void;
  addEventListener?(type: string, listener: (event: any) => void, options?: any): void;
  removeEventListener?(type: string, listener: (event: any) => void, options?: any): void;
}

export interface MessagePortLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start?(): void;
  close?(): void;
  addEventListener(type: "message", listener: (event: MessageEvent<unknown>) => void): void;
  removeEventListener(type: "message", listener: (event: MessageEvent<unknown>) => void): void;
}

export interface MessageChannelLike {
  port1: MessagePortLike;
  port2: MessagePort;
}

type PendingRequest = {
  resolve: (value: unknown) => void;
  reject: (err: unknown) => void;
  timeoutId?: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
};

type CellUpdate = { address: string; value: CellScalar; sheet?: string };

function normalizeCellScalar(value: CellScalar): CellScalar {
  if (typeof value !== "string") return value;
  if (!isFormulaInput(value)) return value;
  return normalizeFormulaText(value);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

const FORMULA_PARSE_OPTIONS_ERROR =
  'options must be { localeId?: string, referenceStyle?: "A1" | "R1C1" } or a ParseOptions object';
const FORMULA_PARSE_OPTIONS_NOT_OBJECT_ERROR = "options must be an object";

function normalizeFormulaParseOptions(options: unknown): unknown | undefined {
  if (options == null) return undefined;
  if (!isPlainObject(options)) {
    throw new Error(FORMULA_PARSE_OPTIONS_NOT_OBJECT_ERROR);
  }

  // Backward compatibility: older call sites may pass the full `ParseOptions` object supported by
  // `crates/formula-wasm` (snake_case keys like `reference_style`).
  if ("locale" in options || "reference_style" in options || "normalize_relative_to" in options) {
    return options;
  }

  const allowed = new Set(["localeId", "referenceStyle"]);
  for (const key of Object.keys(options)) {
    if (!allowed.has(key)) {
      throw new Error(FORMULA_PARSE_OPTIONS_ERROR);
    }
  }

  const out: FormulaParseOptions = {};
  const localeId = (options as any).localeId;
  if (localeId !== undefined) {
    if (typeof localeId !== "string") throw new Error(FORMULA_PARSE_OPTIONS_ERROR);
    out.localeId = localeId;
  }

  const referenceStyle = (options as any).referenceStyle;
  if (referenceStyle !== undefined) {
    if (referenceStyle !== "A1" && referenceStyle !== "R1C1") {
      throw new Error(FORMULA_PARSE_OPTIONS_ERROR);
    }
    out.referenceStyle = referenceStyle;
  }

  return out;
}

function pruneNullsDeep(value: unknown): unknown {
  if (value == null) return value;
  if (Array.isArray(value)) return value.map((item) => pruneNullsDeep(item));
  if (typeof value !== "object") return value;

  const out: Record<string, unknown> = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    if (raw === null || raw === undefined) continue;
    out[key] = pruneNullsDeep(raw);
  }
  return out;
}

function pruneUndefinedShallow<T extends object>(value: T): T {
  const out: Record<string, unknown> = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    if (raw === undefined) continue;
    out[key] = raw;
  }
  return out as T;
}

/**
 * Worker-backed Engine client using a MessagePort RPC transport.
 *
 * This is intentionally similar to the proven implementation in
 * `apps/desktop/src/engine/worker/EngineWorker.ts` so desktop and web share the
 * same transport behavior.
 */
export class EngineWorker {
  private readonly worker: WorkerLike;
  private readonly port: MessagePortLike;
  private portListener: ((event: MessageEvent<unknown>) => void) | null = null;
  private portMessageErrorListener: ((event: any) => void) | null = null;
  private workerErrorListener: ((event: any) => void) | null = null;
  private shuttingDown = false;
  private readonly pending = new Map<number, PendingRequest>();
  private nextId = 1;

  private pendingCellUpdates: CellUpdate[] = [];
  private flushPromise: Promise<void> | null = null;

  private constructor(worker: WorkerLike, port: MessagePortLike) {
    this.worker = worker;
    this.port = port;

    const handler = (event: MessageEvent<unknown>) => {
      this.onMessage(event.data as WorkerOutboundMessage);
    };
    this.portListener = handler;
    this.port.addEventListener("message", handler);
    this.port.start?.();

    // If the underlying transport can't deserialize a message, MessagePort emits a `messageerror`
    // event and the corresponding response is dropped. Treat this as fatal so pending RPCs don't
    // hang forever.
    const portAny = this.port as any;
    if (typeof portAny?.addEventListener === "function") {
      const onMessageError = () => {
        this.shutdown(new Error("port messageerror"));
      };
      this.portMessageErrorListener = onMessageError;
      try {
        portAny.addEventListener("messageerror", onMessageError);
      } catch {
        // ignore
      }
    }

    // If the worker crashes after startup, pending RPC promises would otherwise hang forever unless
    // callers supplied timeouts or AbortSignals. Treat worker "error" as fatal and reject all
    // pending requests to avoid leaks/hangs.
    if (typeof this.worker.addEventListener === "function") {
      const onWorkerError = (event: any) => {
        const message = event && typeof event === "object" && "message" in event ? String((event as any).message) : "";
        this.shutdown(new Error(message ? `worker error: ${message}` : "worker error"));
      };
      this.workerErrorListener = onWorkerError;
      try {
        this.worker.addEventListener("error", onWorkerError);
      } catch {
        // ignore
      }
    }
  }

  static async connect(options: {
    worker: WorkerLike;
    wasmModuleUrl?: string;
    wasmBinaryUrl?: string;
    channelFactory?: () => MessageChannelLike;
    signal?: AbortSignal;
    timeoutMs?: number;
  }): Promise<EngineWorker> {
    const channel = options.channelFactory?.() ?? new MessageChannel();
    const port = channel.port1;
    const worker = options.worker;

    const cleanup = () => {
      try {
        port.close?.();
      } catch {
        // ignore
      }
      try {
        // If the port hasn't been transferred yet (e.g. aborted before posting init), ensure we
        // close it to avoid leaking a MessagePort handle.
        (channel.port2 as any)?.close?.();
      } catch {
        // ignore
      }
      try {
        worker.terminate();
      } catch {
        // ignore
      }
    };

    if (options.signal?.aborted) {
      cleanup();
      throw new Error("EngineWorker.connect aborted");
    }

    const timeoutMsRaw = options.timeoutMs;
    const timeoutMs =
      typeof timeoutMsRaw === "number" && Number.isFinite(timeoutMsRaw) && timeoutMsRaw > 0 ? Math.trunc(timeoutMsRaw) : null;

    let timeoutId: ReturnType<typeof setTimeout> | null = null;
    let settled = false;
    // Use `let` so the `onReady` handler can safely reference it even in the (unlikely)
    // event that a mock Worker posts "ready" synchronously.
    let onAbort: (() => void) | null = null;
    let onReady: ((event: MessageEvent<unknown>) => void) | null = null;
    let onWorkerError: ((event: any) => void) | null = null;

    const removeListeners = () => {
      if (onReady) {
        port.removeEventListener("message", onReady);
      }
      if (options.signal && onAbort) {
        options.signal.removeEventListener("abort", onAbort);
      }
      if (onWorkerError && typeof worker.removeEventListener === "function") {
        try {
          worker.removeEventListener("error", onWorkerError);
        } catch {
          // ignore
        }
      }
      if (timeoutId != null) {
        clearTimeout(timeoutId);
        timeoutId = null;
      }
    };

    const ready = new Promise<void>((resolve, reject) => {
      onReady = (event: MessageEvent<unknown>) => {
        const msg = event.data as WorkerOutboundMessage;
        if (msg && typeof msg === "object" && (msg as any).type === "ready") {
          if (settled) return;
          settled = true;
          removeListeners();
          resolve();
        }
      };
      port.addEventListener("message", onReady);

      onAbort = () => {
        if (settled) return;
        settled = true;
        removeListeners();
        cleanup();
        reject(new Error("EngineWorker.connect aborted"));
      };

      if (options.signal) {
        if (options.signal.aborted) {
          onAbort();
          return;
        }
        options.signal.addEventListener("abort", onAbort, { once: true });
      }

      if (typeof worker.addEventListener === "function") {
        onWorkerError = (event: any) => {
          if (settled) return;
          settled = true;
          removeListeners();
          cleanup();
          const message = event && typeof event === "object" && "message" in event ? String((event as any).message) : "";
          reject(new Error(message ? `EngineWorker.connect worker error: ${message}` : "EngineWorker.connect worker error"));
        };
        try {
          worker.addEventListener("error", onWorkerError, { once: true });
        } catch {
          // Some WorkerLike shims may not support options objects; fall back to the basic signature.
          try {
            worker.addEventListener("error", onWorkerError);
          } catch {
            // ignore
          }
        }
      }

      if (timeoutMs != null) {
        timeoutId = setTimeout(() => {
          if (settled) return;
          settled = true;
          removeListeners();
          cleanup();
          reject(new Error(`EngineWorker.connect timed out after ${timeoutMs}ms`));
        }, timeoutMs);
      }
    });
    // The connect handshake can be aborted/failed before we ever `await ready` (e.g. if the abort
    // signal fires before the init message is posted). Attach a no-op rejection handler so a
    // handshake rejection never surfaces as an unhandled rejection.
    void ready.catch(() => {});

    const initMessage: InitMessage = {
      type: "init",
      port: channel.port2,
      wasmModuleUrl: options.wasmModuleUrl ?? "",
      wasmBinaryUrl: options.wasmBinaryUrl
    };
    port.start?.();
    // If the signal was aborted after we installed the abort listener but before the init message
    // is posted, bail out without transferring the port.
    if (options.signal?.aborted) {
      removeListeners();
      cleanup();
      throw new Error("EngineWorker.connect aborted");
    }
    try {
      worker.postMessage(initMessage, [channel.port2]);
    } catch (err) {
      // If the init message fails to post (e.g. the Worker was already terminated), do not leave
      // an `abort` listener registered on the caller's signal (it would keep the MessagePort +
      // closures alive).
      removeListeners();
      cleanup();
      throw err;
    }

    await ready;
    // EngineWorker constructor installs the general message handler and calls `port.start()` again
    // (idempotent). We delay constructing it until after the ready handshake so aborted/failed
    // connect attempts don't leak a never-returned EngineWorker instance.
    return new EngineWorker(worker, port);
  }

  private shutdown(error: Error | ((id: number) => Error)): void {
    if (this.shuttingDown) {
      return;
    }
    this.shuttingDown = true;

    // If the caller fired-and-forgot `setCell`, we may have a microtask-batched `setCells`
    // flush pending. Clearing the update batch avoids posting a message to a closed port
    // (which would reject the flush promise and can surface as an unhandled rejection).
    this.pendingCellUpdates = [];
    for (const [id, pending] of this.pending) {
      pending.timeoutId && clearTimeout(pending.timeoutId);
      if (pending.signal && pending.abortListener) {
        pending.signal.removeEventListener("abort", pending.abortListener);
      }
      pending.reject(typeof error === "function" ? error(id) : error);
    }
    this.pending.clear();

    if (this.portListener) {
      try {
        this.port.removeEventListener("message", this.portListener);
      } catch {
        // ignore
      }
      this.portListener = null;
    }
    if (this.portMessageErrorListener) {
      const portAny = this.port as any;
      if (typeof portAny?.removeEventListener === "function") {
        try {
          portAny.removeEventListener("messageerror", this.portMessageErrorListener);
        } catch {
          // ignore
        }
      }
      this.portMessageErrorListener = null;
    }
    if (this.workerErrorListener && typeof this.worker.removeEventListener === "function") {
      try {
        this.worker.removeEventListener("error", this.workerErrorListener);
      } catch {
        // ignore
      }
      this.workerErrorListener = null;
    }
    try {
      this.worker.terminate();
    } catch {
      // ignore
    }
    try {
      this.port.close?.();
    } catch {
      // ignore
    }
  }

  terminate(): void {
    this.shutdown((id) => new Error(`worker terminated (request ${id})`));
  }

  /**
   * Simple liveness check for the worker transport.
   *
   * This RPC is independent of any loaded workbook and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async ping(options?: RpcOptions): Promise<string> {
    return (await this.invoke("ping", {}, options)) as string;
  }

  async newWorkbook(options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("newWorkbook", {}, options);
  }

  async loadWorkbookFromJson(json: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("loadFromJson", { json }, options);
  }

  /**
   * Load a workbook from raw `.xlsx` bytes.
   *
   * Note: the payload is transferred to the worker to avoid an extra
   * structured-clone copy.
   *
   * - If `bytes` spans its entire backing buffer (`byteOffset === 0` and
   *   `byteLength === bytes.buffer.byteLength`), the buffer is transferred and
   *   detached on the calling side.
   * - If `bytes` is a view into a larger buffer, we first copy it to a compact
   *   `Uint8Array` so only the relevant range is transferred (leaving the
   *   original buffer intact).
   */
  async loadWorkbookFromXlsxBytes(bytes: Uint8Array, options?: RpcOptions): Promise<void> {
    await this.flush();
    let payload = bytes;
    if (payload.byteOffset !== 0 || payload.byteLength !== payload.buffer.byteLength) {
      // `Uint8Array#slice()` copies into a new backing buffer, but `Buffer#slice()` (Node) returns a
      // view into the same underlying pool. Normalize to a compact standalone `Uint8Array` so we
      // transfer only the relevant bytes.
      payload = payload.slice();
      if (payload.byteOffset !== 0 || payload.byteLength !== payload.buffer.byteLength) {
        payload = new Uint8Array(payload);
      }
    }
    await this.invoke("loadFromXlsxBytes", { bytes: payload }, options, [payload.buffer]);
  }

  /**
   * Load a workbook from Office-encrypted `.xlsx` bytes, decrypting it in WASM with `password`.
   *
   * Note: the payload is transferred to the worker to avoid an extra
   * structured-clone copy (same semantics as `loadWorkbookFromXlsxBytes`).
   */
  async loadWorkbookFromEncryptedXlsxBytes(bytes: Uint8Array, password: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    if (typeof password !== "string") {
      throw new Error("password must be a string");
    }
    let payload = bytes;
    if (payload.byteOffset !== 0 || payload.byteLength !== payload.buffer.byteLength) {
      // See `loadWorkbookFromXlsxBytes` for why we normalize to a standalone buffer.
      payload = payload.slice();
      if (payload.byteOffset !== 0 || payload.byteLength !== payload.buffer.byteLength) {
        payload = new Uint8Array(payload);
      }
    }
    await this.invoke("loadFromEncryptedXlsxBytes", { bytes: payload, password }, options, [payload.buffer]);
  }

  async toJson(options?: RpcOptions): Promise<string> {
    await this.flush();
    return (await this.invoke("toJson", {}, options)) as string;
  }

  async getWorkbookInfo(options?: RpcOptions): Promise<WorkbookInfoDto> {
    await this.flush();
    return (await this.invoke("getWorkbookInfo", {}, options)) as WorkbookInfoDto;
  }

  async getCell(
    address: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellData> {
    await this.flush();
    return (await this.invoke("getCell", { address, sheet }, options)) as CellData;
  }

  async getCellRich(
    address: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellDataRich> {
    await this.flush();
    const raw = (await this.invoke("getCellRich", { address, sheet }, options)) as CellDataRich;
    // WASM-bindgen DTOs sometimes include `null` for missing optional fields (instead of omitting
    // them). Normalize these to better match the Engine TS API (which uses `undefined`/omitted
    // keys for optional fields) and to keep rich values stable across roundtrips.
    return pruneNullsDeep(raw) as CellDataRich;
  }

  async getRange(
    range: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellData[][]> {
    await this.flush();
    return (await this.invoke("getRange", { range, sheet }, options)) as CellData[][];
  }

  async getRangeCompact(
    range: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellDataCompact[][]> {
    await this.flush();
    return (await this.invoke("getRangeCompact", { range, sheet }, options)) as CellDataCompact[][];
  }

  setCell(address: string, value: CellScalar, sheet?: string): Promise<void> {
    this.pendingCellUpdates.push({ address, value: normalizeCellScalar(value), sheet });
    const promise = this.scheduleFlush();
    // `setCell` batching is sometimes fire-and-forget. Attach a no-op rejection handler so a
    // failed flush doesn't surface as an unhandled rejection when callers don't await.
    // (Awaiting the returned promise still observes the rejection.)
    void promise.catch(() => {});
    return promise;
  }

  async setCellRich(address: string, value: CellValueRich | null, sheet?: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setCellRich", { address, value, sheet }, options);
  }

  async setCells(
    updates: Array<{ address: string; value: CellScalar; sheet?: string }>,
    options?: RpcOptions
  ): Promise<void> {
    if (updates.length === 0) {
      return;
    }
    await this.flush();
    const normalized = updates.map((update) => ({ ...update, value: normalizeCellScalar(update.value) }));
    await this.invoke("setCells", { updates: normalized }, options);
  }

  async setRange(
    range: string,
    values: CellScalar[][],
    sheet?: string,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    const normalizedValues = values.map((row) => row.map((value) => normalizeCellScalar(value)));
    await this.invoke("setRange", { range, values: normalizedValues, sheet }, options);
  }

  async internStyle(style: WorkbookStyleDto | null, options?: RpcOptions): Promise<number> {
    await this.flush();
    return (await this.invoke("internStyle", { style }, options)) as number;
  }

  async setColFormatRuns(
    sheet: string,
    col: number,
    runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    await this.invoke("setColFormatRuns", { sheet, col, runs }, options);
  }

  async setLocale(localeId: string, options?: RpcOptions): Promise<boolean> {
    await this.flush();
    return (await this.invoke("setLocale", { localeId }, options)) as boolean;
  }

  async getCalcSettings(options?: RpcOptions): Promise<CalcSettings> {
    await this.flush();
    return (await this.invoke("getCalcSettings", {}, options)) as CalcSettings;
  }

  async setCalcSettings(settings: CalcSettings, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setCalcSettings", { settings }, options);
  }

  async setEngineInfo(info: EngineInfoDto, options?: RpcOptions): Promise<void> {
    await this.flush();
    const normalized = pruneUndefinedShallow(info);
    const memavail = (normalized as any).memavail as unknown;
    if (memavail !== undefined && memavail !== null) {
      if (typeof memavail !== "number" || !Number.isFinite(memavail)) {
        throw new Error("memavail must be a finite number");
      }
    }
    const totmem = (normalized as any).totmem as unknown;
    if (totmem !== undefined && totmem !== null) {
      if (typeof totmem !== "number" || !Number.isFinite(totmem)) {
        throw new Error("totmem must be a finite number");
      }
    }
    await this.invoke("setEngineInfo", { info: normalized }, options);
  }

  async setInfoOrigin(origin: string | null, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setInfoOrigin", { origin }, options);
  }

  async setInfoOriginForSheet(sheet: string, origin: string | null, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setInfoOriginForSheet", { sheet, origin }, options);
  }

  async setWorkbookFileMetadata(
    directory: string | null,
    filename: string | null,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    await this.invoke("setWorkbookFileMetadata", { directory, filename }, options);
  }

  async setCellStyleId(
    address: string,
    styleId: number,
    sheet?: string,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    await this.invoke("setCellStyleId", { sheet, address, styleId }, options);
  }

  async setRowStyleId(sheet: string, row: number, styleId: number | null, options?: RpcOptions): Promise<void>;
  async setRowStyleId(row: number, styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  async setRowStyleId(
    sheetOrRow: string | number,
    rowOrStyleId: number,
    styleIdOrSheet?: number | string | null | RpcOptions,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    let sheet: string | undefined;
    let row: number;
    let styleId: number | null;
    let finalRpcOptions: RpcOptions | undefined;

    if (typeof sheetOrRow === "string") {
      sheet = sheetOrRow;
      row = rowOrStyleId;
      styleId = styleIdOrSheet as number | null;
      finalRpcOptions = options;
    } else {
      row = sheetOrRow;
      styleId = rowOrStyleId;
      if (typeof styleIdOrSheet === "string" || styleIdOrSheet == null) {
        sheet = typeof styleIdOrSheet === "string" ? styleIdOrSheet : undefined;
        finalRpcOptions = options;
      } else {
        // Allow: setRowStyleId(row, styleId, rpcOptions)
        sheet = undefined;
        finalRpcOptions = styleIdOrSheet as RpcOptions;
      }
    }

    await this.invoke("setRowStyleId", { sheet, row, styleId }, finalRpcOptions);
  }

  async setColStyleId(sheet: string, col: number, styleId: number | null, options?: RpcOptions): Promise<void>;
  async setColStyleId(col: number, styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  async setColStyleId(
    sheetOrCol: string | number,
    colOrStyleId: number,
    styleIdOrSheet?: number | string | null | RpcOptions,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    let sheet: string | undefined;
    let col: number;
    let styleId: number | null;
    let finalRpcOptions: RpcOptions | undefined;

    if (typeof sheetOrCol === "string") {
      sheet = sheetOrCol;
      col = colOrStyleId;
      styleId = styleIdOrSheet as number | null;
      finalRpcOptions = options;
    } else {
      col = sheetOrCol;
      styleId = colOrStyleId;
      if (typeof styleIdOrSheet === "string" || styleIdOrSheet == null) {
        sheet = typeof styleIdOrSheet === "string" ? styleIdOrSheet : undefined;
        finalRpcOptions = options;
      } else {
        sheet = undefined;
        finalRpcOptions = styleIdOrSheet as RpcOptions;
      }
    }

    await this.invoke("setColStyleId", { sheet, col, styleId }, finalRpcOptions);
  }

  async setSheetDefaultStyleId(sheet: string, styleId: number | null, options?: RpcOptions): Promise<void>;
  async setSheetDefaultStyleId(styleId: number, sheet?: string, options?: RpcOptions): Promise<void>;
  async setSheetDefaultStyleId(
    sheetOrStyleId: string | number,
    styleIdOrSheet?: number | string | null | RpcOptions,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    let sheet: string | undefined;
    let styleId: number | null;
    let finalRpcOptions: RpcOptions | undefined;

    if (typeof sheetOrStyleId === "string") {
      sheet = sheetOrStyleId;
      styleId = styleIdOrSheet as number | null;
      finalRpcOptions = options;
    } else {
      styleId = sheetOrStyleId;
      if (typeof styleIdOrSheet === "string" || styleIdOrSheet == null) {
        sheet = typeof styleIdOrSheet === "string" ? styleIdOrSheet : undefined;
        finalRpcOptions = options;
      } else {
        sheet = undefined;
        finalRpcOptions = styleIdOrSheet as RpcOptions;
      }
    }

    await this.invoke("setSheetDefaultStyleId", { sheet, styleId }, finalRpcOptions);
  }

  async setFormatRunsByCol(
    sheet: string,
    col: number,
    runs: FormatRun[],
    options?: RpcOptions
  ): Promise<void>;
  async setFormatRunsByCol(
    col: number,
    runs: FormatRun[],
    sheet?: string,
    options?: RpcOptions
  ): Promise<void>;
  async setFormatRunsByCol(
    sheetOrCol: string | number,
    colOrRuns: number | FormatRun[],
    runsOrSheet?: FormatRun[] | string | null | RpcOptions,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    let sheet: string | undefined;
    let col: number;
    let runs: FormatRun[];
    let finalRpcOptions: RpcOptions | undefined;

    if (typeof sheetOrCol === "string") {
      sheet = sheetOrCol;
      col = colOrRuns as number;
      runs = runsOrSheet as FormatRun[];
      finalRpcOptions = options;
    } else {
      col = sheetOrCol;
      runs = colOrRuns as FormatRun[];
      if (typeof runsOrSheet === "string" || runsOrSheet == null) {
        sheet = typeof runsOrSheet === "string" ? runsOrSheet : undefined;
        finalRpcOptions = options;
      } else {
        // Allow: setFormatRunsByCol(col, runs, rpcOptions)
        sheet = undefined;
        finalRpcOptions = runsOrSheet as RpcOptions;
      }
    }

    await this.invoke("setFormatRunsByCol", { sheet, col, runs }, finalRpcOptions);
  }
  /**
   * Set (or clear) a per-column width override.
   *
   * `width` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
   *
   * Prefer `setColWidthChars(sheet, col, widthChars)` for an explicit unit name.
   */
  async setColWidth(col: number, width: number | null, sheet?: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setColWidth", { sheet, col, width }, options);
  }

  async setColHidden(col: number, hidden: boolean, sheet?: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setColHidden", { sheet, col, hidden }, options);
  }

  async recalculate(sheet?: string, options?: RpcOptions): Promise<CellChange[]> {
    await this.flush();
    return (await this.invoke("recalculate", { sheet }, options)) as CellChange[];
  }

  async getPivotSchema(
    sheet: string,
    sourceRangeA1: string,
    sampleSize?: number,
    options?: RpcOptions
  ): Promise<PivotSchema> {
    await this.flush();
    return (await this.invoke("getPivotSchema", { sheet, sourceRangeA1, sampleSize }, options)) as PivotSchema;
  }

  async getPivotFieldItems(sheet: string, sourceRangeA1: string, field: string, options?: RpcOptions): Promise<PivotFieldItems> {
    await this.flush();
    return (await this.invoke("getPivotFieldItems", { sheet, sourceRangeA1, field }, options)) as PivotFieldItems;
  }

  async getPivotFieldItemsPaged(
    sheet: string,
    sourceRangeA1: string,
    field: string,
    offset: number,
    limit: number,
    options?: RpcOptions
  ): Promise<PivotFieldItems> {
    await this.flush();
    return (await this.invoke(
      "getPivotFieldItemsPaged",
      { sheet, sourceRangeA1, field, offset, limit },
      options
    )) as PivotFieldItems;
  }

  async calculatePivot(
    sheet: string,
    sourceRangeA1: string,
    destinationTopLeftA1: string,
    config: PivotConfig,
    options?: RpcOptions
  ): Promise<PivotCalculationResult> {
    await this.flush();
    // `serde_wasm_bindgen` does not accept `undefined` values inside structs; prune optional keys
    // so callers can pass `{ foo?: undefined }` without breaking deserialization.
    const normalizedConfig = pruneNullsDeep(config) as PivotConfig;
    return (await this.invoke(
      "calculatePivot",
      { sheet, sourceRangeA1, destinationTopLeftA1, config: normalizedConfig },
      options
    )) as PivotCalculationResult;
  }

  async goalSeek(request: GoalSeekRequest, options?: RpcOptions): Promise<GoalSeekResponse> {
    await this.flush();
    // `serde_wasm_bindgen` treats `{ foo: undefined }` as an error for `Option<T>` fields. Strip
    // undefined optional tuning keys so callers can pass `GoalSeekRequest` objects constructed from
    // partially-filled UI state.
    //
    // The Rust goal seek solver defaults to a relatively loose tolerance (currently 0.001). For the
    // TS EngineWorker API we default to a stricter tolerance so callers get stable, precise results
    // without having to specify tuning knobs.
    const normalized = pruneUndefinedShallow({ ...request, tolerance: request.tolerance ?? 1e-6 });
    return (await this.invoke("goalSeek", normalized, options)) as GoalSeekResponse;
  }

  async applyOperation(op: EditOp, options?: RpcOptions): Promise<EditResult> {
    await this.flush();
    return (await this.invoke("applyOperation", { op }, options)) as EditResult;
  }

  /**
   * Rewrite formulas as if they were copied by `(deltaRow, deltaCol)`.
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async rewriteFormulasForCopyDelta(
    requests: RewriteFormulaForCopyDeltaRequest[],
    rpcOptions?: RpcOptions
  ): Promise<string[]> {
    return (await this.invoke("rewriteFormulasForCopyDelta", { requests }, rpcOptions)) as string[];
  }

  /**
   * Canonicalize a locale-specific formula string into the engine's persisted form.
   *
   * This call is independent of any loaded workbook.
   */
  async canonicalizeFormula(
    formula: string,
    localeId: string,
    referenceStyle?: "A1" | "R1C1",
    rpcOptions?: RpcOptions
  ): Promise<string> {
    return (await this.invoke(
      "canonicalizeFormula",
      { formula, localeId, referenceStyle },
      rpcOptions
    )) as string;
  }

  /**
   * Localize a canonical (English) formula string for display in `localeId`.
   *
   * This call is independent of any loaded workbook.
   */
  async localizeFormula(
    formula: string,
    localeId: string,
    referenceStyle?: "A1" | "R1C1",
    rpcOptions?: RpcOptions
  ): Promise<string> {
    return (await this.invoke(
      "localizeFormula",
      { formula, localeId, referenceStyle },
      rpcOptions
    )) as string;
  }

  /**
   * Return the list of formula locale ids supported by the underlying engine build.
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async supportedLocaleIds(rpcOptions?: RpcOptions): Promise<string[]> {
    return (await this.invoke("supportedLocaleIds", {}, rpcOptions)) as string[];
  }

  /**
   * Return locale metadata used by formula parsing/rendering (separators, boolean literals, RTL flag, etc).
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async getLocaleInfo(localeId: string, rpcOptions?: RpcOptions): Promise<FormulaLocaleInfo> {
    return (await this.invoke("getLocaleInfo", { localeId }, rpcOptions)) as FormulaLocaleInfo;
  }

  /**
   * Tokenize a formula string for editor tooling (syntax highlighting, etc).
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async lexFormula(formula: string, options?: FormulaParseOptions, rpcOptions?: RpcOptions): Promise<FormulaToken[]>;
  async lexFormula(formula: string, rpcOptions?: RpcOptions): Promise<FormulaToken[]>;
  async lexFormula(
    formula: string,
    optionsOrRpcOptions?: FormulaParseOptions | RpcOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaToken[]> {
    const isFormulaParseOptions = (value: unknown): value is FormulaParseOptions =>
      isPlainObject(value) && !("signal" in value) && !("timeoutMs" in value);
    const isRpcOptions = (value: unknown): value is RpcOptions =>
      Boolean(value && typeof value === "object" && ("signal" in value || "timeoutMs" in value));

    const options = isRpcOptions(optionsOrRpcOptions) && !isFormulaParseOptions(optionsOrRpcOptions)
      ? undefined
      : (optionsOrRpcOptions as FormulaParseOptions | undefined);
    const finalRpcOptions = isRpcOptions(optionsOrRpcOptions) ? optionsOrRpcOptions : rpcOptions;

    const normalizedOptions = normalizeFormulaParseOptions(options);
    return (await this.invoke("lexFormula", { formula, options: normalizedOptions }, finalRpcOptions)) as FormulaToken[];
  }

  /**
   * Best-effort lexer for editor syntax highlighting.
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async lexFormulaPartial(
    formula: string,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialLexResult>;
  async lexFormulaPartial(formula: string, rpcOptions?: RpcOptions): Promise<FormulaPartialLexResult>;
  async lexFormulaPartial(
    formula: string,
    optionsOrRpcOptions?: FormulaParseOptions | RpcOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialLexResult> {
    const isFormulaParseOptions = (value: unknown): value is FormulaParseOptions =>
      isPlainObject(value) && !("signal" in value) && !("timeoutMs" in value);
    const isRpcOptions = (value: unknown): value is RpcOptions =>
      Boolean(value && typeof value === "object" && ("signal" in value || "timeoutMs" in value));

    const options = isRpcOptions(optionsOrRpcOptions) && !isFormulaParseOptions(optionsOrRpcOptions)
      ? undefined
      : (optionsOrRpcOptions as FormulaParseOptions | undefined);
    const finalRpcOptions = isRpcOptions(optionsOrRpcOptions) ? optionsOrRpcOptions : rpcOptions;

    const normalizedOptions = normalizeFormulaParseOptions(options);
    return (await this.invoke("lexFormulaPartial", { formula, options: normalizedOptions }, finalRpcOptions)) as FormulaPartialLexResult;
  }

  /**
   * Best-effort partial parse for editor/autocomplete scenarios.
   *
   * `cursor` (when provided) is expressed as a **UTF-16 code unit** offset (JS
   * string indexing). This matches the span units returned by `lexFormula`.
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async parseFormulaPartial(
    formula: string,
    cursor?: number,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialParseResult>;
  async parseFormulaPartial(
    formula: string,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialParseResult>;
  async parseFormulaPartial(
    formula: string,
    cursorOrOptions?: number | FormulaParseOptions,
    optionsOrRpcOptions?: FormulaParseOptions | RpcOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialParseResult> {
    const isFormulaParseOptions = (value: unknown): value is FormulaParseOptions =>
      isPlainObject(value) && !("signal" in value) && !("timeoutMs" in value);
    const isRpcOptions = (value: unknown): value is RpcOptions =>
      Boolean(value && typeof value === "object" && ("signal" in value || "timeoutMs" in value));

    let cursor: number | undefined;
    let options: FormulaParseOptions | undefined;
    let finalRpcOptions: RpcOptions | undefined;

    if (typeof cursorOrOptions === "number") {
      cursor = cursorOrOptions;
      if (isRpcOptions(optionsOrRpcOptions)) {
        // Allow: parseFormulaPartial(formula, cursor, rpcOptions)
        options = undefined;
        finalRpcOptions = optionsOrRpcOptions;
      } else {
        options = optionsOrRpcOptions as FormulaParseOptions | undefined;
        finalRpcOptions = rpcOptions;
      }
    } else {
      // Cursor omitted. Support both:
      // - parseFormulaPartial(formula, options?, rpcOptions?)
      // - parseFormulaPartial(formula, undefined, options?, rpcOptions?) (legacy call sites)
      if (isFormulaParseOptions(cursorOrOptions)) {
        options = cursorOrOptions;
        finalRpcOptions = (isRpcOptions(optionsOrRpcOptions) ? optionsOrRpcOptions : rpcOptions) as RpcOptions | undefined;
      } else if (cursorOrOptions === undefined) {
        if (isFormulaParseOptions(optionsOrRpcOptions)) {
          options = optionsOrRpcOptions;
          finalRpcOptions = rpcOptions;
        } else if (isRpcOptions(optionsOrRpcOptions)) {
          options = undefined;
          finalRpcOptions = optionsOrRpcOptions;
        } else if (optionsOrRpcOptions != null && typeof optionsOrRpcOptions === "object") {
          // Unknown object: treat it as a parse options bag and let the WASM boundary validate.
          // This ensures common mistakes (e.g. wrong casing on `localeId`) surface as a clear error
          // instead of being silently ignored.
          options = optionsOrRpcOptions as FormulaParseOptions;
          finalRpcOptions = rpcOptions;
        } else {
          options = undefined;
          finalRpcOptions = rpcOptions;
        }
      } else if (isRpcOptions(cursorOrOptions)) {
        // Allow: parseFormulaPartial(formula, rpcOptions)
        options = undefined;
        finalRpcOptions = cursorOrOptions;
      } else {
        // Unknown object; assume it's a parse options bag.
        options = cursorOrOptions as FormulaParseOptions;
        finalRpcOptions = optionsOrRpcOptions as RpcOptions | undefined;
      }
    }

    const normalizedOptions = normalizeFormulaParseOptions(options);
    return (await this.invoke(
      "parseFormulaPartial",
      { formula, cursor, options: normalizedOptions },
      finalRpcOptions
    )) as FormulaPartialParseResult;
  }

  async setSheetDimensions(sheet: string, rows: number, cols: number, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setSheetDimensions", { sheet, rows, cols }, options);
  }

  async getSheetDimensions(sheet: string, options?: RpcOptions): Promise<{ rows: number; cols: number }> {
    await this.flush();
    return (await this.invoke("getSheetDimensions", { sheet }, options)) as { rows: number; cols: number };
  }

  async renameSheet(oldName: string, newName: string, options?: RpcOptions): Promise<boolean> {
    await this.flush();
    return (await this.invoke("renameSheet", { oldName, newName }, options)) as boolean;
  }

  async setSheetOrigin(sheet: string, origin: string | null, options?: RpcOptions): Promise<void> {
    // Origin is UI/view metadata (scroll position + frozen panes) and is independent of pending
    // cell edits. Avoid forcing a `setCells` flush on high-frequency scroll updates.
    await this.invoke("setSheetOrigin", { sheet, origin }, options);
  }

  async setSheetDisplayName(sheetId: string, name: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setSheetDisplayName", { sheetId, name }, options);
  }

  /**
   * Set (or clear) a per-column width override.
   *
   * `widthChars` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
   */
  async setColWidthChars(sheet: string, col: number, widthChars: number | null, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setColWidthChars", { sheet, col, widthChars }, options);
  }

  private async scheduleFlush(): Promise<void> {
    if (this.flushPromise) {
      return this.flushPromise;
    }
    if (this.pendingCellUpdates.length === 0) {
      return;
    }

    this.flushPromise = new Promise((resolve, reject) => {
      queueMicrotask(async () => {
        try {
          await this.flush();
          resolve();
        } catch (err) {
          reject(err);
        } finally {
          this.flushPromise = null;
        }
      });
    });

    return this.flushPromise;
  }

  private async flush(): Promise<void> {
    if (this.pendingCellUpdates.length === 0) {
      return;
    }
    const updates = this.pendingCellUpdates;
    this.pendingCellUpdates = [];
    await this.invoke("setCells", { updates });
  }

  private invoke(
    method: RpcMethod,
    params: unknown,
    options?: RpcOptions,
    transfer?: Transferable[]
  ): Promise<unknown> {
    const id = this.nextId++;
    const request: RpcRequest = { type: "request", id, method, params };

    return new Promise((resolve, reject) => {
      const pending: PendingRequest = { resolve, reject };
      this.pending.set(id, pending);

      if (options?.timeoutMs != null) {
        pending.timeoutId = setTimeout(() => {
          this.cancel(id, new Error(`request timed out after ${options.timeoutMs}ms`));
        }, options.timeoutMs);
      }

      if (options?.signal) {
        pending.signal = options.signal;
        if (options.signal.aborted) {
          this.cancel(id, new Error("request aborted"));
          return;
        }
        pending.abortListener = () => this.cancel(id, new Error("request aborted"));
        options.signal.addEventListener("abort", pending.abortListener, { once: true });
      }

      try {
        if (transfer && transfer.length > 0) {
          this.port.postMessage(request, transfer);
        } else {
          this.port.postMessage(request);
        }
      } catch (err) {
        this.pending.delete(id);
        pending.timeoutId && clearTimeout(pending.timeoutId);
        if (pending.signal && pending.abortListener) {
          pending.signal.removeEventListener("abort", pending.abortListener);
        }
        reject(err);
      }
    });
  }

  private cancel(id: number, error: Error): void {
    const pending = this.pending.get(id);
    if (!pending) {
      return;
    }

    this.pending.delete(id);
    pending.timeoutId && clearTimeout(pending.timeoutId);
    if (pending.signal && pending.abortListener) {
      pending.signal.removeEventListener("abort", pending.abortListener);
    }

    const cancelMessage: RpcCancel = { type: "cancel", id };
    try {
      this.port.postMessage(cancelMessage);
    } catch {
      // Ignore; cancellation is best-effort and may race with worker teardown.
    }
    pending.reject(error);
  }

  private onMessage(msg: WorkerOutboundMessage): void {
    if (msg.type !== "response") {
      return;
    }

    const pending = this.pending.get(msg.id);
    if (!pending) {
      return;
    }

    this.pending.delete(msg.id);
    pending.timeoutId && clearTimeout(pending.timeoutId);
    if (pending.signal && pending.abortListener) {
      pending.signal.removeEventListener("abort", pending.abortListener);
    }

    if ((msg as RpcResponseOk).ok) {
      pending.resolve((msg as RpcResponseOk).result);
      return;
    }

    pending.reject(new Error((msg as RpcResponseErr).error));
  }
}
