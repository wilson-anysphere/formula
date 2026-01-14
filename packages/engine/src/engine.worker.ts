/// <reference lib="webworker" />

import type {
  CellScalar,
  FormulaParseOptions,
  InitMessage,
  RpcCancel,
  RpcRequest,
  SheetUsedRangeDto,
  WorkbookInfoDto,
  WorkerInboundMessage,
  WorkerOutboundMessage,
} from "./protocol.ts";

type WasmWorkbookInstance = {
  getWorkbookInfo?: () => unknown;
  getCell(address: string, sheet?: string): unknown;
  getCellRich?: (address: string, sheet?: string) => unknown;
  goalSeek?: (request: unknown) => unknown;
  getPivotSchema?: (sheet: string, sourceRangeA1: string, sampleSize?: number) => unknown;
  getPivotFieldItems?: (sheet: string, sourceRangeA1: string, field: string) => unknown;
  getPivotFieldItemsPaged?: (
    sheet: string,
    sourceRangeA1: string,
    field: string,
    offset: number,
    limit: number
  ) => unknown;
  calculatePivot?: (
    sheet: string,
    sourceRangeA1: string,
    destinationTopLeftA1: string,
    config: unknown
  ) => unknown;
  setCell(address: string, value: CellScalar, sheet?: string): void;
  setCellRich?: (address: string, value: unknown, sheet?: string) => void;
  setCellPhonetic?: (address: string, phonetic: string | null | undefined, sheet?: string) => void;
  getCellPhonetic?: (address: string, sheet?: string) => string | undefined;
  setCells?: (updates: Array<{ address: string; value: CellScalar; sheet?: string }>) => void;
  internStyle?: (style: unknown) => number;
  setColFormatRuns?: (sheet: string, col: number, runs: unknown) => void;
  setLocale?: (localeId: string) => boolean;
  getTextCodepage?: () => number;
  setTextCodepage?: (codepage: number) => void;
  getCalcSettings?: () => unknown;
  setCalcSettings?: (settings: unknown) => void;
  getRange(range: string, sheet?: string): unknown;
  getRangeCompact?: (range: string, sheet?: string) => unknown;
  setRange(range: string, values: CellScalar[][], sheet?: string): void;
  recalculate(sheet?: string): unknown;
  setEngineInfo?: (info: unknown) => void;
  // Legacy engine-info setters (pre `setEngineInfo`). These are retained by `crates/formula-wasm`
  // for backward compatibility. The worker will fan out `setEngineInfo` calls to these when
  // `setEngineInfo` itself is missing.
  setInfoSystem?: (system: string | null | undefined) => void;
  setInfoDirectory?: (directory: string | null | undefined) => void;
  setInfoOSVersion?: (osversion: string | null | undefined) => void;
  setInfoRelease?: (release: string | null | undefined) => void;
  setInfoVersion?: (version: string | null | undefined) => void;
  setInfoMemAvail?: (memavail: number | null | undefined) => void;
  setInfoTotMem?: (totmem: number | null | undefined) => void;
  setInfoOrigin?: (origin: string | null) => void;
  setInfoOriginForSheet?: (sheet: string, origin: string | null) => void;
  applyOperation?: (op: unknown) => unknown;
  setSheetDimensions?: (sheet: string, rows: number, cols: number) => void;
  getSheetDimensions?: (sheet: string) => { rows: number; cols: number };
  renameSheet?: (oldName: string, newName: string) => boolean;
  setWorkbookFileMetadata?: (directory: string | null, filename: string | null) => void;
  // `crates/formula-wasm` has historically used both a sheet-first and sheet-last signature for
  // `setCellStyleId`. The worker prefers the modern sheet-first form and falls back at runtime
  // when it detects the legacy ordering.
  setCellStyleId?: {
    (sheet: string, address: string, styleId: number): void;
    (address: string, styleId: number, sheet?: string): void;
  };
  setRowStyleId?: (sheet: string, row: number, styleId?: number) => void;
  setColStyleId?: (sheet: string, col: number, styleId?: number) => void;
  setSheetDefaultStyleId?: (sheet: string, styleId?: number) => void;
  setFormatRunsByCol?: (
    sheet: string,
    col: number,
    runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>
  ) => void;
  setColWidth?: (sheet: string, col: number, width: number | null) => void;
  setColHidden?: (sheet: string, col: number, hidden: boolean) => void;
  setSheetOrigin?: (sheet: string, origin: string | null) => void;
  setColWidthChars?: (sheet: string, col: number, widthChars: number | null) => void;
  setSheetDisplayName?: (sheetId: string, name: string) => void;
  toJson(): string;
};

type UsedRangeState = SheetUsedRangeDto;

type EngineWorkbookJson = {
  sheetOrder?: unknown;
  textCodepage?: unknown;
  sheets?: Record<
    string,
    {
      cells?: Record<string, unknown>;
      rowCount?: number;
      colCount?: number;
      visibility?: string;
      tabColor?: unknown;
    }
  >;
};

const DEFAULT_SHEET_NAME = "Sheet1";

function normalizeSheetName(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed === "" ? undefined : trimmed;
}

function sheetNameOrDefault(value: unknown): string {
  return normalizeSheetName(value) ?? DEFAULT_SHEET_NAME;
}

function colNameToIndex(colName: string): number {
  if (colName.trim() === "") {
    throw new Error("Expected a non-empty column name");
  }

  let n = 0;
  for (const ch of colName.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`Invalid column name: ${colName}`);
    }
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

function fromA1(address: string): { row0: number; col0: number } {
  const trimmed = address.trim();
  const match = /^\$?([A-Za-z]+)\$?([1-9][0-9]*)$/.exec(trimmed);
  if (!match) {
    throw new Error(`Invalid A1 address: ${address}`);
  }

  const [, colName, rowStr] = match;
  const row1 = Number(rowStr);
  if (!Number.isInteger(row1) || row1 < 1) {
    throw new Error(`Invalid row in A1 address: ${address}`);
  }

  return { row0: row1 - 1, col0: colNameToIndex(colName) };
}

function updateUsedRange(map: Map<string, UsedRangeState>, sheetId: string, row: number, col: number): void {
  const existing = map.get(sheetId);
  if (!existing) {
    map.set(sheetId, { start_row: row, end_row: row, start_col: col, end_col: col });
    return;
  }
  existing.start_row = Math.min(existing.start_row, row);
  existing.end_row = Math.max(existing.end_row, row);
  existing.start_col = Math.min(existing.start_col, col);
  existing.end_col = Math.max(existing.end_col, col);
}

type WasmModule = {
  default?: (module_or_path?: unknown) => Promise<void> | void;
  supportedLocaleIds?: () => unknown;
  getLocaleInfo?: (localeId: string) => unknown;
  lexFormula: (formula: string, options?: FormulaParseOptions) => unknown;
  parseFormulaPartial: (formula: string, cursor?: number, options?: FormulaParseOptions) => unknown;
  canonicalizeFormula?: (formula: string, localeId: string, referenceStyle?: "A1" | "R1C1") => string;
  localizeFormula?: (formula: string, localeId: string, referenceStyle?: "A1" | "R1C1") => string;
  rewriteFormulasForCopyDelta?: (requests: unknown) => unknown;
  WasmWorkbook: {
    new (): WasmWorkbookInstance;
    fromJson(json: string): WasmWorkbookInstance;
    fromXlsxBytes?: (bytes: Uint8Array) => WasmWorkbookInstance;
    fromEncryptedXlsxBytes?: (bytes: Uint8Array, password: string) => WasmWorkbookInstance;
  };
  lexFormulaPartial?: (formula: string, options?: unknown) => unknown;
};

let port: MessagePort | null = null;
let portListener: ((event: MessageEvent<unknown>) => void) | null = null;
let transportGeneration = 0;
let wasmModuleUrl: string | null = null;
let wasmBinaryUrl: string | null = null;
let workbook: WasmWorkbookInstance | null = null;

function normalizeCellScalar(value: unknown): CellScalar {
  // wasm-bindgen maps `Option<T>` to `T | undefined` in JS. Our public protocol uses `null`
  // for empty cells, so normalize `undefined` at the worker boundary.
  if (value === undefined) return null;
  if (value === null) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return null;
}

function normalizeCellData(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  const cell = value as any;
  if (!("input" in cell) && !("value" in cell)) return value;
  // Mutate in place to avoid allocating a second object per cell.
  if ("input" in cell) cell.input = normalizeCellScalar(cell.input);
  if ("value" in cell) cell.value = normalizeCellScalar(cell.value);
  return cell;
}

function normalizeRangeData(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  // Mutate in place to avoid allocating a second set of arrays/objects.
  for (const row of value) {
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      normalizeCellData(cell);
    }
  }
  return value;
}

function normalizeCellDataCompact(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  // Compact payload shape: [input, value]
  if (value.length < 2) return value;
  // Mutate in place to avoid allocating a second set of arrays before structured-clone.
  (value as any)[0] = normalizeCellScalar((value as any)[0]);
  (value as any)[1] = normalizeCellScalar((value as any)[1]);
  return value;
}

function normalizeRangeDataCompact(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  // Mutate in place so callers don't pay for an extra set of arrays prior to postMessage.
  for (const row of value) {
    if (!Array.isArray(row)) continue;
    for (const cell of row) {
      normalizeCellDataCompact(cell);
    }
  }
  return value;
}

function normalizeCellChanges(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  for (const change of value) {
    if (!change || typeof change !== "object") continue;
    const obj = change as any;
    if (!("value" in obj)) continue;
    obj.value = normalizeCellScalar(obj.value);
  }
  return value;
}

function normalizePivotCalculation(value: unknown): unknown {
  if (!value || typeof value !== "object") return value;
  const obj = value as any;
  if (!Array.isArray(obj.writes)) return value;
  // Mutate in place to avoid allocating a new wrapper object.
  obj.writes = normalizeCellChanges(obj.writes);
  return obj;
}

function cloneToPlainData(value: unknown): unknown {
  // wasm-bindgen APIs can return objects with prototypes that are structured-clone
  // safe but not ideal for RPC consumers. Normalize editor-tooling results into
  // plain objects/arrays at the worker boundary.
  if (value === undefined) return null;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

let cancelledRequests = new Set<number>();
const pendingRequestIds = new Set<number>();

// Cancels can arrive for request IDs that will never be sent (e.g. abort signal
// already fired before the request message is posted). Track those separately
// in a bounded structure so they don't leak forever.
const preCancelledRequestIds = new Set<number>();
const preCancelledRequestQueue: number[] = [];
const MAX_PRE_CANCELLED_REQUEST_IDS = 4096;

// Cancellation messages can arrive after the worker has already responded (e.g.
// main thread aborts while a response is in-flight). Track a bounded set of
// recently completed request IDs so late cancellations can be ignored without
// growing `cancelledRequests` forever.
const completedRequestIds = new Set<number>();
const completedRequestQueue: number[] = [];
const MAX_COMPLETED_REQUEST_IDS = 1024;

function markRequestCompleted(id: number): void {
  pendingRequestIds.delete(id);
  cancelledRequests.delete(id);
  preCancelledRequestIds.delete(id);

  // Defensive: a misbehaving caller (or test harness) can reuse request ids. Ensure this helper is
  // idempotent so we don't push duplicate ids into the completion queue, which would cause the
  // eviction logic to delete ids prematurely (and re-enable late cancellation tracking).
  if (completedRequestIds.has(id)) {
    return;
  }

  completedRequestIds.add(id);
  completedRequestQueue.push(id);
  if (completedRequestQueue.length > MAX_COMPLETED_REQUEST_IDS) {
    const oldest = completedRequestQueue.shift();
    if (oldest != null) {
      completedRequestIds.delete(oldest);
    }
  }
}

function trackCancellation(id: number): void {
  if (completedRequestIds.has(id)) {
    return;
  }

  if (pendingRequestIds.has(id)) {
    cancelledRequests.add(id);
    return;
  }

  if (preCancelledRequestIds.has(id)) {
    return;
  }

  preCancelledRequestIds.add(id);
  preCancelledRequestQueue.push(id);
  if (preCancelledRequestQueue.length > MAX_PRE_CANCELLED_REQUEST_IDS) {
    const oldest = preCancelledRequestQueue.shift();
    if (oldest != null) {
      preCancelledRequestIds.delete(oldest);
    }
  }
}

function freeWorkbook(instance: WasmWorkbookInstance | null): void {
  // wasm-bindgen classes expose an eager `free()` API. Prefer it so `newWorkbook`
  // / `loadFromJson` don't rely on GC timing to release WASM allocations.
  try {
    (instance as any)?.free?.();
  } catch {
    // Ignore failures; worst case the object is left for GC/finalization.
  }
}

let wasmModulePromise: Promise<WasmModule> | null = null;
let wasmModulePromiseUrl: string | null = null;

async function loadWasmModule(moduleUrl: string): Promise<WasmModule> {
  // Vite will try to pre-bundle dynamic imports unless explicitly told not to.
  // The `@vite-ignore` hint prevents Vite from trying to pre-bundle arbitrary URLs.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `moduleUrl` is runtime-defined (Vite dev server / asset URL).
  const mod = (await import(/* @vite-ignore */ moduleUrl)) as WasmModule;
  const init = mod.default;
  if (init) {
    if (wasmBinaryUrl) {
      // wasm-bindgen >=0.2.105 prefers an object parameter, but older versions
      // accepted `module_or_path` directly. Try the modern form first to avoid
      // a noisy console warning, then fall back for compatibility.
      try {
        await init({ module_or_path: wasmBinaryUrl });
      } catch {
        await init(wasmBinaryUrl);
      }
    } else {
      await init();
    }
  }
  return mod;
}

function getWasmModule(moduleUrl: string): Promise<WasmModule> {
  if (wasmModulePromise && wasmModulePromiseUrl === moduleUrl) {
    return wasmModulePromise;
  }

  wasmModulePromiseUrl = moduleUrl;
  wasmModulePromise = loadWasmModule(moduleUrl).catch((err) => {
    // If initialization fails (e.g. transient network error during dev), allow
    // future requests to retry by clearing the cached promise.
    wasmModulePromise = null;
    wasmModulePromiseUrl = null;
    throw err;
  });
  return wasmModulePromise;
}

async function ensureWorkbook(moduleUrl: string, generation: number): Promise<WasmWorkbookInstance> {
  const mod = await getWasmModule(moduleUrl);
  // `ensureWorkbook` can be awaiting while a new init message arrives. Do not create/replace the
  // workbook instance for a stale generation; callers will bail out and the new init will perform
  // its own hydration.
  if (generation !== transportGeneration) {
    throw new Error("stale engine worker generation");
  }
  if (!workbook) {
    workbook = new mod.WasmWorkbook();
  }
  return workbook;
}

function postMessageToMain(msg: WorkerOutboundMessage): void {
  try {
    port?.postMessage(msg);
  } catch (err) {
    // Ignore; the port can be closed/terminated while requests are in-flight, and we still need to
    // clean up request tracking state (pending/cancelled/completed ids) to avoid leaks.
    //
    // If the failure is a structured-clone/DataCloneError for a response message, try to deliver a
    // minimal error response instead so the caller doesn't hang forever waiting for a reply.
    const errName = (err as any)?.name;
    if (errName !== "DataCloneError") {
      return;
    }

    if (!msg || typeof msg !== "object" || (msg as any).type !== "response") {
      return;
    }

    const id = (msg as any).id;
    if (typeof id !== "number") {
      return;
    }

    try {
      port?.postMessage({
        type: "response",
        id,
        ok: false,
        error: "failed to serialize worker response (DataCloneError)",
      } satisfies WorkerOutboundMessage);
    } catch {
      // ignore
    }
  }
}

function isCancelled(id: number): boolean {
  return cancelledRequests.has(id);
}

function consumeCancellation(id: number): boolean {
  if (!cancelledRequests.has(id)) {
    return false;
  }
  cancelledRequests.delete(id);
  return true;
}

async function handleRequest(message: WorkerInboundMessage, generation: number): Promise<void> {
  if (generation !== transportGeneration) {
    return;
  }
  if (message.type === "cancel") {
    trackCancellation((message as RpcCancel).id);
    return;
  }

  const req = message as RpcRequest;
  const id = req.id;

  const moduleUrl = wasmModuleUrl;
  if (!moduleUrl) {
    postMessageToMain({
      type: "response",
      id,
      ok: false,
      error: "worker not initialized",
    });
    markRequestCompleted(id);
    return;
  }

  if (consumeCancellation(id)) {
    markRequestCompleted(id);
    return;
  }

  try {
    const mod = await getWasmModule(moduleUrl);
    if (generation !== transportGeneration) {
      return;
    }

    if (consumeCancellation(id)) {
      markRequestCompleted(id);
      return;
    }

    const params = req.params as any;
    let result: unknown;

    switch (req.method) {
      case "ping":
        result = "pong";
        break;
      case "supportedLocaleIds":
        {
          const supportedLocaleIds = (mod as any).supportedLocaleIds;
          if (typeof supportedLocaleIds !== "function") {
            throw new Error("supportedLocaleIds: wasm module does not export supportedLocaleIds()");
          }
          result = cloneToPlainData(supportedLocaleIds());
        }
        break;
      case "getLocaleInfo":
        {
          const getLocaleInfo = (mod as any).getLocaleInfo;
          if (typeof getLocaleInfo !== "function") {
            throw new Error("getLocaleInfo: wasm module does not export getLocaleInfo()");
          }
          result = cloneToPlainData(getLocaleInfo(params.localeId));
        }
        break;
      case "lexFormula":
        {
          const lexFormula = mod.lexFormula;
          if (typeof lexFormula !== "function") {
            throw new Error("lexFormula: wasm module does not export lexFormula()");
          }
          result = cloneToPlainData(lexFormula(params.formula, params.options));
        }
        break;
      case "canonicalizeFormula":
        {
          const canonicalizeFormula = mod.canonicalizeFormula;
          if (typeof canonicalizeFormula !== "function") {
            throw new Error("canonicalizeFormula: wasm module does not export canonicalizeFormula()");
          }
          result = canonicalizeFormula(params.formula, params.localeId, params.referenceStyle);
        }
        break;
      case "localizeFormula":
        {
          const localizeFormula = mod.localizeFormula;
          if (typeof localizeFormula !== "function") {
            throw new Error("localizeFormula: wasm module does not export localizeFormula()");
          }
          result = localizeFormula(params.formula, params.localeId, params.referenceStyle);
        }
        break;
      case "lexFormulaPartial":
        {
          const lexFormulaPartial = mod.lexFormulaPartial;
          if (typeof lexFormulaPartial !== "function") {
            throw new Error("lexFormulaPartial: wasm module does not export lexFormulaPartial()");
          }
          result = cloneToPlainData(lexFormulaPartial(params.formula, params.options));
        }
        break;
      case "parseFormulaPartial":
        {
          const parseFormulaPartial = mod.parseFormulaPartial;
          if (typeof parseFormulaPartial !== "function") {
            throw new Error("parseFormulaPartial: wasm module does not export parseFormulaPartial()");
          }
          result = cloneToPlainData(parseFormulaPartial(params.formula, params.cursor, params.options));
        }
        break;
      case "rewriteFormulasForCopyDelta":
        {
          const rewrite = mod.rewriteFormulasForCopyDelta;
          if (typeof rewrite !== "function") {
            throw new Error("rewriteFormulasForCopyDelta: wasm module does not export rewriteFormulasForCopyDelta()");
          }
          // This RPC can return large arrays (e.g. paste/fill), so avoid JSON clone overhead.
          result = rewrite(params.requests);
        }
        break;
      case "newWorkbook":
        {
          const next = new mod.WasmWorkbook();
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      case "loadFromJson":
        {
          const next = mod.WasmWorkbook.fromJson(params.json);
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      case "loadFromXlsxBytes":
        {
          const rawBytes = params.bytes as unknown;
          let bytes: Uint8Array;
          if (rawBytes instanceof Uint8Array) {
            bytes = rawBytes;
          } else if (rawBytes instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes);
          } else if (ArrayBuffer.isView(rawBytes) && rawBytes.buffer instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes.buffer, rawBytes.byteOffset, rawBytes.byteLength);
          } else {
            throw new Error(
              "loadFromXlsxBytes: expected params.bytes to be a Uint8Array/ArrayBuffer/ArrayBufferView"
            );
          }

          const fromXlsxBytes = mod.WasmWorkbook.fromXlsxBytes;
          if (typeof fromXlsxBytes !== "function") {
            throw new Error("loadFromXlsxBytes: WasmWorkbook.fromXlsxBytes is not available in this WASM build");
          }

          const next = fromXlsxBytes(bytes);
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      case "loadFromEncryptedXlsxBytes":
        {
          const rawBytes = params.bytes as unknown;
          let bytes: Uint8Array;
          if (rawBytes instanceof Uint8Array) {
            bytes = rawBytes;
          } else if (rawBytes instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes);
          } else if (ArrayBuffer.isView(rawBytes) && rawBytes.buffer instanceof ArrayBuffer) {
            bytes = new Uint8Array(rawBytes.buffer, rawBytes.byteOffset, rawBytes.byteLength);
          } else {
            throw new Error(
              "loadFromEncryptedXlsxBytes: expected params.bytes to be a Uint8Array/ArrayBuffer/ArrayBufferView"
            );
          }

          const password = (params as any).password as unknown;
          if (typeof password !== "string") {
            throw new Error("loadFromEncryptedXlsxBytes: expected params.password to be a string");
          }

          const fromEncryptedXlsxBytes = (mod.WasmWorkbook as any).fromEncryptedXlsxBytes;
          if (typeof fromEncryptedXlsxBytes !== "function") {
            throw new Error(
              "loadFromEncryptedXlsxBytes: WasmWorkbook.fromEncryptedXlsxBytes is not available in this WASM build"
            );
          }

          const next = fromEncryptedXlsxBytes(bytes, password);
          freeWorkbook(workbook);
          workbook = next;
        }
        result = null;
        break;
      default:
        {
          const wb = await ensureWorkbook(moduleUrl, generation);
          if (generation !== transportGeneration) {
            return;
          }

          if (consumeCancellation(id)) {
            markRequestCompleted(id);
            return;
          }

          switch (req.method) {
            case "getWorkbookInfo":
              if (typeof (wb as any).getWorkbookInfo === "function") {
                result = cloneToPlainData((wb as any).getWorkbookInfo());
                break;
              }

              // Backward compatibility: older WASM builds don't export `getWorkbookInfo()`. Fall back
              // to parsing `toJson()` inside the worker so large workbook JSON strings don't cross
              // the postMessage boundary.
              {
                const json = wb.toJson();
                let parsed: EngineWorkbookJson | null = null;
                try {
                  parsed = JSON.parse(json) as EngineWorkbookJson;
                } catch {
                  parsed = null;
                }

                const sheetsRecord =
                  parsed?.sheets && typeof parsed.sheets === "object" ? (parsed.sheets as EngineWorkbookJson["sheets"]) : null;

                const sheetIds = (() => {
                  if (!sheetsRecord) return [];

                  const keys = Object.keys(sheetsRecord);
                  const explicitOrderRaw = Array.isArray(parsed?.sheetOrder) ? (parsed?.sheetOrder as unknown[]) : [];
                  if (explicitOrderRaw.length === 0) return keys;

                  const keySet = new Set(keys);
                  const ordered: string[] = [];
                  const seen = new Set<string>();

                  for (const candidate of explicitOrderRaw) {
                    if (typeof candidate !== "string") continue;
                    if (!candidate) continue;
                    if (!keySet.has(candidate)) continue;
                    if (seen.has(candidate)) continue;
                    ordered.push(candidate);
                    seen.add(candidate);
                  }

                  // Preserve any sheets that aren't listed in `sheetOrder` (for backwards compatibility).
                  for (const key of keys) {
                    if (seen.has(key)) continue;
                    ordered.push(key);
                  }

                  return ordered;
                })();

                const usedRanges = new Map<string, UsedRangeState>();
                for (const sheetId of sheetIds) {
                  const cells = sheetsRecord?.[sheetId]?.cells;
                  if (!cells || typeof cells !== "object") continue;
                  for (const address of Object.keys(cells)) {
                    try {
                      const { row0, col0 } = fromA1(address);
                      updateUsedRange(usedRanges, sheetId, row0, col0);
                    } catch {
                      // Ignore invalid A1 keys; used range tracking is best-effort.
                    }
                  }
                }

                const sheets: WorkbookInfoDto["sheets"] =
                  sheetIds.length > 0
                    ? sheetIds.map((id) => {
                        const sheetMeta = sheetsRecord?.[id];
                        const sheet: any = { id, name: id };
                        if (sheetMeta && typeof sheetMeta === "object") {
                          if (typeof sheetMeta.rowCount === "number") {
                            sheet.rowCount = sheetMeta.rowCount;
                          }
                          if (typeof sheetMeta.colCount === "number") {
                            sheet.colCount = sheetMeta.colCount;
                          }
                          if (
                            typeof (sheetMeta as any).visibility === "string" &&
                            ((sheetMeta as any).visibility === "visible" ||
                              (sheetMeta as any).visibility === "hidden" ||
                              (sheetMeta as any).visibility === "veryHidden")
                          ) {
                            sheet.visibility = (sheetMeta as any).visibility;
                          }
                          const tabColor = (sheetMeta as any).tabColor as unknown;
                          if (tabColor && typeof tabColor === "object" && !Array.isArray(tabColor)) {
                            sheet.tabColor = tabColor;
                          }
                        }

                        const used = usedRanges.get(id);
                        if (used) {
                          sheet.usedRange = used;
                        }
                        return sheet;
                      })
                    : [{ id: DEFAULT_SHEET_NAME, name: DEFAULT_SHEET_NAME }];

                result = { path: null, origin_path: null, sheets } satisfies WorkbookInfoDto;
              }
              break;
            case "toJson":
              result = wb.toJson();
              break;
            case "getCell":
              {
                const sheet = normalizeSheetName(params.sheet);
                result = normalizeCellData(wb.getCell(params.address, sheet));
              }
              break;
            case "getCellPhonetic":
              if (typeof (wb as any).getCellPhonetic !== "function") {
                throw new Error(
                  "getCellPhonetic: WasmWorkbook.getCellPhonetic is not available in this WASM build"
                );
              }
              {
                const sheet = normalizeSheetName(params.sheet);
                result = (wb as any).getCellPhonetic(params.address, sheet) ?? null;
              }
              break;
            case "getCellRich":
              if (typeof (wb as any).getCellRich !== "function") {
                throw new Error("getCellRich: WasmWorkbook.getCellRich is not available in this WASM build");
              }
              result = cloneToPlainData((wb as any).getCellRich(params.address, normalizeSheetName(params.sheet)));
              break;
            case "getRange":
              {
                const sheet = normalizeSheetName(params.sheet);
                result = normalizeRangeData(wb.getRange(params.range, sheet));
              }
              break;
            case "getRangeCompact":
              if (typeof (wb as any).getRangeCompact !== "function") {
                throw new Error(
                  "getRangeCompact: WasmWorkbook.getRangeCompact is not available in this WASM build"
                );
              }
              {
                result = normalizeRangeDataCompact((wb as any).getRangeCompact(params.range, normalizeSheetName(params.sheet)));
              }
              break;
            case "setSheetDimensions":
              if (typeof (wb as any).setSheetDimensions !== "function") {
                throw new Error("setSheetDimensions: not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                (wb as any).setSheetDimensions(sheet, params.rows, params.cols);
              }
              result = null;
              break;
            case "getSheetDimensions":
              if (typeof (wb as any).getSheetDimensions !== "function") {
                throw new Error("getSheetDimensions: not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                result = (wb as any).getSheetDimensions(sheet);
              }
              break;
            case "renameSheet":
              if (typeof (wb as any).renameSheet !== "function") {
                throw new Error("renameSheet: WasmWorkbook.renameSheet is not available in this WASM build");
              }
              {
                const oldName = normalizeSheetName(params.oldName);
                const newName = normalizeSheetName(params.newName);
                // Defensive: sheet names are expected to be non-empty strings. Avoid forwarding
                // whitespace-only names into the engine (older WASM builds may not validate inputs
                // consistently).
                if (!oldName || !newName) {
                  result = false;
                } else {
                  result = Boolean((wb as any).renameSheet(oldName, newName));
                }
              }
              break;
            case "setSheetOrigin":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const origin = (params as any).origin ?? null;
                if (typeof (wb as any).setSheetOrigin === "function") {
                  (wb as any).setSheetOrigin(sheet, origin);
                } else if (typeof (wb as any).setInfoOriginForSheet === "function") {
                  // Backward compatibility: older WASM builds exposed `setInfoOriginForSheet` as
                  // the per-sheet origin setter.
                  (wb as any).setInfoOriginForSheet(sheet, origin);
                } else {
                  throw new Error("setSheetOrigin: not available in this WASM build");
                }
              }
              result = null;
              break;
            case "setColWidthChars":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const widthChars = params.widthChars ?? null;
                if (typeof (wb as any).setColWidthChars === "function") {
                  (wb as any).setColWidthChars(sheet, params.col, widthChars);
                } else if (typeof (wb as any).setColWidth === "function") {
                  // Backward compatibility: older WASM builds used `setColWidth` as the column
                  // width override setter (same Excel "character" unit semantics).
                  (wb as any).setColWidth(sheet, params.col, widthChars);
                } else {
                  throw new Error("setColWidthChars: not available in this WASM build");
                }
              }
              result = null;
              break;
            case "setSheetDisplayName":
              if (typeof (wb as any).setSheetDisplayName !== "function") {
                throw new Error("setSheetDisplayName: not available in this WASM build");
              }
              (wb as any).setSheetDisplayName(sheetNameOrDefault(params.sheetId), params.name);
              result = null;
              break;
            case "setCells":
              for (const update of params.updates as Array<any>) {
                if (typeof update?.sheet === "string") {
                  const trimmed = update.sheet.trim();
                  if (trimmed === "") {
                    delete update.sheet;
                  } else {
                    update.sheet = trimmed;
                  }
                }
              }
              if (typeof (wb as any).setCells === "function") {
                (wb as any).setCells(params.updates);
              } else {
                for (const update of params.updates as Array<any>) {
                  wb.setCell(update.address, update.value, update.sheet);
                }
              }
              result = null;
              break;
            case "setCellRich":
              if (typeof (wb as any).setCellRich !== "function") {
                throw new Error("setCellRich: WasmWorkbook.setCellRich is not available in this WASM build");
              }
              {
                (wb as any).setCellRich(params.address, params.value, normalizeSheetName(params.sheet));
              }
              result = null;
              break;
            case "setCellPhonetic":
              if (typeof (wb as any).setCellPhonetic !== "function") {
                throw new Error(
                  "setCellPhonetic: WasmWorkbook.setCellPhonetic is not available in this WASM build"
                );
              }
              {
                (wb as any).setCellPhonetic(params.address, params.phonetic ?? null, normalizeSheetName(params.sheet));
              }
              result = null;
              break;
            case "setRange":
              {
                wb.setRange(params.range, params.values, normalizeSheetName(params.sheet));
              }
              result = null;
              break;
            case "setLocale":
              if (typeof (wb as any).setLocale === "function") {
                result = (wb as any).setLocale(params.localeId);
              } else {
                result = false;
              }
              break;
            case "getCalcSettings":
              if (typeof (wb as any).getCalcSettings !== "function") {
                throw new Error("getCalcSettings: WasmWorkbook.getCalcSettings is not available in this WASM build");
              }
              result = cloneToPlainData((wb as any).getCalcSettings());
              break;
            case "setCalcSettings":
              if (typeof (wb as any).setCalcSettings !== "function") {
                throw new Error("setCalcSettings: WasmWorkbook.setCalcSettings is not available in this WASM build");
              }
              (wb as any).setCalcSettings(params.settings);
              result = null;
              break;
            case "setEngineInfo":
              if (typeof (wb as any).setEngineInfo === "function") {
                (wb as any).setEngineInfo(params.info);
              } else {
                // Backward compatibility: older WASM builds exposed per-field `setInfo*` setters
                // instead of the batched `setEngineInfo` API.
                const info = params.info as any;
                if (!info || typeof info !== "object") {
                  throw new Error("setEngineInfo: WasmWorkbook.setEngineInfo is not available in this WASM build");
                }

                const hasKey = (key: string): boolean =>
                  Object.prototype.hasOwnProperty.call(info, key) || (typeof Reflect?.has === "function" && Reflect.has(info, key));

                let handled = false;
                const setLegacy = (key: string, methodName: string) => {
                  if (!hasKey(key)) return;
                  const method = (wb as any)[methodName] as unknown;
                  if (typeof method !== "function") return;
                  handled = true;
                  method.call(wb, info[key]);
                };

                setLegacy("system", "setInfoSystem");
                // Some older builds may have exposed `setInfoDirectory`; prefer it when present.
                setLegacy("directory", "setInfoDirectory");
                setLegacy("osversion", "setInfoOSVersion");
                setLegacy("release", "setInfoRelease");
                setLegacy("version", "setInfoVersion");
                setLegacy("memavail", "setInfoMemAvail");
                setLegacy("totmem", "setInfoTotMem");

                if (!handled && Object.keys(info).length > 0) {
                  throw new Error("setEngineInfo: WasmWorkbook.setEngineInfo is not available in this WASM build");
                }
              }
              result = null;
              break;
            case "setInfoOrigin":
              if (typeof (wb as any).setInfoOrigin !== "function") {
                throw new Error("setInfoOrigin: WasmWorkbook.setInfoOrigin is not available in this WASM build");
              }
              (wb as any).setInfoOrigin(params.origin ?? null);
              result = null;
              break;
            case "setInfoOriginForSheet":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const origin = (params as any).origin ?? null;
                if (typeof (wb as any).setInfoOriginForSheet === "function") {
                  (wb as any).setInfoOriginForSheet(sheet, origin);
                } else if (typeof (wb as any).setSheetOrigin === "function") {
                  // Forward compatibility: some WASM builds may expose only the modern
                  // `setSheetOrigin` API. Treat the legacy RPC as an alias.
                  (wb as any).setSheetOrigin(sheet, origin);
                } else {
                  throw new Error("setInfoOriginForSheet: not available in this WASM build");
                }
              }
              result = null;
              break;
            case "setWorkbookFileMetadata":
              // Optional API: older WASM bundles may not expose workbook-level file metadata.
              // Treat missing support as a no-op so callers (and DocumentController hydration)
              // don't fail hard when running against a minimal build.
              if (typeof (wb as any).setWorkbookFileMetadata === "function") {
                (wb as any).setWorkbookFileMetadata(params.directory ?? null, params.filename ?? null);
              }
              result = null;
              break;
            case "setCellStyleId":
              if (typeof (wb as any).setCellStyleId !== "function") {
                throw new Error("setCellStyleId: WasmWorkbook.setCellStyleId is not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                // `crates/formula-wasm` uses a sheet-first signature (`setCellStyleId(sheet, address, styleId)`),
                // but some older builds used a sheet-last form (`setCellStyleId(address, styleId, sheet)`).
                //
                // Prefer sheet-first, but fall back when the call fails with a parse error that clearly
                // indicates we sent the sheet name as the cell address (legacy signature).
                try {
                  (wb as any).setCellStyleId(sheet, params.address, params.styleId);
                } catch (err) {
                  const message =
                    typeof err === "string"
                      ? err
                      : err && typeof err === "object" && typeof (err as any).message === "string"
                        ? String((err as any).message)
                        : null;
                  if (message && message.includes(`invalid cell address: ${sheet}`)) {
                    (wb as any).setCellStyleId(params.address, params.styleId, sheet);
                    // Continue; outer switch case will still set `result=null`.
                    // (Avoid `break` here which would skip setting the RPC result.)
                  } else {
                    throw err;
                  }
                }
              }
              result = null;
              break;
            case "setRowStyleId":
              if (typeof (wb as any).setRowStyleId !== "function") {
                throw new Error("setRowStyleId: WasmWorkbook.setRowStyleId is not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                // Backward compatibility: older WASM builds used a numeric `styleId` where `0`
                // clears the override. Newer builds accept `Option<u32>` and treat both `0` and
                // `null`/`undefined` as clear. Always forwarding a number keeps both working.
                const styleId = params.styleId == null ? 0 : params.styleId;
                (wb as any).setRowStyleId(sheet, params.row, styleId);
              }
              result = null;
              break;
            case "setColStyleId":
              if (typeof (wb as any).setColStyleId !== "function") {
                throw new Error("setColStyleId: WasmWorkbook.setColStyleId is not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const styleId = params.styleId == null ? 0 : params.styleId;
                (wb as any).setColStyleId(sheet, params.col, styleId);
              }
              result = null;
              break;
            case "setFormatRunsByCol":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const runs = params.runs ?? [];
                if (typeof (wb as any).setFormatRunsByCol === "function") {
                  (wb as any).setFormatRunsByCol(sheet, params.col, runs);
                } else if (typeof (wb as any).setColFormatRuns === "function") {
                  // Backward compatibility: older WASM builds exposed `setColFormatRuns` as the
                  // formatting-run setter (same payload shape).
                  (wb as any).setColFormatRuns(sheet, params.col, runs);
                } else {
                  throw new Error(
                    "setFormatRunsByCol: WasmWorkbook.setFormatRunsByCol is not available in this WASM build"
                  );
                }
              }
              result = null;
              break;
            case "setSheetDefaultStyleId":
              if (typeof (wb as any).setSheetDefaultStyleId !== "function") {
                throw new Error(
                  "setSheetDefaultStyleId: WasmWorkbook.setSheetDefaultStyleId is not available in this WASM build"
                );
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const styleId = params.styleId == null ? 0 : params.styleId;
                (wb as any).setSheetDefaultStyleId(sheet, styleId);
              }
              result = null;
              break;
            case "setColWidth":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const width = params.width ?? null;
                if (typeof (wb as any).setColWidth === "function") {
                  (wb as any).setColWidth(sheet, params.col, width);
                } else if (typeof (wb as any).setColWidthChars === "function") {
                  // Backward compatibility: some WASM builds only expose `setColWidthChars`, which
                  // uses the same Excel character-unit semantics as `setColWidth`.
                  (wb as any).setColWidthChars(sheet, params.col, width);
                } else {
                  throw new Error("setColWidth: WasmWorkbook.setColWidth is not available in this WASM build");
                }
              }
              result = null;
              break;
            case "setColHidden":
              if (typeof (wb as any).setColHidden !== "function") {
                throw new Error("setColHidden: WasmWorkbook.setColHidden is not available in this WASM build");
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                (wb as any).setColHidden(sheet, params.col, Boolean(params.hidden));
              }
              result = null;
              break;
            case "setSheetDefaultColWidth":
              if (typeof (wb as any).setSheetDefaultColWidth !== "function") {
                throw new Error(
                  "setSheetDefaultColWidth: WasmWorkbook.setSheetDefaultColWidth is not available in this WASM build",
                );
              }
              {
                const sheet = sheetNameOrDefault(params.sheet);
                (wb as any).setSheetDefaultColWidth(sheet, params.widthChars ?? null);
              }
              result = null;
              break;
            case "setColFormatRuns":
              {
                const sheet = sheetNameOrDefault(params.sheet);
                const runs = params.runs ?? [];
                if (typeof (wb as any).setColFormatRuns === "function") {
                  (wb as any).setColFormatRuns(sheet, params.col, runs);
                } else if (typeof (wb as any).setFormatRunsByCol === "function") {
                  // Backward compatibility: some WASM builds used `setFormatRunsByCol` as the
                  // formatting-run setter.
                  (wb as any).setFormatRunsByCol(sheet, params.col, runs);
                } else {
                  throw new Error(
                    "setColFormatRuns: WasmWorkbook.setColFormatRuns is not available in this WASM build"
                  );
                }
              }
              result = null;
              break;
            case "internStyle":
              if (typeof (wb as any).internStyle !== "function") {
                throw new Error("internStyle: WasmWorkbook.internStyle is not available in this WASM build");
              }
              result = (wb as any).internStyle(params.style);
              break;
            case "recalculate":
              {
                result = normalizeCellChanges(wb.recalculate(normalizeSheetName(params.sheet)));
              }
              break;
            case "applyOperation":
              if (typeof (wb as any).applyOperation === "function") {
                const op = params.op as any;
                if (op && typeof op === "object") {
                  const sheet = (op as any).sheet as unknown;
                  if (typeof sheet === "string") {
                    const trimmed = sheet.trim();
                    // Defensive: avoid creating an empty-named sheet via `ensure_sheet("")` when
                    // callers accidentally pass a blank sheet id.
                    (op as any).sheet = trimmed === "" ? DEFAULT_SHEET_NAME : trimmed;
                  }
                }

                result = cloneToPlainData((wb as any).applyOperation(op));
              } else {
                throw new Error("applyOperation: WasmWorkbook.applyOperation is not available in this WASM build");
              }
              break;
            case "goalSeek":
              if (typeof (wb as any).goalSeek !== "function") {
                throw new Error("goalSeek: WasmWorkbook.goalSeek is not available in this WASM build");
              }
              {
                const raw = (wb as any).goalSeek(params);
                // Goal seek returns `{ result, changes }` where `changes` mirrors the `recalculate()`
                // payload. Normalize `value` scalars to ensure structured-clone safe null/primitive
                // values (and to avoid JSON clone dropping `undefined` keys).
                if (raw && typeof raw === "object") {
                  const obj = raw as any;
                  if ("result" in obj && "changes" in obj) {
                    const normalizedResult =
                      obj.result && typeof obj.result === "object" ? { ...(obj.result as any) } : obj.result;
                    const rawChanges = obj.changes;
                    result = {
                      result: normalizedResult,
                      changes: Array.isArray(rawChanges) ? normalizeCellChanges(rawChanges) : [],
                    };
                  } else if ("success" in obj && "solution" in obj) {
                    // Backward compatibility: older WASM builds returned a flat payload
                    // `{ success, status?, solution, iterations, finalError, finalOutput? }`.
                    const legacy = cloneToPlainData(obj) as any;
                    const status =
                      typeof legacy.status === "string" && legacy.status.trim()
                        ? legacy.status
                        : legacy.success
                          ? "Converged"
                          : "NumericalFailure";
                    const targetValue = typeof params?.targetValue === "number" && Number.isFinite(params.targetValue) ? params.targetValue : null;
                    const finalError = typeof legacy.finalError === "number" ? legacy.finalError : Number(legacy.finalError);
                    const legacyFinalOutput =
                      typeof legacy.finalOutput === "number"
                        ? legacy.finalOutput
                        : legacy.finalOutput === undefined
                          ? null
                          : Number(legacy.finalOutput);
                    const finalOutput =
                      legacyFinalOutput != null && Number.isFinite(legacyFinalOutput)
                        ? legacyFinalOutput
                        : targetValue != null && Number.isFinite(targetValue) && Number.isFinite(finalError)
                          ? targetValue + finalError
                          : NaN;
                    result = {
                      result: {
                        status,
                        solution: typeof legacy.solution === "number" ? legacy.solution : Number(legacy.solution),
                        iterations: typeof legacy.iterations === "number" ? legacy.iterations : Number(legacy.iterations),
                        finalOutput,
                        finalError,
                      },
                      changes: [],
                    };
                  } else {
                    result = cloneToPlainData(obj);
                  }
                } else {
                  result = cloneToPlainData(raw);
                }
              }
              break;
            case "getPivotSchema":
              if (typeof (wb as any).getPivotSchema !== "function") {
                throw new Error("getPivotSchema: WasmWorkbook.getPivotSchema is not available in this WASM build");
              }
              result = cloneToPlainData(
                (wb as any).getPivotSchema(sheetNameOrDefault(params.sheet), params.sourceRangeA1, params.sampleSize)
              );
              break;
            case "getPivotFieldItems":
              if (typeof (wb as any).getPivotFieldItems !== "function") {
                throw new Error(
                  "getPivotFieldItems: WasmWorkbook.getPivotFieldItems is not available in this WASM build"
                );
              }
              result = cloneToPlainData(
                (wb as any).getPivotFieldItems(sheetNameOrDefault(params.sheet), params.sourceRangeA1, params.field)
              );
              break;
            case "getPivotFieldItemsPaged":
              if (typeof (wb as any).getPivotFieldItemsPaged !== "function") {
                throw new Error(
                  "getPivotFieldItemsPaged: WasmWorkbook.getPivotFieldItemsPaged is not available in this WASM build"
                );
              }
              result = cloneToPlainData(
                (wb as any).getPivotFieldItemsPaged(
                  sheetNameOrDefault(params.sheet),
                  params.sourceRangeA1,
                  params.field,
                  params.offset,
                  params.limit
                )
              );
              break;
            case "calculatePivot":
              if (typeof (wb as any).calculatePivot !== "function") {
                throw new Error("calculatePivot: WasmWorkbook.calculatePivot is not available in this WASM build");
              }
              // Normalize before cloning so we preserve wasm-bindgen `Option<T>` -> `undefined`
              // mappings (JSON cloning would drop `undefined` object keys entirely).
              {
                const raw = (wb as any).calculatePivot(
                  sheetNameOrDefault(params.sheet),
                  params.sourceRangeA1,
                  params.destinationTopLeftA1,
                  params.config
                );
                result = cloneToPlainData(normalizePivotCalculation(raw));
              }
              break;
            default:
              throw new Error(`unknown method: ${req.method}`);
          }
        }
    }

    if (isCancelled(id)) {
      // Cancellation might arrive after the request starts; we still perform the work
      // but suppress the response.
      cancelledRequests.delete(id);
      markRequestCompleted(id);
      return;
    }

    if (generation !== transportGeneration) {
      return;
    }
    postMessageToMain({ type: "response", id, ok: true, result });
    markRequestCompleted(id);
  } catch (err) {
    if (isCancelled(id)) {
      cancelledRequests.delete(id);
      markRequestCompleted(id);
      return;
    }

    if (generation !== transportGeneration) {
      return;
    }
    postMessageToMain({
      type: "response",
      id,
      ok: false,
      error: err instanceof Error ? err.message : String(err),
    });
    markRequestCompleted(id);
  }
}

function isWorkerInboundMessage(data: unknown): data is WorkerInboundMessage {
  if (!data || typeof data !== "object" || !("type" in data)) {
    return false;
  }

  const type = (data as any).type;
  if (type !== "request" && type !== "cancel") {
    return false;
  }

  if (typeof (data as any).id !== "number") {
    return false;
  }

  if (type === "request") {
    return typeof (data as any).method === "string";
  }

  return true;
}

let requestQueue: Promise<void> = Promise.resolve();

function resetTransportForInit(): void {
  // This worker is intended to be initialized exactly once, but tests (and some hot-reload/dev
  // environments) can dispatch multiple init messages into the same module instance. Tear down any
  // prior MessagePort/listeners so old ports don't leak and old cancellation state can't poison the
  // new connection.
  transportGeneration += 1;
  const existing = port;
  const existingListener = portListener;
  port = null;
  portListener = null;

  if (existing && existingListener) {
    try {
      existing.removeEventListener("message", existingListener);
    } catch {
      // ignore
    }
  }
  try {
    existing?.close();
  } catch {
    // ignore
  }

  // Drop the workbook instance eagerly so its WASM allocations can be reclaimed.
  freeWorkbook(workbook);
  workbook = null;

  // Clear cancellation + request tracking state so cancels from a previous connection don't apply
  // to a newly initialized port (request ids are reused across connections).
  cancelledRequests = new Set<number>();
  pendingRequestIds.clear();
  preCancelledRequestIds.clear();
  preCancelledRequestQueue.length = 0;
  completedRequestIds.clear();
  completedRequestQueue.length = 0;

  // Reset the serialized request queue so new connections don't block forever on an abandoned/hung
  // request chain.
  requestQueue = Promise.resolve();
}

self.addEventListener("message", (event: MessageEvent<unknown>) => {
  const data = event.data;

  const msg = data as InitMessage;
  if (!msg || typeof msg !== "object" || (msg as any).type !== "init") {
    return;
  }

  resetTransportForInit();

  port = msg.port;
  wasmModuleUrl = msg.wasmModuleUrl;
  wasmBinaryUrl = msg.wasmBinaryUrl ?? null;

  const generation = transportGeneration;
  const listener = (inner: MessageEvent<unknown>) => {
    // If the worker is re-initialized, ignore any queued messages from the previous port.
    if (generation !== transportGeneration) {
      return;
    }
    const inbound = inner.data;
    if (!isWorkerInboundMessage(inbound)) {
      return;
    }

    if (inbound.type === "cancel") {
      // Handle cancels immediately so in-flight requests can observe cancellation.
      trackCancellation(inbound.id);
      return;
    }

    pendingRequestIds.add(inbound.id);
    if (preCancelledRequestIds.has(inbound.id)) {
      preCancelledRequestIds.delete(inbound.id);
      cancelledRequests.add(inbound.id);
    }

    // Serialize request processing to avoid interleaving workbook mutations.
    requestQueue = requestQueue
      .then(() => handleRequest(inbound, generation))
      .catch(() => {
        // `handleRequest` should catch and respond to all errors, but if something
        // escapes we don't want to wedge the queue (or leak pending request ids).
        if (generation !== transportGeneration) {
          return;
        }
        try {
          markRequestCompleted((inbound as any).id);
        } catch {
          // ignore
        }
      });
  };
  portListener = listener;
  port.addEventListener("message", listener);
  port.start?.();

  postMessageToMain({ type: "ready" });
});
