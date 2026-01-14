import { PermissionError, PermissionManager } from "./permission-manager.mjs";
import { validateExtensionManifest } from "./manifest.mjs";

const API_PERMISSIONS = {
  "workbook.getActiveWorkbook": [],
  "workbook.openWorkbook": ["workbook.manage"],
  "workbook.createWorkbook": ["workbook.manage"],
  "workbook.save": ["workbook.manage"],
  "workbook.saveAs": ["workbook.manage"],
  "workbook.close": ["workbook.manage"],
  "sheets.getActiveSheet": [],

  "cells.getSelection": ["cells.read"],
  "cells.getRange": ["cells.read"],
  "cells.getCell": ["cells.read"],
  "cells.setCell": ["cells.write"],
  "cells.setRange": ["cells.write"],

  "sheets.getSheet": ["sheets.manage"],
  "sheets.activateSheet": ["sheets.manage"],
  "sheets.createSheet": ["sheets.manage"],
  "sheets.renameSheet": ["sheets.manage"],
  "sheets.deleteSheet": ["sheets.manage"],

  "commands.registerCommand": ["ui.commands"],
  "commands.unregisterCommand": ["ui.commands"],
  "commands.executeCommand": ["ui.commands"],

  "ui.createPanel": ["ui.panels"],
  "ui.setPanelHtml": ["ui.panels"],
  "ui.postMessageToPanel": ["ui.panels"],
  "ui.disposePanel": ["ui.panels"],
  "ui.registerContextMenu": ["ui.menus"],
  "ui.unregisterContextMenu": ["ui.menus"],
  "ui.showInputBox": [],
  "ui.showQuickPick": [],

  "functions.register": [],
  "functions.unregister": [],

  "dataConnectors.register": [],
  "dataConnectors.unregister": [],

  "network.fetch": ["network"],
  "network.openWebSocket": ["network"],

  "clipboard.readText": ["clipboard"],
  "clipboard.writeText": ["clipboard"],

  "storage.get": ["storage"],
  "storage.set": ["storage"],
  "storage.delete": ["storage"],

  "config.get": ["storage"],
  "config.update": ["storage"]
};

const MAX_TAINTED_RANGES_PER_EXTENSION = 50;

// Extension APIs represent ranges as full 2D JS arrays. With Excel-scale sheets, unbounded
// ranges can allocate millions of entries and OOM the host/worker. Keep reads/writes bounded
// to match the desktop extension API guardrails.
const DEFAULT_EXTENSION_RANGE_CELL_LIMIT = 200000;

function getRangeSize(range) {
  if (!range || typeof range !== "object") return null;
  const startRow = Number(range.startRow);
  const startCol = Number(range.startCol);
  const endRow = Number(range.endRow);
  const endCol = Number(range.endCol);
  if (![startRow, startCol, endRow, endCol].every((v) => Number.isFinite(v))) return null;
  const rows = Math.max(0, Math.max(startRow, endRow) - Math.min(startRow, endRow) + 1);
  const cols = Math.max(0, Math.max(startCol, endCol) - Math.min(startCol, endCol) + 1);
  return { rows, cols, cellCount: rows * cols };
}

function assertExtensionRangeWithinLimits(range, { label, maxCells } = {}) {
  const size = getRangeSize(range);
  if (!size) return null;
  const limit = Number.isFinite(maxCells) ? maxCells : DEFAULT_EXTENSION_RANGE_CELL_LIMIT;
  if (size.cellCount > limit) {
    const name = String(label ?? "Range");
    throw new Error(`${name} is too large (${size.rows}x${size.cols}=${size.cellCount} cells). Limit is ${limit} cells.`);
  }
  return size;
}

function normalizeNonNegativeInt(value, { label }) {
  const raw = value;
  if (typeof raw === "number") {
    if (Number.isInteger(raw) && raw >= 0) return raw;
    throw new Error(`Invalid ${label}: ${raw}`);
  }
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (!/^\d+$/.test(trimmed)) throw new Error(`Invalid ${label}: ${raw}`);
    const n = Number(trimmed);
    if (Number.isInteger(n) && n >= 0) return n;
    throw new Error(`Invalid ${label}: ${raw}`);
  }
  throw new Error(`Invalid ${label}: ${String(raw)}`);
}

function normalizeTaintedRange(range) {
  if (!range || typeof range !== "object") return null;

  const sheetId = typeof range.sheetId === "string" ? range.sheetId.trim() : "";
  if (!sheetId) return null;

  const startRow = Number(range.startRow);
  const startCol = Number(range.startCol);
  const endRow = Number(range.endRow);
  const endCol = Number(range.endCol);
  if (![startRow, startCol, endRow, endCol].every((v) => Number.isFinite(v))) return null;

  const sr = Math.max(0, Math.min(startRow, endRow));
  const er = Math.max(0, Math.max(startRow, endRow));
  const sc = Math.max(0, Math.min(startCol, endCol));
  const ec = Math.max(0, Math.max(startCol, endCol));

  return {
    sheetId,
    startRow: Math.trunc(sr),
    startCol: Math.trunc(sc),
    endRow: Math.trunc(er),
    endCol: Math.trunc(ec)
  };
}

function rangeContains(a, b) {
  return (
    a.sheetId === b.sheetId &&
    a.startRow <= b.startRow &&
    a.endRow >= b.endRow &&
    a.startCol <= b.startCol &&
    a.endCol >= b.endCol
  );
}

function rangesOverlapOrTouch1D(aStart, aEnd, bStart, bEnd) {
  return aStart <= bEnd + 1 && bStart <= aEnd + 1;
}

function canMergeTaintedRanges(a, b) {
  if (a.sheetId !== b.sheetId) return false;

  if (rangeContains(a, b) || rangeContains(b, a)) return true;

  // Merge rectangles only when the union is still a perfect rectangle (no L-shapes),
  // to avoid over-tainting cells that weren't read.
  const sameCols = a.startCol === b.startCol && a.endCol === b.endCol;
  if (sameCols && rangesOverlapOrTouch1D(a.startRow, a.endRow, b.startRow, b.endRow)) return true;

  const sameRows = a.startRow === b.startRow && a.endRow === b.endRow;
  if (sameRows && rangesOverlapOrTouch1D(a.startCol, a.endCol, b.startCol, b.endCol)) return true;

  return false;
}

function unionTaintedRanges(a, b) {
  return {
    sheetId: a.sheetId,
    startRow: Math.min(a.startRow, b.startRow),
    startCol: Math.min(a.startCol, b.startCol),
    endRow: Math.max(a.endRow, b.endRow),
    endCol: Math.max(a.endCol, b.endCol)
  };
}

function compressTaintedRangesToLimit(ranges, limit) {
  const max = Number.isFinite(limit) ? Math.max(0, limit) : MAX_TAINTED_RANGES_PER_EXTENSION;
  if (ranges.length <= max) return ranges;
  // Keep a bounded number of ranges to avoid unbounded memory usage. Prefer keeping
  // the most recently accessed ranges rather than merging into coarse bounding boxes
  // (which can over-taint cells that were never read and introduce false positives).
  return ranges.slice(-max);
}

function addTaintedRangeToList(ranges, nextRange) {
  const normalized = normalizeTaintedRange(nextRange);
  if (!normalized) return ranges;

  const existing = Array.isArray(ranges) ? ranges : [];

  /** @type {Array<any>} */
  const out = [];
  let merged = normalized;
  for (const item of existing) {
    const normalizedItem = normalizeTaintedRange(item);
    if (!normalizedItem) continue;

    if (canMergeTaintedRanges(normalizedItem, merged)) {
      merged = unionTaintedRanges(normalizedItem, merged);
    } else {
      out.push(normalizedItem);
    }
  }

  out.push(merged);
  return compressTaintedRangesToLimit(out, MAX_TAINTED_RANGES_PER_EXTENSION);
}

function normalizeStringArray(value) {
  if (!Array.isArray(value)) return [];
  return value
    .filter((entry) => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

function safeParseUrl(url) {
  const raw = String(url ?? "");
  try {
    if (typeof globalThis?.location?.href === "string") {
      return new URL(raw, globalThis.location.href);
    }
  } catch {
    // ignore
  }

  try {
    return new URL(raw, "http://localhost/");
  } catch {
    return null;
  }
}

function safeGetProp(obj, prop) {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

function isUrlAllowedByHosts(urlString, hosts) {
  const parsed = safeParseUrl(urlString);
  if (!parsed) return false;
  const origin = parsed.origin;
  const host = parsed.hostname;

  for (const rawEntry of normalizeStringArray(hosts)) {
    const entry = rawEntry.trim();
    if (!entry) continue;

    if (entry.includes("://")) {
      if (origin === entry) return true;
      continue;
    }

    if (entry.startsWith("*.")) {
      const suffix = entry.slice(2);
      if (host === suffix) return true;
      if (host.endsWith(`.${suffix}`)) return true;
      continue;
    }

    if (host === entry) return true;
  }

  return false;
}

function serializeError(error) {
  const payload = { message: "Unknown error" };
  try {
    if (error && typeof error === "object" && "message" in error) {
      payload.message = String(error.message);
    } else {
      payload.message = String(error);
    }
  } catch {
    payload.message = "Unknown error";
  }

  try {
    if (error && typeof error === "object" && "stack" in error && error.stack != null) {
      payload.stack = String(error.stack);
    }
  } catch {
    // ignore stack serialization failures
  }

  try {
    if (error && typeof error === "object") {
      if (typeof error.name === "string" && error.name.trim().length > 0) {
        payload.name = String(error.name);
      }
      if (Object.prototype.hasOwnProperty.call(error, "code")) {
        const code = error.code;
        const primitive =
          code == null ||
          typeof code === "string" ||
          typeof code === "number" ||
          typeof code === "boolean";
        payload.code = primitive ? code : String(code);
      }
    }
  } catch {
    // ignore metadata serialization failures
  }

  return payload;
}

function deserializeError(payload) {
  const message = typeof payload === "string" ? payload : String(payload?.message ?? "Unknown error");
  const err = new Error(message);
  if (payload?.stack) err.stack = String(payload.stack);
  if (typeof payload?.name === "string" && payload.name.trim().length > 0) {
    err.name = String(payload.name);
  }
  if (Object.prototype.hasOwnProperty.call(payload ?? {}, "code")) {
    err.code = payload.code;
  }
  return err;
}

function createTimeoutError({ extensionId, operation, timeoutMs }) {
  const err = new Error(`Extension ${extensionId} ${operation} timed out after ${timeoutMs}ms`);
  err.name = "ExtensionTimeoutError";
  err.code = "EXTENSION_TIMEOUT";
  return err;
}

function createWorkerTerminatedError({ extensionId, reason, cause }) {
  const base = cause instanceof Error ? cause.message : String(cause ?? "unknown reason");
  const err = new Error(`Extension worker terminated: ${extensionId}: ${reason}: ${base}`);
  err.name = "ExtensionWorkerTerminatedError";
  err.code = "EXTENSION_WORKER_TERMINATED";
  return err;
}

function createRequestId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function normalizeSandboxOptions(options) {
  const value = options && typeof options === "object" ? options : {};
  return {
    strictImports: value.strictImports !== false,
    disableEval: value.disableEval !== false
  };
}

const STORAGE_PROTO_POLLUTION_KEY = "__proto__";
// Persist `__proto__` under an internal alias so JSON parsing/loading cannot mutate prototypes.
const STORAGE_PROTO_POLLUTION_KEY_ALIAS = "__formula_reserved_key__:__proto__";

function normalizeStorageKey(key) {
  const s = String(key);
  if (s === STORAGE_PROTO_POLLUTION_KEY) return STORAGE_PROTO_POLLUTION_KEY_ALIAS;
  return s;
}

function normalizeExtensionStorageRecord(data) {
  const out = Object.create(null);
  let migrated = false;

  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return { record: out, migrated: false };
  }

  try {
    if (Object.getPrototypeOf(data) !== Object.prototype) {
      migrated = true;
    }
  } catch {
    migrated = true;
  }

  for (const [key, value] of Object.entries(data)) {
    const normalizedKey = normalizeStorageKey(key);
    if (normalizedKey !== key) migrated = true;
    out[normalizedKey] = value;
  }

  return { record: out, migrated };
}

class InMemoryExtensionStorage {
  constructor() {
    this._data = Object.create(null);
  }

  /**
   * @param {string} extensionId
   */
  getExtensionStore(extensionId) {
    const id = String(extensionId);
    if (!this._data[id]) this._data[id] = Object.create(null);
    return this._data[id];
  }

  /**
   * Clears any cached/in-memory storage for the given extension.
   * @param {string} extensionId
   */
  clearExtensionStore(extensionId) {
    delete this._data[String(extensionId)];
  }
}

function getDefaultLocalStorage() {
  try {
    if (typeof globalThis === "undefined") return null;
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

class LocalStorageExtensionStorage {
  /**
   * @param {{ storage?: Storage | null, keyPrefix?: string }} [options]
   */
  constructor({ storage, keyPrefix = "formula.extensionHost.storage." } = {}) {
    this._storage = storage ?? getDefaultLocalStorage();
    this._keyPrefix = String(keyPrefix);
    /** @type {Map<string, { target: Record<string, any>, proxy: any }>} */
    this._stores = new Map();
  }

  _key(extensionId) {
    return `${this._keyPrefix}${extensionId}`;
  }

  _load(extensionId) {
    if (!this._storage) return Object.create(null);
    const key = this._key(extensionId);
    try {
      const raw = this._storage.getItem(key);
      if (raw == null) return Object.create(null);
      const parsed = JSON.parse(raw);
      const { record, migrated } = normalizeExtensionStorageRecord(parsed);
      // If we migrated any data, or the persisted record is empty/noisy, rewrite it. `_persist`
      // removes the key entirely when empty.
      if (migrated || Object.keys(record).length === 0) {
        try {
          this._persist(extensionId, record);
        } catch {
          // ignore migration write failures
        }
      }
      return record;
    } catch {
      // If the stored value is corrupted (invalid JSON), remove it so the next access starts from
      // a clean slate.
      try {
        if (typeof this._storage.removeItem === "function") {
          this._storage.removeItem(key);
        }
      } catch {
        // ignore
      }
      return Object.create(null);
    }
  }

  _persist(extensionId, target) {
    if (!this._storage) return;
    const key = this._key(extensionId);
    // When the per-extension store becomes empty, remove the record entirely so localStorage
    // doesn't accumulate noisy `"{}"` entries.
    if (Object.keys(target).length === 0) {
      if (typeof this._storage.removeItem === "function") {
        this._storage.removeItem(key);
      } else {
        this._storage.setItem(key, JSON.stringify(target));
      }
      return;
    }
    this._storage.setItem(key, JSON.stringify(target));
  }

  /**
   * @param {string} extensionId
   */
  getExtensionStore(extensionId) {
    const id = String(extensionId);
    const existing = this._stores.get(id);
    if (existing) return existing.proxy;

    const target = this._load(id);
    const persist = () => {
      this._persist(id, target);
    };

    const proxy = new Proxy(target, {
      set: (obj, prop, value) => {
        const key = String(prop);
        const prev = Object.prototype.hasOwnProperty.call(obj, key) ? obj[key] : undefined;
        const hadPrev = Object.prototype.hasOwnProperty.call(obj, key);
        obj[key] = value;
        try {
          persist();
        } catch (err) {
          if (!hadPrev) {
            delete obj[key];
          } else {
            obj[key] = prev;
          }
          throw err;
        }
        return true;
      },
      deleteProperty: (obj, prop) => {
        const key = String(prop);
        const hadPrev = Object.prototype.hasOwnProperty.call(obj, key);
        const prev = obj[key];
        delete obj[key];
        try {
          persist();
        } catch (err) {
          if (hadPrev) obj[key] = prev;
          throw err;
        }
        return true;
      }
    });

    this._stores.set(id, { target, proxy });
    return proxy;
  }

  /**
   * Clears persisted + in-memory storage for the given extension.
   * @param {string} extensionId
   */
  clearExtensionStore(extensionId) {
    const id = String(extensionId);
    this._stores.delete(id);
    if (!this._storage) return;
    try {
      this._storage.removeItem(this._key(id));
    } catch {
      // ignore
    }
  }
}

class BrowserExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionStorage,
    permissionStorageKey,
    spreadsheetApi,
    clipboardApi,
    clipboardWriteGuard,
    storageApi,
    uiApi,
    activationTimeoutMs,
    commandTimeoutMs = 5000,
    customFunctionTimeoutMs = 5000,
    dataConnectorTimeoutMs = 5000,
    sandbox
  } = {}) {
    if (!spreadsheetApi) {
      throw new Error("BrowserExtensionHost requires a spreadsheetApi implementation");
    }

    this._engineVersion = engineVersion;
    this._extensions = new Map();
    this._commands = new Map();
    this._panelContributions = new Map();
    this._panels = new Map();
    this._contextMenus = new Map();
    this._customFunctions = new Map();
    this._dataConnectors = new Map();
    this._messages = [];
    this._uiApi = uiApi ?? null;
    const isNodeRuntime = typeof process !== "undefined" && typeof process?.versions?.node === "string";
    const defaultActivationTimeoutMs = isNodeRuntime ? 20000 : 5000;
    this._activationTimeoutMs = Number.isFinite(activationTimeoutMs)
      ? Math.max(0, activationTimeoutMs)
      : defaultActivationTimeoutMs;
    this._commandTimeoutMs = Number.isFinite(commandTimeoutMs) ? Math.max(0, commandTimeoutMs) : 5000;
    this._customFunctionTimeoutMs = Number.isFinite(customFunctionTimeoutMs)
      ? Math.max(0, customFunctionTimeoutMs)
      : 5000;
    this._dataConnectorTimeoutMs = Number.isFinite(dataConnectorTimeoutMs)
      ? Math.max(0, dataConnectorTimeoutMs)
      : 5000;

    this._sandbox = normalizeSandboxOptions(sandbox);

    this._spreadsheet = spreadsheetApi;
    this._workbook = { name: "MockWorkbook", path: null };
    this._sheets = [{ id: "sheet1", name: "Sheet1" }];
    this._nextSheetId = 2;
    this._activeSheetId = "sheet1";
    this._spreadsheetDisposables = [];

    this._clipboardText = "";
    this._clipboardApi = clipboardApi ?? {
      readText: async () => this._clipboardText,
      writeText: async (text) => {
        this._clipboardText = String(text ?? "");
      }
    };
    this._clipboardWriteGuard = typeof clipboardWriteGuard === "function" ? clipboardWriteGuard : null;

    if (storageApi) {
      this._storageApi = storageApi;
    } else {
      const ls = getDefaultLocalStorage();
      this._storageApi = ls ? new LocalStorageExtensionStorage({ storage: ls }) : new InMemoryExtensionStorage();
    }

    this._permissionManager = new PermissionManager({
      prompt: permissionPrompt,
      storage: permissionStorage,
      storageKey: permissionStorageKey
    });

    const trackSpreadsheetDisposable = (candidate) => {
      if (!candidate) return;
      if (typeof candidate === "function") {
        this._spreadsheetDisposables.push(candidate);
        return;
      }
      if (candidate && typeof candidate === "object" && typeof candidate.dispose === "function") {
        this._spreadsheetDisposables.push(() => {
          try {
            candidate.dispose();
          } catch {
            // ignore
          }
        });
      }
    };

    trackSpreadsheetDisposable(this._spreadsheet.onSelectionChanged?.((e) => this._broadcastEvent("selectionChanged", e)));
    trackSpreadsheetDisposable(this._spreadsheet.onCellChanged?.((e) => this._broadcastEvent("cellChanged", e)));
    trackSpreadsheetDisposable(this._spreadsheet.onSheetActivated?.((e) => this._broadcastEvent("sheetActivated", e)));
    trackSpreadsheetDisposable(this._spreadsheet.onWorkbookOpened?.((e) => this._broadcastEvent("workbookOpened", e)));
    trackSpreadsheetDisposable(this._spreadsheet.onBeforeSave?.((e) => this._broadcastEvent("beforeSave", e)));
  }

  get engineVersion() {
    return this._engineVersion;
  }

  get spreadsheet() {
    return this._spreadsheet;
  }

  async _getActiveWorkbook() {
    if (typeof this._spreadsheet.getActiveWorkbook === "function") {
      try {
        const workbook = await this._spreadsheet.getActiveWorkbook();
        if (workbook && typeof workbook === "object") {
          const next = { ...this._workbook };
          // `Workbook.name` is required by the API contract, but treat empty/missing values
          // as "no update" to preserve a stable snapshot in partially-implemented hosts.
          let nameSet = false;
          try {
            if (Object.prototype.hasOwnProperty.call(workbook, "name")) {
              const rawName = workbook.name;
              const trimmed = rawName == null ? "" : String(rawName).trim();
              if (trimmed) {
                next.name = trimmed;
                nameSet = true;
              }
            }
          } catch {
            // ignore
          }

          // `path` is optional; importantly, `null` means "unsaved workbook". Preserve the
          // existing path only when the host omits `path` (undefined), so transitions from
          // saved -> unsaved are reflected in extension snapshots/events.
          try {
            if (Object.prototype.hasOwnProperty.call(workbook, "path")) {
              const rawPath = workbook.path;
              if (rawPath === undefined) {
                // keep previous
              } else if (rawPath == null) {
                next.path = null;
              } else {
                const str = String(rawPath);
                const trimmed = str.trim();
                next.path = trimmed.length > 0 ? str : null;
                if (!nameSet && trimmed.length > 0) {
                  // Best-effort: if the host omits `name` but provides a file path,
                  // derive the workbook name from the basename.
                  next.name = trimmed.split(/[/\\]/).pop() ?? trimmed;
                  nameSet = true;
                }
              }
            }
          } catch {
            // ignore
          }

          this._workbook = { name: next.name, path: next.path ?? null };
        }
      } catch {
        // ignore
      }
    }
    return this._getWorkbookSnapshot();
  }

  _getWorkbookSnapshot() {
    const externalSheets =
      typeof this._spreadsheet.listSheets === "function" ? this._spreadsheet.listSheets() : null;
    const sheets = Array.isArray(externalSheets)
      ? externalSheets.map((s) => ({ id: s.id, name: s.name }))
      : Array.isArray(this._sheets)
        ? this._sheets.map((s) => ({ id: s.id, name: s.name }))
        : [];

    const externalActiveSheet =
      typeof this._spreadsheet.getActiveSheet === "function" ? this._spreadsheet.getActiveSheet() : null;
    const activeSheet =
      externalActiveSheet && typeof externalActiveSheet === "object"
        ? { id: externalActiveSheet.id, name: externalActiveSheet.name }
        : sheets.find((s) => s.id === this._activeSheetId) ??
          sheets[0] ??
          { id: "sheet1", name: "Sheet1" };

    return {
      name: this._workbook.name,
      path: this._workbook.path,
      sheets,
      activeSheet
    };
  }

  openWorkbook(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    const name =
      workbookPathStr == null || workbookPathStr.trim().length === 0
        ? "MockWorkbook"
        : workbookPathStr.split(/[/\\]/).pop() ?? workbookPathStr;
    this._workbook = { name, path: workbookPathStr };
    const workbook = this._getWorkbookSnapshot();
    this._broadcastEvent("workbookOpened", { workbook });
    return workbook;
  }

  saveWorkbook() {
    // Browser host stub mirrors the Node host: emit a stable payload for beforeSave.
    this._broadcastEvent("beforeSave", { workbook: this._getWorkbookSnapshot() });
    return true;
  }

  saveWorkbookAs(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    if (workbookPathStr == null || workbookPathStr.trim().length === 0) {
      throw new Error("Workbook path must be a non-empty string");
    }

    const name = workbookPathStr.split(/[/\\]/).pop() ?? workbookPathStr;
    this._workbook = { name, path: workbookPathStr };
    this._broadcastEvent("beforeSave", { workbook: this._getWorkbookSnapshot() });
    return true;
  }

  closeWorkbook() {
    this.openWorkbook(null);
    return true;
  }

  async loadExtensionFromUrl(manifestUrl) {
    if (!manifestUrl) throw new Error("manifestUrl is required");
    const resolvedUrl = new URL(
      String(manifestUrl),
      globalThis.location?.href ?? "http://localhost/"
    ).toString();

    const response = await fetch(resolvedUrl);
    if (!response.ok) {
      throw new Error(`Failed to fetch extension manifest: ${resolvedUrl} (${response.status})`);
    }

    const parsed = await response.json();
    const manifest = validateExtensionManifest(parsed, {
      engineVersion: this._engineVersion,
      enforceEngine: true
    });
    const extensionId = `${manifest.publisher}.${manifest.name}`;
    if (/[/\\]/.test(extensionId) || extensionId.includes("\0")) {
      throw new Error(
        `Invalid extension id: ${extensionId} (publisher/name must not contain path separators)`
      );
    }

    const baseUrl = new URL("./", resolvedUrl);
    const entry = manifest.browser ?? manifest.module ?? manifest.main;
    if (!entry) {
      throw new Error(`Extension manifest missing entrypoint (main/module/browser): ${extensionId}`);
    }

    const mainUrl = new URL(entry, baseUrl);
    if (!mainUrl.href.startsWith(baseUrl.href)) {
      throw new Error(
        `Extension entrypoint must resolve inside extension base URL: '${entry}' resolved to '${mainUrl.href}'`
      );
    }

    await this.loadExtension({
      extensionId,
      extensionPath: baseUrl.toString(),
      manifest,
      mainUrl: mainUrl.toString()
    });

    return extensionId;
  }

  async loadExtension({ extensionId, extensionPath, manifest, mainUrl }) {
    if (!extensionId) throw new Error("extensionId is required");
    if (!manifest || typeof manifest !== "object") throw new Error("manifest is required");
    if (!mainUrl) throw new Error("mainUrl is required");

    if (this._extensions.has(extensionId)) {
      throw new Error(`Extension already loaded: ${extensionId}`);
    }

    const extension = {
      id: extensionId,
      path: extensionPath,
      manifest,
      mainUrl,
      worker: null,
      active: false,
      registeredCommands: new Set(),
      pendingRequests: new Map(),
      taintedRanges: [],
      workerData: {
        extensionId,
        extensionPath,
        mainUrl,
        extensionUri: extensionPath,
        sandbox: { ...this._sandbox },
        // Browser host has no filesystem; provide stable identifiers so extensions
        // can key storage off these paths similarly to the desktop host.
        globalStoragePath: `memory://formula/extensions/${extensionId}/globalStorage`,
        workspaceStoragePath: `memory://formula/extensions/${extensionId}/workspaceStorage`
      }
    };

    for (const cmd of manifest.contributes.commands ?? []) {
      const existing = this._commands.get(cmd.command);
      if (existing && existing !== extensionId) {
        throw new Error(`Command id already contributed: ${cmd.command} (existing: ${existing}, new: ${extensionId})`);
      }
      this._commands.set(cmd.command, extensionId);
    }

    for (const panel of manifest.contributes.panels ?? []) {
      const existing = this._panelContributions.get(panel.id);
      if (existing && existing !== extensionId) {
        throw new Error(`Panel id already contributed: ${panel.id} (existing: ${existing}, new: ${extensionId})`);
      }
      this._panelContributions.set(panel.id, extensionId);
    }

    for (const fn of manifest.contributes.customFunctions ?? []) {
      if (this._customFunctions.has(fn.name)) {
        throw new Error(`Duplicate custom function name: ${fn.name}`);
      }
      this._customFunctions.set(fn.name, extensionId);
    }

    for (const connector of manifest.contributes.dataConnectors ?? []) {
      if (this._dataConnectors.has(connector.id)) {
        throw new Error(`Duplicate data connector id: ${connector.id}`);
      }
      this._dataConnectors.set(connector.id, extensionId);
    }

    this._extensions.set(extensionId, extension);
    this._spawnWorker(extension);
    return extensionId;
  }

  /**
   * Clears all persisted state owned by an extension (permission grants + extension storage).
   *
   * Intended to be called by install/uninstall managers so a reinstall behaves like a clean
   * install.
   *
   * @param {string} extensionId
   */
  async resetExtensionState(extensionId) {
    const id = String(extensionId);

    // Best-effort: these should never throw for the default localStorage-backed implementations,
    // but we intentionally do not let failures prevent uninstall flows from completing.
    try {
      await this._permissionManager.revokePermissions(id);
    } catch {
      // ignore
    }

    try {
      const storageApi = this._storageApi;
      if (!storageApi) return;

      if (typeof storageApi.clearExtensionStore === "function") {
        try {
          await storageApi.clearExtensionStore(id);
          return;
        } catch {
          // Fall back to clearing keys below.
        }
      }

      // Fallback: clear all keys from the per-extension store.
      if (typeof storageApi.getExtensionStore === "function") {
        const store = storageApi.getExtensionStore(id);
        if (store && typeof store === "object") {
          for (const key of Object.keys(store)) {
            try {
              delete store[key];
            } catch {
              // ignore
            }
          }
        }
      }
    } catch {
      // ignore
    }
  }

  async updateExtension({ extensionId, extensionPath, manifest, mainUrl }) {
    const id = String(extensionId);
    if (this._extensions.has(id)) {
      await this.unloadExtension(id);
    }
    return this.loadExtension({ extensionId: id, extensionPath, manifest, mainUrl });
  }

  async startup() {
    const tasks = [];
    for (const extension of this._extensions.values()) {
      if ((extension.manifest.activationEvents ?? []).includes("onStartupFinished")) {
        tasks.push(this._activateExtension(extension, "onStartupFinished"));
      }
    }
    await Promise.all(tasks);

    // Mirror the Node host behavior: extensions that activate on startup should receive the
    // initial workbook event so they can initialize state.
    this._broadcastEvent("workbookOpened", { workbook: await this._getActiveWorkbook() });
  }

  async startupExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    if ((extension.manifest.activationEvents ?? []).includes("onStartupFinished") && !extension.active) {
      await this._activateExtension(extension, "onStartupFinished");
      const workbook = await this._getActiveWorkbook();
      this._sendEventToExtension(extension, "workbookOpened", { workbook });
    }
  }

  async activateView(viewId) {
    const id = String(viewId ?? "");
    const activationEvent = `onView:${id}`;
    const targets = [];
    for (const extension of this._extensions.values()) {
      if ((extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        targets.push(extension);
      }
    }

    // Activation events are used to decide which extensions should be activated when a view is
    // shown (similar to `onCommand:*`). The `viewActivated` event itself is broadcast to all
    // active extensions so any running extension can subscribe via `formula.events.onViewActivated`.
    //
    // Broadcast before activating `onView:*` targets so an unrelated activation failure cannot
    // prevent already-running extensions from receiving the view activation notification.
    this._broadcastEvent("viewActivated", { viewId: id });

    for (const extension of targets) {
      if (extension.active) continue;
      try {
        // eslint-disable-next-line no-await-in-loop
        await this._activateExtension(extension, activationEvent);

        // Extensions activated as a result of this view activation need to receive the event
        // after they finish starting up.
        this._sendEventToExtension(extension, "viewActivated", { viewId: id });
      } catch (error) {
        // Best-effort: one broken extension should not prevent view activation notifications
        // from reaching other extensions.
        // eslint-disable-next-line no-console
        console.error(
          `[formula][extensions] Failed to activate extension ${extension.id} for ${activationEvent}: ${String(
            error?.message ?? error
          )}`
        );
      }
    }
  }

  async activateCustomFunction(functionName) {
    const activationEvent = `onCustomFunction:${functionName}`;
    const tasks = [];
    for (const extension of this._extensions.values()) {
      if (extension.active) continue;
      if ((extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        tasks.push(this._activateExtension(extension, activationEvent));
      }
    }
    await Promise.all(tasks);
  }

  async executeCommand(commandId, ...args) {
    const extensionId = this._commands.get(commandId);
    if (!extensionId) throw new Error(`Unknown command: ${commandId}`);

    const extension = this._extensions.get(extensionId);
    if (!extension) throw new Error(`Extension not loaded: ${extensionId}`);

    // Executing a contributed command is a user-facing UI surface. Even if the extension has
    // already activated (and registered its command handlers), re-check `ui.commands` so revoking
    // permissions takes effect immediately without requiring a full host restart.
    await this._permissionManager.ensurePermissions(
      {
        extensionId: extension.id,
        displayName: extension.manifest.displayName ?? extension.manifest.name,
        declaredPermissions: extension.manifest.permissions ?? []
      },
      ["ui.commands"],
      { apiKey: "commands.executeCommand" }
    );

    if (!extension.active) {
      const activationEvent = `onCommand:${commandId}`;
      if (!(extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        throw new Error(`Extension ${extensionId} is not activated for ${activationEvent}`);
      }
      await this._activateExtension(extension, activationEvent);
    }

    await this._ensureWorker(extension);
    return this._requestFromWorker(
      extension,
      {
        type: "execute_command",
        commandId,
        args
      },
      {
        timeoutMs: this._commandTimeoutMs,
        operation: `command '${commandId}'`
      }
    );
  }

  async invokeCustomFunction(functionName, ...args) {
    const name = String(functionName);
    const extensionId = this._customFunctions.get(name);
    if (!extensionId) throw new Error(`Unknown custom function: ${name}`);

    const extension = this._extensions.get(extensionId);
    if (!extension) throw new Error(`Extension not loaded: ${extensionId}`);

    if (!extension.active) {
      const activationEvent = `onCustomFunction:${name}`;
      if (!(extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        throw new Error(`Extension ${extensionId} is not activated for ${activationEvent}`);
      }
      await this._activateExtension(extension, activationEvent);
    }

    await this._ensureWorker(extension);
    return this._requestFromWorker(
      extension,
      {
        type: "invoke_custom_function",
        functionName: name,
        args
      },
      {
        timeoutMs: this._customFunctionTimeoutMs,
        operation: `custom function '${name}'`
      }
    );
  }

  async invokeDataConnector(connectorId, method, ...args) {
    const id = String(connectorId);
    const extensionId = this._dataConnectors.get(id);
    if (!extensionId) throw new Error(`Unknown data connector: ${id}`);

    const extension = this._extensions.get(extensionId);
    if (!extension) throw new Error(`Extension not loaded: ${extensionId}`);

    const methodName = String(method);
    if (methodName.trim().length === 0) throw new Error("Data connector method must be a non-empty string");

    if (!extension.active) {
      const activationEvent = `onDataConnector:${id}`;
      if (!(extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        throw new Error(`Extension ${extensionId} is not activated for ${activationEvent}`);
      }
      await this._activateExtension(extension, activationEvent);
    }

    await this._ensureWorker(extension);
    return this._requestFromWorker(
      extension,
      {
        type: "invoke_data_connector",
        connectorId: id,
        method: methodName,
        args
      },
      {
        timeoutMs: this._dataConnectorTimeoutMs,
        operation: `data connector '${id}' (${methodName})`
      }
    );
  }

  getMessages() {
    return [...this._messages];
  }

  getPanel(panelId) {
    return this._panels.get(panelId);
  }

  getPanelOutgoingMessages(panelId) {
    const panel = this._panels.get(panelId);
    if (!panel) return [];
    return [...(panel.outgoingMessages ?? [])];
  }

  dispatchPanelMessage(panelId, message) {
    const panel = this._panels.get(panelId);
    if (!panel) throw new Error(`Unknown panel: ${panelId}`);
    const extension = this._extensions.get(panel.extensionId);
    if (!extension) throw new Error(`Extension not loaded: ${panel.extensionId}`);
    if (!extension.worker) {
      throw new Error(`Extension worker not running: ${panel.extensionId}`);
    }
    extension.worker.postMessage({ type: "panel_message", panelId, message });
  }

  listExtensions() {
    return [...this._extensions.values()].map((ext) => ({
      id: ext.id,
      path: ext.path,
      active: ext.active,
      manifest: ext.manifest
    }));
  }

  async getGrantedPermissions(extensionId) {
    return this._permissionManager.getGrantedPermissions(String(extensionId));
  }

  async revokePermissions(extensionId, permissions) {
    return this._permissionManager.revokePermissions(String(extensionId), permissions);
  }

  async resetPermissions(extensionId) {
    const id = String(extensionId);
    await this._permissionManager.resetPermissions(id);

    // Resetting an extension's permissions should take effect immediately.
    //
    // Many permissions (e.g. `ui.commands`) are requested during activation when the extension
    // registers commands/panels. If the worker stays alive after a reset, the extension won't
    // re-run its activation path and therefore won't re-request permissions. Restart the worker
    // so future command/view activations behave like a fresh load.
    const extension = this._extensions.get(id);
    if (extension) {
      this._terminateWorker(extension, {
        reason: "permissions_reset",
        cause: new Error("Extension permissions reset"),
      });
    }
  }

  async resetAllPermissions() {
    await this._permissionManager.resetAllPermissions();
    for (const extension of this._extensions.values()) {
      this._terminateWorker(extension, {
        reason: "permissions_reset",
        cause: new Error("Extension permissions reset"),
      });
    }
  }

  /**
   * Best-effort clearing of extension storage/config state.
   * This is intended for uninstall flows so a reinstall starts fresh.
   *
   * @param {string} extensionId
   */
  async clearExtensionStorage(extensionId) {
    const id = String(extensionId);
    const api = this._storageApi;
    if (!api) return;
    if (typeof api.clearExtensionStore === "function") {
      try {
        await api.clearExtensionStore(id);
        return;
      } catch {
        // fall through to clearing keys
      }
    }

    // Back-compat fallback: clear all keys from the per-extension store.
    // This does not guarantee the backing record is removed (depends on the storage backend),
    // but ensures the next install sees an empty store.
    if (typeof api.getExtensionStore === "function") {
      try {
        const store = api.getExtensionStore(id);
        if (store && typeof store === "object") {
          for (const key of Object.keys(store)) {
            try {
              delete store[key];
            } catch {
              // ignore
            }
          }
        }
      } catch {
        // ignore
      }
    }
  }

  getContributedCommands() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const cmd of extension.manifest.contributes?.commands ?? []) {
        const keywords = Array.isArray(cmd.keywords)
          ? cmd.keywords.filter((kw) => typeof kw === "string" && kw.trim().length > 0)
          : null;
        out.push({
          extensionId: extension.id,
          command: cmd.command,
          title: cmd.title,
          category: cmd.category ?? null,
          icon: cmd.icon ?? null,
          description: typeof cmd.description === "string" ? cmd.description : null,
          keywords
        });
      }
    }
    return out;
  }

  getContributedPanels() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const panel of extension.manifest.contributes?.panels ?? []) {
        out.push({
          extensionId: extension.id,
          id: panel.id,
          title: panel.title,
          icon: panel.icon ?? null
        });
      }
    }
    return out;
  }

  getContributedKeybindings() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const kb of extension.manifest.contributes?.keybindings ?? []) {
        out.push({
          extensionId: extension.id,
          command: kb.command,
          key: kb.key,
          mac: kb.mac ?? null,
          when: kb.when ?? null
        });
      }
    }
    return out;
  }

  /**
   * @param {string} menuId
   */
  getContributedMenu(menuId) {
    const id = String(menuId);
    const out = [];
    for (const extension of this._extensions.values()) {
      const menus = extension.manifest.contributes?.menus ?? {};
      const items = menus[id] ?? [];
      for (const item of items) {
        out.push({
          extensionId: extension.id,
          command: item.command,
          when: item.when ?? null,
          group: item.group ?? null
        });
      }
    }
    for (const record of this._contextMenus.values()) {
      if (record.menuId !== id) continue;
      for (const item of record.items ?? []) {
        out.push({
          extensionId: record.extensionId,
          command: item.command,
          when: item.when ?? null,
          group: item.group ?? null
        });
      }
    }
    return out;
  }

  getContributedMenus() {
    const menuIds = new Set();
    for (const extension of this._extensions.values()) {
      const menus = extension.manifest.contributes?.menus ?? {};
      for (const key of Object.keys(menus)) menuIds.add(key);
    }
    for (const record of this._contextMenus.values()) {
      if (record?.menuId) menuIds.add(record.menuId);
    }

    const out = {};
    for (const id of [...menuIds].sort()) {
      out[id] = this.getContributedMenu(id);
    }
    return out;
  }

  getContributedCustomFunctions() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const fn of extension.manifest.contributes?.customFunctions ?? []) {
        out.push({
          extensionId: extension.id,
          name: fn.name,
          description: fn.description ?? null
        });
      }
    }
    return out;
  }

  getContributedDataConnectors() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const connector of extension.manifest.contributes?.dataConnectors ?? []) {
        out.push({
          extensionId: extension.id,
          id: connector.id,
          name: connector.name,
          icon: connector.icon ?? null
        });
      }
    }
    return out;
  }

  async dispose() {
    // Best-effort: ensure spreadsheetApi event subscriptions are cleaned up so callers that
    // create/destroy multiple hosts (tests, hot reload) do not leak listeners.
    try {
      const disposables = Array.isArray(this._spreadsheetDisposables) ? this._spreadsheetDisposables : [];
      this._spreadsheetDisposables = [];
      for (const dispose of disposables) {
        try {
          if (typeof dispose === "function") dispose();
        } catch {
          // ignore
        }
      }
    } catch {
      // ignore
    }

    const extensions = [...this._extensions.values()];
    for (const ext of extensions) {
      this._terminateWorker(ext, { reason: "dispose", cause: new Error("BrowserExtensionHost disposed") });
    }

    this._extensions.clear();
    this._commands.clear();
    this._panelContributions.clear();
    this._panels.clear();
    this._contextMenus.clear();
    this._customFunctions.clear();
    this._dataConnectors.clear();
    this._messages = [];
  }

  async reloadExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    this._terminateWorker(extension, { reason: "reload", cause: new Error("Extension reloaded") });
    this._spawnWorker(extension);
  }

  async unloadExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    this._terminateWorker(extension, { reason: "unload", cause: new Error("Extension unloaded") });

    try {
      for (const cmd of extension.manifest.contributes?.commands ?? []) {
        if (this._commands.get(cmd.command) === extension.id) this._commands.delete(cmd.command);
      }
    } catch {
      // ignore
    }

    try {
      for (const panel of extension.manifest.contributes?.panels ?? []) {
        if (this._panelContributions.get(panel.id) === extension.id) {
          this._panelContributions.delete(panel.id);
        }
      }
    } catch {
      // ignore
    }

    try {
      for (const fn of extension.manifest.contributes?.customFunctions ?? []) {
        if (this._customFunctions.get(fn.name) === extension.id) this._customFunctions.delete(fn.name);
      }
    } catch {
      // ignore
    }

    try {
      for (const connector of extension.manifest.contributes?.dataConnectors ?? []) {
        if (this._dataConnectors.get(connector.id) === extension.id) this._dataConnectors.delete(connector.id);
      }
    } catch {
      // ignore
    }

    // Panels created by the extension are cleaned up in `_terminateWorker()`, which also notifies
    // the UI via `onPanelDisposed`. (Keeping the disposal logic centralized avoids leaving dangling
    // UI panels when the worker crashes/timeouts.)

    for (const [registrationId, record] of this._contextMenus.entries()) {
      if (record?.extensionId === extension.id) this._contextMenus.delete(registrationId);
    }

    this._extensions.delete(id);
  }

  _spawnWorker(extension) {
    if (extension.worker) return;

    // Keep this inline `new Worker(new URL(...))` shape so Vite (and other bundlers)
    // can statically analyze the worker entry during production builds.
    const worker = new Worker(new URL("./extension-worker.mjs", import.meta.url), { type: "module" });
    extension.worker = worker;

    worker.addEventListener("message", (event) =>
      this._handleWorkerMessage(extension, worker, event.data)
    );
    worker.addEventListener("error", (event) =>
      this._handleWorkerCrash(extension, worker, new Error(String(event?.message ?? "Worker error")))
    );

    try {
      worker.postMessage({ type: "init", ...extension.workerData });
    } catch {
      // ignore
    }
  }

  async _ensureWorker(extension) {
    if (extension.worker) return;
    // A fresh worker always starts inactive until we successfully activate.
    extension.active = false;
    this._spawnWorker(extension);
  }

  _requestFromWorker(extension, message, { timeoutMs, operation }) {
    const worker = extension.worker;
    if (!worker) {
      throw new Error(`Extension worker not running: ${extension.id}`);
    }

    const id = createRequestId();
    const payload = { ...message, id };

    return new Promise((resolve, reject) => {
      const timeout =
        timeoutMs > 0
          ? setTimeout(() => {
              const pending = extension.pendingRequests.get(id);
              if (!pending) return;
              extension.pendingRequests.delete(id);
              const err = createTimeoutError({
                extensionId: extension.id,
                operation,
                timeoutMs
              });
              pending.reject(err);
              this._terminateWorker(extension, { reason: "timeout", cause: err });
            }, timeoutMs)
          : null;

      extension.pendingRequests.set(id, { resolve, reject, timeout });

      try {
        worker.postMessage(payload);
      } catch (error) {
        extension.pendingRequests.delete(id);
        if (timeout) clearTimeout(timeout);
        reject(error);
      }
    });
  }

  _terminateWorker(extension, { reason, cause }) {
    const worker = extension.worker;

    extension.active = false;
    // Intentionally preserve `taintedRanges` across worker restarts/crashes. Extensions can persist
    // spreadsheet data outside worker memory (e.g. storage/IndexedDB) and later attempt clipboard
    // writes from a fresh worker. Clearing taint on termination would allow bypassing clipboard DLP
    // enforcement via a deliberate crash/timeout/restart.

    // Best-effort cleanup for runtime registrations that were tied to the crashed worker.
    // Contributed commands stay registered so the app can still route them for future activations.
    try {
      const contributedCommands = new Set((extension.manifest.contributes.commands ?? []).map((c) => c.command));
      for (const cmd of extension.registeredCommands) {
        if (contributedCommands.has(cmd)) continue;
        if (this._commands.get(cmd) === extension.id) this._commands.delete(cmd);
      }
      extension.registeredCommands.clear();
    } catch {
      // ignore
    }

    const disposedPanelIds = [];
    for (const [panelId, panel] of this._panels.entries()) {
      if (panel?.extensionId !== extension.id) continue;
      this._panels.delete(panelId);
      disposedPanelIds.push(panelId);
    }
    if (disposedPanelIds.length > 0) {
      for (const panelId of disposedPanelIds) {
        try {
          this._uiApi?.onPanelDisposed?.(panelId);
        } catch {
          // ignore
        }
      }
    }

    // Remove context menus registered by this extension worker; they would otherwise leak.
    for (const [registrationId, record] of this._contextMenus.entries()) {
      if (record?.extensionId === extension.id) this._contextMenus.delete(registrationId);
    }

    for (const [reqId, pending] of extension.pendingRequests.entries()) {
      if (pending.timeout) clearTimeout(pending.timeout);
      pending.reject(createWorkerTerminatedError({ extensionId: extension.id, reason, cause }));
      extension.pendingRequests.delete(reqId);
    }

    if (!worker) return;
    extension.worker = null;

    try {
      worker.terminate();
    } catch {
      // ignore
    }
  }

  async _handleApiCall(extension, message) {
    const worker = extension.worker;
    const id = message?.id;
    // Treat malformed `args` payloads as "no args" to avoid TypeError crashes when untrusted
    // extension workers post malformed messages.
    const args = Array.isArray(message?.args) ? message.args : [];
    let apiKey = "";
    let permissions = [];
    let isNetworkApi = false;
    let networkUrl = null;
    let namespace = "";
    let method = "";
    const safeCoerceString = (value) => {
      try {
        return String(value ?? "");
      } catch {
        return "";
      }
    };

    try {
      namespace = typeof message?.namespace === "string" ? message.namespace : safeCoerceString(message?.namespace);
      method = typeof message?.method === "string" ? message.method : safeCoerceString(message?.method);
      apiKey = `${namespace}.${method}`;
      permissions = API_PERMISSIONS[apiKey] ?? [];
      isNetworkApi = apiKey === "network.fetch" || apiKey === "network.openWebSocket";
      networkUrl = isNetworkApi ? safeCoerceString(args?.[0]) : null;

      // Validate obvious argument errors before prompting for permissions to avoid
      // spurious permission prompts for calls that will fail fast anyway.
      if (apiKey === "workbook.openWorkbook" || apiKey === "workbook.saveAs") {
        const workbookPath = args?.[0];
        if (typeof workbookPath !== "string" || workbookPath.trim().length === 0) {
          throw new Error("Workbook path must be a non-empty string");
        }
      }
      await this._permissionManager.ensurePermissions(
        {
          extensionId: extension.id,
          displayName: extension.manifest.displayName ?? extension.manifest.name,
          declaredPermissions: extension.manifest.permissions ?? []
        },
        permissions,
        {
          apiKey,
          ...(isNetworkApi ? { network: { url: networkUrl } } : {})
        }
      );

      if (isNetworkApi) {
        const grants = await this._permissionManager.getGrantedPermissions(extension.id);
        const policy = grants?.network;
        if (!policy) {
          throw new PermissionError("Permission denied: network");
        }
        if (policy.mode === "deny") {
          throw new PermissionError("Permission denied: network");
        }
        if (policy.mode === "allowlist" && !isUrlAllowedByHosts(networkUrl, policy.hosts ?? [])) {
          const host = safeParseUrl(networkUrl)?.hostname ?? null;
          throw new PermissionError(
            host ? `Permission denied: network (${host})` : `Permission denied: network (${networkUrl})`
          );
        }
      }

      const result = await this._executeApi(namespace, method, args, extension);

      try {
        worker?.postMessage({
          type: "api_result",
          id,
          result
        });
      } catch {
        // ignore
      }
    } catch (error) {
      try {
        worker?.postMessage({
          type: "api_error",
          id,
          error: serializeError(error)
        });
      } catch {
        // ignore
      }
    }
  }

  async _executeApi(namespace, method, args, extension) {
    switch (`${namespace}.${method}`) {
      case "workbook.getActiveWorkbook":
        return this._getActiveWorkbook();
      case "workbook.openWorkbook":
        {
          const workbookPath = args?.[0];
          if (typeof workbookPath !== "string" || workbookPath.trim().length === 0) {
            throw new Error("Workbook path must be a non-empty string");
          }

          if (typeof this._spreadsheet.openWorkbook === "function") {
            // Best-effort: update internal workbook metadata from the input path so
            // `_getActiveWorkbook` has reasonable defaults even if the host doesn't
            // implement `getActiveWorkbook`.
            const name = workbookPath.split(/[/\\]/).pop() ?? workbookPath;
            const prev = this._workbook;
            this._workbook = { name, path: workbookPath };

            try {
              await this._spreadsheet.openWorkbook(workbookPath);
              const workbook = await this._getActiveWorkbook();
              this._broadcastEvent("workbookOpened", { workbook });
              return workbook;
            } catch (err) {
              this._workbook = prev;
              throw err;
            }
          }
          return this.openWorkbook(workbookPath);
        }
      case "workbook.createWorkbook":
        if (typeof this._spreadsheet.createWorkbook === "function") {
          // Mirror the browser-host stub: reset workbook metadata before creation so
          // consumers relying on the in-memory snapshot get deterministic fields.
          const prev = this._workbook;
          this._workbook = { name: "MockWorkbook", path: null };
          try {
            await this._spreadsheet.createWorkbook();
            const workbook = await this._getActiveWorkbook();
            this._broadcastEvent("workbookOpened", { workbook });
            return workbook;
          } catch (err) {
            this._workbook = prev;
            throw err;
          }
        }
        return this.openWorkbook(null);
      case "workbook.save":
        if (typeof this._spreadsheet.saveWorkbook === "function") {
          const workbook = await this._getActiveWorkbook();
          // If the workbook already has a path we can emit `beforeSave` immediately.
          // For pathless workbooks, the host app may prompt for a Save As path and can
          // cancel the operation. Avoid emitting `beforeSave` until the path is known.
          const workbookPath = workbook?.path;
          const hasPath = typeof workbookPath === "string" && workbookPath.trim().length > 0;
          if (hasPath) {
            this._broadcastEvent("beforeSave", { workbook });
          }
          await this._spreadsheet.saveWorkbook();
          // Saving can update the workbook path/name (e.g. Save prompting for a path).
          // Refresh internal workbook metadata so subsequent snapshots are accurate.
          await this._getActiveWorkbook();
          return null;
        }
        this.saveWorkbook();
        return null;
      case "workbook.saveAs":
        {
          const workbookPath = args?.[0];
          if (typeof workbookPath !== "string" || workbookPath.trim().length === 0) {
            throw new Error("Workbook path must be a non-empty string");
          }

          if (typeof this._spreadsheet.saveWorkbookAs === "function") {
            const name = workbookPath.split(/[/\\]/).pop() ?? workbookPath;
            const prev = this._workbook;
            this._workbook = { name, path: workbookPath };
            this._broadcastEvent("beforeSave", { workbook: this._getWorkbookSnapshot() });
            try {
              await this._spreadsheet.saveWorkbookAs(workbookPath);
              await this._getActiveWorkbook();
              return null;
            } catch (err) {
              this._workbook = prev;
              throw err;
            }
          }

          this.saveWorkbookAs(workbookPath);
          return null;
        }
      case "workbook.close":
        if (typeof this._spreadsheet.closeWorkbook === "function") {
          // Mirror browser-host semantics: after close, treat the session as a fresh workbook
          // (close is modeled as "open empty workbook").
          const prev = this._workbook;
          this._workbook = { name: "MockWorkbook", path: null };
          try {
            await this._spreadsheet.closeWorkbook();
            const workbook = await this._getActiveWorkbook();
            this._broadcastEvent("workbookOpened", { workbook });
            return null;
          } catch (err) {
            this._workbook = prev;
            throw err;
          }
        }
        this.closeWorkbook();
        return null;
      case "sheets.getActiveSheet":
        if (typeof this._spreadsheet.getActiveSheet === "function") {
          const sheet = await this._spreadsheet.getActiveSheet();
          const id = sheet && typeof sheet === "object" && sheet.id != null ? String(sheet.id).trim() : "";
          if (id) this._activeSheetId = id;
          return sheet;
        }
        return this._sheets.find((s) => s.id === this._activeSheetId) ?? { id: "sheet1", name: "Sheet1" };

      case "cells.getSelection": {
        const sheetId = await this._resolveActiveSheetId();
        const result = await this._spreadsheet.getSelection();
        // Best-effort: keep selection reads consistent with the desktop extension API range cap.
        // (We can't prevent the spreadsheet implementation from allocating `values`, but we can
        // still fail fast before returning huge payloads to extension workers.)
        try {
          let coords = null;
          if (result && typeof result === "object") {
            coords = {
              startRow: result.startRow,
              startCol: result.startCol,
              endRow: result.endRow,
              endCol: result.endCol
            };
            const hasNumeric = Object.values(coords).every((v) => Number.isFinite(Number(v)));
            if (!hasNumeric) {
              coords = null;
            }
            if (!coords && typeof result?.address === "string" && result.address.trim()) {
              try {
                const { ref: addrRef } = this._splitSheetQualifier(result.address);
                coords = this._parseA1RangeRef(addrRef);
              } catch {
                coords = null;
              }
            }
          }
          if (coords) assertExtensionRangeWithinLimits(coords, { label: "Selection" });
        } catch (err) {
          // Match desktop semantics: `cells.getSelection` throws if the selection is too large.
          throw err;
        }
        if (result && typeof result === "object") {
          const direct = normalizeTaintedRange({
            sheetId,
            startRow: result.startRow,
            startCol: result.startCol,
            endRow: result.endRow,
            endCol: result.endCol
          });
          if (direct) {
            this._taintExtensionRange(extension, direct);
          } else {
            // Best-effort: if the host returns a canonical A1 `address` but omits numeric coords,
            // fall back to parsing it so clipboard DLP still reflects what the extension read.
            try {
              const addr = typeof result.address === "string" ? result.address : null;
              if (addr) {
                const { sheetName: addrSheetName, ref: addrRef } = this._splitSheetQualifier(addr);
                const parsed = this._parseA1RangeRef(addrRef);
                const rangeSheetId = addrSheetName != null ? await this._resolveSheetId(addrSheetName) : sheetId;
                this._taintExtensionRange(extension, { sheetId: rangeSheetId, ...parsed });
              }
            } catch {
              // ignore
            }
          }
        }
        return result;
      }
      case "cells.getRange": {
        const ref = args[0];
        const { sheetName, ref: a1Ref } = this._splitSheetQualifier(ref);
        // Fail fast on unbounded A1 ranges so we don't ask the spreadsheet implementation to
        // materialize huge 2D arrays.
        {
          let coords = null;
          try {
            coords = this._parseA1RangeRef(a1Ref);
          } catch {
            coords = null;
          }
          if (coords) {
            assertExtensionRangeWithinLimits(coords, { label: "Range" });
          }
        }
        const sheetId = await this._resolveSheetId(sheetName);
        const hasGetRange = typeof this._spreadsheet.getRange === "function";
        /** @type {any} */
        let result;
        if (hasGetRange) {
          result = await this._spreadsheet.getRange(ref);
        } else {
          const activeSheetId = await this._resolveActiveSheetId();
          // `getCell` is scoped to the active sheet in most host implementations; without a
          // sheet-aware `getRange` implementation we cannot safely service cross-sheet reads.
          if (sheetName != null && sheetId !== activeSheetId) {
            throw new Error(
              `Sheet-qualified ranges require spreadsheetApi.getRange (cannot read '${String(
                sheetName
              )}' while active sheet is '${activeSheetId}')`
            );
          }
          result = await this._defaultGetRange(a1Ref);
        }

        if (result && typeof result === "object") {
          // Prefer tainting based on the returned A1 `address` when available. This covers
          // named range lookups and lets sheet-qualified addresses override the sheet inferred
          // from the input ref.
          let rangeSheetId = sheetId;
          let coordsFromAddress = null;
          try {
            const addr = typeof result.address === "string" ? result.address.trim() : "";
            if (addr) {
              const { sheetName: addrSheetName, ref: addrRef } = this._splitSheetQualifier(addr);
              if (addrSheetName != null) {
                try {
                  const addrSheetId = await this._resolveSheetId(addrSheetName);
                  if (typeof addrSheetId === "string" && addrSheetId.trim()) {
                    rangeSheetId = addrSheetId.trim();
                  }
                } catch {
                  // ignore
                }
              }
              try {
                coordsFromAddress = this._parseA1RangeRef(addrRef);
              } catch {
                coordsFromAddress = null;
              }
            }
          } catch {
            coordsFromAddress = null;
          }

          if (coordsFromAddress) {
            this._taintExtensionRange(extension, { sheetId: rangeSheetId, ...coordsFromAddress });
          } else {
            const direct = normalizeTaintedRange({
              sheetId: rangeSheetId,
              startRow: result.startRow,
              startCol: result.startCol,
              endRow: result.endRow,
              endCol: result.endCol
            });
            if (direct) {
              this._taintExtensionRange(extension, direct);
            } else {
              try {
                const parsed = this._parseA1RangeRef(a1Ref);
                this._taintExtensionRange(extension, { sheetId: rangeSheetId, ...parsed });
              } catch {
                // ignore
              }
            }
          }
        }

        return result;
      }
      case "cells.getCell": {
        const row = normalizeNonNegativeInt(args[0], { label: "row" });
        const col = normalizeNonNegativeInt(args[1], { label: "col" });
        const sheetId = await this._resolveActiveSheetId();
        const value = await this._spreadsheet.getCell(row, col);
        this._taintExtensionRange(extension, { sheetId, startRow: row, startCol: col, endRow: row, endCol: col });
        return value;
      }
      case "cells.setCell":
        await this._spreadsheet.setCell(args[0], args[1], args[2]);
        return null;
      case "cells.setRange":
        // Fail fast on unbounded A1 ranges so we don't ask the spreadsheet implementation to
        // apply multi-million-cell writes.
        {
          let coords = null;
          try {
            const { ref: a1Ref } = this._splitSheetQualifier(args[0]);
            coords = this._parseA1RangeRef(a1Ref);
          } catch {
            coords = null;
          }
          if (coords) {
            assertExtensionRangeWithinLimits(coords, { label: "Range" });
          }
        }
        if (typeof this._spreadsheet.setRange === "function") {
          await this._spreadsheet.setRange(args[0], args[1]);
          return null;
        }
        await this._defaultSetRange(args[0], args[1]);
        return null;

      case "sheets.getSheet":
        if (typeof this._spreadsheet.getSheet === "function") {
          return this._spreadsheet.getSheet(args[0]);
        }
        return this._defaultGetSheet(args[0]);
      case "sheets.activateSheet":
        if (typeof this._spreadsheet.activateSheet === "function") {
          const sheet = await this._spreadsheet.activateSheet(args[0]);
          const id = sheet && typeof sheet === "object" && sheet.id != null ? String(sheet.id).trim() : "";
          if (id) this._activeSheetId = id;
          return sheet;
        }
        return this._defaultActivateSheet(args[0]);
      case "sheets.createSheet":
        if (typeof this._spreadsheet.createSheet === "function") {
          const sheet = await this._spreadsheet.createSheet(args[0]);
          const id = sheet && typeof sheet === "object" && sheet.id != null ? String(sheet.id).trim() : "";
          if (id) this._activeSheetId = id;
          return sheet;
        }
        return this._defaultCreateSheet(args[0]);
      case "sheets.renameSheet":
        if (typeof this._spreadsheet.renameSheet === "function") {
          await this._spreadsheet.renameSheet(args[0], args[1]);
          return null;
        }
        this._defaultRenameSheet(args[0], args[1]);
        return null;
      case "sheets.deleteSheet":
        if (typeof this._spreadsheet.deleteSheet === "function") {
          await this._spreadsheet.deleteSheet(args[0]);
          return null;
        }
        this._defaultDeleteSheet(args[0]);
        return null;

      case "commands.registerCommand":
        if (this._commands.has(args[0]) && this._commands.get(args[0]) !== extension.id) {
          throw new Error(`Command already registered by another extension: ${args[0]}`);
        }
        this._commands.set(args[0], extension.id);
        extension.registeredCommands.add(args[0]);
        return null;
      case "commands.unregisterCommand":
        if (this._commands.get(args[0]) === extension.id) this._commands.delete(args[0]);
        extension.registeredCommands.delete(args[0]);
        return null;
      case "commands.executeCommand": {
        const [commandId, ...rest] = args;
        return this.executeCommand(String(commandId), ...rest);
      }

      case "ui.showMessage":
        this._messages.push({ message: args[0], type: args[1] });
        try {
          await this._uiApi?.showMessage?.(args[0], args[1]);
        } catch {
          // ignore UI errors
        }
        return null;

      case "ui.showInputBox": {
        const options = args[0] ?? {};
        if (this._uiApi?.showInputBox) {
          try {
            return await this._uiApi.showInputBox(options);
          } catch {
            return null;
          }
        }
        // Back-compat placeholder implementation (tests may rely on this fallback).
        const value = options && typeof options === "object" ? options.value : undefined;
        return typeof value === "string" ? value : null;
      }

      case "ui.showQuickPick": {
        const items = Array.isArray(args[0]) ? args[0] : [];
        const options = args[1] ?? {};
        if (this._uiApi?.showQuickPick) {
          try {
            return await this._uiApi.showQuickPick(items, options);
          } catch {
            return null;
          }
        }
        // Back-compat placeholder implementation: choose the first entry.
        if (items.length === 0) return null;
        const first = items[0];
        if (first && typeof first === "object" && Object.prototype.hasOwnProperty.call(first, "value")) {
          return first.value;
        }
        return first;
      }

      case "ui.registerContextMenu": {
        const menuIdRaw = String(args[0]);
        const menuId = menuIdRaw.trim();
        const items = Array.isArray(args[1]) ? args[1] : [];
        if (menuId.trim().length === 0) throw new Error("Menu id must be a non-empty string");

        const normalized = [];
        for (const [idx, item] of items.entries()) {
          if (!item || typeof item !== "object") {
            throw new Error(`Menu item at index ${idx} must be an object`);
          }
          const commandRaw = String(item.command ?? "");
          const command = commandRaw.trim();
          if (command.length === 0) {
            throw new Error(`Menu item at index ${idx} must include a non-empty command`);
          }
          const when = item.when === undefined ? null : item.when === null ? null : String(item.when);
          const group = item.group === undefined ? null : item.group === null ? null : String(item.group);
          normalized.push({ command, when, group });
        }

        const id = createRequestId();
        this._contextMenus.set(id, { id, extensionId: extension.id, menuId, items: normalized });
        return { id };
      }

      case "ui.unregisterContextMenu": {
        const id = String(args[0]);
        const record = this._contextMenus.get(id);
        if (record && record.extensionId === extension.id) {
          this._contextMenus.delete(id);
        }
        return null;
      }

      case "ui.createPanel": {
        const panelId = String(args[0]);
        const options = args[1] ?? {};
        const title = String(options.title ?? panelId);
        const existing = this._panels.get(panelId);
        if (existing && existing.extensionId !== extension.id) {
          throw new Error(`Panel already created by another extension: ${panelId}`);
        }
        const panel = existing ?? {
          id: panelId,
          title,
          icon: options.icon ?? null,
          position: options.position ?? null,
          html: "",
          extensionId: extension.id,
          outgoingMessages: []
        };
        panel.title = title;
        panel.icon = options.icon ?? panel.icon ?? null;
        panel.position = options.position ?? panel.position ?? null;
        this._panels.set(panelId, panel);
        try {
          this._uiApi?.onPanelCreated?.(panel);
        } catch {
          // ignore
        }
        return { id: panelId };
      }
      case "ui.setPanelHtml": {
        const panelId = String(args[0]);
        const html = String(args[1]);
        const panel = this._panels.get(panelId);
        if (!panel) throw new Error(`Unknown panel: ${panelId}`);
        if (panel.extensionId !== extension.id) {
          throw new Error(`Panel ${panelId} does not belong to extension ${extension.id}`);
        }
        panel.html = html;
        try {
          this._uiApi?.onPanelHtmlUpdated?.(panelId, html);
        } catch {
          // ignore
        }
        return null;
      }
      case "ui.postMessageToPanel": {
        const panelId = String(args[0]);
        const message = args[1];
        const panel = this._panels.get(panelId);
        if (!panel) throw new Error(`Unknown panel: ${panelId}`);
        if (panel.extensionId !== extension.id) {
          throw new Error(`Panel ${panelId} does not belong to extension ${extension.id}`);
        }
        panel.outgoingMessages.push(message);
        try {
          this._uiApi?.onPanelMessage?.(panelId, message);
        } catch {
          // ignore
        }
        return null;
      }
      case "ui.disposePanel": {
        const panelId = String(args[0]);
        const panel = this._panels.get(panelId);
        if (!panel) return null;
        if (panel.extensionId !== extension.id) {
          throw new Error(`Panel ${panelId} does not belong to extension ${extension.id}`);
        }
        this._panels.delete(panelId);
        try {
          this._uiApi?.onPanelDisposed?.(panelId);
        } catch {
          // ignore
        }
        return null;
      }

      case "functions.register": {
        const fnName = String(args[0]);
        const contributed = (extension.manifest.contributes.customFunctions ?? []).some(
          (f) => f.name === fnName
        );
        if (!contributed) {
          throw new Error(`Custom function not declared in manifest: ${fnName}`);
        }
        this._customFunctions.set(fnName, extension.id);
        return null;
      }
      case "functions.unregister": {
        const fnName = String(args[0]);
        if (this._customFunctions.get(fnName) === extension.id) this._customFunctions.delete(fnName);
        return null;
      }

      case "dataConnectors.register": {
        const connectorId = String(args[0]);
        const contributed = (extension.manifest.contributes.dataConnectors ?? []).some(
          (c) => c.id === connectorId
        );
        if (!contributed) {
          throw new Error(`Data connector not declared in manifest: ${connectorId}`);
        }
        if (this._dataConnectors.has(connectorId) && this._dataConnectors.get(connectorId) !== extension.id) {
          throw new Error(`Data connector already registered by another extension: ${connectorId}`);
        }
        this._dataConnectors.set(connectorId, extension.id);
        return null;
      }
      case "dataConnectors.unregister": {
        const connectorId = String(args[0]);
        if (this._dataConnectors.get(connectorId) === extension.id) this._dataConnectors.delete(connectorId);
        return null;
      }

      case "network.fetch": {
        const rawUrl = String(args[0]);
        const init = args[1];

        // When running inside the Tauri desktop shell, prefer proxying outbound
        // HTTP(S) requests through the Rust backend via a Tauri command. This:
        //
        // - avoids relying on CORS headers for the `tauri://…` origin
        // - keeps networking behavior consistent/hardened across WebViews
        //
        // Note: extension network access is enforced by Formula's permission
        // model + worker guardrails (which replace `fetch`/`WebSocket` inside the
        // extension worker). CSP is defense-in-depth, not the primary enforcement
        // mechanism here.
        /** @type {any} */
        let tauriInvoke = null;
        /** @type {any} */
        let tauriInvokeOwner = null;
        const tauri = safeGetProp(globalThis, "__TAURI__") ?? null;
        const core = safeGetProp(tauri, "core");
        const coreInvoke = safeGetProp(core, "invoke");
        if (typeof coreInvoke === "function") {
          tauriInvoke = coreInvoke;
          tauriInvokeOwner = core;
        } else {
          const legacyInvoke = safeGetProp(tauri, "invoke");
          if (typeof legacyInvoke === "function") {
            tauriInvoke = legacyInvoke;
            tauriInvokeOwner = tauri;
          }
        }
        if (typeof tauriInvoke === "function") {
          const resolved = safeParseUrl(rawUrl);
          const url = resolved ? resolved.toString() : rawUrl;
          return tauriInvoke.call(tauriInvokeOwner, "network_fetch", { url, init: init ?? null });
        }

        if (typeof fetch !== "function") {
          throw new Error("Network fetch is not available in this runtime");
        }

        const response = await fetch(rawUrl, init);
        const bodyText = await response.text();
        const headers = Array.from(response.headers.entries());

        return {
          ok: response.ok,
          status: response.status,
          statusText: response.statusText,
          url: response.url,
          headers,
          bodyText
        };
      }

      case "network.openWebSocket":
        // WebSocket connections are initiated directly inside the worker after the host
        // confirms the extension has been granted the `network` permission.
        return null;

      case "clipboard.readText":
        return this._clipboardApi.readText();
      case "clipboard.writeText": {
        if (this._clipboardWriteGuard) {
          const taintedRanges = Array.isArray(extension?.taintedRanges) ? extension.taintedRanges : [];
          await this._clipboardWriteGuard({ extensionId: extension.id, taintedRanges: taintedRanges.map((r) => ({ ...r })) });
        }
        await this._clipboardApi.writeText(String(args[0] ?? ""));
        return null;
      }

      case "storage.get": {
        const store = this._storageApi.getExtensionStore(extension.id);
        const key = normalizeStorageKey(args[0]);
        return store[key];
      }
      case "storage.set": {
        const key = normalizeStorageKey(args[0]);
        const value = args[1];
        const store = this._storageApi.getExtensionStore(extension.id);
        store[key] = value;
        return null;
      }
      case "storage.delete": {
        const key = normalizeStorageKey(args[0]);
        const store = this._storageApi.getExtensionStore(extension.id);
        delete store[key];
        return null;
      }

      case "config.get": {
        const key = String(args[0]);
        const store = this._storageApi.getExtensionStore(extension.id);
        const stored = store[`__config__:${key}`];
        if (stored !== undefined) return stored;

        const schema = extension.manifest.contributes.configuration;
        const defaultValue = schema?.properties?.[key]?.default;
        if (defaultValue !== undefined) return defaultValue;
        return undefined;
      }
      case "config.update": {
        const configKey = String(args[0]);
        const schema = extension.manifest.contributes.configuration;
        if (!schema?.properties || !Object.prototype.hasOwnProperty.call(schema.properties, configKey)) {
          throw new Error(`Configuration key not declared in manifest: ${configKey}`);
        }

        const key = `__config__:${configKey}`;
        const value = args[1];
        const store = this._storageApi.getExtensionStore(extension.id);
        store[key] = value;
        this._sendEventToExtension(extension, "configChanged", { key: configKey, value });
        return null;
      }

      default:
        throw new Error(`Unknown API method: ${namespace}.${method}`);
    }
  }

  _splitSheetQualifier(input) {
    const s = String(input ?? "").trim();

    const quoted = s.match(/^'((?:[^']|'')+)'!(.+)$/);
    if (quoted) {
      return {
        sheetName: quoted[1].replace(/''/g, "'"),
        ref: quoted[2]
      };
    }

    const unquoted = s.match(/^([^!]+)!(.+)$/);
    if (unquoted) {
      return { sheetName: unquoted[1], ref: unquoted[2] };
    }

    return { sheetName: null, ref: s };
  }

  async _resolveActiveSheetId() {
    if (typeof this._spreadsheet.getActiveSheet === "function") {
      try {
        const sheet = await this._spreadsheet.getActiveSheet();
        const id =
          sheet && typeof sheet === "object" && sheet.id != null ? String(sheet.id).trim() : "";
        if (id) {
          this._activeSheetId = id;
          return id;
        }
      } catch {
        // ignore
      }
    }

    const fallback = typeof this._activeSheetId === "string" && this._activeSheetId ? this._activeSheetId : null;
    if (fallback) return fallback;

    const sheet = Array.isArray(this._sheets) ? this._sheets[0] : null;
    const sheetId = sheet && typeof sheet?.id === "string" ? sheet.id : null;
    return sheetId ?? "sheet1";
  }

  async _resolveSheetId(sheetName) {
    if (sheetName == null) return this._resolveActiveSheetId();

    const name = String(sheetName).trim();
    if (!name) return this._resolveActiveSheetId();

    if (typeof this._spreadsheet.getSheet === "function") {
      try {
        const sheet = await this._spreadsheet.getSheet(name);
        const id =
          sheet && typeof sheet === "object" && sheet.id != null ? String(sheet.id).trim() : "";
        if (id) return id;
      } catch {
        // ignore
      }
    }

    const candidateLists = [
      typeof this._spreadsheet.listSheets === "function" ? this._spreadsheet.listSheets() : null,
      Array.isArray(this._sheets) ? this._sheets : null
    ];

    for (const list of candidateLists) {
      if (!Array.isArray(list)) continue;
      for (const sheet of list) {
        if (!sheet || typeof sheet !== "object") continue;
        const id = sheet.id != null ? String(sheet.id).trim() : "";
        const sheetDisplayName = sheet.name != null ? String(sheet.name).trim() : "";
        if (!id) continue;
        if (sheetDisplayName === name || id === name) return id;
      }
    }

    return name;
  }

  _resolveSheetIdSync(sheetName) {
    if (sheetName == null) return this._activeSheetId ?? "sheet1";
    const name = String(sheetName).trim();
    if (!name) return this._activeSheetId ?? "sheet1";

    const candidateLists = [
      typeof this._spreadsheet.listSheets === "function" ? this._spreadsheet.listSheets() : null,
      Array.isArray(this._sheets) ? this._sheets : null
    ];

    for (const list of candidateLists) {
      if (!Array.isArray(list)) continue;
      for (const sheet of list) {
        if (!sheet || typeof sheet !== "object") continue;
        const id = sheet.id != null ? String(sheet.id).trim() : "";
        const sheetDisplayName = sheet.name != null ? String(sheet.name).trim() : "";
        if (!id) continue;
        if (sheetDisplayName === name || id === name) return id;
      }
    }

    return name;
  }

  _taintExtensionRange(extension, range) {
    try {
      extension.taintedRanges = addTaintedRangeToList(extension.taintedRanges, range);
    } catch {
      // ignore - taint tracking should never interfere with extension execution
    }
  }

  _parseA1CellRef(ref) {
    const match = /^\s*\$?([A-Za-z]+)\$?(\d+)\s*$/.exec(String(ref ?? ""));
    if (!match) throw new Error(`Invalid A1 cell reference: ${ref}`);
    const [, colLetters, rowDigits] = match;
    const row = Number.parseInt(rowDigits, 10) - 1;
    if (!Number.isFinite(row) || row < 0) throw new Error(`Invalid row in A1 reference: ${ref}`);

    const cleaned = colLetters.toUpperCase();
    let col = 0;
    for (const ch of cleaned) {
      col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    col -= 1;
    if (col < 0) throw new Error(`Invalid column in A1 reference: ${ref}`);
    return { row, col };
  }

  _parseA1RangeRef(ref) {
    const raw = String(ref ?? "");
    const parts = raw.split(":");
    if (parts.length > 2) throw new Error(`Invalid A1 range reference: ${ref}`);
    const start = this._parseA1CellRef(parts[0]);
    const end = parts.length === 2 ? this._parseA1CellRef(parts[1]) : start;
    return {
      startRow: Math.min(start.row, end.row),
      startCol: Math.min(start.col, end.col),
      endRow: Math.max(start.row, end.row),
      endCol: Math.max(start.col, end.col)
    };
  }

  async _defaultGetRange(ref) {
    const { startRow, startCol, endRow, endCol } = this._parseA1RangeRef(ref);
    assertExtensionRangeWithinLimits({ startRow, startCol, endRow, endCol }, { label: "Range" });
    const values = [];
    for (let r = startRow; r <= endRow; r++) {
      const row = [];
      for (let c = startCol; c <= endCol; c++) {
        // eslint-disable-next-line no-await-in-loop
        row.push(await this._spreadsheet.getCell(r, c));
      }
      values.push(row);
    }
    return { startRow, startCol, endRow, endCol, values };
  }

  async _defaultSetRange(ref, values) {
    const { startRow, startCol, endRow, endCol } = this._parseA1RangeRef(ref);
    assertExtensionRangeWithinLimits({ startRow, startCol, endRow, endCol }, { label: "Range" });
    const expectedRows = endRow - startRow + 1;
    const expectedCols = endCol - startCol + 1;

    if (!Array.isArray(values) || values.length !== expectedRows) {
      throw new Error(
        `Range values must be a ${expectedRows}x${expectedCols} array (got ${Array.isArray(values) ? values.length : 0} rows)`
      );
    }

    for (let r = 0; r < expectedRows; r++) {
      const rowValues = values[r];
      if (!Array.isArray(rowValues) || rowValues.length !== expectedCols) {
        throw new Error(
          `Range values must be a ${expectedRows}x${expectedCols} array (row ${r} has ${Array.isArray(rowValues) ? rowValues.length : 0} cols)`
        );
      }
      for (let c = 0; c < expectedCols; c++) {
        // eslint-disable-next-line no-await-in-loop
        await this._spreadsheet.setCell(startRow + r, startCol + c, rowValues[c]);
      }
    }
  }

  _defaultGetSheet(name) {
    const sheets = Array.isArray(this._sheets) ? this._sheets : [];
    const sheet = sheets.find((s) => s.name === String(name));
    if (!sheet) return undefined;
    return { id: sheet.id, name: sheet.name };
  }

  _defaultCreateSheet(name) {
    const sheetName = String(name);
    if (sheetName.trim().length === 0) throw new Error("Sheet name must be a non-empty string");
    this._sheets = this._sheets ?? [{ id: "sheet1", name: "Sheet1" }];
    this._nextSheetId = this._nextSheetId ?? 2;
    if (this._sheets.some((s) => s.name === sheetName)) {
      throw new Error(`Sheet already exists: ${sheetName}`);
    }

    const sheet = { id: `sheet${this._nextSheetId++}`, name: sheetName };
    this._sheets.push(sheet);
    this._activeSheetId = sheet.id;
    this._broadcastEvent("sheetActivated", { sheet: { id: sheet.id, name: sheet.name } });
    return { id: sheet.id, name: sheet.name };
  }

  _defaultRenameSheet(oldName, newName) {
    const from = String(oldName);
    const to = String(newName);
    if (to.trim().length === 0) throw new Error("New sheet name must be a non-empty string");
    this._sheets = this._sheets ?? [{ id: "sheet1", name: "Sheet1" }];
    if (this._sheets.some((s) => s.name === to)) {
      throw new Error(`Sheet already exists: ${to}`);
    }
    const sheet = this._sheets.find((s) => s.name === from);
    if (!sheet) throw new Error(`Unknown sheet: ${from}`);
    sheet.name = to;
  }

  _defaultActivateSheet(name) {
    const sheetName = String(name);
    const sheet = this._sheets.find((s) => s.name === sheetName);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
    if (sheet.id === this._activeSheetId) {
      return { id: sheet.id, name: sheet.name };
    }
    this._activeSheetId = sheet.id;
    this._broadcastEvent("sheetActivated", { sheet: { id: sheet.id, name: sheet.name } });
    return { id: sheet.id, name: sheet.name };
  }

  _defaultDeleteSheet(name) {
    const sheetName = String(name);
    const idx = this._sheets.findIndex((s) => s.name === sheetName);
    if (idx === -1) throw new Error(`Unknown sheet: ${sheetName}`);
    if (this._sheets.length === 1) {
      throw new Error("Cannot delete the last remaining sheet");
    }

    const sheet = this._sheets[idx];
    const wasActive = sheet.id === this._activeSheetId;
    this._sheets.splice(idx, 1);

    if (wasActive) {
      this._activeSheetId = this._sheets[0].id;
      const active = this._sheets.find((s) => s.id === this._activeSheetId) ?? this._sheets[0];
      this._broadcastEvent("sheetActivated", { sheet: { id: active.id, name: active.name } });
    }
  }

  _handleWorkerMessage(extension, worker, message) {
    if (worker !== extension.worker) return;
    if (!message || typeof message !== "object") return;

    switch (message.type) {
      case "api_call":
        this._handleApiCall(extension, message);
        return;
      case "activate_result": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.resolve(true);
        return;
      }
      case "activate_error": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.reject(deserializeError(message.error));
        return;
      }
      case "command_result": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.resolve(message.result);
        return;
      }
      case "command_error": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.reject(deserializeError(message.error));
        return;
      }
      case "custom_function_result": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.resolve(message.result);
        return;
      }
      case "custom_function_error": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.reject(deserializeError(message.error));
        return;
      }
      case "data_connector_result": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.resolve(message.result);
        return;
      }
      case "data_connector_error": {
        const pending = extension.pendingRequests.get(message.id);
        if (!pending) return;
        extension.pendingRequests.delete(message.id);
        if (pending.timeout) clearTimeout(pending.timeout);
        pending.reject(deserializeError(message.error));
        return;
      }
      case "log":
        return;
      default:
        return;
    }
  }

  _handleWorkerCrash(extension, worker, error) {
    if (worker !== extension.worker) return;
    this._terminateWorker(extension, { reason: "crash", cause: error });
  }

  async _activateExtension(extension, reason) {
    if (extension.active) return;
    await this._ensureWorker(extension);
    await this._requestFromWorker(extension, { type: "activate", reason }, {
      timeoutMs: this._activationTimeoutMs,
      operation: `activation (${reason})`
    });
    extension.active = true;
  }

  _maybeTaintEventPayload(extension, event, data) {
    try {
      if (!extension || typeof extension !== "object") return;
      const evt = String(event ?? "");
      if (!evt) return;

      if (evt === "selectionChanged") {
        const selection = data?.selection;
        if (!selection || typeof selection !== "object") return;

        // Only taint when the event payload actually includes spreadsheet data. Some hosts may
        // emit selectionChanged events with empty matrices (e.g. large selections) to avoid
        // catastrophic allocations; those payloads do not expose cell values/formulas and should
        // not contribute to clipboard DLP enforcement.
        let valuesBounds = null;
        let formulasBounds = null;
        try {
          if (Object.prototype.hasOwnProperty.call(selection, "values")) {
            const values = selection.values;
            if (Array.isArray(values)) {
              let maxRow = -1;
              let maxCol = -1;
              for (let r = 0; r < values.length; r++) {
                const row = values[r];
                if (!Array.isArray(row) || row.length === 0) continue;
                maxRow = Math.max(maxRow, r);
                maxCol = Math.max(maxCol, row.length - 1);
              }
              if (maxRow >= 0 && maxCol >= 0) {
                valuesBounds = { rows: maxRow + 1, cols: maxCol + 1 };
              }
            }
          }
        } catch {
          // ignore
        }

        try {
          if (Object.prototype.hasOwnProperty.call(selection, "formulas")) {
            const formulas = selection.formulas;
            if (Array.isArray(formulas)) {
              let maxRow = -1;
              let maxCol = -1;
              for (let r = 0; r < formulas.length; r++) {
                const row = formulas[r];
                if (!Array.isArray(row) || row.length === 0) continue;
                for (let c = 0; c < row.length; c++) {
                  const cell = row[c];
                  if (typeof cell !== "string" || cell.trim().length === 0) continue;
                  maxRow = Math.max(maxRow, r);
                  maxCol = Math.max(maxCol, c);
                }
              }
              if (maxRow >= 0 && maxCol >= 0) {
                formulasBounds = { rows: maxRow + 1, cols: maxCol + 1 };
              }
            }
          }
        } catch {
          // ignore
        }

        if (!valuesBounds && !formulasBounds) return;

        const looksLikeRange =
          Object.prototype.hasOwnProperty.call(selection, "startRow") &&
          Object.prototype.hasOwnProperty.call(selection, "startCol") &&
          Object.prototype.hasOwnProperty.call(selection, "endRow") &&
          Object.prototype.hasOwnProperty.call(selection, "endCol");

        let range = null;
        if (looksLikeRange) {
          range = {
            startRow: selection.startRow,
            startCol: selection.startCol,
            endRow: selection.endRow,
            endCol: selection.endCol
          };
        } else if (typeof selection?.address === "string" && selection.address.trim()) {
          try {
            const { sheetName: addrSheetName, ref: addrRef } = this._splitSheetQualifier(selection.address);
            range = this._parseA1RangeRef(addrRef);
            if (addrSheetName != null) {
              const resolved = this._resolveSheetIdSync(addrSheetName);
              if (typeof resolved === "string" && resolved.trim()) {
                // eslint-disable-next-line no-param-reassign
                data = { ...(data ?? {}), sheetId: resolved.trim() };
              }
            }
          } catch {
            range = null;
          }
        }
        if (!range) return;

        // Some hosts may include truncated/partial matrices that do not fully match the declared
        // selection range. Align taint tracking with the data that was actually delivered by
        // clamping the range to the matrix bounds (best-effort).
        let sr = null;
        let sc = null;
        let er = null;
        let ec = null;
        try {
          if (
            Number.isFinite(Number(range.startRow)) &&
            Number.isFinite(Number(range.startCol)) &&
            Number.isFinite(Number(range.endRow)) &&
            Number.isFinite(Number(range.endCol))
          ) {
            sr = Math.min(Number(range.startRow), Number(range.endRow));
            sc = Math.min(Number(range.startCol), Number(range.endCol));
            er = Math.max(Number(range.startRow), Number(range.endRow));
            ec = Math.max(Number(range.startCol), Number(range.endCol));
          }
        } catch {
          // ignore
        }
        if (sr == null || sc == null || er == null || ec == null) return;

        /** @type {Array<{ startRow: number, startCol: number, endRow: number, endCol: number }>} */
        const deliveredRanges = [];
        try {
          const seen = new Set();
          for (const bounds of [valuesBounds, formulasBounds]) {
            if (!bounds) continue;
            const boundedEndRow = Math.min(er, sr + bounds.rows - 1);
            const boundedEndCol = Math.min(ec, sc + bounds.cols - 1);
            const key = `${sr},${sc},${boundedEndRow},${boundedEndCol}`;
            if (seen.has(key)) continue;
            seen.add(key);
            deliveredRanges.push({ startRow: sr, startCol: sc, endRow: boundedEndRow, endCol: boundedEndCol });
          }
        } catch {
          // ignore
        }
        if (deliveredRanges.length === 0) return;

        const sheetId =
          (typeof data?.sheetId === "string" && data.sheetId.trim()) ||
          (typeof selection?.sheetId === "string" && selection.sheetId.trim()) ||
          (typeof this._activeSheetId === "string" && this._activeSheetId.trim()) ||
          null;
        if (!sheetId) return;
        this._activeSheetId = sheetId;

        for (const delivered of deliveredRanges) {
          this._taintExtensionRange(extension, {
            sheetId,
            startRow: delivered.startRow,
            startCol: delivered.startCol,
            endRow: delivered.endRow,
            endCol: delivered.endCol
          });
        }
        return;
      }

      if (evt === "cellChanged") {
        if (!data || typeof data !== "object") return;
        // Only taint when the payload includes a cell value (even if null). If a host emits
        // coordinate-only cellChanged events, those do not expose data and should not taint.
        if (!Object.prototype.hasOwnProperty.call(data, "value")) return;
        if (data.value === undefined) return;
        const hasRow = Object.prototype.hasOwnProperty.call(data, "row");
        const hasCol = Object.prototype.hasOwnProperty.call(data, "col");

        let row = hasRow ? data.row : null;
        let col = hasCol ? data.col : null;

        if ((row == null || col == null) && typeof data?.address === "string" && data.address.trim()) {
          try {
            const { sheetName: addrSheetName, ref: addrRef } = this._splitSheetQualifier(data.address);
            const parsed = this._parseA1CellRef(addrRef);
            row = parsed.row;
            col = parsed.col;
            if (addrSheetName != null) {
              const resolved = this._resolveSheetIdSync(addrSheetName);
              if (typeof resolved === "string" && resolved.trim()) {
                // eslint-disable-next-line no-param-reassign
                data = { ...(data ?? {}), sheetId: resolved.trim() };
              }
            }
          } catch {
            // ignore
          }
        }

        if (row == null || col == null) return;

        const sheetId =
          (typeof data?.sheetId === "string" && data.sheetId.trim()) ||
          (typeof this._activeSheetId === "string" && this._activeSheetId.trim()) ||
          null;
        if (!sheetId) return;
        this._activeSheetId = sheetId;

        this._taintExtensionRange(extension, {
          sheetId,
          startRow: row,
          startCol: col,
          endRow: row,
          endCol: col
        });
      }
    } catch {
      // ignore - event-based taint tracking must never interfere with extension execution
    }
  }

  _syncHostStateFromEvent(event, data) {
    // Best-effort: keep internal book/sheet metadata in sync based on incoming events.
    try {
      const evt = String(event ?? "");
      if (!evt) return;

      if (evt === "sheetActivated") {
        const id = data?.sheet?.id;
        if (typeof id === "string" && id.trim()) {
          this._activeSheetId = id.trim();
        }
        return;
      }

      if (evt === "workbookOpened" || evt === "beforeSave") {
        const wb = data?.workbook;
        if (!wb || typeof wb !== "object") return;

        const next = { ...this._workbook };
        let nameSet = false;

        // Merge name/path using the same semantics as `_getActiveWorkbook`.
        try {
          if (Object.prototype.hasOwnProperty.call(wb, "name")) {
            const rawName = wb.name;
            const trimmed = rawName == null ? "" : String(rawName).trim();
            if (trimmed) {
              next.name = trimmed;
              nameSet = true;
            }
          }
        } catch {
          // ignore
        }

        try {
          if (Object.prototype.hasOwnProperty.call(wb, "path")) {
            const rawPath = wb.path;
            if (rawPath === undefined) {
              // keep previous
            } else if (rawPath == null) {
              next.path = null;
            } else {
              const str = String(rawPath);
              const trimmed = str.trim();
              next.path = trimmed.length > 0 ? str : null;
              if (!nameSet && trimmed.length > 0) {
                next.name = trimmed.split(/[/\\]/).pop() ?? trimmed;
                nameSet = true;
              }
            }
          }
        } catch {
          // ignore
        }

        // Best-effort: update internal sheet metadata from the workbook snapshot so
        // `getActiveWorkbook` stays accurate even when the host doesn't implement
        // `listSheets` / `getActiveSheet`.
        try {
          if (Object.prototype.hasOwnProperty.call(wb, "sheets") && Array.isArray(wb.sheets)) {
            const normalized = wb.sheets
              .map((sheet) => {
                if (!sheet || typeof sheet !== "object") return null;
                const id = typeof sheet.id === "string" ? sheet.id.trim() : String(sheet.id ?? "").trim();
                const name = typeof sheet.name === "string" ? sheet.name.trim() : String(sheet.name ?? "").trim();
                if (!id || !name) return null;
                return { id, name };
              })
              .filter(Boolean);
            if (normalized.length > 0) {
              this._sheets = normalized;
            }
          }
        } catch {
          // ignore
        }

        try {
          if (
            Object.prototype.hasOwnProperty.call(wb, "activeSheet") &&
            wb.activeSheet &&
            typeof wb.activeSheet === "object"
          ) {
            const idRaw = wb.activeSheet.id;
            const id = typeof idRaw === "string" ? idRaw.trim() : String(idRaw ?? "").trim();
            if (id) this._activeSheetId = id;
          }
        } catch {
          // ignore
        }

        this._workbook = { name: next.name, path: next.path ?? null };
      }
    } catch {
      // ignore
    }
  }

  _sanitizeEventPayload(event, data) {
    try {
      const evt = String(event ?? "");
      if (evt !== "selectionChanged") return data;
      if (!data || typeof data !== "object") return data;
      const selection = data.selection;
      if (!selection || typeof selection !== "object") return data;

      // Prefer numeric coords, but fall back to parsing A1 address when needed.
      let coords = null;
      const looksLikeRange =
        Object.prototype.hasOwnProperty.call(selection, "startRow") &&
        Object.prototype.hasOwnProperty.call(selection, "startCol") &&
        Object.prototype.hasOwnProperty.call(selection, "endRow") &&
        Object.prototype.hasOwnProperty.call(selection, "endCol");
      if (looksLikeRange) {
        coords = {
          startRow: selection.startRow,
          startCol: selection.startCol,
          endRow: selection.endRow,
          endCol: selection.endCol
        };
      } else if (typeof selection?.address === "string" && selection.address.trim()) {
        try {
          const { ref: addrRef } = this._splitSheetQualifier(selection.address);
          coords = this._parseA1RangeRef(addrRef);
        } catch {
          coords = null;
        }
      }
      if (!coords) return data;

      const size = getRangeSize(coords);
      if (!size || size.cellCount <= DEFAULT_EXTENSION_RANGE_CELL_LIMIT) return data;

      // Mirror desktop and Node-host guardrails: `selectionChanged` payloads may include a full
      // 2D values/formulas matrix. For Excel-scale selections this can allocate hundreds of
      // thousands of cells and OOM the extension worker.
      //
      // If the host explicitly marks the payload as `truncated`, allow small partial matrices
      // to pass through (e.g. a preview cell). Otherwise, strip the matrices and mark the payload
      // as truncated.
      const matrixCellCount = (matrix) => {
        if (!Array.isArray(matrix)) return 0;
        let total = 0;
        for (const row of matrix) {
          if (!Array.isArray(row) || row.length === 0) continue;
          total += row.length;
          if (total > DEFAULT_EXTENSION_RANGE_CELL_LIMIT) return total;
        }
        return total;
      };

      const deliveredValuesCells = matrixCellCount(selection.values);
      const deliveredFormulasCells = matrixCellCount(selection.formulas);
      const deliveredCells = Math.max(deliveredValuesCells, deliveredFormulasCells);

      if (selection.truncated === true && deliveredCells <= DEFAULT_EXTENSION_RANGE_CELL_LIMIT) {
        return data;
      }

      return {
        ...(data ?? {}),
        selection: {
          ...(selection ?? {}),
          values: [],
          formulas: [],
          truncated: true
        }
      };
    } catch {
      return data;
    }
  }

  _sendEventToExtension(extension, event, data, options = {}) {
    try {
      try {
        if (!options?.skipHostSync) {
          this._syncHostStateFromEvent(event, data);
        }
      } catch {
        // ignore
      }

      this._maybeTaintEventPayload(extension, event, data);
      extension.worker.postMessage({ type: "event", event, data });
    } catch {
      // ignore
    }
  }

  _broadcastEvent(event, data) {
    // Best-effort: track the active sheet id from events so that downstream logic
    // (including event-based clipboard taint tracking) can fall back to it when
    // payloads omit `sheetId`.
    try {
      this._syncHostStateFromEvent(event, data);
    } catch {
      // ignore
    }

    let payload = data;
    try {
      payload = this._sanitizeEventPayload(event, data);
    } catch {
      payload = data;
    }

    for (const extension of this._extensions.values()) {
      // Only deliver events to active extensions. Workers are spawned eagerly on load, but
      // the extension entrypoint is not imported/executed until activation. Broadcasting
      // selection/cell events to inactive extensions would taint ranges that the extension
      // never had a chance to read, causing false positives in clipboard DLP enforcement.
      if (!extension.active) continue;
      this._sendEventToExtension(extension, event, payload, { skipHostSync: true });
    }
  }
}

export { BrowserExtensionHost, API_PERMISSIONS };
