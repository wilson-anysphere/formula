const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { Worker } = require("node:worker_threads");
const { pathToFileURL } = require("node:url");

const { validateExtensionManifest } = require("./manifest");
const { PermissionManager } = require("./permission-manager");
const { writeFileAtomic } = require("./write-file-atomic");
const { InMemorySpreadsheet } = require("./spreadsheet-mock");
const { installExtensionFromDirectory, uninstallExtension, listInstalledExtensions } = require("./installer");

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

function columnLettersToIndex(letters) {
  const cleaned = String(letters ?? "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(cleaned)) {
    throw new Error(`Invalid column letters: ${letters}`);
  }

  let index = 0;
  for (const ch of cleaned) {
    index = index * 26 + (ch.charCodeAt(0) - 64); // A=1
  }
  return index - 1; // 0-based
}

function parseA1CellRef(ref) {
  const match = /^\s*\$?([A-Za-z]+)\$?(\d+)\s*$/.exec(String(ref ?? ""));
  if (!match) throw new Error(`Invalid A1 cell reference: ${ref}`);
  const [, colLetters, rowDigits] = match;
  const row = Number.parseInt(rowDigits, 10) - 1; // 0-based
  const col = columnLettersToIndex(colLetters);
  if (!Number.isFinite(row) || row < 0) throw new Error(`Invalid row in A1 reference: ${ref}`);
  return { row, col };
}

function parseSheetPrefix(raw) {
  const str = String(raw ?? "");
  const bang = str.indexOf("!");
  if (bang === -1) return { sheetName: null, a1Ref: str };
  const sheetPart = str.slice(0, bang).trim();
  const rest = str.slice(bang + 1);
  if (sheetPart.length === 0) throw new Error(`Invalid sheet-qualified reference: ${raw}`);
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'") && sheetPart.length >= 2) {
    const unquoted = sheetPart.slice(1, -1).replace(/''/g, "'");
    return { sheetName: unquoted, a1Ref: rest };
  }
  return { sheetName: sheetPart, a1Ref: rest };
}

function parseA1RangeRef(ref) {
  const { sheetName, a1Ref } = parseSheetPrefix(ref);
  const parts = String(a1Ref ?? "").split(":");
  if (parts.length > 2) throw new Error(`Invalid A1 range reference: ${ref}`);
  const start = parseA1CellRef(parts[0]);
  const end = parts.length === 2 ? parseA1CellRef(parts[1]) : start;
  return {
    sheetName,
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col)
  };
}

function sanitizeEventPayload(event, data) {
  const evt = String(event ?? "");
  if (evt !== "selectionChanged") return data;
  if (!data || typeof data !== "object") return data;
  const selection = data.selection;
  if (!selection || typeof selection !== "object") return data;

  let size = getRangeSize(selection);
  if (!size && typeof selection.address === "string" && selection.address.trim()) {
    try {
      size = getRangeSize(parseA1RangeRef(selection.address));
    } catch {
      size = null;
    }
  }
  if (!size || size.cellCount <= DEFAULT_EXTENSION_RANGE_CELL_LIMIT) return data;

  // Mirror desktop and browser-host guardrails: `selectionChanged` payloads may include a full
  // 2D values/formulas matrix. For Excel-scale selections this can allocate hundreds of thousands
  // of cells and OOM the extension worker.
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
}

const STORAGE_PROTO_POLLUTION_KEY = "__proto__";
// Persist `__proto__` under an internal alias so JSON parsing/loading cannot mutate prototypes.
// This preserves round-trip behavior for extensions that (intentionally or accidentally) use the key.
const STORAGE_PROTO_POLLUTION_KEY_ALIAS = "__formula_reserved_key__:__proto__";

function normalizeStorageKey(key) {
  const s = String(key);
  if (s === STORAGE_PROTO_POLLUTION_KEY) return STORAGE_PROTO_POLLUTION_KEY_ALIAS;
  return s;
}

function normalizeExtensionStorageStore(data) {
  const out = Object.create(null);
  let migrated = false;

  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return { store: out, migrated: false };
  }

  try {
    if (Object.getPrototypeOf(data) !== Object.prototype) {
      migrated = true;
    }
  } catch {
    migrated = true;
  }

  for (const [extensionId, record] of Object.entries(data)) {
    if (!record || typeof record !== "object" || Array.isArray(record)) {
      migrated = true;
      continue;
    }

    try {
      if (Object.getPrototypeOf(record) !== Object.prototype) {
        migrated = true;
      }
    } catch {
      migrated = true;
    }

    const normalizedRecord = Object.create(null);
    for (const [key, value] of Object.entries(record)) {
      const normalizedKey = normalizeStorageKey(key);
      if (normalizedKey !== key) migrated = true;
      normalizedRecord[normalizedKey] = value;
    }

    // Avoid persisting noisy empty records.
    if (Object.keys(normalizedRecord).length === 0) {
      migrated = true;
      continue;
    }

    out[extensionId] = normalizedRecord;
  }

  return { store: out, migrated };
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

function isEagainWorkerInitError(error) {
  if (!error || typeof error !== "object") return false;
  const msg = typeof error.message === "string" ? error.message : String(error.message ?? "");
  // Node wraps uv_thread_create failures for Worker threads as `ERR_WORKER_INIT_FAILED` with a
  // message like "EAGAIN". Under heavy CI/agent load this can be transient (system-wide thread
  // exhaustion), so we treat it as retryable.
  return error.code === "ERR_WORKER_INIT_FAILED" && /\bEAGAIN\b/.test(msg);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

class ExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionsStoragePath = path.join(process.cwd(), ".formula", "permissions.json"),
    extensionStoragePath = path.join(process.cwd(), ".formula", "storage.json"),
    auditDbPath = null,
    // Extension worker spawn + VM initialization can be slow under load (especially in CI).
    // Keep a conservative default so legitimate extensions don't hit spurious timeouts.
    activationTimeoutMs = 15000,
    commandTimeoutMs = 5000,
    customFunctionTimeoutMs = 5000,
    dataConnectorTimeoutMs = 5000,
    memoryMb = 256,
    spreadsheet = new InMemorySpreadsheet()
  } = {}) {
    this._engineVersion = engineVersion;
    this._extensions = new Map();
    this._commands = new Map(); // commandId -> extensionId
    this._panels = new Map(); // panelId -> { id, title, html }
    this._contextMenus = new Map(); // registrationId -> { id, extensionId, menuId, items }
    this._customFunctions = new Map(); // functionName -> extensionId
    this._dataConnectors = new Map(); // connectorId -> extensionId
    this._panelHtmlWaiters = new Map(); // panelId -> Set<{ resolve, reject, timeout }>
    this._messages = [];
    this._clipboardText = "";
    this._activationTimeoutMs = Number.isFinite(activationTimeoutMs)
      ? Math.max(0, activationTimeoutMs)
      : 15000;
    this._commandTimeoutMs = Number.isFinite(commandTimeoutMs) ? Math.max(0, commandTimeoutMs) : 5000;
    this._customFunctionTimeoutMs = Number.isFinite(customFunctionTimeoutMs)
      ? Math.max(0, customFunctionTimeoutMs)
      : 5000;
    this._dataConnectorTimeoutMs = Number.isFinite(dataConnectorTimeoutMs)
      ? Math.max(0, dataConnectorTimeoutMs)
      : 5000;

    // Best-effort sandboxing: V8 resource limits prevent a single extension worker from
    // consuming unbounded heap memory, but they do not cover all native/externals.
    // Set to null/0 to disable.
    this._memoryMb = Number.isFinite(memoryMb) ? Math.max(0, memoryMb) : 0;
    this._extensionStoragePath = extensionStoragePath;
    this._extensionDataRoot = path.join(
      path.dirname(path.resolve(extensionStoragePath)),
      "extension-data"
    );
    this._spreadsheet = spreadsheet;
    this._spreadsheetDisposables = [];
    this._workbook = { name: "MockWorkbook", path: null };

    // Security baseline: host-level permissions and audit logging for extension runtime.
    // The security package is ESM; extension-host is CJS, so we load it lazily.
    this._securityModulePromise = null;
    this._securityModule = null;
    this._securityPermissionManager = null;
    this._securityAuditLogger = null;
    this._securityAuditDbPath =
      auditDbPath ?? path.join(path.dirname(path.resolve(permissionsStoragePath)), "audit.sqlite");

    this._permissionManager = new PermissionManager({
      storagePath: permissionsStoragePath,
      prompt: permissionPrompt
    });

    this._workerScriptPath = path.resolve(__dirname, "../worker/extension-worker.js");
    // Load the extension API runtime directly inside the VM sandbox. The public CommonJS entrypoint
    // (`packages/extension-api/index.js`) may `require()` sibling files, but the sandboxed API loader
    // intentionally runs with a locked-down `require` implementation.
    this._apiModulePath = path.resolve(__dirname, "../../extension-api/src/runtime.js");

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
  }

  _waitForPanelHtml(panelId, timeoutMs = 5000) {
    const id = String(panelId);
    const existing = this._panels.get(id);
    if (existing && typeof existing.html === "string" && existing.html.length > 0) {
      return Promise.resolve(existing.html);
    }

    return new Promise((resolve, reject) => {
      const timeout =
        timeoutMs > 0
          ? setTimeout(() => {
              const waiters = this._panelHtmlWaiters.get(id);
              if (waiters) {
                for (const waiter of waiters) {
                  if (waiter.resolve === resolve) {
                    waiters.delete(waiter);
                    break;
                  }
                }
                if (waiters.size === 0) this._panelHtmlWaiters.delete(id);
              }
              reject(new Error(`Timed out waiting for panel HTML: ${id}`));
            }, timeoutMs)
          : null;

      const record = { resolve, reject, timeout };
      let waiters = this._panelHtmlWaiters.get(id);
      if (!waiters) {
        waiters = new Set();
        this._panelHtmlWaiters.set(id, waiters);
      }
      waiters.add(record);
    });
  }

  _resolvePanelHtmlWaiters(panelId) {
    const id = String(panelId);
    const panel = this._panels.get(id);
    if (!panel || typeof panel.html !== "string" || panel.html.length === 0) return;
    const waiters = this._panelHtmlWaiters.get(id);
    if (!waiters) return;
    this._panelHtmlWaiters.delete(id);
    for (const waiter of waiters) {
      if (waiter.timeout) clearTimeout(waiter.timeout);
      waiter.resolve(panel.html);
    }
  }

  _rejectPanelHtmlWaiters(panelId, error) {
    const id = String(panelId);
    const waiters = this._panelHtmlWaiters.get(id);
    if (!waiters) return;
    this._panelHtmlWaiters.delete(id);
    for (const waiter of waiters) {
      if (waiter.timeout) clearTimeout(waiter.timeout);
      waiter.reject(error);
    }
  }

  get spreadsheet() {
    return this._spreadsheet;
  }

  async loadExtension(extensionPath) {
    const manifestPath = path.join(extensionPath, "package.json");
    const raw = await fs.readFile(manifestPath, "utf8");
    const parsed = JSON.parse(raw);
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

    const extensionRoot = path.resolve(extensionPath);
    const globalStoragePath = path.join(this._extensionDataRoot, extensionId, "globalStorage");
    const workspaceStoragePath = path.join(this._extensionDataRoot, extensionId, "workspaceStorage");
    await fs.mkdir(globalStoragePath, { recursive: true });
    await fs.mkdir(workspaceStoragePath, { recursive: true });
    const extensionUri = pathToFileURL(extensionRoot).href;

    const mainPath = path.resolve(extensionRoot, manifest.main);
    if (!mainPath.startsWith(extensionRoot + path.sep)) {
      throw new Error(
        `Extension entrypoint must resolve inside extension folder: '${manifest.main}' resolved to '${mainPath}'`
      );
    }
    const stat = await fs.stat(mainPath).catch((error) => {
      if (error && error.code === "ENOENT") {
        throw new Error(`Extension entrypoint not found: ${manifest.main}`);
      }
      throw error;
    });
    if (!stat.isFile()) {
      throw new Error(`Extension entrypoint is not a file: ${manifest.main}`);
    }

    // Prevent symlink escapes: the entrypoint path may be lexically inside the extension folder
    // but still point elsewhere when resolved via realpath.
    const extensionRootReal = await fs.realpath(extensionRoot).catch(() => extensionRoot);
    const mainPathReal = await fs.realpath(mainPath).catch(() => mainPath);
    const relativeReal = path.relative(extensionRootReal, mainPathReal);
    if (relativeReal === "" || relativeReal === ".." || relativeReal.startsWith(".." + path.sep) || path.isAbsolute(relativeReal)) {
      throw new Error(
        `Extension entrypoint must resolve inside extension folder: '${manifest.main}' resolved to '${mainPath}' (realpath '${mainPathReal}')`
      );
    }
    const extension = {
      id: extensionId,
      path: extensionPath,
      mainPath,
      manifest,
      worker: null,
      active: false,
      registeredCommands: new Set(),
      pendingRequests: new Map(),
      workerSpawn: null,
      workerTermination: null,
      workerData: {
        extensionId,
        extensionPath,
        mainPath,
        apiModulePath: this._apiModulePath,
        extensionUri,
        globalStoragePath,
        workspaceStoragePath
      }
    };

    // Register contributed commands so the app can route executions before activation.
    for (const cmd of manifest.contributes.commands ?? []) {
      this._commands.set(cmd.command, extensionId);
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
    await this._spawnWorker(extension);

    return extensionId;
  }

  async startup() {
    const tasks = [];
    for (const extension of this._extensions.values()) {
      if ((extension.manifest.activationEvents ?? []).includes("onStartupFinished")) {
        tasks.push(this._activateExtension(extension, "onStartupFinished"));
      }
    }
    await Promise.all(tasks);

    // Extensions that activate on startup should receive an initial workbook event.
    this._broadcastEvent("workbookOpened", { workbook: this._getWorkbookSnapshot() });
  }

  async startupExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    if ((extension.manifest.activationEvents ?? []).includes("onStartupFinished") && !extension.active) {
      await this._activateExtension(extension, "onStartupFinished");
      // Mirror the behavior of `startup()` but only for this extension: newly-activated
      // startup extensions should receive the initial workbook event, without re-broadcasting
      // it to every other already-running extension.
      this._sendEventToExtension(extension, "workbookOpened", { workbook: this._getWorkbookSnapshot() });
    }
  }

  _getWorkbookSnapshot() {
    const sheets =
      typeof this._spreadsheet.listSheets === "function" ? this._spreadsheet.listSheets() : [];
    const normalizedSheets = Array.isArray(sheets) ? sheets : [];
    const activeSheet =
      typeof this._spreadsheet.getActiveSheet === "function"
        ? this._spreadsheet.getActiveSheet()
        : normalizedSheets[0] ?? { id: "sheet1", name: "Sheet1" };

    return {
      name: this._workbook.name,
      path: this._workbook.path,
      sheets: normalizedSheets,
      activeSheet
    };
  }

  openWorkbook(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    const name =
      workbookPathStr == null || workbookPathStr.trim().length === 0
        ? "MockWorkbook"
        : path.basename(workbookPathStr);
    this._workbook = { name, path: workbookPathStr };
    const workbook = this._getWorkbookSnapshot();
    this._broadcastEvent("workbookOpened", { workbook });
    return workbook;
  }

  saveWorkbook() {
    // Stub implementation: the application will eventually perform real persistence
    // and may await extensions for edits. For now we just emit a stable event payload.
    this._broadcastEvent("beforeSave", { workbook: this._getWorkbookSnapshot() });
    return true;
  }

  saveWorkbookAs(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    if (workbookPathStr == null || workbookPathStr.trim().length === 0) {
      throw new Error("Workbook path must be a non-empty string");
    }

    const name = path.basename(workbookPathStr);
    this._workbook = { name, path: workbookPathStr };
    this._broadcastEvent("beforeSave", { workbook: this._getWorkbookSnapshot() });
    return true;
  }

  closeWorkbook() {
    this.openWorkbook(null);
    return true;
  }

  async activateView(viewId) {
    const id = String(viewId ?? "");

    // `viewActivated` is exposed as `formula.events.onViewActivated`. Like other workbook/grid
    // events, it should be delivered to any active extensions immediately (even if the view's
    // owning extension later requests permissions during activation).
    this._broadcastEvent("viewActivated", { viewId: id });

    const activationEvent = `onView:${id}`;
    const targets = [];
    for (const extension of this._extensions.values()) {
      if ((extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        targets.push(extension);
      }
    }

    const panelReadyTasks = [];
    for (const extension of targets) {
      // Extensions that were already active received the broadcast above. Extensions that are
      // activated as a result of this view activation need to be notified after activation.
      const wasActive = extension.active;
      if (!wasActive) {
        await this._activateExtension(extension, activationEvent);
        this._sendEventToExtension(extension, "viewActivated", { viewId: id });
      }
      // If the view corresponds to a contributed panel, wait for it to render so callers can
      // treat `activateView()` as "ready to show" (important for flaky/shared test runners).
      if ((extension.manifest.contributes.panels ?? []).some((p) => p.id === id)) {
        panelReadyTasks.push(this._waitForPanelHtml(id, this._activationTimeoutMs));
      }
    }

    if (panelReadyTasks.length > 0) {
      await Promise.all(panelReadyTasks);
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

  getClipboardText() {
    return this._clipboardText;
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
    return this._permissionManager.resetPermissions(String(extensionId));
  }

  async resetAllPermissions() {
    return this._permissionManager.resetAllPermissions();
  }

  getContributedCommands() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const cmd of extension.manifest.contributes.commands ?? []) {
        out.push({
          extensionId: extension.id,
          command: cmd.command,
          title: cmd.title,
          category: cmd.category ?? null,
          icon: cmd.icon ?? null,
          description: cmd.description ?? null,
          keywords: Array.isArray(cmd.keywords) ? cmd.keywords : null
        });
      }
    }
    return out;
  }

  getContributedPanels() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const panel of extension.manifest.contributes.panels ?? []) {
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
      for (const kb of extension.manifest.contributes.keybindings ?? []) {
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
      const menus = extension.manifest.contributes.menus ?? {};
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
      for (const item of record.items) {
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
      const menus = extension.manifest.contributes.menus ?? {};
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
      for (const fn of extension.manifest.contributes.customFunctions ?? []) {
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
      for (const connector of extension.manifest.contributes.dataConnectors ?? []) {
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
    // Best-effort: ensure spreadsheet event subscriptions are cleaned up so callers that
    // create/destroy multiple hosts do not leak listeners.
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
    this._extensions.clear();
    this._commands.clear();
    this._panels.clear();
    this._contextMenus.clear();
    this._customFunctions.clear();
    this._dataConnectors.clear();
    // Best-effort: reject any pending panel render waits so dispose doesn't hang.
    for (const [panelId, waiters] of this._panelHtmlWaiters.entries()) {
      for (const waiter of waiters) {
        if (waiter.timeout) clearTimeout(waiter.timeout);
        waiter.reject(new Error(`ExtensionHost disposed while waiting for panel HTML: ${panelId}`));
      }
    }
    this._panelHtmlWaiters.clear();
    this._messages = [];

    await Promise.allSettled(
      extensions.map(async (ext) =>
        this._terminateWorker(ext, { reason: "dispose", cause: new Error("ExtensionHost disposed") })
      )
    );
  }

  async reloadExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    await this._terminateWorker(extension, {
      reason: "reload",
      cause: new Error("Extension reloaded")
    });

    await this._spawnWorker(extension);
  }

  async unloadExtension(extensionId) {
    const id = String(extensionId);
    const extension = this._extensions.get(id);
    if (!extension) throw new Error(`Extension not loaded: ${id}`);

    await this._terminateWorker(extension, {
      reason: "unload",
      cause: new Error("Extension unloaded")
    });

    try {
      for (const cmd of extension.manifest.contributes.commands ?? []) {
        if (this._commands.get(cmd.command) === extension.id) this._commands.delete(cmd.command);
      }
    } catch {
      // ignore
    }

    try {
      for (const fn of extension.manifest.contributes.customFunctions ?? []) {
        if (this._customFunctions.get(fn.name) === extension.id) this._customFunctions.delete(fn.name);
      }
    } catch {
      // ignore
    }

    try {
      for (const connector of extension.manifest.contributes.dataConnectors ?? []) {
        if (this._dataConnectors.get(connector.id) === extension.id) this._dataConnectors.delete(connector.id);
      }
    } catch {
      // ignore
    }

    // Remove panels owned by this extension.
    for (const [panelId, panel] of this._panels.entries()) {
      if (panel?.extensionId === extension.id) this._panels.delete(panelId);
    }

    // Remove context menus owned by this extension.
    for (const [registrationId, record] of this._contextMenus.entries()) {
      if (record?.extensionId === extension.id) this._contextMenus.delete(registrationId);
    }

    this._extensions.delete(id);
  }

  /**
   * Clears persisted state owned by a given extension id (permissions + storage).
   *
   * This is intended to be called by installers/uninstall flows so a reinstall behaves
   * like a clean install.
   *
   * @param {string} extensionId
   */
  async resetExtensionState(extensionId) {
    const id = String(extensionId);
    if (!id || id === "." || id === "..") return;

    // Best-effort: do not fail uninstall flows because persistence is unavailable.
    try {
      await this._permissionManager.revokePermissions(id);
    } catch {
      // ignore
    }

    try {
      const store = await this._loadExtensionStorage();
      if (store && typeof store === "object" && !Array.isArray(store)) {
        if (Object.prototype.hasOwnProperty.call(store, id)) {
          delete store[id];
          await this._saveExtensionStorage(store);
        }
      }
    } catch {
      // ignore
    }

    // Best-effort: clear any on-disk extension data directories.
    // While extensions cannot currently access the filesystem directly, the host
    // already provisions `globalStoragePath`/`workspaceStoragePath` folders. Clearing
    // them on uninstall ensures future file-backed storage additions behave like a
    // clean reinstall.
    try {
      if (/[/\\]/.test(id) || id.includes("\0")) return;
      const root = path.resolve(this._extensionDataRoot);
      const target = path.resolve(path.join(root, id));
      const rel = path.relative(root, target);
      const inside = rel === "" || (!rel.startsWith(".." + path.sep) && rel !== ".." && !path.isAbsolute(rel));
      if (!inside) return;
      await fs.rm(target, { recursive: true, force: true });
    } catch {
      // ignore
    }
  }

  _getWorkerResourceLimits() {
    if (!this._memoryMb) return undefined;
    const oldGeneration = Math.floor(this._memoryMb);
    if (oldGeneration <= 0) return undefined;
    const youngGeneration = Math.min(oldGeneration, Math.min(64, Math.max(16, Math.floor(oldGeneration / 4))));
    return {
      // Best-effort: caps V8 heap size inside the extension worker thread.
      maxOldGenerationSizeMb: oldGeneration,
      maxYoungGenerationSizeMb: youngGeneration
    };
  }

  _spawnWorker(extension) {
    if (extension.worker) return Promise.resolve();
    if (extension.workerSpawn) return extension.workerSpawn;

    extension.workerSpawn = (async () => {
      const resourceLimits = this._getWorkerResourceLimits();
      const optionsBase = {
        workerData: extension.workerData,
        ...(resourceLimits ? { resourceLimits } : {})
      };

      // Retry transient EAGAIN failures. Keep the delays small: this is only meant to smooth over
      // temporary system-wide thread exhaustion on shared runners.
      const maxAttempts = 8;
      let delayMs = 10;

      for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
        try {
          const worker = new Worker(this._workerScriptPath, optionsBase);
          extension.worker = worker;

          worker.on("message", (msg) => this._handleWorkerMessage(extension, worker, msg));
          worker.on("error", (err) => this._handleWorkerCrash(extension, worker, err));
          worker.on("exit", (code) => {
            if (worker !== extension.worker) return;
            const err =
              code === 0
                ? new Error("Worker exited")
                : new Error(`Worker exited with code ${code ?? "unknown"}`);
            void this._terminateWorker(extension, { reason: "exit", cause: err }).catch(() => {
              // Best-effort: avoid unhandled rejections from worker termination bookkeeping.
            });
          });

          return;
        } catch (error) {
          if (!isEagainWorkerInitError(error) || attempt === maxAttempts) {
            throw error;
          }
          await sleep(delayMs);
          delayMs = Math.min(250, delayMs * 2);
        }
      }
    })();

    extension.workerSpawn = extension.workerSpawn.finally(() => {
      // Clear the spawn lock for future restarts.
      extension.workerSpawn = null;
    });

    return extension.workerSpawn;
  }

  async _ensureWorker(extension) {
    if (extension.worker) return;
    if (extension.workerTermination) {
      try {
        await extension.workerTermination;
      } catch {
        // ignore
      }
    }

    // A fresh worker always starts inactive until we successfully activate.
    extension.active = false;
    await this._spawnWorker(extension);
  }

  _requestFromWorker(extension, message, { timeoutMs, operation }) {
    const worker = extension.worker;
    if (!worker) {
      throw new Error(`Extension worker not running: ${extension.id}`);
    }

    const id = crypto.randomUUID();
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
              void this._terminateWorker(extension, {
                reason: "timeout",
                cause: err
              }).catch(() => {
                // Best-effort: avoid unhandled rejections from worker termination bookkeeping.
              });
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

  async _terminateWorker(extension, { reason, cause }) {
    const worker = extension.worker;

    extension.active = false;

    // Best-effort cleanup for runtime registrations that were tied to the crashed worker.
    // Contributed commands stay registered so the app can still route them for future activations.
    try {
      const contributedCommands = new Set(
        (extension.manifest.contributes.commands ?? []).map((c) => c.command)
      );

      for (const cmd of extension.registeredCommands) {
        if (contributedCommands.has(cmd)) continue;
        if (this._commands.get(cmd) === extension.id) this._commands.delete(cmd);
      }
      extension.registeredCommands.clear();
    } catch {
      // ignore
    }

    // Remove panels owned by this extension worker; they can no longer receive messages.
    for (const [panelId, panel] of this._panels.entries()) {
      if (panel?.extensionId === extension.id) {
        this._panels.delete(panelId);
        this._rejectPanelHtmlWaiters(panelId, createWorkerTerminatedError({ extensionId: extension.id, reason, cause }));
      }
    }

    // Remove context menus registered by this extension worker; they would otherwise leak.
    for (const [registrationId, record] of this._contextMenus.entries()) {
      if (record?.extensionId === extension.id) this._contextMenus.delete(registrationId);
    }

    // Reject outstanding requests bound for this worker.
    for (const [id, pending] of extension.pendingRequests.entries()) {
      if (pending.timeout) clearTimeout(pending.timeout);
      pending.reject(
        createWorkerTerminatedError({
          extensionId: extension.id,
          reason,
          cause
        })
      );
      extension.pendingRequests.delete(id);
    }

    if (!worker) return extension.workerTermination ?? null;

    // Mark worker dead immediately so new activations can spawn a fresh one.
    extension.worker = null;

    if (extension.workerTermination) return extension.workerTermination;

    const termination = worker
      .terminate()
      .catch(() => {
        // ignore
      })
      .finally(() => {
        if (extension.workerTermination === termination) extension.workerTermination = null;
      });

    extension.workerTermination = termination;

    return extension.workerTermination;
  }

  async _ensureSecurity() {
    if (this._securityModule) return;

    if (!this._securityModulePromise) {
      const fallbackPath = pathToFileURL(path.resolve(__dirname, "../../security/src/index.js")).href;
      this._securityModulePromise = import("@formula/security")
        .catch(async (error) => {
          // In monorepo/test environments we may not have workspace packages linked into node_modules.
          // Fall back to importing the source package directly so the extension host can run without
          // a full package manager install.
          if (error && typeof error === "object" && error.code && error.code !== "ERR_MODULE_NOT_FOUND") {
            throw error;
          }

          return import(fallbackPath);
        })
        .then((security) => {
          const { AuditLogger, SqliteAuditLogStore } = security;
          const store = new SqliteAuditLogStore({ path: this._securityAuditDbPath });
          const auditLogger = new AuditLogger({ store });

          this._securityModule = security;
          this._securityAuditLogger = auditLogger;
          return security;
        });
    }

    await this._securityModulePromise;
  }

  async _createSecurityPermissionManager(extension, apiKey) {
    await this._ensureSecurity();
    const principal = { type: "extension", id: extension.id };
    const security = this._securityModule;
    if (!security?.PermissionManager) {
      throw new Error("Security module is not available");
    }

    const permissionManager = new security.PermissionManager({ auditLogger: this._securityAuditLogger });
    const grants = await this._permissionManager.getGrantedPermissions(extension.id);
    const network = grants?.network;

    if (network?.mode === "full") {
      permissionManager.grant(principal, { network: { mode: "full" } }, {
        source: "extension-host.permission-manager",
        apiKey
      });
    } else if (network?.mode === "allowlist") {
      permissionManager.grant(
        principal,
        { network: { mode: "allowlist", allowlist: Array.isArray(network.hosts) ? network.hosts : [] } },
        {
          source: "extension-host.permission-manager",
          apiKey
        }
      );
    }

    return { principal, permissionManager };
  }

  async _handleApiCall(extension, message) {
    const worker = extension.worker;
    const id = message?.id;
    // Treat malformed `args` payloads as "no args" to avoid TypeError crashes when untrusted
    // extension workers post malformed messages.
    const args = Array.isArray(message?.args) ? message.args : [];
    const securityPrincipal = { type: "extension", id: extension.id };
    let apiKey = "";
    /** @type {string[]} */
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
      if (isNetworkApi) {
        await this._ensureSecurity();
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
      if (isNetworkApi) {
        try {
          await this._ensureSecurity();
          this._securityAuditLogger?.log({
            eventType: "security.permission.denied",
            actor: securityPrincipal,
            success: false,
            metadata: {
              apiKey,
              permissions,
              url: networkUrl ?? undefined,
              message: String(error?.message ?? error)
            }
          });
        } catch {
          // ignore
        }
      }

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
        return this._getWorkbookSnapshot();
      case "workbook.openWorkbook":
        {
          const workbookPath = args?.[0];
          if (typeof workbookPath !== "string" || workbookPath.trim().length === 0) {
            throw new Error("Workbook path must be a non-empty string");
          }
          return this.openWorkbook(workbookPath);
        }
      case "workbook.createWorkbook":
        return this.openWorkbook(null);
      case "workbook.save":
        this.saveWorkbook();
        return null;
      case "workbook.saveAs":
        {
          const workbookPath = args?.[0];
          if (typeof workbookPath !== "string" || workbookPath.trim().length === 0) {
            throw new Error("Workbook path must be a non-empty string");
          }
          this.saveWorkbookAs(workbookPath);
          return null;
        }
      case "workbook.close":
        this.closeWorkbook();
        return null;
      case "sheets.getActiveSheet":
        return this._spreadsheet.getActiveSheet?.() ?? { id: "sheet1", name: "Sheet1" };
      case "sheets.getSheet":
        return this._spreadsheet.getSheet(args[0]);
      case "sheets.activateSheet":
        return this._spreadsheet.activateSheet(args[0]);
      case "sheets.createSheet":
        return this._spreadsheet.createSheet(args[0]);
      case "sheets.renameSheet":
        this._spreadsheet.renameSheet(args[0], args[1]);
        return null;
      case "sheets.deleteSheet":
        this._spreadsheet.deleteSheet(args[0]);
        return null;

      case "cells.getSelection": {
        const result = await this._spreadsheet.getSelection();
        assertExtensionRangeWithinLimits(result, { label: "Selection" });
        return result;
      }
      case "cells.getRange": {
        // Fail fast on unbounded A1 ranges so we don't ask the spreadsheet implementation to
        // materialize huge 2D arrays. Ignore parse failures so non-A1 identifiers (e.g. named
        // ranges) can still be handled by the host spreadsheet implementation.
        let coords = null;
        try {
          coords = parseA1RangeRef(args[0]);
        } catch {
          coords = null;
        }
        if (coords) assertExtensionRangeWithinLimits(coords, { label: "Range" });
        return this._spreadsheet.getRange(args[0]);
      }
      case "cells.getCell":
        return this._spreadsheet.getCell(args[0], args[1]);
      case "cells.setCell":
        this._spreadsheet.setCell(args[0], args[1], args[2]);
        return null;
      case "cells.setRange": {
        let coords = null;
        try {
          coords = parseA1RangeRef(args[0]);
        } catch {
          coords = null;
        }
        if (coords) assertExtensionRangeWithinLimits(coords, { label: "Range" });
        this._spreadsheet.setRange(args[0], args[1]);
        return null;
      }

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
        return null;

      case "ui.showInputBox": {
        // Placeholder implementation until the desktop UI wires actual input prompts.
        const options = args[0] ?? {};
        const value = options && typeof options === "object" ? options.value : undefined;
        return typeof value === "string" ? value : null;
      }

      case "ui.showQuickPick": {
        // Placeholder implementation: choose the first entry.
        const items = Array.isArray(args[0]) ? args[0] : [];
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

        const id = crypto.randomUUID();
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
        this._panels.set(panelId, {
          id: panelId,
          title,
          html: "",
          extensionId: extension.id,
          outgoingMessages: []
        });
        return { id: panelId };
      }
      case "ui.setPanelHtml": {
        const panelId = String(args[0]);
        const html = String(args[1]);
        const panel = this._panels.get(panelId);
        if (!panel) throw new Error(`Unknown panel: ${panelId}`);
        panel.html = html;
        this._resolvePanelHtmlWaiters(panelId);
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
        return null;
      }
      case "ui.disposePanel": {
        const panelId = String(args[0]);
        this._panels.delete(panelId);
        this._rejectPanelHtmlWaiters(panelId, new Error(`Panel disposed: ${panelId}`));
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
        if (typeof fetch !== "function") {
          throw new Error("Network fetch is not available in this runtime");
        }

        const apiKey = "network.fetch";
        const { principal, permissionManager } = await this._createSecurityPermissionManager(extension, apiKey);
        const secureFetch = this._securityModule.createSecureFetch({
          principal,
          permissionManager,
          auditLogger: this._securityAuditLogger,
          promptIfDenied: false
        });

        const url = String(args[0]);
        const init = args[1];
        const response = await secureFetch(url, init);
        const bodyText = await response.text();
        const headers = [];
        response.headers.forEach((value, key) => {
          headers.push([key, value]);
        });

        return {
          ok: response.ok,
          status: response.status,
          statusText: response.statusText,
          url: response.url,
          headers,
          bodyText
        };
      }

      case "network.openWebSocket": {
        // Used by worker runtimes to permission-gate WebSocket connections before they
        // call the platform WebSocket implementation directly.
        const apiKey = "network.openWebSocket";
        const url = String(args[0]);
        const { principal, permissionManager } = await this._createSecurityPermissionManager(extension, apiKey);
        await permissionManager.ensure(principal, { kind: "network", url }, { promptIfDenied: false });
        this._securityAuditLogger?.log({
          eventType: "security.network.websocket.open",
          actor: principal,
          success: true,
          metadata: { url }
        });
        return null;
      }

      case "clipboard.readText":
        return this._clipboardText;
      case "clipboard.writeText":
        this._clipboardText = String(args[0] ?? "");
        return null;

      case "storage.get": {
        const store = await this._loadExtensionStorage();
        const key = normalizeStorageKey(args[0]);
        return store[extension.id]?.[key];
      }
      case "storage.set": {
        const key = normalizeStorageKey(args[0]);
        const value = args[1];
        const store = await this._loadExtensionStorage();
        store[extension.id] = store[extension.id] ?? Object.create(null);
        store[extension.id][key] = value;
        await this._saveExtensionStorage(store);
        return null;
      }
      case "storage.delete": {
        const key = normalizeStorageKey(args[0]);
        const store = await this._loadExtensionStorage();
        if (store[extension.id]) {
          delete store[extension.id][key];
          if (Object.keys(store[extension.id]).length === 0) {
            delete store[extension.id];
          }
          await this._saveExtensionStorage(store);
        }
        return null;
      }

      case "config.get": {
        const key = String(args[0]);
        const store = await this._loadExtensionStorage();
        const stored = store[extension.id]?.[`__config__:${key}`];
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
        const store = await this._loadExtensionStorage();
        store[extension.id] = store[extension.id] ?? Object.create(null);
        store[extension.id][key] = value;
        await this._saveExtensionStorage(store);
        this._sendEventToExtension(extension, "configChanged", { key: configKey, value });
        return null;
      }

      default:
        throw new Error(`Unknown API method: ${namespace}.${method}`);
    }
  }

  async _loadExtensionStorage() {
    try {
      const raw = await fs.readFile(this._extensionStoragePath, "utf8");
      const parsed = JSON.parse(raw);
      const { store, migrated } = normalizeExtensionStorageStore(parsed);
      if (migrated) {
        try {
          await this._saveExtensionStorage(store);
        } catch {
          // ignore migration write failures
        }
      }
      return store;
    } catch {
      return Object.create(null);
    }
  }

  async _saveExtensionStorage(store) {
    await fs.mkdir(path.dirname(this._extensionStoragePath), { recursive: true });
    await writeFileAtomic(this._extensionStoragePath, JSON.stringify(store, null, 2), "utf8");
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
      case "log": {
        // We keep this as a hook for future UI integration.
        return;
      }
      case "audit": {
        try {
          this._securityAuditLogger?.log(message.event);
        } catch {
          // Audit logging should not break extension execution.
        }
        return;
      }
      default:
        return;
    }
  }

  _handleWorkerCrash(extension, worker, error) {
    if (worker !== extension.worker) return;
    void this._terminateWorker(extension, {
      reason: "crash",
      cause: error
    }).catch(() => {
      // Best-effort: avoid unhandled rejections from worker termination bookkeeping.
    });
  }

  async _activateExtension(extension, reason) {
    if (extension.active) return;
    await this._ensureWorker(extension);
    await this._requestFromWorker(
      extension,
      { type: "activate", reason },
      { timeoutMs: this._activationTimeoutMs, operation: `activation (${reason})` }
    );
    extension.active = true;
  }

  _sendEventToExtension(extension, event, data) {
    try {
      extension.worker.postMessage({ type: "event", event, data });
    } catch {
      // ignore
    }
  }

  _broadcastEvent(event, data) {
    const payload = sanitizeEventPayload(event, data);
    for (const extension of this._extensions.values()) {
      // Only deliver events to active extensions. Workers are spawned eagerly on load, but the
      // extension entrypoint is not imported/executed until activation. Broadcasting workbook/grid
      // events to inactive extensions is wasted work and can cause false positives if the host ever
      // adds event-based taint tracking (as the browser host does for clipboard DLP enforcement).
      if (!extension.active) continue;
      try {
        extension.worker.postMessage({
          type: "event",
          event,
          data: payload
        });
      } catch {
        // ignore
      }
    }
  }
}

module.exports = {
  ExtensionHost,
  API_PERMISSIONS,

  installExtensionFromDirectory,
  uninstallExtension,
  listInstalledExtensions
};
