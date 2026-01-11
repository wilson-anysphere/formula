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
  if (error instanceof Error) {
    return { message: error.message, stack: error.stack };
  }
  return { message: String(error) };
}

function deserializeError(payload) {
  const message = typeof payload === "string" ? payload : String(payload?.message ?? "Unknown error");
  const err = new Error(message);
  if (payload?.stack) err.stack = String(payload.stack);
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

class InMemoryExtensionStorage {
  constructor() {
    this._data = {};
  }

  /**
   * @param {string} extensionId
   */
  getExtensionStore(extensionId) {
    if (!this._data[extensionId]) this._data[extensionId] = {};
    return this._data[extensionId];
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
    if (!this._storage) return {};
    try {
      const raw = this._storage.getItem(this._key(extensionId));
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return {};
      return parsed;
    } catch {
      return {};
    }
  }

  _persist(extensionId, target) {
    if (!this._storage) return;
    this._storage.setItem(this._key(extensionId), JSON.stringify(target));
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
}

class BrowserExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionStorage,
    permissionStorageKey,
    spreadsheetApi,
    clipboardApi,
    storageApi,
    activationTimeoutMs = 5000,
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
    this._panels = new Map();
    this._contextMenus = new Map();
    this._customFunctions = new Map();
    this._dataConnectors = new Map();
    this._messages = [];
    this._activationTimeoutMs = Number.isFinite(activationTimeoutMs)
      ? Math.max(0, activationTimeoutMs)
      : 5000;
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

    this._clipboardText = "";
    this._clipboardApi = clipboardApi ?? {
      readText: async () => this._clipboardText,
      writeText: async (text) => {
        this._clipboardText = String(text ?? "");
      }
    };

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

    this._spreadsheet.onSelectionChanged?.((e) => this._broadcastEvent("selectionChanged", e));
    this._spreadsheet.onCellChanged?.((e) => this._broadcastEvent("cellChanged", e));
    this._spreadsheet.onSheetActivated?.((e) => this._broadcastEvent("sheetActivated", e));
  }

  get spreadsheet() {
    return this._spreadsheet;
  }

  async _getActiveWorkbook() {
    if (typeof this._spreadsheet.getActiveWorkbook === "function") {
      try {
        const workbook = await this._spreadsheet.getActiveWorkbook();
        if (workbook && typeof workbook === "object") {
          this._workbook = {
            name: String(workbook.name ?? this._workbook.name),
            path: workbook.path ?? this._workbook.path ?? null
          };
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
    const manifest = validateExtensionManifest(parsed, { engineVersion: this._engineVersion });
    const extensionId = `${manifest.publisher}.${manifest.name}`;

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
    this._spawnWorker(extension);
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

    // Mirror the Node host behavior: extensions that activate on startup should receive the
    // initial workbook event so they can initialize state.
    this._broadcastEvent("workbookOpened", { workbook: await this._getActiveWorkbook() });
  }

  async activateView(viewId) {
    const activationEvent = `onView:${viewId}`;
    const targets = [];
    for (const extension of this._extensions.values()) {
      if ((extension.manifest.activationEvents ?? []).includes(activationEvent)) {
        targets.push(extension);
      }
    }

    for (const extension of targets) {
      if (!extension.active) {
        await this._activateExtension(extension, activationEvent);
      }
      this._sendEventToExtension(extension, "viewActivated", { viewId });
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

  async resetAllPermissions() {
    return this._permissionManager.resetAllPermissions();
  }

  getContributedCommands() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const cmd of extension.manifest.contributes?.commands ?? []) {
        out.push({
          extensionId: extension.id,
          command: cmd.command,
          title: cmd.title,
          category: cmd.category ?? null,
          icon: cmd.icon ?? null
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
          mac: kb.mac ?? null
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
    const extensions = [...this._extensions.values()];
    this._extensions.clear();
    this._commands.clear();
    this._panels.clear();
    this._contextMenus.clear();
    this._customFunctions.clear();
    this._dataConnectors.clear();
    this._messages = [];

    for (const ext of extensions) {
      this._terminateWorker(ext, { reason: "dispose", cause: new Error("BrowserExtensionHost disposed") });
    }
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
      for (const fn of extension.manifest.contributes?.customFunctions ?? []) {
        if (this._customFunctions.get(fn.name) === extension.id) this._customFunctions.delete(fn.name);
      }
    } catch {
      // ignore
    }

    for (const [panelId, panel] of this._panels.entries()) {
      if (panel?.extensionId === extension.id) this._panels.delete(panelId);
    }

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

    for (const [panelId, panel] of this._panels.entries()) {
      if (panel?.extensionId === extension.id) this._panels.delete(panelId);
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
    const { id, namespace, method, args } = message;
    const apiKey = `${namespace}.${method}`;
    const worker = extension.worker;

    const permissions = API_PERMISSIONS[apiKey] ?? [];
    const isNetworkApi = apiKey === "network.fetch" || apiKey === "network.openWebSocket";
    const networkUrl = isNetworkApi ? String(args?.[0] ?? "") : null;

    try {
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
        return this.openWorkbook(args[0]);
      case "workbook.createWorkbook":
        return this.openWorkbook(null);
      case "workbook.save":
        this.saveWorkbook();
        return null;
      case "workbook.saveAs":
        this.saveWorkbookAs(args[0]);
        return null;
      case "workbook.close":
        this.closeWorkbook();
        return null;
      case "sheets.getActiveSheet":
        if (typeof this._spreadsheet.getActiveSheet === "function") {
          return this._spreadsheet.getActiveSheet();
        }
        return this._sheets.find((s) => s.id === this._activeSheetId) ?? { id: "sheet1", name: "Sheet1" };

      case "cells.getSelection":
        return this._spreadsheet.getSelection();
      case "cells.getRange":
        if (typeof this._spreadsheet.getRange === "function") {
          return this._spreadsheet.getRange(args[0]);
        }
        return this._defaultGetRange(args[0]);
      case "cells.getCell":
        return this._spreadsheet.getCell(args[0], args[1]);
      case "cells.setCell":
        await this._spreadsheet.setCell(args[0], args[1], args[2]);
        return null;
      case "cells.setRange":
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
          return this._spreadsheet.activateSheet(args[0]);
        }
        return this._defaultActivateSheet(args[0]);
      case "sheets.createSheet":
        if (typeof this._spreadsheet.createSheet === "function") {
          return this._spreadsheet.createSheet(args[0]);
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
        const menuId = String(args[0]);
        const items = Array.isArray(args[1]) ? args[1] : [];
        if (menuId.trim().length === 0) throw new Error("Menu id must be a non-empty string");

        const normalized = [];
        for (const [idx, item] of items.entries()) {
          if (!item || typeof item !== "object") {
            throw new Error(`Menu item at index ${idx} must be an object`);
          }
          const command = String(item.command ?? "");
          if (command.trim().length === 0) {
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

        const url = String(args[0]);
        const init = args[1];
        const response = await fetch(url, init);
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
      case "clipboard.writeText":
        await this._clipboardApi.writeText(String(args[0] ?? ""));
        return null;

      case "storage.get": {
        const store = this._storageApi.getExtensionStore(extension.id);
        return store[String(args[0])];
      }
      case "storage.set": {
        const key = String(args[0]);
        const value = args[1];
        const store = this._storageApi.getExtensionStore(extension.id);
        store[key] = value;
        return null;
      }
      case "storage.delete": {
        const key = String(args[0]);
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

  _sendEventToExtension(extension, event, data) {
    try {
      extension.worker.postMessage({ type: "event", event, data });
    } catch {
      // ignore
    }
  }

  _broadcastEvent(event, data) {
    for (const extension of this._extensions.values()) {
      try {
        extension.worker.postMessage({
          type: "event",
          event,
          data
        });
      } catch {
        // ignore
      }
    }
  }
}

export { BrowserExtensionHost, API_PERMISSIONS };
