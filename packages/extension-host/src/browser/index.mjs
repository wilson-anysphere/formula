import { PermissionManager } from "./permission-manager.mjs";
import { validateExtensionManifest } from "./manifest.mjs";

const API_PERMISSIONS = {
  "workbook.getActiveWorkbook": [],
  "workbook.openWorkbook": ["workbook.manage"],
  "workbook.createWorkbook": ["workbook.manage"],
  "sheets.getActiveSheet": [],

  "cells.getSelection": ["cells.read"],
  "cells.getRange": ["cells.read"],
  "cells.getCell": ["cells.read"],
  "cells.setCell": ["cells.write"],
  "cells.setRange": ["cells.write"],

  "sheets.getSheet": ["sheets.manage"],
  "sheets.createSheet": ["sheets.manage"],
  "sheets.renameSheet": ["sheets.manage"],

  "commands.registerCommand": ["ui.commands"],
  "commands.unregisterCommand": ["ui.commands"],
  "commands.executeCommand": ["ui.commands"],

  "ui.createPanel": ["ui.panels"],
  "ui.setPanelHtml": ["ui.panels"],
  "ui.postMessageToPanel": ["ui.panels"],
  "ui.disposePanel": ["ui.panels"],

  "functions.register": [],
  "functions.unregister": [],

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

function serializeError(error) {
  if (error instanceof Error) {
    return { message: error.message, stack: error.stack };
  }
  return { message: String(error) };
}

function createRequestId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
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

class BrowserExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    spreadsheetApi,
    clipboardApi,
    storageApi
  } = {}) {
    if (!spreadsheetApi) {
      throw new Error("BrowserExtensionHost requires a spreadsheetApi implementation");
    }

    this._engineVersion = engineVersion;
    this._extensions = new Map();
    this._commands = new Map();
    this._panels = new Map();
    this._customFunctions = new Map();
    this._messages = [];
    this._pendingWorkerRequests = new Map();

    this._spreadsheet = spreadsheetApi;
    this._workbook = { name: "Workbook", path: null };
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

    this._storageApi = storageApi ?? new InMemoryExtensionStorage();

    this._permissionManager = new PermissionManager({ prompt: permissionPrompt });

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
    return this._workbook;
  }

  openWorkbook(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    const name =
      workbookPathStr == null || workbookPathStr.trim().length === 0
        ? "MockWorkbook"
        : workbookPathStr.split(/[/\\]/).pop() ?? workbookPathStr;
    this._workbook = { name, path: workbookPathStr };
    this._broadcastEvent("workbookOpened", { workbook: this._workbook });
    return this._workbook;
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

    const baseUrl = new URL("./", resolvedUrl).toString();
    const entry = manifest.browser ?? manifest.module ?? manifest.main;
    if (!entry) {
      throw new Error(`Extension manifest missing entrypoint (main/module/browser): ${extensionId}`);
    }

    const mainUrl = new URL(entry, baseUrl).toString();

    await this.loadExtension({
      extensionId,
      extensionPath: baseUrl,
      manifest,
      mainUrl
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

    const workerUrl = new URL("./extension-worker.mjs", import.meta.url);
    const worker = new Worker(workerUrl, { type: "module" });

    const extension = {
      id: extensionId,
      path: extensionPath,
      manifest,
      mainUrl,
      worker,
      active: false,
      registeredCommands: new Set()
    };

    worker.addEventListener("message", (event) => this._handleWorkerMessage(extension, event.data));
    worker.addEventListener("error", (event) =>
      this._handleWorkerCrash(extension, new Error(String(event?.message ?? "Worker error")))
    );

    worker.postMessage({
      type: "init",
      extensionId,
      extensionPath,
      mainUrl
    });

    for (const cmd of manifest.contributes.commands ?? []) {
      this._commands.set(cmd.command, extensionId);
    }

    for (const fn of manifest.contributes.customFunctions ?? []) {
      if (this._customFunctions.has(fn.name)) {
        throw new Error(`Duplicate custom function name: ${fn.name}`);
      }
      this._customFunctions.set(fn.name, extensionId);
    }

    this._extensions.set(extensionId, extension);
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

    const id = createRequestId();
    const promise = new Promise((resolve, reject) => {
      this._pendingWorkerRequests.set(id, { resolve, reject });
    });

    extension.worker.postMessage({
      type: "execute_command",
      id,
      commandId,
      args
    });

    return promise;
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

    const id = createRequestId();
    const promise = new Promise((resolve, reject) => {
      this._pendingWorkerRequests.set(id, { resolve, reject });
    });

    extension.worker.postMessage({
      type: "invoke_custom_function",
      id,
      functionName: name,
      args
    });

    return promise;
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
    extension.worker.postMessage({ type: "panel_message", panelId, message });
  }

  async dispose() {
    const extensions = [...this._extensions.values()];
    this._extensions.clear();
    this._commands.clear();
    this._panels.clear();
    this._customFunctions.clear();
    this._messages = [];

    for (const ext of extensions) {
      try {
        ext.worker.terminate();
      } catch {
        // ignore
      }
    }
  }

  async _handleApiCall(extension, message) {
    const { id, namespace, method, args } = message;
    const apiKey = `${namespace}.${method}`;

    const permissions = API_PERMISSIONS[apiKey] ?? [];

    try {
      await this._permissionManager.ensurePermissions(
        {
          extensionId: extension.id,
          displayName: extension.manifest.displayName ?? extension.manifest.name,
          declaredPermissions: extension.manifest.permissions ?? []
        },
        permissions
      );

      const result = await this._executeApi(namespace, method, args, extension);

      extension.worker.postMessage({
        type: "api_result",
        id,
        result
      });
    } catch (error) {
      extension.worker.postMessage({
        type: "api_error",
        id,
        error: serializeError(error)
      });
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

  _handleWorkerMessage(extension, message) {
    if (!message || typeof message !== "object") return;

    switch (message.type) {
      case "api_call":
        this._handleApiCall(extension, message);
        return;
      case "activate_result": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.resolve(true);
        return;
      }
      case "activate_error": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.reject(new Error(String(message.error?.message ?? message.error)));
        return;
      }
      case "command_result": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.resolve(message.result);
        return;
      }
      case "command_error": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.reject(new Error(String(message.error?.message ?? message.error)));
        return;
      }
      case "custom_function_result": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.resolve(message.result);
        return;
      }
      case "custom_function_error": {
        const pending = this._pendingWorkerRequests.get(message.id);
        if (!pending) return;
        this._pendingWorkerRequests.delete(message.id);
        pending.reject(new Error(String(message.error?.message ?? message.error)));
        return;
      }
      case "log":
        return;
      default:
        return;
    }
  }

  _handleWorkerCrash(extension, error) {
    for (const [id, pending] of this._pendingWorkerRequests.entries()) {
      pending.reject(new Error(`Extension worker crashed: ${extension.id}: ${error.message}`));
      this._pendingWorkerRequests.delete(id);
    }
  }

  async _activateExtension(extension, reason) {
    const id = createRequestId();
    const promise = new Promise((resolve, reject) => {
      this._pendingWorkerRequests.set(id, { resolve, reject });
    });

    extension.worker.postMessage({ type: "activate", id, reason });
    await promise;
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
