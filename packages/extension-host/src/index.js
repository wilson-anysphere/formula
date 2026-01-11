const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { Worker } = require("node:worker_threads");
const { pathToFileURL } = require("node:url");

const { validateExtensionManifest } = require("./manifest");
const { PermissionManager } = require("./permission-manager");
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

class ExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionsStoragePath = path.join(process.cwd(), ".formula", "permissions.json"),
    extensionStoragePath = path.join(process.cwd(), ".formula", "storage.json"),
    auditDbPath = null,
    activationTimeoutMs = 5000,
    commandTimeoutMs = 5000,
    customFunctionTimeoutMs = 5000,
    memoryMb = 256,
    spreadsheet = new InMemorySpreadsheet()
  } = {}) {
    this._engineVersion = engineVersion;
    this._extensions = new Map();
    this._commands = new Map(); // commandId -> extensionId
    this._panels = new Map(); // panelId -> { id, title, html }
    this._customFunctions = new Map(); // functionName -> extensionId
    this._messages = [];
    this._clipboardText = "";
    this._activationTimeoutMs = Number.isFinite(activationTimeoutMs)
      ? Math.max(0, activationTimeoutMs)
      : 5000;
    this._commandTimeoutMs = Number.isFinite(commandTimeoutMs) ? Math.max(0, commandTimeoutMs) : 5000;
    this._customFunctionTimeoutMs = Number.isFinite(customFunctionTimeoutMs)
      ? Math.max(0, customFunctionTimeoutMs)
      : 5000;

    // Best-effort sandboxing: V8 resource limits prevent a single extension worker from
    // consuming unbounded heap memory, but they do not cover all native/externals.
    // Set to null/0 to disable.
    this._memoryMb = Number.isFinite(memoryMb) ? Math.max(0, memoryMb) : 0;
    this._extensionStoragePath = extensionStoragePath;
    this._spreadsheet = spreadsheet;
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
    this._apiModulePath = path.resolve(__dirname, "../../extension-api/index.js");

    this._spreadsheet.onSelectionChanged?.((e) => this._broadcastEvent("selectionChanged", e));
    this._spreadsheet.onCellChanged?.((e) => this._broadcastEvent("cellChanged", e));
    this._spreadsheet.onSheetActivated?.((e) => this._broadcastEvent("sheetActivated", e));
  }

  get spreadsheet() {
    return this._spreadsheet;
  }

  async loadExtension(extensionPath) {
    const manifestPath = path.join(extensionPath, "package.json");
    const raw = await fs.readFile(manifestPath, "utf8");
    const parsed = JSON.parse(raw);
    const manifest = validateExtensionManifest(parsed, { engineVersion: this._engineVersion });

    const extensionId = `${manifest.publisher}.${manifest.name}`;

    const extensionRoot = path.resolve(extensionPath);
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
    const extension = {
      id: extensionId,
      path: extensionPath,
      mainPath,
      manifest,
      worker: null,
      active: false,
      registeredCommands: new Set(),
      pendingRequests: new Map(),
      workerTermination: null,
      workerData: {
        extensionId,
        extensionPath,
        mainPath,
        apiModulePath: this._apiModulePath
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

    // Extensions that activate on startup should receive an initial workbook event.
    this._broadcastEvent("workbookOpened", { workbook: this._workbook });
  }

  openWorkbook(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    const name =
      workbookPathStr == null || workbookPathStr.trim().length === 0
        ? "MockWorkbook"
        : path.basename(workbookPathStr);
    this._workbook = { name, path: workbookPathStr };
    this._broadcastEvent("workbookOpened", { workbook: this._workbook });
    return this._workbook;
  }

  saveWorkbook() {
    // Stub implementation: the application will eventually perform real persistence
    // and may await extensions for edits. For now we just emit a stable event payload.
    this._broadcastEvent("beforeSave", { workbook: this._workbook });
    return true;
  }

  saveWorkbookAs(workbookPath) {
    const workbookPathStr = workbookPath == null ? null : String(workbookPath);
    if (workbookPathStr == null || workbookPathStr.trim().length === 0) {
      throw new Error("Workbook path must be a non-empty string");
    }

    const name = path.basename(workbookPathStr);
    this._workbook = { name, path: workbookPathStr };
    this._broadcastEvent("beforeSave", { workbook: this._workbook });
    return true;
  }

  closeWorkbook() {
    this.openWorkbook(null);
    return true;
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

  getContributedCommands() {
    const out = [];
    for (const extension of this._extensions.values()) {
      for (const cmd of extension.manifest.contributes.commands ?? []) {
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
    const extensions = [...this._extensions.values()];
    this._extensions.clear();
    this._commands.clear();
    this._panels.clear();
    this._customFunctions.clear();
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

    this._spawnWorker(extension);
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
    if (extension.worker) return;

    const options = {
      workerData: extension.workerData
    };

    const resourceLimits = this._getWorkerResourceLimits();
    if (resourceLimits) options.resourceLimits = resourceLimits;

    const worker = new Worker(this._workerScriptPath, options);
    extension.worker = worker;

    worker.on("message", (msg) => this._handleWorkerMessage(extension, worker, msg));
    worker.on("error", (err) => this._handleWorkerCrash(extension, worker, err));
    worker.on("exit", (code) => {
      if (worker !== extension.worker) return;
      const err =
        code === 0
          ? new Error("Worker exited")
          : new Error(`Worker exited with code ${code ?? "unknown"}`);
      void this._terminateWorker(extension, { reason: "exit", cause: err });
    });
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
    this._spawnWorker(extension);
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
    if (this._securityPermissionManager) return;

    if (!this._securityModulePromise) {
      const fallbackPath = pathToFileURL(path.resolve(__dirname, "../../security/src/index.js")).href;
      this._securityModulePromise = import("@formula/security")
        .catch(() => import(fallbackPath))
        .then((security) => {
          const { AuditLogger, PermissionManager, SqliteAuditLogStore } = security;
          const store = new SqliteAuditLogStore({ path: this._securityAuditDbPath });
          const auditLogger = new AuditLogger({ store });
          const permissionManager = new PermissionManager({ auditLogger });

          this._securityModule = security;
          this._securityAuditLogger = auditLogger;
          this._securityPermissionManager = permissionManager;
          return security;
        });
    }

    await this._securityModulePromise;
  }

  async _handleApiCall(extension, message) {
    const { id, namespace, method, args } = message;
    const apiKey = `${namespace}.${method}`;
    const worker = extension.worker;

    const permissions = API_PERMISSIONS[apiKey] ?? [];
    const securityPrincipal = { type: "extension", id: extension.id };
    const isNetworkFetch = apiKey === "network.fetch";

    try {
      if (isNetworkFetch) {
        await this._ensureSecurity();
      }

      await this._permissionManager.ensurePermissions(
        {
          extensionId: extension.id,
          displayName: extension.manifest.displayName ?? extension.manifest.name,
          declaredPermissions: extension.manifest.permissions ?? []
        },
        permissions
      );

      if (isNetworkFetch) {
        // Mirror the extension-host permission grant into the unified security manager so
        // audit logging and future allowlist enforcement happen in one place.
        this._securityPermissionManager.grant(
          securityPrincipal,
          { network: { mode: "full" } },
          { source: "extension-host.permission-manager", apiKey }
        );
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
      if (isNetworkFetch) {
        try {
          await this._ensureSecurity();
          this._securityAuditLogger?.log({
            eventType: "security.permission.denied",
            actor: securityPrincipal,
            success: false,
            metadata: { apiKey, permissions, message: String(error?.message ?? error) }
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
        return this._workbook;
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

      case "cells.getSelection":
        return this._spreadsheet.getSelection();
      case "cells.getRange":
        return this._spreadsheet.getRange(args[0]);
      case "cells.getCell":
        return this._spreadsheet.getCell(args[0], args[1]);
      case "cells.setCell":
        this._spreadsheet.setCell(args[0], args[1], args[2]);
        return null;
      case "cells.setRange":
        this._spreadsheet.setRange(args[0], args[1]);
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

        await this._ensureSecurity();
        const principal = { type: "extension", id: extension.id };
        const secureFetch = this._securityModule.createSecureFetch({
          principal,
          permissionManager: this._securityPermissionManager,
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

      case "network.openWebSocket":
        // Used by worker runtimes to permission-gate WebSocket connections before they
        // call the platform WebSocket implementation directly.
        return null;

      case "clipboard.readText":
        return this._clipboardText;
      case "clipboard.writeText":
        this._clipboardText = String(args[0] ?? "");
        return null;

      case "storage.get": {
        const store = await this._loadExtensionStorage();
        return store[extension.id]?.[String(args[0])];
      }
      case "storage.set": {
        const key = String(args[0]);
        const value = args[1];
        const store = await this._loadExtensionStorage();
        store[extension.id] = store[extension.id] ?? {};
        store[extension.id][key] = value;
        await this._saveExtensionStorage(store);
        return null;
      }
      case "storage.delete": {
        const key = String(args[0]);
        const store = await this._loadExtensionStorage();
        if (store[extension.id]) delete store[extension.id][key];
        await this._saveExtensionStorage(store);
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
        store[extension.id] = store[extension.id] ?? {};
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
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  }

  async _saveExtensionStorage(store) {
    await fs.mkdir(path.dirname(this._extensionStoragePath), { recursive: true });
    await fs.writeFile(this._extensionStoragePath, JSON.stringify(store, null, 2), "utf8");
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

module.exports = {
  ExtensionHost,
  API_PERMISSIONS,

  installExtensionFromDirectory,
  uninstallExtension,
  listInstalledExtensions
};
