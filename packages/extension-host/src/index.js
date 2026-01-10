const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { Worker } = require("node:worker_threads");

const { validateExtensionManifest } = require("./manifest");
const { PermissionManager } = require("./permission-manager");
const { InMemorySpreadsheet } = require("./spreadsheet-mock");
const { installExtensionFromDirectory, uninstallExtension, listInstalledExtensions } = require("./installer");

const API_PERMISSIONS = {
  "workbook.getActiveWorkbook": [],
  "sheets.getActiveSheet": [],

  "cells.getSelection": ["cells.read"],
  "cells.getCell": ["cells.read"],
  "cells.setCell": ["cells.write"],

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

class ExtensionHost {
  constructor({
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionsStoragePath = path.join(process.cwd(), ".formula", "permissions.json"),
    extensionStoragePath = path.join(process.cwd(), ".formula", "storage.json"),
    auditDbPath = null,
    spreadsheet = new InMemorySpreadsheet()
  } = {}) {
    this._engineVersion = engineVersion;
    this._extensions = new Map();
    this._commands = new Map(); // commandId -> extensionId
    this._panels = new Map(); // panelId -> { id, title, html }
    this._customFunctions = new Map(); // functionName -> extensionId
    this._messages = [];
    this._clipboardText = "";
    this._pendingWorkerRequests = new Map();
    this._extensionStoragePath = extensionStoragePath;
    this._spreadsheet = spreadsheet;

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

    this._spreadsheet.onSelectionChanged?.((e) => this._broadcastEvent("selectionChanged", e));
    this._spreadsheet.onCellChanged?.((e) => this._broadcastEvent("cellChanged", e));
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

    const mainPath = path.resolve(extensionPath, manifest.main);
    const workerScript = path.resolve(__dirname, "../worker/extension-worker.js");
    const apiModulePath = path.resolve(__dirname, "../../extension-api/index.js");

    const worker = new Worker(workerScript, {
      workerData: {
        extensionId,
        extensionPath,
        mainPath,
        apiModulePath
      }
    });

    const extension = {
      id: extensionId,
      path: extensionPath,
      mainPath,
      manifest,
      worker,
      active: false,
      registeredCommands: new Set()
    };

    worker.on("message", (msg) => this._handleWorkerMessage(extension, msg));
    worker.on("error", (err) => this._handleWorkerCrash(extension, err));
    worker.on("exit", (code) => {
      if (code !== 0) {
        this._handleWorkerCrash(extension, new Error(`Worker exited with code ${code}`));
      }
    });

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

    const id = crypto.randomUUID();
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

    const id = crypto.randomUUID();
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
      extensions.map(async (ext) => {
        try {
          await ext.worker.terminate();
        } catch {
          // ignore
        }
      })
    );
  }

  async _ensureSecurity() {
    if (this._securityPermissionManager) return;

    if (!this._securityModulePromise) {
      this._securityModulePromise = import("@formula/security").then((security) => {
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

      extension.worker.postMessage({
        type: "api_result",
        id,
        result
      });
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
        return { name: "MockWorkbook", path: null };
      case "sheets.getActiveSheet":
        return { id: "sheet1", name: "Sheet1" };

      case "cells.getSelection":
        return this._spreadsheet.getSelection();
      case "cells.getCell":
        return this._spreadsheet.getCell(args[0], args[1]);
      case "cells.setCell":
        this._spreadsheet.setCell(args[0], args[1], args[2]);
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

  _handleWorkerCrash(extension, error) {
    for (const [id, pending] of this._pendingWorkerRequests.entries()) {
      pending.reject(new Error(`Extension worker crashed: ${extension.id}: ${error.message}`));
      this._pendingWorkerRequests.delete(id);
    }
  }

  async _activateExtension(extension, reason) {
    const id = crypto.randomUUID();
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

module.exports = {
  ExtensionHost,
  API_PERMISSIONS,

  installExtensionFromDirectory,
  uninstallExtension,
  listInstalledExtensions
};
