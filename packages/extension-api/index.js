const GLOBAL_STATE_KEY = Symbol.for("formula.extensionApi.state");
const state = globalThis[GLOBAL_STATE_KEY] ?? {
  transport: null,
  currentContext: {
    extensionId: "",
    extensionPath: "",
    extensionUri: "",
    globalStoragePath: "",
    workspaceStoragePath: ""
  },
  nextRequestId: 1,
  pendingRequests: new Map(),
  commandHandlers: new Map(),
  eventHandlers: new Map(),
  panelMessageHandlers: new Map(),
  customFunctionHandlers: new Map()
};

// Ensure we always reuse the same shared state even if the package is loaded via both
// `require` (CJS) and `import` (ESM) within the same runtime (eg: Node workers).
globalThis[GLOBAL_STATE_KEY] = state;

function __setTransport(nextTransport) {
  state.transport = nextTransport;
}

function __setContext(ctx) {
  state.currentContext = {
    extensionId: String(ctx?.extensionId ?? ""),
    extensionPath: String(ctx?.extensionPath ?? ""),
    extensionUri: String(ctx?.extensionUri ?? ""),
    globalStoragePath: String(ctx?.globalStoragePath ?? ""),
    workspaceStoragePath: String(ctx?.workspaceStoragePath ?? "")
  };
}

function getTransportOrThrow() {
  if (!state.transport || typeof state.transport.postMessage !== "function") {
    throw new Error(
      "Extension API transport not initialized. This module must be run inside an extension host worker."
    );
  }
  return state.transport;
}

function createRequestId() {
  return String(state.nextRequestId++);
}

function rpcCall(namespace, method, args) {
  const t = getTransportOrThrow();
  const id = createRequestId();

  t.postMessage({
    type: "api_call",
    id,
    namespace,
    method,
    args
  });

  return new Promise((resolve, reject) => {
    state.pendingRequests.set(id, { resolve, reject });
  });
}

function notifyError(err) {
  try {
    // eslint-disable-next-line no-console
    console.error(err);
  } catch {
    // ignore
  }
}

function __handleMessage(message) {
  if (!message || typeof message !== "object") return;

  switch (message.type) {
    case "api_result": {
      const pending = state.pendingRequests.get(message.id);
      if (!pending) return;
      state.pendingRequests.delete(message.id);
      pending.resolve(message.result);
      return;
    }
    case "api_error": {
      const pending = state.pendingRequests.get(message.id);
      if (!pending) return;
      state.pendingRequests.delete(message.id);
      const errorMessage =
        typeof message.error === "string"
          ? message.error
          : String(message.error?.message ?? "Unknown error");
      const err = new Error(errorMessage);
      if (message.error?.stack) err.stack = String(message.error.stack);
      pending.reject(err);
      return;
    }
    case "execute_command": {
      const { id, commandId, args } = message;
      Promise.resolve()
        .then(async () => {
          const handler = state.commandHandlers.get(commandId);
          if (!handler) {
            throw new Error(`Command not registered: ${commandId}`);
          }
          return handler(...(Array.isArray(args) ? args : []));
        })
        .then(
          (result) => {
            getTransportOrThrow().postMessage({
              type: "command_result",
              id,
              result
            });
          },
          (error) => {
            getTransportOrThrow().postMessage({
              type: "command_error",
              id,
              error: { message: String(error?.message ?? error), stack: error?.stack }
            });
          }
        );
      return;
    }
    case "invoke_custom_function": {
      const { id, functionName, args } = message;
      Promise.resolve()
        .then(async () => {
          const handler = state.customFunctionHandlers.get(functionName);
          if (!handler) {
            throw new Error(`Custom function not registered: ${functionName}`);
          }
          return handler(...(Array.isArray(args) ? args : []));
        })
        .then(
          (result) => {
            getTransportOrThrow().postMessage({
              type: "custom_function_result",
              id,
              result
            });
          },
          (error) => {
            getTransportOrThrow().postMessage({
              type: "custom_function_error",
              id,
              error: { message: String(error?.message ?? error), stack: error?.stack }
            });
          }
        );
      return;
    }
    case "panel_message": {
      const panelId = String(message.panelId ?? "");
      const handlers = state.panelMessageHandlers.get(panelId);
      if (!handlers) return;
      for (const handler of handlers) {
        try {
          handler(message.message);
        } catch (err) {
          notifyError(err);
        }
      }
      return;
    }
    case "event": {
      const handlers = state.eventHandlers.get(message.event);
      if (!handlers) return;
      for (const handler of handlers) {
        try {
          handler(message.data);
        } catch (err) {
          notifyError(err);
        }
      }
      return;
    }
    default:
      return;
  }
}

function addEventHandler(event, handler) {
  if (typeof handler !== "function") {
    throw new Error("Event handler must be a function");
  }
  const key = String(event);
  if (!state.eventHandlers.has(key)) state.eventHandlers.set(key, new Set());
  const set = state.eventHandlers.get(key);
  set.add(handler);
  return new DisposableImpl(() => {
    set.delete(handler);
    if (set.size === 0) state.eventHandlers.delete(key);
  });
}

function attachNonEnumerableMethods(target, methods) {
  if (!target || typeof target !== "object") return target;
  if (!methods || typeof methods !== "object") return target;
  for (const [name, fn] of Object.entries(methods)) {
    if (typeof fn !== "function") continue;
    if (Object.prototype.hasOwnProperty.call(target, name)) continue;
    Object.defineProperty(target, name, { value: fn, enumerable: false });
  }
  return target;
}

function enhanceWorkbook(workbook) {
  if (!workbook || typeof workbook !== "object") return workbook;
  const obj = { ...workbook };

  if (Array.isArray(obj.sheets)) {
    obj.sheets = obj.sheets.map((sheet) => enhanceSheet(sheet));
  }
  if (obj.activeSheet && typeof obj.activeSheet === "object") {
    obj.activeSheet = enhanceSheet(obj.activeSheet);
  }

  return attachNonEnumerableMethods(obj, {
    async save() {
      await rpcCall("workbook", "save", []);
    },
    async saveAs(workbookPath) {
      await rpcCall("workbook", "saveAs", [String(workbookPath)]);
      const updated = await rpcCall("workbook", "getActiveWorkbook", []);
      if (updated && typeof updated === "object") {
        obj.name = updated.name;
        obj.path = updated.path;
        if (Array.isArray(updated.sheets)) {
          obj.sheets = updated.sheets.map((sheet) => enhanceSheet(sheet));
        }
        if (updated.activeSheet && typeof updated.activeSheet === "object") {
          obj.activeSheet = enhanceSheet(updated.activeSheet);
        }
      }
    },
    async close() {
      await rpcCall("workbook", "close", []);
      const updated = await rpcCall("workbook", "getActiveWorkbook", []);
      if (updated && typeof updated === "object") {
        obj.name = updated.name;
        obj.path = updated.path;
        if (Array.isArray(updated.sheets)) {
          obj.sheets = updated.sheets.map((sheet) => enhanceSheet(sheet));
        }
        if (updated.activeSheet && typeof updated.activeSheet === "object") {
          obj.activeSheet = enhanceSheet(updated.activeSheet);
        }
      }
    }
  });
}

function enhanceSheet(sheet) {
  if (!sheet || typeof sheet !== "object") return sheet;
  const obj = { ...sheet };

  return attachNonEnumerableMethods(obj, {
    async getRange(ref) {
      return rpcCall("cells", "getRange", [`${obj.name}!${String(ref)}`]);
    },
    async setRange(ref, values) {
      await rpcCall("cells", "setRange", [`${obj.name}!${String(ref)}`, values]);
    },
    async activate() {
      const updated = await rpcCall("sheets", "activateSheet", [obj.name]);
      if (updated && typeof updated === "object") {
        obj.id = updated.id;
        obj.name = updated.name;
      }
      return obj;
    },
    async rename(newName) {
      const from = obj.name;
      const to = String(newName);
      await rpcCall("sheets", "renameSheet", [from, to]);
      obj.name = to;
      return obj;
    }
  });
}

class DisposableImpl {
  constructor(disposeFn) {
    this._disposeFn = disposeFn;
  }

  dispose() {
    if (!this._disposeFn) return;
    const fn = this._disposeFn;
    this._disposeFn = null;
    fn();
  }
}

class PanelImpl {
  constructor(id) {
    this.id = id;
    const panelId = id;

    this.webview = {
      get html() {
        return "";
      },
      set html(value) {
        // Fire-and-forget; errors get surfaced to host logs.
        rpcCall("ui", "setPanelHtml", [panelId, String(value)]).catch(notifyError);
      },
      async setHtml(html) {
        await rpcCall("ui", "setPanelHtml", [panelId, String(html)]);
      },
      async postMessage(message) {
        await rpcCall("ui", "postMessageToPanel", [panelId, message]);
      },
      onDidReceiveMessage(handler) {
        if (typeof handler !== "function") {
          throw new Error("onDidReceiveMessage handler must be a function");
        }
        if (!state.panelMessageHandlers.has(panelId))
          state.panelMessageHandlers.set(panelId, new Set());
        const set = state.panelMessageHandlers.get(panelId);
        set.add(handler);
        return new DisposableImpl(() => {
          set.delete(handler);
          if (set.size === 0) state.panelMessageHandlers.delete(panelId);
        });
      }
    };
  }

  dispose() {
    rpcCall("ui", "disposePanel", [this.id]).catch(notifyError);
  }
}

const cells = {
  async getSelection() {
    return rpcCall("cells", "getSelection", []);
  },

  async getRange(ref) {
    return rpcCall("cells", "getRange", [String(ref)]);
  },

  async getCell(row, col) {
    return rpcCall("cells", "getCell", [row, col]);
  },

  async setCell(row, col, value) {
    await rpcCall("cells", "setCell", [row, col, value]);
  },

  async setRange(ref, values) {
    await rpcCall("cells", "setRange", [String(ref), values]);
  }
};

const workbook = {
  async getActiveWorkbook() {
    const result = await rpcCall("workbook", "getActiveWorkbook", []);
    return enhanceWorkbook(result);
  },

  async openWorkbook(workbookPath) {
    const result = await rpcCall("workbook", "openWorkbook", [String(workbookPath)]);
    return enhanceWorkbook(result);
  },

  async createWorkbook() {
    const result = await rpcCall("workbook", "createWorkbook", []);
    return enhanceWorkbook(result);
  },

  async save() {
    await rpcCall("workbook", "save", []);
  },

  async saveAs(workbookPath) {
    await rpcCall("workbook", "saveAs", [String(workbookPath)]);
  },

  async close() {
    await rpcCall("workbook", "close", []);
  }
};

const sheets = {
  async getActiveSheet() {
    const result = await rpcCall("sheets", "getActiveSheet", []);
    return enhanceSheet(result);
  },

  async getSheet(name) {
    const result = await rpcCall("sheets", "getSheet", [String(name)]);
    if (!result) return undefined;
    return enhanceSheet(result);
  },

  async activateSheet(name) {
    const result = await rpcCall("sheets", "activateSheet", [String(name)]);
    return enhanceSheet(result);
  },

  async createSheet(name) {
    const result = await rpcCall("sheets", "createSheet", [String(name)]);
    return enhanceSheet(result);
  },

  async renameSheet(oldName, newName) {
    await rpcCall("sheets", "renameSheet", [String(oldName), String(newName)]);
  },

  async deleteSheet(name) {
    await rpcCall("sheets", "deleteSheet", [String(name)]);
  }
};

const commands = {
  async registerCommand(id, handler) {
    if (typeof id !== "string" || id.trim().length === 0) {
      throw new Error("Command id must be a non-empty string");
    }
    if (typeof handler !== "function") {
      throw new Error("Command handler must be a function");
    }

    state.commandHandlers.set(id, handler);
    await rpcCall("commands", "registerCommand", [id]);

    return new DisposableImpl(() => {
      state.commandHandlers.delete(id);
      rpcCall("commands", "unregisterCommand", [id]).catch(notifyError);
    });
  },

  async executeCommand(id, ...args) {
    return rpcCall("commands", "executeCommand", [String(id), ...args]);
  }
};

const functions = {
  async register(name, def) {
    const fnName = String(name);
    if (fnName.trim().length === 0) throw new Error("Function name must be a non-empty string");
    if (!def || typeof def !== "object") throw new Error("Function definition must be an object");
    if (typeof def.handler !== "function") throw new Error("Function definition must include handler()");

    state.customFunctionHandlers.set(fnName, def.handler);
    await rpcCall("functions", "register", [
      fnName,
      {
        description: def.description,
        parameters: def.parameters,
        result: def.result,
        isAsync: def.isAsync,
        returnsArray: def.returnsArray
      }
    ]);

    return new DisposableImpl(() => {
      state.customFunctionHandlers.delete(fnName);
      rpcCall("functions", "unregister", [fnName]).catch(notifyError);
    });
  }
};

const ui = {
  async showMessage(message, type = "info") {
    await rpcCall("ui", "showMessage", [String(message), String(type)]);
  },

  async showInputBox(options) {
    const result = await rpcCall("ui", "showInputBox", [options ?? {}]);
    if (result === null || result === undefined) return undefined;
    return String(result);
  },

  async showQuickPick(items, options) {
    const result = await rpcCall("ui", "showQuickPick", [Array.isArray(items) ? items : [], options ?? {}]);
    if (result === null || result === undefined) return undefined;
    return result;
  },

  async registerContextMenu(menuId, items) {
    const id = String(menuId);
    if (id.trim().length === 0) throw new Error("Menu id must be a non-empty string");
    if (!Array.isArray(items)) throw new Error("Menu items must be an array");

    const result = await rpcCall("ui", "registerContextMenu", [id, items]);
    const registrationId = String(result?.id ?? "");
    if (registrationId.trim().length === 0) {
      throw new Error("Failed to register context menu: missing registration id");
    }

    return new DisposableImpl(() => {
      rpcCall("ui", "unregisterContextMenu", [registrationId]).catch(notifyError);
    });
  },

  async createPanel(id, options) {
    const result = await rpcCall("ui", "createPanel", [String(id), options ?? {}]);
    return new PanelImpl(result?.id ?? String(id));
  }
};

const storage = {
  async get(key) {
    return rpcCall("storage", "get", [String(key)]);
  },

  async set(key, value) {
    await rpcCall("storage", "set", [String(key), value]);
  },

  async delete(key) {
    await rpcCall("storage", "delete", [String(key)]);
  }
};

const config = {
  async get(key) {
    return rpcCall("config", "get", [String(key)]);
  },
  async update(key, value) {
    await rpcCall("config", "update", [String(key), value]);
  },
  onDidChange(callback) {
    return addEventHandler("configChanged", callback);
  }
};

function createHeaders(entries) {
  const map = new Map(Array.isArray(entries) ? entries : []);
  return {
    get(name) {
      if (typeof name !== "string") return undefined;
      return map.get(name.toLowerCase()) ?? map.get(name) ?? undefined;
    }
  };
}

function createFetchResponse(payload) {
  const bodyText = String(payload?.bodyText ?? "");
  const headers = createHeaders(
    (payload?.headers ?? []).map(([k, v]) => [String(k).toLowerCase(), String(v)])
  );

  return {
    ok: Boolean(payload?.ok),
    status: Number(payload?.status ?? 0),
    statusText: String(payload?.statusText ?? ""),
    url: String(payload?.url ?? ""),
    headers,
    async text() {
      return bodyText;
    },
    async json() {
      return JSON.parse(bodyText);
    }
  };
}

const network = {
  async fetch(url, init) {
    const result = await rpcCall("network", "fetch", [String(url), init]);
    return createFetchResponse(result);
  },

  /**
   * Used by worker runtimes to permission-gate WebSocket connections before they
   * call the platform WebSocket implementation directly.
   *
   * @param {string} url
   * @returns {Promise<void>}
   */
  async openWebSocket(url) {
    await rpcCall("network", "openWebSocket", [String(url)]);
  }
};

const clipboard = {
  async readText() {
    return rpcCall("clipboard", "readText", []);
  },
  async writeText(text) {
    await rpcCall("clipboard", "writeText", [String(text)]);
  }
};

const events = {
  onSelectionChanged(callback) {
    return addEventHandler("selectionChanged", callback);
  },
  onCellChanged(callback) {
    return addEventHandler("cellChanged", callback);
  },
  onSheetActivated(callback) {
    return addEventHandler("sheetActivated", (e) => {
      if (e && typeof e === "object" && e.sheet && typeof e.sheet === "object") {
        callback({ ...e, sheet: enhanceSheet(e.sheet) });
        return;
      }
      callback(e);
    });
  },
  onWorkbookOpened(callback) {
    return addEventHandler("workbookOpened", (e) => {
      if (e && typeof e === "object" && e.workbook && typeof e.workbook === "object") {
        callback({ ...e, workbook: enhanceWorkbook(e.workbook) });
        return;
      }
      callback(e);
    });
  },
  onBeforeSave(callback) {
    return addEventHandler("beforeSave", (e) => {
      if (e && typeof e === "object" && e.workbook && typeof e.workbook === "object") {
        callback({ ...e, workbook: enhanceWorkbook(e.workbook) });
        return;
      }
      callback(e);
    });
  },
  onViewActivated(callback) {
    return addEventHandler("viewActivated", callback);
  }
};

const context = new Proxy(
  {},
  {
    get(_target, prop) {
      if (prop === "extensionId") return state.currentContext.extensionId;
      if (prop === "extensionPath") return state.currentContext.extensionPath;
      if (prop === "extensionUri") return state.currentContext.extensionUri;
      if (prop === "globalStoragePath") return state.currentContext.globalStoragePath;
      if (prop === "workspaceStoragePath") return state.currentContext.workspaceStoragePath;
      return undefined;
    }
  }
);

module.exports = {
  workbook,
  sheets,
  cells,
  commands,
  functions,
  network,
  clipboard,
  ui,
  storage,
  config,
  events,
  context,

  __setTransport,
  __setContext,
  __handleMessage
};
