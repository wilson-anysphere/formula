const GLOBAL_STATE_KEY = Symbol.for("formula.extensionApi.state");
const state = globalThis[GLOBAL_STATE_KEY] ?? {
  transport: null,
  currentContext: { extensionId: "", extensionPath: "" },
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
    extensionPath: String(ctx?.extensionPath ?? "")
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
    return rpcCall("workbook", "getActiveWorkbook", []);
  },

  async openWorkbook(workbookPath) {
    return rpcCall("workbook", "openWorkbook", [String(workbookPath)]);
  },

  async createWorkbook() {
    return rpcCall("workbook", "createWorkbook", []);
  }
};

const sheets = {
  async getActiveSheet() {
    return rpcCall("sheets", "getActiveSheet", []);
  },

  async getSheet(name) {
    return rpcCall("sheets", "getSheet", [String(name)]);
  },

  async createSheet(name) {
    return rpcCall("sheets", "createSheet", [String(name)]);
  },

  async renameSheet(oldName, newName) {
    await rpcCall("sheets", "renameSheet", [String(oldName), String(newName)]);
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
    return addEventHandler("sheetActivated", callback);
  },
  onWorkbookOpened(callback) {
    return addEventHandler("workbookOpened", callback);
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
