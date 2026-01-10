let transport = null;
let currentContext = { extensionId: "", extensionPath: "" };

let nextRequestId = 1;
const pendingRequests = new Map();

const commandHandlers = new Map();
const eventHandlers = new Map();
const panelMessageHandlers = new Map();
const customFunctionHandlers = new Map();

function __setTransport(nextTransport) {
  transport = nextTransport;
}

function __setContext(ctx) {
  currentContext = {
    extensionId: String(ctx?.extensionId ?? ""),
    extensionPath: String(ctx?.extensionPath ?? "")
  };
}

function getTransportOrThrow() {
  if (!transport || typeof transport.postMessage !== "function") {
    throw new Error(
      "Extension API transport not initialized. This module must be run inside an extension host worker."
    );
  }
  return transport;
}

function createRequestId() {
  return String(nextRequestId++);
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
    pendingRequests.set(id, { resolve, reject });
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
      const pending = pendingRequests.get(message.id);
      if (!pending) return;
      pendingRequests.delete(message.id);
      pending.resolve(message.result);
      return;
    }
    case "api_error": {
      const pending = pendingRequests.get(message.id);
      if (!pending) return;
      pendingRequests.delete(message.id);
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
          const handler = commandHandlers.get(commandId);
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
          const handler = customFunctionHandlers.get(functionName);
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
      const handlers = panelMessageHandlers.get(panelId);
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
      const handlers = eventHandlers.get(message.event);
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
  if (!eventHandlers.has(key)) eventHandlers.set(key, new Set());
  const set = eventHandlers.get(key);
  set.add(handler);
  return new DisposableImpl(() => {
    set.delete(handler);
    if (set.size === 0) eventHandlers.delete(key);
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
        if (!panelMessageHandlers.has(panelId)) panelMessageHandlers.set(panelId, new Set());
        const set = panelMessageHandlers.get(panelId);
        set.add(handler);
        return new DisposableImpl(() => {
          set.delete(handler);
          if (set.size === 0) panelMessageHandlers.delete(panelId);
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

  async getCell(row, col) {
    return rpcCall("cells", "getCell", [row, col]);
  },

  async setCell(row, col, value) {
    await rpcCall("cells", "setCell", [row, col, value]);
  }
};

const workbook = {
  async getActiveWorkbook() {
    return rpcCall("workbook", "getActiveWorkbook", []);
  }
};

const sheets = {
  async getActiveSheet() {
    return rpcCall("sheets", "getActiveSheet", []);
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

    commandHandlers.set(id, handler);
    await rpcCall("commands", "registerCommand", [id]);

    return new DisposableImpl(() => {
      commandHandlers.delete(id);
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

    customFunctionHandlers.set(fnName, def.handler);
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
      customFunctionHandlers.delete(fnName);
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

const events = {
  onSelectionChanged(callback) {
    return addEventHandler("selectionChanged", callback);
  },
  onCellChanged(callback) {
    return addEventHandler("cellChanged", callback);
  }
};

const context = new Proxy(
  {},
  {
    get(_target, prop) {
      if (prop === "extensionId") return currentContext.extensionId;
      if (prop === "extensionPath") return currentContext.extensionPath;
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
  ui,
  storage,
  config,
  events,
  context,

  __setTransport,
  __setContext,
  __handleMessage
};
