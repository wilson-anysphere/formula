import * as formulaApi from "@formula/extension-api";

formulaApi.__setTransport({
  postMessage: (message) => postMessage(message)
});

let workerData = null;
let extensionModule = null;
let activated = false;
let activationPromise = null;

let nextInternalRequestId = 1;
const internalPending = new Map();

function createInternalRequestId() {
  return `__internal__${nextInternalRequestId++}`;
}

function deserializeError(payload) {
  const message =
    typeof payload === "string" ? payload : String(payload?.message ?? "Unknown error");
  const err = new Error(message);
  if (payload?.stack) err.stack = String(payload.stack);
  return err;
}

function internalRpcCall(namespace, method, args) {
  const id = createInternalRequestId();
  postMessage({
    type: "api_call",
    id,
    namespace,
    method,
    args
  });

  return new Promise((resolve, reject) => {
    internalPending.set(id, { resolve, reject });
  });
}

function handleInternalResponse(message) {
  if (!message || typeof message !== "object") return false;
  if (message.type !== "api_result" && message.type !== "api_error") return false;
  const pending = internalPending.get(message.id);
  if (!pending) return false;
  internalPending.delete(message.id);
  if (message.type === "api_result") {
    pending.resolve(message.result);
  } else {
    pending.reject(deserializeError(message.error));
  }
  return true;
}

function init(data) {
  workerData = {
    extensionId: String(data?.extensionId ?? ""),
    extensionPath: String(data?.extensionPath ?? ""),
    extensionUri: String(data?.extensionUri ?? data?.extensionPath ?? ""),
    globalStoragePath: String(data?.globalStoragePath ?? ""),
    workspaceStoragePath: String(data?.workspaceStoragePath ?? ""),
    mainUrl: String(data?.mainUrl ?? "")
  };

  formulaApi.__setContext({
    extensionId: workerData.extensionId,
    extensionPath: workerData.extensionPath,
    extensionUri: workerData.extensionUri,
    globalStoragePath: workerData.globalStoragePath,
    workspaceStoragePath: workerData.workspaceStoragePath
  });
}

if (typeof globalThis.fetch === "function") {
  globalThis.fetch = async (input, init) => {
    return formulaApi.network.fetch(String(input), init);
  };
}

if (typeof globalThis.WebSocket === "function") {
  const NativeWebSocket = globalThis.WebSocket;

  class PermissionedWebSocket {
    static CONNECTING = 0;
    static OPEN = 1;
    static CLOSING = 2;
    static CLOSED = 3;

    CONNECTING = 0;
    OPEN = 1;
    CLOSING = 2;
    CLOSED = 3;

    constructor(url, protocols) {
      this._url = String(url ?? "");
      this._protocols = protocols;
      this._ws = null;
      this._readyState = PermissionedWebSocket.CONNECTING;
      this._binaryType = "blob";
      this._protocol = "";
      this._extensions = "";
      this._bufferedAmount = 0;
      this._pendingClose = null;
      /** @type {Map<string, Set<Function>>} */
      this._listeners = new Map();

      this.onopen = null;
      this.onmessage = null;
      this.onerror = null;
      this.onclose = null;

      void this._start();
    }

    get url() {
      return this._ws ? this._ws.url : this._url;
    }

    get readyState() {
      return this._ws ? this._ws.readyState : this._readyState;
    }

    get bufferedAmount() {
      return this._ws ? this._ws.bufferedAmount : this._bufferedAmount;
    }

    get extensions() {
      return this._ws ? this._ws.extensions : this._extensions;
    }

    get protocol() {
      return this._ws ? this._ws.protocol : this._protocol;
    }

    get binaryType() {
      return this._ws ? this._ws.binaryType : this._binaryType;
    }

    set binaryType(value) {
      this._binaryType = value;
      if (this._ws) this._ws.binaryType = value;
    }

    addEventListener(type, listener) {
      if (typeof listener !== "function") return;
      const key = String(type);
      let set = this._listeners.get(key);
      if (!set) {
        set = new Set();
        this._listeners.set(key, set);
      }
      set.add(listener);
    }

    removeEventListener(type, listener) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      set.delete(listener);
      if (set.size === 0) this._listeners.delete(String(type));
    }

    dispatchEvent(event) {
      if (!event || typeof event.type !== "string") return true;
      this._emit(event.type, event);
      return true;
    }

    send(data) {
      if (this.readyState !== PermissionedWebSocket.OPEN || !this._ws) {
        throw new Error("WebSocket is not open");
      }

      this._ws.send(data);
    }

    close(code, reason) {
      if (!this._ws) {
        this._pendingClose = { code, reason };
        this._readyState = PermissionedWebSocket.CLOSING;
        return;
      }
      this._ws.close(code, reason);
    }

    async _start() {
      try {
        await internalRpcCall("network", "openWebSocket", [this._url]);
      } catch (err) {
        this._fail(err);
        return;
      }

      let ws;
      try {
        ws =
          this._protocols === undefined ? new NativeWebSocket(this._url) : new NativeWebSocket(this._url, this._protocols);
      } catch (err) {
        this._fail(err);
        return;
      }

      this._ws = ws;
      try {
        ws.binaryType = this._binaryType;
      } catch {
        // ignore
      }

      ws.addEventListener("open", () => {
        this._readyState = PermissionedWebSocket.OPEN;
        this._protocol = ws.protocol;
        this._extensions = ws.extensions;
        this._emit("open", { type: "open", target: this });

        const pendingClose = this._pendingClose;
        if (pendingClose) {
          this._pendingClose = null;
          try {
            ws.close(pendingClose.code, pendingClose.reason);
          } catch {
            // ignore
          }
        }
      });

      ws.addEventListener("message", (event) => {
        this._emit("message", { type: "message", data: event.data, target: this });
      });

      ws.addEventListener("error", () => {
        this._emit("error", { type: "error", target: this });
      });

      ws.addEventListener("close", (event) => {
        this._readyState = PermissionedWebSocket.CLOSED;
        this._emit("close", {
          type: "close",
          code: event.code,
          reason: event.reason,
          wasClean: event.wasClean,
          target: this
        });
      });
    }

    _fail(err) {
      this._readyState = PermissionedWebSocket.CLOSED;
      this._emit("error", { type: "error", error: err, target: this });
      this._emit("close", { type: "close", code: 1008, reason: String(err?.message ?? err), wasClean: false, target: this });
    }

    _emit(type, event) {
      const evt = event && typeof event === "object" ? event : { type };
      const propHandler = this[`on${type}`];
      if (typeof propHandler === "function") {
        try {
          propHandler.call(this, evt);
        } catch {
          // ignore
        }
      }

      const set = this._listeners.get(String(type));
      if (!set) return;
      for (const listener of [...set]) {
        try {
          listener.call(this, evt);
        } catch {
          // ignore
        }
      }
    }
  }

  globalThis.WebSocket = PermissionedWebSocket;
}

async function activateExtension() {
  if (activated) return;
  if (activationPromise) return activationPromise;
  if (!workerData) throw new Error("Extension worker not initialized");

  activationPromise = (async () => {
    if (!extensionModule) {
      extensionModule = await import(workerData.mainUrl);
    }

    const activateFn = extensionModule.activate || extensionModule.default?.activate;
    if (typeof activateFn !== "function") {
      throw new Error(`Extension entrypoint does not export an activate() function`);
    }

    const context = {
      extensionId: workerData.extensionId,
      extensionPath: workerData.extensionPath,
      extensionUri: workerData.extensionUri,
      globalStoragePath: workerData.globalStoragePath,
      workspaceStoragePath: workerData.workspaceStoragePath,
      subscriptions: []
    };

    await activateFn(context);
    activated = true;
  })();

  try {
    await activationPromise;
  } finally {
    if (!activated) activationPromise = null;
  }
}

self.addEventListener("message", async (event) => {
  const message = event.data;
  if (!message || typeof message !== "object") return;

  if (message.type === "init") {
    init(message);
    return;
  }

  if (message.type === "activate") {
    try {
      await activateExtension();
      postMessage({ type: "activate_result", id: message.id });
    } catch (error) {
      postMessage({
        type: "activate_error",
        id: message.id,
        error: { message: String(error?.message ?? error), stack: error?.stack }
      });
    }
    return;
  }

  if (handleInternalResponse(message)) {
    return;
  }

  formulaApi.__handleMessage(message);
});
