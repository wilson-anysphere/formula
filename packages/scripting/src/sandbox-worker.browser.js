import { buildModuleRunnerJavaScript, buildSandboxedScript, serializeError, transpileTypeScript } from "./sandbox.js";

let nextRpcId = 1;
/** @type {Map<number, { resolve: (value: any) => void, reject: (err: any) => void }>} */
const pendingRpc = new Map();

const originalFetch = self.fetch?.bind(self);
const OriginalWebSocket = self.WebSocket;

/**
 * @param {string} method
 * @param {any} params
 */
function hostRpc(method, params) {
  const id = nextRpcId++;
  return new Promise((resolve, reject) => {
    pendingRpc.set(id, { resolve, reject });
    self.postMessage({ type: "rpc", id, method, params });
  });
}

function installConsoleCapture() {
  const original = self.console;
  const send = (level, args) => {
    try {
      self.postMessage({ type: "console", level, message: args.map(String).join(" ") });
    } catch {
      // ignore
    }
    try {
      original?.[level]?.(...args);
    } catch {
      // ignore
    }
  };
  self.console = {
    log: (...args) => send("log", args),
    info: (...args) => send("info", args),
    warn: (...args) => send("warn", args),
    error: (...args) => send("error", args),
  };
}

/**
 * Best-effort permission enforcement for the browser worker.
 * In Node, enforcement happens in @formula/security's sandbox worker.
 *
 * @param {any} permissions
 */
function installPermissionGuards(permissions) {
  const mode = permissions?.network?.mode ?? "none";
  const allowlist = permissions?.network?.allowlist ?? [];

  // Prevent sandbox escapes by disabling powerful primitives that can spin up
  // nested execution contexts or bypass fetch/WebSocket guards. These are
  // best-effort overrides: in some environments these globals may be missing or
  // non-writable.
  for (const [key, value] of [
    ["Worker", undefined],
    ["XMLHttpRequest", undefined],
    ["importScripts", undefined],
  ]) {
    try {
      // eslint-disable-next-line no-param-reassign
      self[key] = value;
    } catch {
      try {
        Object.defineProperty(self, key, { value, configurable: true });
      } catch {
        // ignore
      }
    }
  }

  const isAllowed = (urlString) => {
    const url = new URL(urlString, self.location?.href ?? "https://localhost");
    const origin = url.origin;
    const host = url.hostname;

    for (const rawEntry of allowlist) {
      const entry = String(rawEntry ?? "").trim();
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
  };

  if (mode === "none") {
    self.fetch = async (url) => {
      const urlString = String(url);
      throw new Error(`Network access denied for ${urlString}`);
    };

    self.WebSocket = class BlockedWebSocket {
      constructor(url) {
        throw new Error(`Network access denied for ${String(url)}`);
      }
    };
    return;
  }

  if (mode === "allowlist") {
    self.fetch = async (url, init) => {
      if (!originalFetch) {
        throw new Error("fetch is not available in this environment");
      }
      const urlString = String(url);
      if (!isAllowed(urlString)) {
        throw new Error(`Network access denied for ${urlString}`);
      }
      return originalFetch(urlString, init);
    };

    self.WebSocket = class AllowlistWebSocket {
      constructor(url, protocols) {
        if (!OriginalWebSocket) {
          throw new Error("WebSocket is not available in this environment");
        }
        const urlString = String(url);
        if (!isAllowed(urlString)) {
          throw new Error(`Network access denied for ${urlString}`);
        }
        return new OriginalWebSocket(urlString, protocols);
      }
    };
    return;
  }

  if (originalFetch) {
    self.fetch = originalFetch;
  }
  if (OriginalWebSocket) {
    self.WebSocket = OriginalWebSocket;
  }
}

self.onmessage = (event) => {
  void (async () => {
    const message = event.data;
    if (!message || typeof message !== "object") return;

    if (message.type === "rpcResult") {
      const pending = pendingRpc.get(message.id);
      if (pending) {
        pendingRpc.delete(message.id);
        pending.resolve(message.result);
      }
      return;
    }

    if (message.type === "rpcError") {
      const pending = pendingRpc.get(message.id);
      if (pending) {
        pendingRpc.delete(message.id);
        pending.reject(message.error);
      }
      return;
    }

    if (message.type === "event") {
      try {
        const dispatch = self.__formulaDispatchEvent;
        if (typeof dispatch === "function") {
          dispatch(message.eventType, message.payload);
        }
      } catch (err) {
        self.postMessage({ type: "error", error: serializeError(err) });
      }
      return;
    }

    if (message.type === "cancel") {
      self.postMessage({ type: "error", error: { name: "AbortError", message: "Script execution cancelled" } });
      self.close();
      return;
    }

    if (message.type !== "run") return;

    try {
      installConsoleCapture();
      installPermissionGuards(message.permissions);
      // Expose host RPC as a capability for the generated bootstrap.
      self.__hostRpc = hostRpc;

      const { bootstrap, ts, moduleKind, kind } = buildSandboxedScript({
        code: String(message.code ?? ""),
        activeSheetName: String(message.activeSheetName),
        selection: message.selection,
      });

      const { js } = transpileTypeScript(ts, { moduleKind });
      const userScript = kind === "module" ? buildModuleRunnerJavaScript({ moduleJs: js }) : js;
      const fullScript = `${bootstrap}\n${userScript}`;
      const result = (0, eval)(fullScript);
      await result;
      self.postMessage({ type: "result", result: null });
    } catch (err) {
      self.postMessage({ type: "error", error: serializeError(err) });
    }
  })().catch((err) => {
    // `onmessage` handlers ignore returned promises; ensure we always terminate the
    // promise chain to avoid unhandled rejections.
    try {
      self.postMessage({ type: "error", error: serializeError(err) });
    } catch {
      // ignore
    }
  });
};
