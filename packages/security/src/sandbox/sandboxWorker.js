import vm from "node:vm";
import { parentPort } from "node:worker_threads";

import { createSandboxSecureApis } from "../secureApis/createSecureApis.js";

if (!parentPort) {
  throw new Error("sandboxWorker must be run as a worker thread");
}

const MAX_ERROR_MESSAGE_BYTES = 16 * 1024;
const MAX_ERROR_STACK_BYTES = 64 * 1024;

function truncateUtf8(text, maxBytes) {
  const str = typeof text === "string" ? text : String(text);
  if (Buffer.byteLength(str) <= maxBytes) return str;

  let end = Math.min(str.length, maxBytes);
  let truncated = str.slice(0, end);
  while (end > 0 && Buffer.byteLength(truncated) > maxBytes) {
    end = Math.floor(end * 0.9);
    truncated = str.slice(0, end);
  }
  return `${truncated}…`;
}

function serializeError(error) {
  if (!error || typeof error !== "object") return { name: "Error", message: String(error) };

  return {
    name: typeof error.name === "string" ? error.name : "Error",
    message:
      typeof error.message === "string"
        ? truncateUtf8(error.message, MAX_ERROR_MESSAGE_BYTES)
        : truncateUtf8(String(error), MAX_ERROR_MESSAGE_BYTES),
    code: typeof error.code === "string" ? error.code : undefined,
    stack: typeof error.stack === "string" ? truncateUtf8(error.stack, MAX_ERROR_STACK_BYTES) : undefined,
    principal: error.principal,
    request: error.request,
    reason: error.reason
  };
}

function deserializeHostError(payload) {
  const err = new Error(payload?.message ?? "RPC error");
  err.name = payload?.name ?? "Error";
  if (payload?.code) err.code = payload.code;
  if (payload?.stack) err.stack = payload.stack;
  err.principal = payload?.principal;
  err.request = payload?.request;
  err.reason = payload?.reason;
  return err;
}

function sendAudit(event) {
  parentPort.postMessage({
    type: "audit",
    event
  });
}

function formatConsoleArgs(args) {
  return args
    .map((value) => {
      if (typeof value === "string") return value;
      if (value instanceof Error) return value.stack ?? value.message;
      try {
        return JSON.stringify(value);
      } catch {
        return String(value);
      }
    })
    .join(" ");
}

function hardenFunction(fn) {
  if (typeof fn !== "function") return fn;
  Object.setPrototypeOf(fn, null);
  return fn;
}

function hardenObject(obj) {
  if (!obj || typeof obj !== "object") return obj;
  Object.setPrototypeOf(obj, null);
  return Object.freeze(obj);
}

function createSandboxConsole() {
  const make = (method, stream) =>
    hardenFunction((...args) => {
      parentPort.postMessage({
        type: "output",
        stream,
        text: `${formatConsoleArgs(args)}\n`,
        metadata: { method }
      });
    });

  return hardenObject({
    log: make("log", "stdout"),
    info: make("info", "stdout"),
    warn: make("warn", "stderr"),
    error: make("error", "stderr")
  });
}

function createSandboxTimers() {
  let nextId = 1;
  const active = new Map();

  const setTimeoutSafe = hardenFunction((callback, delay, ...args) => {
    const id = nextId++;
    const handle = setTimeout(() => {
      active.delete(id);
      try {
        callback(...args);
      } catch {
        // Timer callback errors are surfaced via the script promise / parent timeout.
      }
    }, delay);
    active.set(id, handle);
    return id;
  });

  const clearTimeoutSafe = hardenFunction((id) => {
    const handle = active.get(id);
    if (!handle) return;
    active.delete(id);
    clearTimeout(handle);
  });

  const cleanup = () => {
    for (const handle of active.values()) clearTimeout(handle);
    active.clear();
  };

  return { setTimeout: setTimeoutSafe, clearTimeout: clearTimeoutSafe, cleanup };
}

let nextRpcId = 1;
const pendingRpc = new Map();

/** @type {vm.Context | null} */
let currentContext = null;

function makeHostRpc() {
  return hardenFunction((method, params) => {
    const id = nextRpcId++;
    return new Promise((resolve, reject) => {
      pendingRpc.set(id, { resolve, reject });
      parentPort.postMessage({ type: "rpc", id, method, params });
    });
  });
}

parentPort.on("message", (message) => {
  void (async () => {
    if (!message || typeof message !== "object") return;

    if (message.type === "rpcResult") {
      const pending = pendingRpc.get(message.id);
      if (!pending) return;
      pendingRpc.delete(message.id);
      pending.resolve(message.result);
      return;
    }

    if (message.type === "rpcError") {
      const pending = pendingRpc.get(message.id);
      if (!pending) return;
      pendingRpc.delete(message.id);
      pending.reject(deserializeHostError(message.error));
      return;
    }

    if (message.type === "event") {
      if (!currentContext) return;
      try {
        const dispatch = currentContext.__formulaDispatchEvent;
        if (typeof dispatch === "function") {
          dispatch(message.eventType, message.payload);
        }
      } catch (error) {
        try {
          parentPort.postMessage({ type: "error", error: serializeError(error) });
        } catch {
          // ignore
        }
      }
      return;
    }

    if (message.type !== "run") return;

    const {
      principal,
      permissions,
      code,
      timeoutMs,
      memoryMb,
      enableHostRpc = false,
      captureConsole = true,
      wrapAsyncIife = true
    } = message;

    /** @type {ReturnType<typeof createSandboxTimers> | null} */
    let timers = null;
    let memoryInterval = null;
    let memoryLimitSent = false;

    try {
      const auditLogger = {
        log(event) {
          sendAudit(event);
          return "";
        }
      };

      const secureApis = createSandboxSecureApis({
        principal,
        permissionSnapshot: permissions,
        auditLogger
      });

      const consoleShim = captureConsole ? createSandboxConsole() : console;
      timers = createSandboxTimers();

      if (Number.isFinite(memoryMb) && memoryMb > 0) {
        const memoryLimitBytes = memoryMb * 1024 * 1024;
        const triggerBytes = Math.floor(memoryLimitBytes * 0.9);
        memoryInterval = setInterval(() => {
          if (memoryLimitSent) return;
          const { heapUsed, external } = process.memoryUsage();
          const usedBytes = heapUsed + external;
          if (usedBytes > triggerBytes) {
            memoryLimitSent = true;
            try {
              parentPort.postMessage({
                type: "limit",
                limit: "memory",
                memoryMb,
                usedMb: Math.round(usedBytes / 1024 / 1024),
                heapUsedMb: Math.round(heapUsed / 1024 / 1024),
                externalMb: Math.round(external / 1024 / 1024)
              });
            } catch {
              // ignore
            }
          }
        }, 25);
        memoryInterval.unref();
      }

      const sandbox = Object.create(null);
      sandbox.SecureApis = secureApis;
      sandbox.fetch = secureApis.fetch;
      sandbox.fs = secureApis.fs;
      sandbox.console = consoleShim;
      sandbox.setTimeout = timers.setTimeout;
      sandbox.clearTimeout = timers.clearTimeout;

      if (enableHostRpc) {
        sandbox.__hostRpc = makeHostRpc();
      }

      const context = vm.createContext(sandbox, {
        name: "formula-sandbox",
        codeGeneration: { strings: false, wasm: false }
      });

      currentContext = context;
      try {
        const source = wrapAsyncIife ? `(async () => {\n${code}\n})()` : String(code);
        const script = new vm.Script(source, { filename: "sandboxed-script.js" });

        // vm's timeout only applies to the initial synchronous execution of the script.
        // The parent thread additionally enforces an overall timeout and will terminate
        // the worker if the promise does not settle.
        const result = await script.runInContext(context, { timeout: timeoutMs });
        try {
          parentPort.postMessage({ type: "result", result });
        } catch {
          // ignore
        }
      } catch (error) {
        try {
          parentPort.postMessage({ type: "error", error: serializeError(error) });
        } catch {
          // ignore
        }
      } finally {
        currentContext = null;
      }
    } catch (error) {
      try {
        parentPort.postMessage({ type: "error", error: serializeError(error) });
      } catch {
        // ignore
      }
      currentContext = null;
    } finally {
      try {
        timers?.cleanup?.();
      } catch {
        // ignore
      }
      try {
        if (memoryInterval) clearInterval(memoryInterval);
      } catch {
        // ignore
      }
    }
  })().catch((error) => {
    // Message handlers ignore returned promises; swallow errors so we never
    // produce unhandled rejections in the worker.
    try {
      parentPort.postMessage({ type: "error", error: serializeError(error) });
    } catch {
      // ignore
    }
  });
});
