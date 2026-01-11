import vm from "node:vm";
import { parentPort } from "node:worker_threads";

import { createSandboxSecureApis } from "../secureApis/createSecureApis.js";

function serializeError(error) {
  if (!error || typeof error !== "object") return { name: "Error", message: String(error) };

  return {
    name: typeof error.name === "string" ? error.name : "Error",
    message: typeof error.message === "string" ? error.message : String(error),
    code: typeof error.code === "string" ? error.code : undefined,
    stack: typeof error.stack === "string" ? error.stack : undefined,
    principal: error.principal,
    request: error.request,
    reason: error.reason
  };
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

function createSandboxConsole({ principal }) {
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
        // Timer callback errors are surfaced via the sandbox promise / parent timeout.
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

parentPort.on("message", async (message) => {
  if (!message || typeof message !== "object" || message.type !== "run") return;

  const { principal, permissions, code, timeoutMs, memoryMb } = message;

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

  const consoleShim = createSandboxConsole({ principal });
  const timers = createSandboxTimers();

  let memoryInterval = null;
  let memoryLimitSent = false;
  if (Number.isFinite(memoryMb) && memoryMb > 0) {
    const memoryLimitBytes = memoryMb * 1024 * 1024;
    const triggerBytes = Math.floor(memoryLimitBytes * 0.9);
    memoryInterval = setInterval(() => {
      if (memoryLimitSent) return;
      const heapUsed = process.memoryUsage().heapUsed;
      if (heapUsed > triggerBytes) {
        memoryLimitSent = true;
        parentPort.postMessage({
          type: "limit",
          limit: "memory",
          memoryMb,
          usedMb: Math.round(heapUsed / 1024 / 1024)
        });
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
  sandbox.process = undefined;
  sandbox.require = undefined;
  sandbox.module = undefined;

  const context = vm.createContext(sandbox, {
    name: "formula-sandbox",
    codeGeneration: { strings: false, wasm: false }
  });

  try {
    const wrapped = `(async () => {\n${code}\n})()`;
    const script = new vm.Script(wrapped, { filename: "sandboxed-script.js" });

    // vm's timeout only applies to the initial synchronous execution of the script.
    // The parent thread additionally enforces an overall timeout and will terminate
    // the worker if the promise does not settle.
    const result = await script.runInContext(context, { timeout: timeoutMs });
    parentPort.postMessage({ type: "result", result });
  } catch (error) {
    parentPort.postMessage({ type: "error", error: serializeError(error) });
  } finally {
    timers.cleanup();
    if (memoryInterval) clearInterval(memoryInterval);
  }
});
