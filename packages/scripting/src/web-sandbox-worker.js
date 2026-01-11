/* eslint-disable no-console */
import ts from "typescript";

// This file runs in a module WebWorker context (no Node APIs).
//
// Responsibilities:
// - Compile TypeScript user code to JavaScript inside the worker
// - Provide a minimal `ctx` object matching the Node scripting runtime
// - Forward workbook operations to the host via RPC
// - Capture console output in a structured form
// - Enforce a minimal network permission model (fetch/WebSocket)

let nextRpcId = 1;
const pendingRpc = new Map();
let runToken = null;
/** @type {MessagePort | null} */
let controlPort = null;

const postMessageToHost = self.postMessage.bind(self);

const originalFetch = self.fetch?.bind(self);
const OriginalWebSocket = self.WebSocket;

function postControlMessage(message) {
  if (!runToken) return;
  const payload = { ...message, token: runToken };
  if (controlPort) {
    controlPort.postMessage(payload);
  } else {
    postMessageToHost(payload);
  }
}

function rpc(method, params) {
  const id = nextRpcId++;
  return new Promise((resolve, reject) => {
    pendingRpc.set(id, { resolve, reject });
    postControlMessage({ type: "rpc", id, method, params });
  });
}

function serializeError(err) {
  if (err instanceof Error) {
    return { message: err.message, name: err.name, stack: err.stack };
  }
  if (typeof err === "string") {
    return { message: err };
  }
  return { message: "Unknown error" };
}

function formatConsoleArgs(args) {
  return args
    .map((arg) => {
      if (typeof arg === "string") return arg;
      if (arg instanceof Error) return arg.stack ?? arg.message;
      try {
        return JSON.stringify(arg);
      } catch {
        return String(arg);
      }
    })
    .join(" ");
}

function postConsole(level, args) {
  postControlMessage({ type: "console", level, message: formatConsoleArgs(args) });
}

const safeConsole = {
  log: (...args) => postConsole("log", args),
  info: (...args) => postConsole("info", args),
  warn: (...args) => postConsole("warn", args),
  error: (...args) => postConsole("error", args),
};

function applyNetworkSandbox(permissions) {
  const mode = permissions?.network ?? "none";

  if (mode === "none") {
    self.fetch = async () => {
      throw new Error("Network access is not permitted");
    };

    self.WebSocket = class BlockedWebSocket {
      constructor() {
        throw new Error("Network access is not permitted");
      }
    };
    return;
  }

  if (mode === "allowlist") {
    const allowlist = new Set(permissions?.networkAllowlist ?? []);
    self.fetch = async (input, init) => {
      if (!originalFetch) {
        throw new Error("fetch is not available in this environment");
      }
      const url = typeof input === "string" ? input : input?.url;
      const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
      if (!allowlist.has(hostname)) {
        throw new Error(`Network access to ${hostname} is not permitted`);
      }
      return originalFetch(input, init);
    };

    self.WebSocket = class AllowlistWebSocket {
      constructor(url, protocols) {
        if (!OriginalWebSocket) {
          throw new Error("WebSocket is not available in this environment");
        }
        const hostname = new URL(url, self.location?.href ?? "https://localhost").hostname;
        if (!allowlist.has(hostname)) {
          throw new Error(`Network access to ${hostname} is not permitted`);
        }
        return new OriginalWebSocket(url, protocols);
      }
    };
    return;
  }

  // full access
  if (originalFetch) {
    self.fetch = originalFetch;
  }
  if (OriginalWebSocket) {
    self.WebSocket = OriginalWebSocket;
  }
}

function createRangeProxy(sheetName, address) {
  return {
    address,
    getValues: () => rpc("range.getValues", { sheetName, address }),
    setValues: (values) => rpc("range.setValues", { sheetName, address, values }),
    getValue: () => rpc("range.getValue", { sheetName, address }),
    setValue: (value) => rpc("range.setValue", { sheetName, address, value }),
    getFormat: () => rpc("range.getFormat", { sheetName, address }),
    setFormat: (format) => rpc("range.setFormat", { sheetName, address, format }),
  };
}

function createSheetProxy(name) {
  return {
    name,
    getRange: (address) => createRangeProxy(name, address),
  };
}

function createWorkbookProxy() {
  return {
    getSheet: (name) => createSheetProxy(name),
    getActiveSheetName: () => rpc("workbook.getActiveSheetName", null),
    getSelection: () => rpc("workbook.getSelection", null),
    setSelection: (sheetName, address) => rpc("workbook.setSelection", { sheetName, address }),
  };
}

function createUiProxy() {
  return {
    log: (...args) => safeConsole.log(...args),
    alert: (message) => rpc("ui.alert", { message }),
    confirm: (message) => rpc("ui.confirm", { message }),
    prompt: (message, defaultValue) => rpc("ui.prompt", { message, defaultValue }),
  };
}

function compileTypeScript(tsSource) {
  // Wrap user code in an async entrypoint so scripts can freely use top-level
  // `await` (relative to the script body).
  //
  // Note: the host awaits the returned promise. Do not call the function in the
  // emitted source; instead return it from the runner so we can `await` it.
  const wrapped = `async function __formulaUserMain(ctx) {\n${tsSource}\n}\n`;
  const result = ts.transpileModule(wrapped, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.None,
    },
    reportDiagnostics: true,
    fileName: "user-script.ts",
  });

  const diagnostics = (result.diagnostics || []).filter((d) => d.category === ts.DiagnosticCategory.Error);
  if (diagnostics.length > 0) {
    const formatHost = {
      getCanonicalFileName: (f) => f,
      getCurrentDirectory: () => "",
      getNewLine: () => "\n",
    };
    const message = ts.formatDiagnostics(diagnostics, formatHost);
    throw new Error(message);
  }

  return result.outputText;
}

function compileTypeScriptModule(tsSource) {
  const result = ts.transpileModule(tsSource, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.CommonJS,
    },
    reportDiagnostics: true,
    fileName: "user-script.ts",
  });

  const diagnostics = (result.diagnostics || []).filter((d) => d.category === ts.DiagnosticCategory.Error);
  if (diagnostics.length > 0) {
    const formatHost = {
      getCanonicalFileName: (f) => f,
      getCurrentDirectory: () => "",
      getNewLine: () => "\n",
    };
    const message = ts.formatDiagnostics(diagnostics, formatHost);
    throw new Error(message);
  }

  return result.outputText;
}

async function runUserScript({ code, activeSheetName, selection, permissions }) {
  applyNetworkSandbox(permissions ?? {});

  const isModule = /\bexport\s+default\b/.test(code);
  const jsSource = isModule
    ? `${compileTypeScriptModule(code)}\n//# sourceURL=user-script.js`
    : `${compileTypeScript(code)}\n//# sourceURL=user-script.js\nreturn __formulaUserMain(ctx);`;

  const ctx = {
    workbook: createWorkbookProxy(),
    activeSheet: createSheetProxy(activeSheetName),
    selection: createRangeProxy(selection.sheetName, selection.address),
    ui: createUiProxy(),
    alert: (message) => rpc("ui.alert", { message }),
    confirm: (message) => rpc("ui.confirm", { message }),
    prompt: (message, defaultValue) => rpc("ui.prompt", { message, defaultValue }),
    fetch: self.fetch?.bind(self),
    console: safeConsole,
  };

  const runner = isModule
    ? new Function(
        "ctx",
        "console",
        "setTimeout",
        "clearTimeout",
        "setInterval",
        "clearInterval",
        `"use strict";\nconst exports = {};\nconst module = { exports };\n${jsSource}\nconst __formulaMain = module.exports?.default ?? exports.default;\nif (typeof __formulaMain !== \"function\") { throw new Error(\"Script must export a default function\"); }\nreturn __formulaMain(ctx);`,
      )
    : new Function(
        "ctx",
        "console",
        "setTimeout",
        "clearTimeout",
        "setInterval",
        "clearInterval",
        `"use strict";\n${jsSource}`,
      );

  const result = runner(ctx, safeConsole, setTimeout, clearTimeout, setInterval, clearInterval);
  await result;
}

function handleHostMessage(message) {
  if (!message || message.token !== runToken) return;

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
      const error = new Error(message.error?.message || "RPC error");
      error.name = message.error?.name || error.name;
      error.stack = message.error?.stack || error.stack;
      pending.reject(error);
    }
  }
}

self.onmessage = async (event) => {
  const message = event.data;

  if (message && message.type === "run") {
    runToken = message.token;
    controlPort = message.controlPort ?? event.ports?.[0] ?? null;
    if (controlPort) {
      controlPort.onmessage = (portEvent) => handleHostMessage(portEvent.data);
    }

    try {
      await runUserScript(message);
      postControlMessage({ type: "result" });
    } catch (err) {
      postControlMessage({ type: "error", error: serializeError(err) });
    }
    return;
  }

  if (!controlPort) {
    handleHostMessage(message);
  }
};
