/* eslint-disable no-console */
const { parentPort, workerData } = require("node:worker_threads");
const vm = require("node:vm");
const util = require("node:util");
let ts = null;
try {
  // The TypeScript compiler is a devDependency in the monorepo; when running in
  // minimal environments (like CI/unit tests without `pnpm install`) it may be
  // absent. In that case we still allow scripts that are valid JavaScript to run
  // by skipping transpilation.
  // eslint-disable-next-line global-require
  ts = require("typescript");
} catch {
  ts = null;
}

if (!parentPort) {
  throw new Error("sandbox-worker must be run as a worker thread");
}

const originalFetch = globalThis.fetch;
const OriginalWebSocket = globalThis.WebSocket;

let nextRpcId = 1;
const pendingRpc = new Map();

function rpc(method, params) {
  const id = nextRpcId++;
  return new Promise((resolve, reject) => {
    pendingRpc.set(id, { resolve, reject });
    parentPort.postMessage({ type: "rpc", id, method, params });
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

function postConsole(level, args) {
  parentPort.postMessage({
    type: "console",
    level,
    message: util.format(...args),
  });
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
    return {
      fetch: async () => {
        throw new Error("Network access is not permitted");
      },
      WebSocket: class BlockedWebSocket {
        constructor() {
          throw new Error("Network access is not permitted");
        }
      },
    };
  }

  if (mode === "allowlist") {
    const allowlist = new Set(permissions?.networkAllowlist ?? []);
    return {
      fetch: async (input, init) => {
        if (!originalFetch) {
          throw new Error("fetch is not available in this environment");
        }
        const url = typeof input === "string" ? input : input?.url;
        const hostname = new URL(url, "https://localhost").hostname;
        if (!allowlist.has(hostname)) {
          throw new Error(`Network access to ${hostname} is not permitted`);
        }
        return originalFetch(input, init);
      },
      WebSocket: class AllowlistWebSocket {
        constructor(url, protocols) {
          if (!OriginalWebSocket) {
            throw new Error("WebSocket is not available in this environment");
          }
          const hostname = new URL(url, "https://localhost").hostname;
          if (!allowlist.has(hostname)) {
            throw new Error(`Network access to ${hostname} is not permitted`);
          }
          return new OriginalWebSocket(url, protocols);
        }
      },
    };
  }

  // full access
  return {
    fetch: originalFetch,
    WebSocket: OriginalWebSocket,
  };
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
  const wrapped = `async function __formulaUserMain(ctx) {\n${tsSource}\n}\n__formulaUserMain(ctx)`;
  if (!ts) return wrapped;
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

function isModuleScript(tsSource) {
  if (ts) {
    const sourceFile = ts.createSourceFile("user-script.ts", tsSource, ts.ScriptTarget.ES2022, true);
    for (const stmt of sourceFile.statements) {
      if (
        ts.isImportDeclaration(stmt) ||
        ts.isImportEqualsDeclaration(stmt) ||
        ts.isExportAssignment(stmt) ||
        ts.isExportDeclaration(stmt)
      ) {
        return true;
      }
      const mods = stmt.modifiers;
      if (mods && mods.some((m) => m.kind === ts.SyntaxKind.ExportKeyword || m.kind === ts.SyntaxKind.DefaultKeyword)) {
        return true;
      }
    }
    return false;
  }

  // Best-effort fallback. When `typescript` isn't available we only support the
  // simplest module scripts (an `export default` function with no imports), but
  // we still want to classify `import`/`export` usage as module code so we can
  // emit a clear "imports are not supported" error instead of a syntax error.
  return /^\s*(import|export)\b/m.test(tsSource);
}

function findDynamicImportSpecifier(tsSource) {
  if (ts) {
    const sourceFile = ts.createSourceFile("user-script.ts", tsSource, ts.ScriptTarget.ES2022, true);
    let found = null;

    function visit(node) {
      if (found) return;
      if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
        const arg = node.arguments[0];
        if (arg && ts.isStringLiteral(arg)) {
          found = arg.text;
        } else {
          found = "<dynamic>";
        }
        return;
      }
      ts.forEachChild(node, visit);
    }

    visit(sourceFile);
    return found;
  }

  // Best-effort fallback (only used in minimal environments without TypeScript).
  if (/\bimport\s*\(/.test(tsSource)) return "<dynamic>";
  return null;
}

function compileTypeScriptModule(tsSource) {
  if (!ts) {
    // Fallback: support scripts that use ESM `export default` but do not require
    // TypeScript syntax (i.e. are valid JavaScript aside from the export keyword).
    //
    // This keeps recorded macros runnable in minimal environments where the
    // `typescript` dependency is unavailable.
    const importMatch =
      /^\s*import\s+[^"']*["']([^"']+)["']/m.exec(tsSource) ??
      /^\s*import\s*\(\s*["']([^"']+)["']\s*\)/m.exec(tsSource);
    if (importMatch) {
      throw new Error(`Imports are not supported in scripts (attempted to import ${importMatch[1]})`);
    }

    if (/\bexport\s+(?!default\b)/.test(tsSource)) {
      throw new Error("Only `export default` is supported in module scripts when the TypeScript compiler is unavailable");
    }

    return tsSource.replace(/\bexport\s+default\s+/, "exports.default = ");
  }

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

async function runUserScript(tsSource) {
  const isModule = isModuleScript(tsSource);
  const dynamicImportSpecifier = findDynamicImportSpecifier(tsSource);
  if (dynamicImportSpecifier) {
    if (dynamicImportSpecifier === "<dynamic>") {
      throw new Error("Imports are not supported in scripts (attempted to use dynamic import())");
    }
    throw new Error(
      `Imports are not supported in scripts (attempted to use dynamic import(${JSON.stringify(dynamicImportSpecifier)}))`,
    );
  }
  const jsSource = isModule ? compileTypeScriptModule(tsSource) : compileTypeScript(tsSource);

  const networkSandbox = applyNetworkSandbox(workerData.permissions ?? {});

  const ctx = {
    workbook: createWorkbookProxy(),
    activeSheet: createSheetProxy(workerData.activeSheetName),
    selection: createRangeProxy(workerData.selection.sheetName, workerData.selection.address),
    ui: createUiProxy(),
    alert: (message) => rpc("ui.alert", { message }),
    confirm: (message) => rpc("ui.confirm", { message }),
    prompt: (message, defaultValue) => rpc("ui.prompt", { message, defaultValue }),
    fetch: networkSandbox.fetch,
    console: safeConsole,
  };

  const sandbox = {
    exports: {},
    require: (specifier) => {
      throw new Error(`Imports are not supported in scripts (attempted to require ${String(specifier)})`);
    },
    ctx,
    console: safeConsole,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    fetch: networkSandbox.fetch,
    WebSocket: networkSandbox.WebSocket,
  };
  sandbox.module = { exports: sandbox.exports };
  sandbox.globalThis = sandbox;

  const context = vm.createContext(sandbox, {
    codeGeneration: { strings: false, wasm: false },
  });

  const script = new vm.Script(jsSource, { filename: "user-script.js" });
  const result = script.runInContext(context);

  if (!isModule) {
    await result;
    return;
  }

  // Module-style script: `export default async function main(ctx) { ... }`.
  const exported = sandbox.module?.exports?.default ?? sandbox.exports?.default;
  if (typeof exported !== "function") {
    throw new Error("Script must export a default function");
  }

  await exported(ctx);
}

parentPort.on("message", async (message) => {
  if (message && message.type === "run") {
    try {
      await runUserScript(message.code);
      parentPort.postMessage({ type: "result" });
    } catch (err) {
      parentPort.postMessage({ type: "error", error: serializeError(err) });
    }
    return;
  }

  if (message && message.type === "rpcResult") {
    const pending = pendingRpc.get(message.id);
    if (pending) {
      pendingRpc.delete(message.id);
      pending.resolve(message.result);
    }
    return;
  }

  if (message && message.type === "rpcError") {
    const pending = pendingRpc.get(message.id);
    if (pending) {
      pendingRpc.delete(message.id);
      const error = new Error(message.error?.message || "RPC error");
      error.name = message.error?.name || error.name;
      error.stack = message.error?.stack || error.stack;
      pending.reject(error);
    }
  }
});
