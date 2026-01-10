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

async function runUserScript(tsSource) {
  const jsSource = compileTypeScript(tsSource);

  const ctx = {
    workbook: createWorkbookProxy(),
    activeSheet: createSheetProxy(workerData.activeSheetName),
    selection: createRangeProxy(workerData.selection.sheetName, workerData.selection.address),
    ui: {
      log: (...args) => safeConsole.log(...args),
    },
  };

  const sandbox = {
    ctx,
    console: safeConsole,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
  };
  sandbox.globalThis = sandbox;

  const context = vm.createContext(sandbox, {
    codeGeneration: { strings: false, wasm: false },
  });

  const script = new vm.Script(jsSource, { filename: "user-script.js" });
  const result = script.runInContext(context);
  await result;
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
