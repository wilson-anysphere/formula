const { parentPort, workerData } = require("node:worker_threads");
const Module = require("node:module");
const path = require("node:path");

const formulaApi = require(workerData.apiModulePath);
formulaApi.__setTransport({
  postMessage: (message) => parentPort.postMessage(message)
});
formulaApi.__setContext({
  extensionId: workerData.extensionId,
  extensionPath: workerData.extensionPath
});

// Replace global fetch with a permission-gated proxy through the Formula API.
// This provides a VS Code-like "declare + prompt" model for outbound network access.
if (typeof globalThis.fetch === "function") {
  globalThis.fetch = async (input, init) => {
    return formulaApi.network.fetch(String(input), init);
  };
}

// Provide a VS Code-like virtual module for extension authors.
const originalLoad = Module._load;
const extensionRoot = path.resolve(workerData.extensionPath);
const deniedBuiltins = new Set([
  "fs",
  "child_process",
  "worker_threads",
  "cluster",
  "net",
  "tls",
  "dgram",
  "dns",
  "http",
  "https",
  "module",
  "vm"
]);

Module._load = function (request, parent, isMain) {
  if (request === "@formula/extension-api" || request === "formula") {
    return formulaApi;
  }

  const parentFilename = parent?.filename ? path.resolve(parent.filename) : null;
  const isExtensionRequest = parentFilename ? parentFilename.startsWith(extensionRoot + path.sep) : false;

  if (isExtensionRequest) {
    const normalized = typeof request === "string" && request.startsWith("node:")
      ? request.slice("node:".length)
      : request;

    if (Module.builtinModules.includes(normalized) && deniedBuiltins.has(normalized)) {
      throw new Error(`Access to Node builtin module '${normalized}' is not allowed in extensions`);
    }

    const resolved = Module._resolveFilename(request, parent, isMain);
    if (typeof resolved === "string" && path.isAbsolute(resolved)) {
      const resolvedPath = path.resolve(resolved);
      if (!resolvedPath.startsWith(extensionRoot + path.sep)) {
        throw new Error(
          `Extensions cannot require modules outside their extension folder: '${request}' resolved to '${resolvedPath}'`
        );
      }
    }
  }

  return originalLoad.call(this, request, parent, isMain);
};

function safeSerializeLogArg(arg) {
  if (arg instanceof Error) {
    return { error: { message: arg.message, stack: arg.stack } };
  }
  if (typeof arg === "string") return arg;
  if (typeof arg === "number" || typeof arg === "boolean" || arg === null) return arg;
  try {
    return JSON.parse(JSON.stringify(arg));
  } catch {
    return String(arg);
  }
}

for (const level of ["log", "info", "warn", "error"]) {
  const original = console[level];
  console[level] = (...args) => {
    try {
      parentPort.postMessage({
        type: "log",
        level,
        args: args.map(safeSerializeLogArg)
      });
    } catch {
      // ignore
    }
    return original.apply(console, args);
  };
}

let extensionModule = null;
let activated = false;

async function activateExtension() {
  if (activated) return;
  if (!extensionModule) {
    extensionModule = require(workerData.mainPath);
  }

  const activateFn = extensionModule.activate || extensionModule.default?.activate;
  if (typeof activateFn !== "function") {
    throw new Error(`Extension entrypoint does not export an activate() function`);
  }

  const context = {
    extensionId: workerData.extensionId,
    extensionPath: workerData.extensionPath,
    subscriptions: []
  };

  await activateFn(context);
  activated = true;
}

parentPort.on("message", async (message) => {
  if (!message || typeof message !== "object") return;

  if (message.type === "activate") {
    try {
      await activateExtension();
      parentPort.postMessage({ type: "activate_result", id: message.id });
    } catch (error) {
      parentPort.postMessage({
        type: "activate_error",
        id: message.id,
        error: { message: String(error?.message ?? error), stack: error?.stack }
      });
    }
    return;
  }

  formulaApi.__handleMessage(message);
});
