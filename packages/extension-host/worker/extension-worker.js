const { parentPort, workerData } = require("node:worker_threads");
const Module = require("node:module");

const formulaApi = require(workerData.apiModulePath);
formulaApi.__setTransport({
  postMessage: (message) => parentPort.postMessage(message)
});
formulaApi.__setContext({
  extensionId: workerData.extensionId,
  extensionPath: workerData.extensionPath
});

// Provide a VS Code-like virtual module for extension authors.
const originalLoad = Module._load;
Module._load = function (request, parent, isMain) {
  if (request === "@formula/extension-api" || request === "formula") {
    return formulaApi;
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
