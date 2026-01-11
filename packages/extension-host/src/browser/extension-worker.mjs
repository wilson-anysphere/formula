import * as formulaApi from "@formula/extension-api";

formulaApi.__setTransport({
  postMessage: (message) => postMessage(message)
});

let workerData = null;
let extensionModule = null;
let activated = false;

function init(data) {
  workerData = {
    extensionId: String(data?.extensionId ?? ""),
    extensionPath: String(data?.extensionPath ?? ""),
    mainUrl: String(data?.mainUrl ?? "")
  };

  formulaApi.__setContext({
    extensionId: workerData.extensionId,
    extensionPath: workerData.extensionPath
  });
}

if (typeof globalThis.fetch === "function") {
  globalThis.fetch = async (input, init) => {
    return formulaApi.network.fetch(String(input), init);
  };
}

if (typeof globalThis.WebSocket === "function") {
  globalThis.WebSocket = class WebSocketBlocked {
    constructor() {
      throw new Error("WebSocket is not available to extensions in this environment");
    }
  };
}

async function activateExtension() {
  if (activated) return;
  if (!workerData) throw new Error("Extension worker not initialized");
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
    subscriptions: []
  };

  await activateFn(context);
  activated = true;
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

  formulaApi.__handleMessage(message);
});

