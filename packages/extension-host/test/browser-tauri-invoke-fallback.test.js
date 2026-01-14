const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

function restoreProperty(target, prop, descriptor) {
  if (descriptor) {
    Object.defineProperty(target, prop, descriptor);
    return;
  }
  // eslint-disable-next-line no-param-reassign
  delete target[prop];
}

test("BrowserExtensionHost: network.fetch falls back to legacy __TAURI__.invoke when __TAURI__.core is blocked (throwing getter)", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const originalTauri = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const originalFetch = Object.getOwnPropertyDescriptor(globalThis, "fetch");

  const tauri = {};
  const invoke = async function (cmd, args) {
    assert.equal(this, tauri);
    assert.equal(cmd, "network_fetch");
    assert.equal(typeof args?.url, "string");
    return { ok: true, status: 200, statusText: "OK", url: args.url, headers: [], bodyText: "tauri" };
  };
  tauri.invoke = invoke;
  Object.defineProperty(tauri, "core", {
    configurable: true,
    get() {
      throw new Error("Blocked core access");
    }
  });

  Object.defineProperty(globalThis, "__TAURI__", { configurable: true, value: tauri, writable: true });
  globalThis.fetch = async () => {
    throw new Error("fetch() should not be called when Tauri invoke is available");
  };

  t.after(() => {
    restoreProperty(globalThis, "__TAURI__", originalTauri);
    restoreProperty(globalThis, "fetch", originalFetch);
  });

  const scenarios = [{ onPostMessage() {} }];
  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.tauriInvokeFallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "tauriInvokeFallback",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["network"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  const resultPromise = new Promise((resolve, reject) => {
    extension.worker._scenario.onPostMessage = (msg) => {
      if (msg?.type === "api_result" && msg.id === "req") resolve(msg.result);
      if (msg?.type === "api_error" && msg.id === "req") reject(new Error(msg.error?.message ?? msg.error));
    };
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req",
    namespace: "network",
    method: "fetch",
    args: ["https://allowed.example/"]
  });

  const result = await resultPromise;
  assert.equal(result?.bodyText, "tauri");
});

