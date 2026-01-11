const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

function createWorkerCtor(scenarios) {
  return class FakeWorker {
    constructor(_url, _options) {
      this._listeners = new Map();
      this._terminated = false;
      this._scenario = scenarios.shift() ?? {};
    }

    addEventListener(type, listener) {
      const key = String(type);
      if (!this._listeners.has(key)) this._listeners.set(key, new Set());
      this._listeners.get(key).add(listener);
    }

    removeEventListener(type, listener) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      set.delete(listener);
      if (set.size === 0) this._listeners.delete(String(type));
    }

    postMessage(message) {
      if (this._terminated) return;
      try {
        this._scenario.onPostMessage?.(message, this);
      } catch (err) {
        this._emit("error", { message: String(err?.message ?? err) });
      }
    }

    terminate() {
      this._terminated = true;
    }

    emitMessage(message) {
      if (this._terminated) return;
      this._emit("message", { data: message });
    }

    _emit(type, event) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      for (const listener of [...set]) {
        try {
          listener(event);
        } catch {
          // ignore
        }
      }
    }
  };
}

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: data connector registration succeeds when declared", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {(value?: unknown) => void} */
  let resolveApiResult;
  const apiResult = new Promise((resolve) => {
    resolveApiResult = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_result" && msg.id === "req1") resolveApiResult();
      }
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);
  t.after(() => {
    globalThis.Worker = PrevWorker;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.data-connectors";
  const connectorId = "test.connector";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://connectors/",
    mainUrl: "memory://connectors/main.mjs",
    manifest: {
      name: "data-connectors",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [],
        customFunctions: [],
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      },
      activationEvents: [],
      permissions: []
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "dataConnectors",
    method: "register",
    args: [connectorId]
  });

  await apiResult;
});

test("BrowserExtensionHost: data connector registration rejected when not declared", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {(value?: unknown) => void} */
  let resolveApiError;
  const apiError = new Promise((resolve) => {
    resolveApiError = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") resolveApiError(msg.error);
      }
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);
  t.after(() => {
    globalThis.Worker = PrevWorker;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.data-connectors-invalid";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://connectors/",
    mainUrl: "memory://connectors/main.mjs",
    manifest: {
      name: "data-connectors-invalid",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: []
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "dataConnectors",
    method: "register",
    args: ["test.undeclared"]
  });

  const error = await apiError;
  assert.match(String(error?.message ?? error), /Data connector not declared in manifest/i);
});

test("BrowserExtensionHost: invokeDataConnector activates the extension and returns results", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const connectorId = "test.connector";
  const scenarios = [
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "invoke_data_connector") {
          worker.emitMessage({
            type: "data_connector_result",
            id: msg.id,
            result: { columns: ["x"], rows: [[1]] }
          });
        }
      }
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);
  t.after(() => {
    globalThis.Worker = PrevWorker;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.invoke-data-connector";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://connectors/",
    mainUrl: "memory://connectors/main.mjs",
    manifest: {
      name: "invoke-data-connector",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [],
        customFunctions: [],
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      },
      activationEvents: [`onDataConnector:${connectorId}`],
      permissions: []
    }
  });

  const result = await host.invokeDataConnector(connectorId, "query", {}, {});
  assert.deepEqual(result, { columns: ["x"], rows: [[1]] });
  assert.equal(host._extensions.get(extensionId).active, true);
});

test("BrowserExtensionHost: invokeDataConnector requires onDataConnector activation event when extension is inactive", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [{ onPostMessage(msg) { if (msg?.type === "init") return; } }];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);
  t.after(() => {
    globalThis.Worker = PrevWorker;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const connectorId = "test.connector";
  await host.loadExtension({
    extensionId: "test.no-activation",
    extensionPath: "memory://connectors/",
    mainUrl: "memory://connectors/main.mjs",
    manifest: {
      name: "no-activation",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [], dataConnectors: [{ id: connectorId, name: "Test" }] },
      activationEvents: [],
      permissions: []
    }
  });

  await assert.rejects(
    () => host.invokeDataConnector(connectorId, "browse", {}),
    /not activated for onDataConnector/i
  );
});

test("BrowserExtensionHost: data connector timeout terminates worker and allows restart", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const connectorId = "test.connector";
  const scenarios = [
    // First worker: activation succeeds; connector invocations hang forever.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "invoke_data_connector") {
          return;
        }
      }
    },
    // Second worker: activation succeeds; connector returns results.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "invoke_data_connector") {
          worker.emitMessage({ type: "data_connector_result", id: msg.id, result: { ok: true } });
        }
      }
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);
  t.after(() => {
    globalThis.Worker = PrevWorker;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    activationTimeoutMs: 1000,
    dataConnectorTimeoutMs: 50
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.data-connector-timeout";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://connectors/",
    mainUrl: "memory://connectors/main.mjs",
    manifest: {
      name: "data-connector-timeout",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [], dataConnectors: [{ id: connectorId, name: "Test" }] },
      activationEvents: [`onDataConnector:${connectorId}`],
      permissions: []
    }
  });

  const hangPromise = host.invokeDataConnector(connectorId, "query", {}, {});
  await new Promise((r) => setTimeout(r, 10));
  const pendingPromise = host.invokeDataConnector(connectorId, "browse", {});

  await assert.rejects(() => hangPromise, /timed out/i);
  await assert.rejects(() => pendingPromise, /worker terminated/i);
  assert.equal(host._extensions.get(extensionId).active, false);

  assert.deepEqual(await host.invokeDataConnector(connectorId, "browse", {}), { ok: true });
});

