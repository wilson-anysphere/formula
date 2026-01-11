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

function createMemoryStorage() {
  const map = new Map();
  return {
    getItem(key) {
      return map.has(key) ? map.get(key) : null;
    },
    setItem(key, value) {
      map.set(String(key), String(value));
    },
    removeItem(key) {
      map.delete(String(key));
    }
  };
}

test("BrowserExtensionHost: network allowlist blocks non-allowlisted hosts", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.allowlist";
  const extensionId = "test.allowlist";

  storage.setItem(
    storageKey,
    JSON.stringify({
      [extensionId]: {
        network: { mode: "allowlist", hosts: ["allowed.example"] }
      }
    })
  );

  const scenarios = [
    {
      onPostMessage() {
        // noop; test drives messages manually.
      }
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);

  const prevFetch = globalThis.fetch;
  globalThis.fetch = async (url) => {
    const u = new URL(String(url));
    if (u.hostname !== "allowed.example") {
      throw new Error(`Unexpected fetch to ${u.hostname}`);
    }
    return new Response("ok", { status: 200, headers: { "content-type": "text/plain" } });
  };

  t.after(async () => {
    globalThis.Worker = PrevWorker;
    globalThis.fetch = prevFetch;
  });

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("network")) return false;
      return true;
    },
    permissionStorage: storage,
    permissionStorageKey: storageKey
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "allowlist",
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

  const allowed = new Promise((resolve, reject) => {
    extension.worker._scenario.onPostMessage = (msg) => {
      if (msg?.type === "api_result" && msg.id === "req-allowed") resolve(msg.result);
      if (msg?.type === "api_error" && msg.id === "req-allowed") reject(new Error(msg.error?.message ?? msg.error));
    };
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req-allowed",
    namespace: "network",
    method: "fetch",
    args: ["https://allowed.example/"]
  });

  const allowedResult = await allowed;
  assert.equal(allowedResult?.bodyText, "ok");

  const blocked = new Promise((resolve, reject) => {
    extension.worker._scenario.onPostMessage = (msg) => {
      if (msg?.type === "api_error" && msg.id === "req-blocked") resolve(msg.error);
      if (msg?.type === "api_result" && msg.id === "req-blocked") reject(new Error("Expected api_error"));
    };
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req-blocked",
    namespace: "network",
    method: "fetch",
    args: ["https://blocked.example/"]
  });

  const blockedError = await blocked;
  assert.match(String(blockedError?.message ?? blockedError), /Permission denied/i);
});

test("BrowserExtensionHost: revokePermissions forces network to be re-prompted", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    {
      onPostMessage() {}
    }
  ];

  const PrevWorker = globalThis.Worker;
  globalThis.Worker = createWorkerCtor(scenarios);

  const prevFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response("ok", { status: 200 });

  t.after(() => {
    globalThis.Worker = PrevWorker;
    globalThis.fetch = prevFetch;
  });

  const extensionId = "test.revoke";
  let networkPrompts = 0;
  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("network")) {
        networkPrompts += 1;
        return networkPrompts === 1;
      }
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "revoke",
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

  const first = new Promise((resolve, reject) => {
    extension.worker._scenario.onPostMessage = (msg) => {
      if (msg?.type === "api_result" && msg.id === "req1") resolve(msg.result);
      if (msg?.type === "api_error" && msg.id === "req1") reject(new Error(msg.error?.message ?? msg.error));
    };
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "network",
    method: "fetch",
    args: ["https://allowed.example/"]
  });
  await first;
  assert.equal(networkPrompts, 1);

  await host.revokePermissions(extensionId, ["network"]);
  const after = await host.getGrantedPermissions(extensionId);
  assert.ok(!after.network, "Expected network permission to be revoked");

  const second = new Promise((resolve, reject) => {
    extension.worker._scenario.onPostMessage = (msg) => {
      if (msg?.type === "api_error" && msg.id === "req2") resolve(msg.error);
      if (msg?.type === "api_result" && msg.id === "req2") reject(new Error("Expected api_error"));
    };
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req2",
    namespace: "network",
    method: "fetch",
    args: ["https://allowed.example/"]
  });

  const err = await second;
  assert.match(String(err?.message ?? err), /Permission denied/i);
  assert.equal(networkPrompts, 2);
});

