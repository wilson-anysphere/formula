const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

function createMemoryStorage() {
  const map = new Map();
  return {
    getItem(key) {
      return map.has(String(key)) ? map.get(String(key)) : null;
    },
    setItem(key, value) {
      map.set(String(key), String(value));
    },
    removeItem(key) {
      map.delete(String(key));
    }
  };
}

test("browser extension storage: __proto__ keys round-trip across reload without prototype pollution", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const storage = createMemoryStorage();
  const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  Object.defineProperty(globalThis, "localStorage", {
    value: storage,
    configurable: true,
    writable: true
  });

  t.after(() => {
    try {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
        return;
      }
      delete globalThis.localStorage;
    } catch {
      // ignore
    }
  });

  const extension = { id: "pub.ext" };
  const storageKey = "formula.extensionHost.storage.pub.ext";

  const host1 = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  await host1._executeApi("storage", "set", ["__proto__", { polluted: true }], extension);

  assert.deepEqual(await host1._executeApi("storage", "get", ["__proto__"], extension), { polluted: true });
  assert.equal(await host1._executeApi("storage", "get", ["polluted"], extension), undefined);

  const persistedRaw = storage.getItem(storageKey);
  assert.ok(persistedRaw, "expected LocalStorageExtensionStorage to persist the record");
  assert.equal(/"__proto__"\s*:/.test(persistedRaw), false);

  // Simulate a host reload.
  const host2 = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  assert.deepEqual(await host2._executeApi("storage", "get", ["__proto__"], extension), { polluted: true });
  assert.equal(await host2._executeApi("storage", "get", ["polluted"], extension), undefined);
});

test("browser extension storage: migrates legacy persisted __proto__ keys to the safe alias", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const storage = createMemoryStorage();
  const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  Object.defineProperty(globalThis, "localStorage", {
    value: storage,
    configurable: true,
    writable: true
  });

  t.after(() => {
    try {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
        return;
      }
      delete globalThis.localStorage;
    } catch {
      // ignore
    }
  });

  const storageKey = "formula.extensionHost.storage.pub.ext";
  // Seed a raw record with a literal `__proto__` key.
  storage.setItem(storageKey, '{"__proto__":{"polluted":true},"foo":"bar"}');

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  const extension = { id: "pub.ext" };

  assert.deepEqual(await host._executeApi("storage", "get", ["__proto__"], extension), { polluted: true });
  assert.equal(await host._executeApi("storage", "get", ["polluted"], extension), undefined);
  assert.equal(await host._executeApi("storage", "get", ["foo"], extension), "bar");

  const migratedRaw = storage.getItem(storageKey);
  assert.ok(migratedRaw, "expected storage record to remain persisted after migration");
  assert.equal(/"__proto__"\s*:/.test(migratedRaw), false);
});

