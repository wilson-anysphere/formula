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
    },
    _dump() {
      return [...map.entries()];
    }
  };
}

test("browser BrowserExtensionHost.resetExtensionState clears LocalStorageExtensionStorage records", async (t) => {
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

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  const extensionId = "pub.ext";

  // Simulate extension storage writes via the host's storage API implementation.
  host._storageApi.getExtensionStore(extensionId).foo = "bar";

  assert.ok(
    storage.getItem(`formula.extensionHost.storage.${extensionId}`),
    "expected LocalStorageExtensionStorage to persist under the default key prefix"
  );

  await host.resetExtensionState(extensionId);

  assert.equal(storage.getItem(`formula.extensionHost.storage.${extensionId}`), null);

  // Ensure the in-memory cache was cleared, so subsequent installs see a clean store.
  const storeAfter = host._storageApi.getExtensionStore(extensionId);
  assert.equal(storeAfter.foo, undefined);
});

test("browser LocalStorageExtensionStorage removes the record when the last key is deleted", async (t) => {
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

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  const extension = { id: "pub.ext" };
  const storageKey = "formula.extensionHost.storage.pub.ext";

  await host._executeApi("storage", "set", ["foo", "bar"], extension);
  assert.ok(storage.getItem(storageKey), "expected storage key to be persisted");

  await host._executeApi("storage", "delete", ["foo"], extension);
  assert.equal(storage.getItem(storageKey), null);
});
