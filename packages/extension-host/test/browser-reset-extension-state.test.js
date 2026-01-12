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

function createStorageApiWithClear() {
  const data = new Map();
  /** @type {string[]} */
  const cleared = [];
  return {
    cleared,
    getExtensionStore(extensionId) {
      const id = String(extensionId);
      if (!data.has(id)) data.set(id, {});
      return data.get(id);
    },
    clearExtensionStore(extensionId) {
      const id = String(extensionId);
      cleared.push(id);
      data.delete(id);
    }
  };
}

function createStorageApiWithoutClear() {
  const data = new Map();
  return {
    getExtensionStore(extensionId) {
      const id = String(extensionId);
      if (!data.has(id)) data.set(id, {});
      return data.get(id);
    }
  };
}

test("browser resetExtensionState clears injected permission storage and calls storageApi.clearExtensionStore when available", async () => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const permissionStorage = createMemoryStorage();
  const permissionStorageKey = "formula.test.permissions.custom";
  permissionStorage.setItem(
    permissionStorageKey,
    JSON.stringify({
      "pub.ext": { storage: true, network: { mode: "full" } },
      "pub.other": { clipboard: true }
    })
  );

  const storageApi = createStorageApiWithClear();
  const storeBefore = storageApi.getExtensionStore("pub.ext");
  storeBefore.foo = "bar";

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    permissionStorage,
    permissionStorageKey,
    storageApi
  });

  await host.resetExtensionState("pub.ext");

  const perms = JSON.parse(permissionStorage.getItem(permissionStorageKey));
  assert.equal(perms["pub.ext"], undefined);
  assert.deepEqual(perms["pub.other"], { clipboard: true });

  assert.deepEqual(storageApi.cleared, ["pub.ext"]);
  const storeAfter = storageApi.getExtensionStore("pub.ext");
  assert.notEqual(storeAfter, storeBefore);
  assert.equal(storeAfter.foo, undefined);
});

test("browser resetExtensionState falls back to clearing keys when storageApi lacks clearExtensionStore", async () => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const storageApi = createStorageApiWithoutClear();
  const store = storageApi.getExtensionStore("pub.ext");
  store.foo = "bar";
  store.nested = { ok: true };

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    storageApi
  });

  await host.resetExtensionState("pub.ext");

  // Same object instance (no clearExtensionStore), but keys should be deleted.
  const after = storageApi.getExtensionStore("pub.ext");
  assert.equal(after, store);
  assert.deepEqual(after, {});
});

