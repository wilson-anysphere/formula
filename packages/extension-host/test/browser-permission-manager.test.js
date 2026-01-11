const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

async function importBrowserPermissionManager() {
  const moduleUrl = pathToFileURL(
    path.resolve(__dirname, "../src/browser/permission-manager.mjs")
  ).href;
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
    },
    _dump() {
      return [...map.entries()];
    }
  };
}

test("browser PermissionManager: persists grants via injected storage backend", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions";

  let promptCalls = 0;
  const pm1 = new PermissionManager({
    storage,
    storageKey,
    prompt: async ({ permissions }) => {
      promptCalls += 1;
      assert.deepEqual(permissions, ["network"]);
      return true;
    }
  });

  await pm1.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["network"]
    },
    ["network"]
  );

  assert.equal(promptCalls, 1);
  assert.ok(storage._dump().length > 0, "Expected permissions to be stored");

  const pm2 = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => {
      throw new Error("prompt should not be called once grants are persisted");
    }
  });

  await pm2.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["network"]
    },
    ["network"]
  );
});

