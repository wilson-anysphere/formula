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

test("browser PermissionManager: migrates legacy string-array grants to v2 permission records", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.migrate";
  storage.setItem(
    storageKey,
    JSON.stringify({
      "pub.ext": ["network", "clipboard"]
    })
  );

  const pm = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => true
  });

  const granted = await pm.getGrantedPermissions("pub.ext");
  assert.deepEqual(granted, {
    network: { mode: "full" },
    clipboard: true
  });

  const storedRaw = JSON.parse(storage.getItem(storageKey));
  assert.deepEqual(storedRaw, {
    "pub.ext": {
      network: { mode: "full" },
      clipboard: true
    }
  });
});

test("browser PermissionManager: revokePermissions removes persisted grants for a single extension", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.revoke";

  const pm = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["network"]
    },
    ["network"]
  );

  await pm.ensurePermissions(
    {
      extensionId: "pub.other",
      displayName: "Other",
      declaredPermissions: ["clipboard"]
    },
    ["clipboard"]
  );

  await pm.revokePermissions("pub.ext");

  const stored = JSON.parse(storage.getItem(storageKey));
  assert.ok(!stored["pub.ext"], "Expected revoked extension id to be removed from permission store");
  assert.deepEqual(stored["pub.other"], { clipboard: true });

  let promptCalls = 0;
  const pm2 = new PermissionManager({
    storage,
    storageKey,
    prompt: async ({ permissions }) => {
      promptCalls += 1;
      assert.deepEqual(permissions, ["network"]);
      return true;
    }
  });

  // Should prompt again because the persisted grant was removed.
  await pm2.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["network"]
    },
    ["network"]
  );

  assert.equal(promptCalls, 1);
});

test("browser PermissionManager: revokePermissions + resetPermissions clear grants", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.reset";
  const extensionId = "pub.ext";

  const pm = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId,
      displayName: "Ext",
      declaredPermissions: ["network", "clipboard"]
    },
    ["network", "clipboard"]
  );

  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {
    network: { mode: "full" },
    clipboard: true
  });

  await pm.revokePermissions(extensionId, ["clipboard"]);
  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {
    network: { mode: "full" }
  });

  await pm.resetPermissions(extensionId);
  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {});
});

test("browser PermissionManager: resetAllPermissions clears all extensions", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.resetAll";

  const pm = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId: "pub.one",
      displayName: "One",
      declaredPermissions: ["clipboard"]
    },
    ["clipboard"]
  );

  await pm.ensurePermissions(
    {
      extensionId: "pub.two",
      displayName: "Two",
      declaredPermissions: ["network"]
    },
    ["network"]
  );

  await pm.resetAllPermissions();
  assert.deepEqual(await pm.getGrantedPermissions("pub.one"), {});
  assert.deepEqual(await pm.getGrantedPermissions("pub.two"), {});
});

test("browser PermissionManager: accepts object-form declared permissions", async () => {
  const { PermissionManager } = await importBrowserPermissionManager();

  const storage = createMemoryStorage();
  const storageKey = "formula.test.permissions.objectDeclared";
  const extensionId = "pub.obj";

  const pm = new PermissionManager({
    storage,
    storageKey,
    prompt: async () => true,
  });

  await pm.ensurePermissions(
    {
      extensionId,
      displayName: "Obj",
      declaredPermissions: [{ network: { mode: "allowlist", hosts: ["example.com"] } }, { clipboard: true }],
    },
    ["network", "clipboard"],
  );

  // `ensurePermissions` only uses declared permissions as a boolean gate; it should still store
  // the v2 record format for the granted permissions.
  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {
    network: { mode: "full" },
    clipboard: true,
  });
});
