const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

test("extension storage: __proto__ keys do not create prototype properties and persist under a safe alias", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-storage-proto-"));
  const extensionStoragePath = path.join(dir, "storage.json");
  const permissionsStoragePath = path.join(dir, "permissions.json");

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath,
    extensionStoragePath,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose().catch(() => {});
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  const extension = { id: "pub.ext" };

  await host._executeApi("storage", "set", ["__proto__", { polluted: true }], extension);

  assert.deepEqual(await host._executeApi("storage", "get", ["__proto__"], extension), { polluted: true });
  assert.equal(await host._executeApi("storage", "get", ["polluted"], extension), undefined);

  const store = await host._loadExtensionStorage();
  assert.equal(Object.getPrototypeOf(store), null);
  assert.equal(Object.getPrototypeOf(store[extension.id]), null);

  const raw = await fs.readFile(extensionStoragePath, "utf8");
  // Persisted JSON should not contain a raw `"__proto__": ...` entry that could be mis-parsed by runtimes.
  assert.equal(/"__proto__"\s*:/.test(raw), false);

  const persisted = JSON.parse(raw);
  assert.equal(Object.prototype.hasOwnProperty.call(persisted[extension.id], "__proto__"), false);
});

test("extension storage: migrates legacy persisted __proto__ keys to the safe alias while preserving access", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-storage-proto-migrate-"));
  const extensionStoragePath = path.join(dir, "storage.json");
  const permissionsStoragePath = path.join(dir, "permissions.json");

  await fs.writeFile(
    extensionStoragePath,
    // Write raw JSON so the `__proto__` key is preserved as a real entry.
    '{"pub.ext":{"__proto__":{"polluted":true},"foo":"bar"}}',
    "utf8"
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath,
    extensionStoragePath,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose().catch(() => {});
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  });

  const extension = { id: "pub.ext" };

  assert.deepEqual(await host._executeApi("storage", "get", ["__proto__"], extension), { polluted: true });
  assert.equal(await host._executeApi("storage", "get", ["polluted"], extension), undefined);
  assert.equal(await host._executeApi("storage", "get", ["foo"], extension), "bar");

  const migratedRaw = await fs.readFile(extensionStoragePath, "utf8");
  assert.equal(/"__proto__"\s*:/.test(migratedRaw), false);

  const migrated = JSON.parse(migratedRaw);
  assert.equal(Object.prototype.hasOwnProperty.call(migrated["pub.ext"], "__proto__"), false);
  assert.equal(migrated["pub.ext"].foo, "bar");
});
