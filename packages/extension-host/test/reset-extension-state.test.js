const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

test("resetExtensionState removes permissions, extension storage, and extension-data dirs", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-reset-state-"));
  const permissionsStoragePath = path.join(dir, "permissions.json");
  const extensionStoragePath = path.join(dir, "storage.json");

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

  const extensionId = "test.ext";

  // Seed persistence files with an entry for the extension.
  await fs.writeFile(
    permissionsStoragePath,
    JSON.stringify({ [extensionId]: { "ui.commands": true }, other: { clipboard: true } }, null, 2),
    "utf8"
  );
  await fs.writeFile(
    extensionStoragePath,
    JSON.stringify({ [extensionId]: { foo: "bar" }, other: { baz: 1 } }, null, 2),
    "utf8"
  );

  // Seed the extension-data folder provisioned by the host.
  const dataDir = path.join(dir, "extension-data", extensionId);
  await fs.mkdir(path.join(dataDir, "globalStorage"), { recursive: true });
  await fs.mkdir(path.join(dataDir, "workspaceStorage"), { recursive: true });

  await host.resetExtensionState(extensionId);

  // permissions.json entry removed
  const permissions = JSON.parse(await fs.readFile(permissionsStoragePath, "utf8"));
  assert.equal(permissions[extensionId], undefined);
  assert.deepEqual(permissions.other, { clipboard: true });

  // storage.json entry removed
  const storage = JSON.parse(await fs.readFile(extensionStoragePath, "utf8"));
  assert.equal(storage[extensionId], undefined);
  assert.deepEqual(storage.other, { baz: 1 });

  // extension-data directory removed
  await assert.rejects(() => fs.stat(dataDir), /ENOENT/);
});

