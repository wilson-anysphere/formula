const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeExtension(dir, manifest, entrypointSource) {
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(dir, manifest.main), entrypointSource);
}

test("resource limits: memoryMb configures worker_threads resourceLimits", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-resource-limits-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "resource-limits",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" }
    },
    `
      module.exports.activate = async () => {};
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    memoryMb: 64
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = await host.loadExtension(extDir);
  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  assert.equal(extension.worker.resourceLimits.maxOldGenerationSizeMb, 64);
  assert.equal(extension.worker.resourceLimits.maxYoungGenerationSizeMb, 16);
});

test("resource limits: memoryMb=0 disables worker heap limits", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-resource-limits-off-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "resource-limits-off",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" }
    },
    `
      module.exports.activate = async () => {};
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    memoryMb: 0
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = await host.loadExtension(extDir);
  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  // Node's `Worker.resourceLimits` starts out reflecting the constructor options
  // (unset limits show as -1), then updates asynchronously to reflect the actual
  // runtime defaults once the worker thread finishes bootstrapping.
  //
  // We only care that Formula didn't *apply* a low heap cap when `memoryMb=0`.
  const oldMb = extension.worker.resourceLimits.maxOldGenerationSizeMb;
  const youngMb = extension.worker.resourceLimits.maxYoungGenerationSizeMb;
  assert.ok(oldMb === -1 || oldMb > 64, `expected maxOldGenerationSizeMb to be unset/default, got ${oldMb}`);
  assert.ok(youngMb === -1 || youngMb > 16, `expected maxYoungGenerationSizeMb to be unset/default, got ${youngMb}`);
});
