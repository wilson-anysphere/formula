const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), "utf8");
}

test("loadExtension rejects entrypoints that escape the extension folder", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-entrypoint-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  const extDir = path.join(tmp, "ext");
  await fs.mkdir(extDir, { recursive: true });
  await writeJson(path.join(extDir, "package.json"), {
    name: "evil",
    version: "1.0.0",
    publisher: "evil",
    main: "../escape.js",
    engines: { formula: "^1.0.0" },
    permissions: []
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmp, "permissions.json"),
    extensionStoragePath: path.join(tmp, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(() => host.loadExtension(extDir), /resolve inside extension folder/i);
});

test("loadExtension rejects missing entrypoint files", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-entrypoint-missing-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  const extDir = path.join(tmp, "ext");
  await fs.mkdir(extDir, { recursive: true });
  await writeJson(path.join(extDir, "package.json"), {
    name: "missing",
    version: "1.0.0",
    publisher: "evil",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    permissions: []
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmp, "permissions.json"),
    extensionStoragePath: path.join(tmp, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(() => host.loadExtension(extDir), /entrypoint not found/i);
});

