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

test("timeouts: activation timeout terminates a hanging extension", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-timeout-activate-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "hang-activate",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onStartupFinished"]
    },
    `
      module.exports.activate = async () => {
        await new Promise(() => {});
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    activationTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.startup(), /activation.*timed out/i);
});

test("timeouts: command timeout terminates a hanging command handler", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-timeout-command-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "hang-command",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.hangCommand"],
      contributes: {
        commands: [
          {
            command: "test.hangCommand",
            title: "Hang Command"
          }
        ]
      },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.commands.registerCommand("test.hangCommand", async () => {
          return new Promise(() => {});
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    activationTimeoutMs: 1000,
    commandTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand("test.hangCommand"), /command.*timed out/i);
  const extInfo = host.listExtensions().find((e) => e.id === "test.hang-command");
  assert.ok(extInfo);
  assert.equal(extInfo.active, false);
});

