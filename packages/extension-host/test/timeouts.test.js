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
    // Worker startup can be slow under heavy CI load; keep this generous so this
    // test exercises the command timeout rather than flaking on activation.
    activationTimeoutMs: 20_000,
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

test("timeouts: terminating a hung worker rejects other in-flight requests and allows restart", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-timeout-cleanup-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "hang-cleanup",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.quick", "onCommand:test.hangLoop"],
      contributes: {
        commands: [
          { command: "test.quick", title: "Quick Command" },
          { command: "test.hangLoop", title: "Hang Loop" }
        ]
      },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.commands.registerCommand("test.quick", async () => "ok");
        await formula.commands.registerCommand("test.hangLoop", async () => {
          // Block the worker thread event loop so subsequent requests cannot be processed.
          // This simulates a truly misbehaving extension.
          // eslint-disable-next-line no-constant-condition
          while (true) {}
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI load; keep activation timeout generous so this
    // test exercises worker termination + cleanup rather than flaking on activation.
    activationTimeoutMs: 20_000,
    commandTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  // Ensure the extension is activated before starting concurrent command execution.
  assert.equal(await host.executeCommand("test.quick"), "ok");

  const hangPromise = host.executeCommand("test.hangLoop");
  await new Promise((r) => setTimeout(r, 10));
  const pendingPromise = host.executeCommand("test.quick");

  await assert.rejects(() => hangPromise, /timed out/i);
  await assert.rejects(() => pendingPromise, /worker terminated/i);

  // The next command should automatically spin up a fresh worker and re-activate the extension.
  assert.equal(await host.executeCommand("test.quick"), "ok");
});

test("timeouts: custom function timeout terminates a hanging handler", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-timeout-custom-fn-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "hang-custom-fn",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCustomFunction:TEST_HANG"],
      contributes: {
        customFunctions: [
          {
            name: "TEST_HANG",
            description: "Always hangs",
            parameters: [],
            result: { type: "number" }
          }
        ]
      }
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.functions.register("TEST_HANG", {
          handler: async () => {
            await new Promise(() => {});
          }
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI load; keep activation timeout generous so this
    // test exercises the custom-function timeout rather than flaking on activation.
    activationTimeoutMs: 20_000,
    customFunctionTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.invokeCustomFunction("TEST_HANG"), /custom function.*timed out/i);
});

test("timeouts: worker termination clears runtime context menus and can re-register after restart", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-timeout-menus-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "timeout-menus",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.quick", "onCommand:test.hangLoop"],
      contributes: {
        commands: [
          { command: "test.quick", title: "Quick" },
          { command: "test.hangLoop", title: "Hang Loop" }
        ]
      },
      permissions: ["ui.commands", "ui.menus"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.ui.registerContextMenu("cell/context", [{ command: "test.hangLoop" }]);

        await formula.commands.registerCommand("test.quick", async () => "ok");
        await formula.commands.registerCommand("test.hangLoop", async () => {
          // eslint-disable-next-line no-constant-condition
          while (true) {}
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI load; keep activation timeout generous so this
    // test exercises command timeouts + cleanup rather than flaking on activation.
    activationTimeoutMs: 20_000,
    commandTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  assert.equal(await host.executeCommand("test.quick"), "ok");
  assert.deepEqual(host.getContributedMenu("cell/context"), [
    { extensionId: "test.timeout-menus", command: "test.hangLoop", when: null, group: null }
  ]);

  await assert.rejects(() => host.executeCommand("test.hangLoop"), /timed out/i);
  assert.deepEqual(host.getContributedMenu("cell/context"), []);

  assert.equal(await host.executeCommand("test.quick"), "ok");
  assert.deepEqual(host.getContributedMenu("cell/context"), [
    { extensionId: "test.timeout-menus", command: "test.hangLoop", when: null, group: null }
  ]);
});
