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

test("crash: worker crash marks extension inactive and next command spawns a fresh worker", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-crash-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "crash-test",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.crash", "onCommand:test.ok"],
      contributes: {
        commands: [
          { command: "test.crash", title: "Crash" },
          { command: "test.ok", title: "Ok" }
        ]
      },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.commands.registerCommand("test.ok", async () => "ok");
        await formula.commands.registerCommand("test.crash", async () => {
          // Trigger an unhandled exception on the worker event loop after responding.
          setTimeout(() => {
            Promise.reject(new Error("boom"));
          }, 25);
          return "scheduled";
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI/agent load; keep this generous so the
    // test focuses on crash recovery rather than the default 5s activation SLA.
    activationTimeoutMs: 20_000,
    commandTimeoutMs: 20_000
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = await host.loadExtension(extDir);

  assert.equal(await host.executeCommand("test.crash"), "scheduled");

  const deadline = Date.now() + 2000;
  while (Date.now() < deadline) {
    const info = host.listExtensions().find((e) => e.id === extensionId);
    if (info && info.active === false) break;
    await new Promise((r) => setTimeout(r, 10));
  }

  const info = host.listExtensions().find((e) => e.id === extensionId);
  assert.ok(info);
  assert.equal(info.active, false);

  assert.equal(await host.executeCommand("test.ok"), "ok");
});
