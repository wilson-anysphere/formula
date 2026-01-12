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

test("reloadExtension recycles the worker and clears runtime registrations", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-reload-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "reload-cleanup",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onStartupFinished"],
      permissions: ["ui.commands", "ui.panels"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        await formula.commands.registerCommand("test.dynamic", async () => "ok");
        const panel = await formula.ui.createPanel("test.panel", { title: "Test" });
        await panel.webview.setHtml("<h1>Test Panel</h1>");
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep this test focused on reload behavior
    // rather than the default 5s activation SLA.
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = await host.loadExtension(extDir);
  await host.startup();

  assert.equal(await host.executeCommand("test.dynamic"), "ok");
  assert.ok(host.getPanel("test.panel"));

  await host.reloadExtension(extensionId);
  assert.equal(host.getPanel("test.panel"), undefined);
  await assert.rejects(() => host.executeCommand("test.dynamic"), /Unknown command/);

  await host.startup();
  assert.equal(await host.executeCommand("test.dynamic"), "ok");
  assert.ok(host.getPanel("test.panel"));
});
