const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

test("contributions: host exposes manifest contributes data for UI integration", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-contrib-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  const commands = host.getContributedCommands().map((c) => c.command);
  assert.ok(commands.includes("sampleHello.sumSelection"));
  assert.ok(commands.includes("sampleHello.openPanel"));

  const panels = host.getContributedPanels().map((p) => p.id);
  assert.ok(panels.includes("sampleHello.panel"));

  const customFunctions = host.getContributedCustomFunctions().map((f) => f.name);
  assert.ok(customFunctions.includes("SAMPLEHELLO_DOUBLE"));

  const menuItems = host.getContributedMenu("cell/context");
  assert.deepEqual(menuItems, [
    {
      extensionId: "formula.sample-hello",
      command: "sampleHello.sumSelection",
      when: "hasSelection",
      group: "extensions@1"
    },
    {
      extensionId: "formula.sample-hello",
      command: "sampleHello.openPanel",
      when: null,
      group: "extensions@2"
    }
  ]);
});
