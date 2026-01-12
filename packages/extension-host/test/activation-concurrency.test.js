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

test("activation is de-duplicated when multiple commands trigger it concurrently", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-activation-concurrency-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "activation-concurrency",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:test.a", "onCommand:test.b"],
      contributes: {
        commands: [
          { command: "test.a", title: "A" },
          { command: "test.b", title: "B" }
        ]
      },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        globalThis.__activationCount = (globalThis.__activationCount ?? 0) + 1;
        await new Promise((r) => setTimeout(r, 50));

        await formula.commands.registerCommand("test.a", async () => globalThis.__activationCount);
        await formula.commands.registerCommand("test.b", async () => globalThis.__activationCount);
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI/agent load; keep this generous so the
    // test focuses on activation de-duplication rather than the default 5s SLA.
    activationTimeoutMs: 20_000,
    commandTimeoutMs: 20_000
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  const [a, b] = await Promise.all([host.executeCommand("test.a"), host.executeCommand("test.b")]);
  assert.equal(a, 1);
  assert.equal(b, 1);
});
