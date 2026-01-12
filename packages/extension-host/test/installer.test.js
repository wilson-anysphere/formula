const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const {
  ExtensionHost,
  installExtensionFromDirectory,
  uninstallExtension,
  listInstalledExtensions
} = require("../src");

test("installer workflow: install -> list -> load -> uninstall", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-install-"));
  const installRoot = path.join(tmpDir, "extensions");

  const source = path.resolve(__dirname, "../../../extensions/sample-hello");
  const installedPath = await installExtensionFromDirectory(source, installRoot);

  const installed = await listInstalledExtensions(installRoot);
  assert.deepEqual(installed, ["formula.sample-hello"]);

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(tmpDir, "permissions.json"),
    extensionStoragePath: path.join(tmpDir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep this flow test focused on the
    // installer behavior rather than the default 5s activation SLA.
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  await host.loadExtension(installedPath);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setCell(1, 0, 3);
  host.spreadsheet.setCell(1, 1, 4);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

  const result = await host.executeCommand("sampleHello.sumSelection");
  assert.equal(result, 10);

  await host.dispose();

  await uninstallExtension(installRoot, "formula.sample-hello");
  assert.deepEqual(await listInstalledExtensions(installRoot), []);
});

test("installExtensionFromDirectory rejects manifests with unsafe publisher/name segments", async (t) => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-install-unsafe-"));
  t.after(async () => {
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  const source = path.join(tmpDir, "src");
  const installRoot = path.join(tmpDir, "extensions");
  await fs.mkdir(path.join(source, "dist"), { recursive: true });
  await fs.writeFile(
    path.join(source, "package.json"),
    JSON.stringify(
      {
        name: "evil",
        publisher: "../escape",
        version: "1.0.0",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
      },
      null,
      2
    )
  );
  await fs.writeFile(path.join(source, "dist", "extension.js"), "module.exports = {};\n");

  await assert.rejects(
    () => installExtensionFromDirectory(source, installRoot),
    /Invalid extension id/
  );
});
