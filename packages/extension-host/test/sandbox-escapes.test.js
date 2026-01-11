const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeExtension(dir, { name, publisher, main, activateBody }) {
  await fs.mkdir(dir, { recursive: true });

  const manifest = {
    name,
    displayName: name,
    version: "1.0.0",
    publisher,
    main,
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${publisher}.${name}.activate`],
    contributes: {
      commands: [
        {
          command: `${publisher}.${name}.activate`,
          title: "Activate"
        }
      ]
    },
    permissions: []
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(
    path.join(dir, main),
    `module.exports.activate = async () => {\n${activateBody}\n};\n`,
    "utf8"
  );

  return manifest.contributes.commands[0].command;
}

test("sandbox: blocks dynamic import('node:fs/promises')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-import-fs-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-import-fs",
    publisher: "formula",
    main: "extension.js",
    activateBody: "await import('node:fs/promises');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(
    () => host.executeCommand(commandId),
    /Dynamic import is not allowed in extensions.*node:fs\/promises/
  );
});

test("sandbox: blocks process.binding('fs')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-binding-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-binding",
    publisher: "formula",
    main: "extension.js",
    activateBody: "process.binding('fs');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /process\.binding\(\) is not allowed/);
});

test("sandbox: blocks dynamic import('node:http2')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-import-http2-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-import-http2",
    publisher: "formula",
    main: "extension.js",
    activateBody: "await import('node:http2');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(
    () => host.executeCommand(commandId),
    /Dynamic import is not allowed in extensions.*node:http2/
  );
});

