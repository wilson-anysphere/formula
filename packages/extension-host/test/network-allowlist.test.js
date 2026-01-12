const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeExtensionFixture(extensionDir, manifest, entrypointCode) {
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });
  await fs.writeFile(path.join(extensionDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), entrypointCode, "utf8");
}

test("network allowlist: blocks non-allowlisted hosts but allows allowed hosts", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-net-allowlist-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "netAllowlist.fetch";
  const manifest = {
    name: "net-allowlist-ext",
    version: "1.0.0",
    publisher: "formula-test",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${commandId}`],
    contributes: { commands: [{ command: commandId, title: "Net Allowlist" }] },
    permissions: ["ui.commands", "network"]
  };

  await writeExtensionFixture(
    extDir,
    manifest,
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async (url) => {
          const res = await fetch(String(url));
          return await res.text();
        }));
      };
    `
  );

  const extensionId = `${manifest.publisher}.${manifest.name}`;
  const permissionsPath = path.join(dir, "permissions.json");
  await fs.writeFile(
    permissionsPath,
    JSON.stringify(
      {
        [extensionId]: {
          network: { mode: "allowlist", hosts: ["allowed.example"] }
        }
      },
      null,
      2
    ),
    "utf8"
  );

  const prevFetch = globalThis.fetch;
  globalThis.fetch = async (url) => {
    const u = new URL(String(url));
    if (u.hostname !== "allowed.example") {
      throw new Error(`Unexpected network call to ${u.hostname}`);
    }
    return new Response("ok", { status: 200, headers: { "content-type": "text/plain" } });
  };

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: permissionsPath,
    extensionStoragePath: path.join(dir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep this test focused on the
    // allowlist behavior rather than the default 5s activation SLA.
    activationTimeoutMs: 20_000,
    permissionPrompt: async ({ permissions }) => {
      // The allowlisted host should not prompt; blocked hosts should prompt and be denied.
      if (permissions.includes("network")) return false;
      return true;
    }
  });

  t.after(async () => {
    globalThis.fetch = prevFetch;
    await host.dispose();
  });

  await host.loadExtension(extDir);

  assert.equal(await host.executeCommand(commandId, "https://allowed.example/"), "ok");
  await assert.rejects(() => host.executeCommand(commandId, "https://blocked.example/"), /Permission denied/i);
});

test("permissions: revokePermissions blocks subsequent network calls until re-granted", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-net-revoke-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "netRevoke.fetch";
  const manifest = {
    name: "net-revoke-ext",
    version: "1.0.0",
    publisher: "formula-test",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${commandId}`],
    contributes: { commands: [{ command: commandId, title: "Net Revoke" }] },
    permissions: ["ui.commands", "network"]
  };

  await writeExtensionFixture(
    extDir,
    manifest,
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async (url) => {
          const res = await fetch(String(url));
          return await res.text();
        }));
      };
    `
  );

  const extensionId = `${manifest.publisher}.${manifest.name}`;
  const prevFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response("ok", { status: 200 });

  let networkPrompts = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("network")) {
        networkPrompts += 1;
        return networkPrompts === 1;
      }
      return true;
    }
  });

  t.after(async () => {
    globalThis.fetch = prevFetch;
    await host.dispose();
  });

  await host.loadExtension(extDir);

  assert.equal(await host.executeCommand(commandId, "https://allowed.example/"), "ok");
  assert.equal(networkPrompts, 1);

  await host.revokePermissions(extensionId, ["network"]);
  const after = await host.getGrantedPermissions(extensionId);
  assert.ok(!after.network, "Expected network permission to be revoked");

  await assert.rejects(() => host.executeCommand(commandId, "https://allowed.example/"), /Permission denied/i);
  assert.equal(networkPrompts, 2);
});
