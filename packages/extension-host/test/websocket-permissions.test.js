const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

// Worker startup can be slow under heavy CI load. Keep the test focused on permission correctness
// rather than strict 5s activation/command SLAs.
const WEBSOCKET_TEST_TIMEOUT_MS = 20_000;

async function writeExtensionFixture(extensionDir, manifest, entrypointCode) {
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });
  await fs.writeFile(path.join(extensionDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), entrypointCode, "utf8");
}

test("permissions: WebSocket connections are blocked when network permission is denied", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-ws-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "wsExt.connectDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "ws-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "WebSocket Denied" }] },
      permissions: ["ui.commands", "network"]
    },
    `
      const formula = require("@formula/extension-api");

      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
           return await new Promise((resolve) => {
             const ws = new WebSocket("ws://example.invalid/");
            const timer = setTimeout(() => resolve({ status: "timeout" }), 2000);
             ws.addEventListener("close", (e) => {
               clearTimeout(timer);
               resolve({ status: "closed", code: e.code, reason: e.reason, wasClean: e.wasClean });
             });
           });
        }));
      };
    `
  );

  let sawNetworkPrompt = false;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: WEBSOCKET_TEST_TIMEOUT_MS,
    commandTimeoutMs: WEBSOCKET_TEST_TIMEOUT_MS,
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("network")) {
        sawNetworkPrompt = true;
        return false;
      }
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId);
  assert.equal(sawNetworkPrompt, true);
  assert.equal(result.status, "closed");
  assert.match(String(result.reason ?? ""), /Permission denied/);
});
