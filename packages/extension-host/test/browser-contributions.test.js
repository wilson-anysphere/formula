const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: exposes manifest contributions for UI integration", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, []);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.contrib";
  const manifest = {
    name: "contrib",
    publisher: "test",
    version: "1.0.0",
    engines: { formula: "^1.0.0" },
    main: "./dist/extension.mjs",
    activationEvents: [],
    permissions: [],
    contributes: {
      commands: [{ command: "test.cmd", title: "Test Command", category: "Test", keywords: ["  foo  ", "", "bar"] }],
      panels: [{ id: "test.panel", title: "Test Panel" }],
      keybindings: [{ command: "test.cmd", key: "ctrl+shift+t", mac: "cmd+shift+t" }],
      menus: {
        "cell/context": [{ command: "test.cmd", when: "cellHasValue", group: "extensions" }]
      },
      customFunctions: [
        { name: "TEST_FUNC", description: "Test function", parameters: [], result: { type: "number" } }
      ],
      dataConnectors: [{ id: "test.connector", name: "Test Connector" }]
    }
  };

  await host.loadExtension({
    extensionId,
    extensionPath: "memory://contrib/",
    manifest,
    mainUrl: "memory://contrib/main.mjs"
  });

  assert.deepEqual(host.listExtensions().map((e) => e.id), [extensionId]);

  assert.deepEqual(host.getContributedCommands(), [
    {
      extensionId,
      command: "test.cmd",
      title: "Test Command",
      category: "Test",
      icon: null,
      description: null,
      keywords: ["foo", "bar"]
    }
  ]);

  assert.deepEqual(host.getContributedPanels(), [
    {
      extensionId,
      id: "test.panel",
      title: "Test Panel",
      icon: null
    }
  ]);

  assert.deepEqual(host.getContributedKeybindings(), [
    { extensionId, command: "test.cmd", key: "ctrl+shift+t", mac: "cmd+shift+t", when: null }
  ]);

  host._contextMenus.set("runtime", {
    id: "runtime",
    extensionId,
    menuId: "cell/context",
    items: [{ command: "test.runtime", when: null, group: "runtime" }]
  });

  assert.deepEqual(host.getContributedMenu("cell/context"), [
    {
      extensionId,
      command: "test.cmd",
      when: "cellHasValue",
      group: "extensions"
    },
    {
      extensionId,
      command: "test.runtime",
      when: null,
      group: "runtime"
    }
  ]);

  assert.deepEqual(host.getContributedCustomFunctions(), [
    { extensionId, name: "TEST_FUNC", description: "Test function" }
  ]);

  assert.deepEqual(host.getContributedDataConnectors(), [
    { extensionId, id: "test.connector", name: "Test Connector", icon: null }
  ]);
});

test("BrowserExtensionHost: unloadExtension clears contributed panel reservations", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{ onPostMessage(msg) { if (msg?.type === "init") return; } }, { onPostMessage(msg) { if (msg?.type === "init") return; } }]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId: "test.a",
    extensionPath: "memory://a/",
    mainUrl: "memory://a/main.mjs",
    manifest: {
      name: "a",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      activationEvents: [],
      permissions: [],
      contributes: { panels: [{ id: "shared.panel", title: "Shared Panel" }] }
    }
  });

  await host.unloadExtension("test.a");

  // If unloadExtension does not clear _panelContributions, this second load would throw
  // because the panel id is still considered owned by the first (now-unloaded) extension.
  await host.loadExtension({
    extensionId: "test.b",
    extensionPath: "memory://b/",
    mainUrl: "memory://b/main.mjs",
    manifest: {
      name: "b",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      activationEvents: [],
      permissions: [],
      contributes: { panels: [{ id: "shared.panel", title: "Shared Panel" }] }
    }
  });

  assert.deepEqual(host.listExtensions().map((e) => e.id).sort(), ["test.b"]);
});
