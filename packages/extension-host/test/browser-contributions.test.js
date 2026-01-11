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
      commands: [{ command: "test.cmd", title: "Test Command", category: "Test" }],
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
      icon: null
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
