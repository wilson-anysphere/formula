const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: dispose notifies uiApi.onPanelDisposed for created panels", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, []);

  /** @type {string[]} */
  const disposedPanelIds = [];

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    uiApi: {
      onPanelDisposed: (panelId) => disposedPanelIds.push(String(panelId)),
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.dispose-panels";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://dispose-panels/",
    mainUrl: "memory://dispose-panels/main.mjs",
    manifest: {
      name: "dispose-panels",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      main: "./dist/extension.mjs",
      activationEvents: [],
      permissions: [],
      contributes: {},
    },
  });

  // Simulate a panel created via `ui.createPanel` while the extension is active.
  host._panels.set("test.panel", {
    id: "test.panel",
    title: "Test Panel",
    html: "<!doctype html><html><body>Test</body></html>",
    icon: null,
    position: null,
    extensionId,
    outgoingMessages: [],
  });

  await host.dispose();

  assert.deepEqual(disposedPanelIds, ["test.panel"]);
});

