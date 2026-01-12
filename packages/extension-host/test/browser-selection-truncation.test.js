const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: selectionChanged payloads are truncated for huge selections", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {(msg: any) => void} */
  let resolveSelection;
  const selectionEvent = new Promise((resolve) => {
    resolveSelection = resolve;
  });

  installFakeWorker(t, [
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "event" && msg.event === "selectionChanged") {
          resolveSelection(msg);
        }
      }
    }
  ]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-truncation";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-truncation",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await host.startup();

  // 10,000 rows x 26 cols = 260,000 cells (> 200,000 cap).
  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: { startRow: 0, startCol: 0, endRow: 9999, endCol: 25, values: [[1]] }
  });

  const msg = await selectionEvent;
  assert.deepEqual(msg.data.selection.values, []);
  assert.deepEqual(msg.data.selection.formulas, []);
  assert.equal(msg.data.selection.truncated, true);
});

