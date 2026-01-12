const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: rejects huge getRange/setRange before calling spreadsheet API", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  let getRangeCalls = 0;
  let setRangeCalls = 0;

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getRange() {
        getRangeCalls += 1;
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async setRange() {
        setRangeCalls += 1;
      }
    },
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.range-limits";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "range-limits",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["cells.read", "cells.write"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  await assert.rejects(() => host._executeApi("cells", "getRange", ["A1:Z10000"], extension), /too large/i);
  await assert.rejects(() => host._executeApi("cells", "setRange", ["A1:Z10000", []], extension), /too large/i);
  assert.equal(getRangeCalls, 0);
  assert.equal(setRangeCalls, 0);
});

