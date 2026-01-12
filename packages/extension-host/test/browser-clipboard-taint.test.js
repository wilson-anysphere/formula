const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

function sortRanges(ranges) {
  return [...(ranges ?? [])].sort((a, b) => {
    const keyA = `${a.sheetId}:${a.startRow},${a.startCol}:${a.endRow},${a.endCol}`;
    const keyB = `${b.sheetId}:${b.startRow},${b.startCol}:${b.endRow},${b.endCol}`;
    return keyA.localeCompare(keyB);
  });
}

test("BrowserExtensionHost: records read taint for cells.getSelection/getCell/getRange", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getSheet(name) {
        if (String(name) === "Sheet2") return { id: "sheet2", name: "Sheet2" };
        return undefined;
      },
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 1, values: [[1, 2]] };
      },
      async getCell(_row, _col) {
        return 42;
      },
      async getRange(ref) {
        if (String(ref) !== "Sheet2!A1:B2") throw new Error(`Unexpected range ref: ${ref}`);
        return { startRow: 0, startCol: 0, endRow: 1, endCol: 1, values: [[1, 2], [3, 4]] };
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.taint";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "taint",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["cells.read"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  await host._executeApi("cells", "getSelection", [], extension);
  await host._executeApi("cells", "getCell", [2, 3], extension);
  await host._executeApi("cells", "getRange", ["Sheet2!A1:B2"], extension);

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
    { sheetId: "sheet1", startRow: 2, startCol: 3, endRow: 2, endCol: 3 },
    { sheetId: "sheet2", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: clipboard.writeText invokes guard with extensionId + tainted ranges", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  /** @type {any} */
  let guardArgs = null;
  let sawGuardBeforeWrite = false;
  const writes = [];

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return "secret";
      },
      async setCell() {},
    },
    clipboardApi: {
      readText: async () => "",
      writeText: async (text) => {
        assert.equal(sawGuardBeforeWrite, true);
        writes.push(String(text));
      },
    },
    clipboardWriteGuard: async (args) => {
      guardArgs = args;
      sawGuardBeforeWrite = true;
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.clipboard-guard";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "clipboard-guard",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["cells.read", "clipboard"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  await host._executeApi("cells", "getCell", [5, 6], extension);
  await host._executeApi("clipboard", "writeText", ["hello"], extension);

  assert.deepEqual(guardArgs, {
    extensionId,
    taintedRanges: [{ sheetId: "sheet1", startRow: 5, startCol: 6, endRow: 5, endCol: 6 }],
  });
  assert.deepEqual(writes, ["hello"]);
});

test("BrowserExtensionHost: clipboard.writeText blocks when guard throws", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const writes = [];

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return "secret";
      },
      async setCell() {},
    },
    clipboardApi: {
      readText: async () => "",
      writeText: async (text) => {
        writes.push(String(text));
      },
    },
    clipboardWriteGuard: async () => {
      throw new Error("blocked");
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.clipboard-block";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "clipboard-block",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["cells.read", "clipboard"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  await host._executeApi("cells", "getCell", [0, 0], extension);
  await assert.rejects(() => host._executeApi("clipboard", "writeText", ["nope"], extension), /blocked/);
  assert.deepEqual(writes, []);
});

