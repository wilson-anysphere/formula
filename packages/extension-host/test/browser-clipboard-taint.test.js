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

test("BrowserExtensionHost: trims sheetId when normalizing tainted ranges", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.taint-trim-sheet-id";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "taint-trim-sheet-id",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);

  host._taintExtensionRange(extension, { sheetId: "  sheet1  ", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });

  assert.deepEqual(extension.taintedRanges, [{ sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
});

test("BrowserExtensionHost: cells.getCell taints numeric-string coordinates", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  /** @type {any} */
  const calls = [];

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getCell(row, col) {
        calls.push({ row, col });
        return 123;
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.cell-coerce";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "cell-coerce",
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

  await host._executeApi("cells", "getCell", ["2", "3"], extension);

  assert.deepEqual(calls, [{ row: 2, col: 3 }]);

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 2, startCol: 3, endRow: 2, endCol: 3 },
  ]);
});

test("BrowserExtensionHost: cells.getRange taints using result.address when ref is not A1", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getRange(ref) {
        if (String(ref) !== "NamedRange") throw new Error(`Unexpected range ref: ${ref}`);
        return { address: "A1:B2", values: [[1, 2], [3, 4]] };
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.range-address-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "range-address-fallback",
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

  await host._executeApi("cells", "getRange", ["NamedRange"], extension);

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: cells.getRange taints numeric-string coords when address is missing", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getRange(ref) {
        if (String(ref) !== "NamedRange") throw new Error(`Unexpected range ref: ${ref}`);
        return { startRow: "0", startCol: "0", endRow: "1", endCol: "1", values: [[1, 2], [3, 4]] };
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.range-string-coords";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "range-string-coords",
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

  await host._executeApi("cells", "getRange", ["NamedRange"], extension);

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: cells.getRange prefers sheet-qualified result.address when provided", async (t) => {
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
      async getRange(ref) {
        if (String(ref) !== "NamedRange") throw new Error(`Unexpected range ref: ${ref}`);
        return {
          address: "Sheet2!A1:B2",
          startRow: 0,
          startCol: 0,
          endRow: 1,
          endCol: 1,
          values: [[1, 2], [3, 4]],
        };
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.range-sheet-qualified-address";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "range-sheet-qualified-address",
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

  await host._executeApi("cells", "getRange", ["NamedRange"], extension);

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet2", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: cells.getSelection taints using result.address when coords are missing", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getSelection() {
        return { address: "A1:B2", values: [[1, 2], [3, 4]] };
      },
      async getCell() {
        return null;
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-address-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-address-fallback",
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

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: does not taint from events for inactive extensions", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.event-taint-inactive";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "event-taint-inactive",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  assert.equal(extension.active, false);

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 1, values: [[1, 2]] },
  });

  assert.deepEqual(extension.taintedRanges, []);
});

test("BrowserExtensionHost: does not taint selectionChanged when values are empty (truncated payload)", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.event-taint-truncated";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "event-taint-truncated",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 999,
      endCol: 999,
      values: [],
      truncated: true,
    },
  });

  assert.deepEqual(extension.taintedRanges, []);
});

test("BrowserExtensionHost: does not taint cellChanged when value is missing", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.event-taint-cell-missing-value";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "event-taint-cell-missing-value",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("cellChanged", { sheetId: "sheet1", row: 0, col: 0 });

  assert.deepEqual(extension.taintedRanges, []);
});

test("BrowserExtensionHost: does not taint cellChanged when value is undefined", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.event-taint-cell-undefined";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "event-taint-cell-undefined",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("cellChanged", { sheetId: "sheet1", row: 0, col: 0, value: undefined });

  assert.deepEqual(extension.taintedRanges, []);
});

test("BrowserExtensionHost: records taint for selectionChanged events (and passes it to clipboard guard)", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  /** @type {any} */
  let guardArgs = null;

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async setCell() {},
    },
    clipboardApi: {
      readText: async () => "",
      writeText: async () => {},
    },
    clipboardWriteGuard: async (args) => {
      guardArgs = args;
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-taint";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-taint",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["clipboard"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 1, values: [[1, 2]] },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
  ]);

  await host._executeApi("clipboard", "writeText", ["hello"], extension);

  assert.deepEqual(guardArgs, {
    extensionId,
    taintedRanges: [{ sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
  });
});

test("BrowserExtensionHost: selectionChanged taints when formulas matrix includes non-empty strings", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-formulas";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-formulas",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
      values: [],
      formulas: [["=A1"]],
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
  ]);
});

test("BrowserExtensionHost: selectionChanged taints formulas that extend beyond values matrix bounds", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-values-and-formulas";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-values-and-formulas",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 10,
      values: [[1]],
      formulas: [["=A1", "=B1"]],
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: selectionChanged does not over-taint when values/formulas bounds form an L-shape", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-l-shape";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-l-shape",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  // Values cover A1:A2, formulas cover A1:B1. The union would be A1:B2, but that would over-taint
  // B2, which was never delivered. Keep the L-shape as two rectangles.
  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 10,
      endCol: 10,
      values: [[1], [2]],
      formulas: [["=A1", "=B1"]],
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 0 },
  ]);
});

test("BrowserExtensionHost: taints selectionChanged events when payload is truncated but includes values", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-truncated";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-truncated",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
      values: [[1]],
      truncated: true,
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
  ]);
});

test("BrowserExtensionHost: taints truncated selectionChanged events when values are included", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-truncated-with-values";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-truncated-with-values",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    sheetId: "sheet1",
    selection: {
      startRow: 10,
      startCol: 20,
      endRow: 10,
      endCol: 30,
      values: [[1, 2, 3]],
      truncated: true,
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 10, startCol: 20, endRow: 10, endCol: 22 },
  ]);
});

test("BrowserExtensionHost: selectionChanged taints using active-sheet fallback when sheetId omitted", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-taint-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-taint-fallback",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("sheetActivated", { sheet: { id: "sheet2", name: "Sheet2" } });
  host._broadcastEvent("selectionChanged", {
    selection: {
      startRow: 1,
      startCol: 2,
      endRow: 3,
      endCol: 4,
      values: [
        [null, null, null],
        [null, null, null],
        [null, null, null],
      ],
    },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet2", startRow: 1, startCol: 2, endRow: 3, endCol: 4 },
  ]);
});

test("BrowserExtensionHost: selectionChanged taints using selection.address when coords are missing", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.selection-event-address-taint";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "selection-event-address-taint",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("selectionChanged", {
    selection: { address: "A1:B2", values: [[1, 2], [3, 4]] },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: cellChanged taints using active-sheet fallback when sheetId omitted", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.cell-event-taint-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "cell-event-taint-fallback",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("sheetActivated", { sheet: { id: "sheet2", name: "Sheet2" } });
  host._broadcastEvent("cellChanged", { row: 9, col: 8, value: "x" });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet2", startRow: 9, startCol: 8, endRow: 9, endCol: 8 },
  ]);
});

test("BrowserExtensionHost: cellChanged taints using address when row/col are missing", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.cell-event-address-taint";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "cell-event-address-taint",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("cellChanged", { address: "B2", value: "x" });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 1, startCol: 1, endRow: 1, endCol: 1 },
  ]);
});

test("BrowserExtensionHost: sheets.activateSheet updates active-sheet id for event taint fallback", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  let activeSheetId = "sheet1";
  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async activateSheet(name) {
        if (String(name) === "Sheet2") {
          activeSheetId = "sheet2";
          return { id: "sheet2", name: "Sheet2" };
        }
        return { id: activeSheetId, name: "Sheet1" };
      },
      async getActiveSheet() {
        return { id: activeSheetId, name: activeSheetId === "sheet2" ? "Sheet2" : "Sheet1" };
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.activate-sheet-taint-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "activate-sheet-taint-fallback",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  await host._executeApi("sheets", "activateSheet", ["Sheet2"], extension);
  host._broadcastEvent("selectionChanged", {
    selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] },
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet2", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
  ]);
});

test("BrowserExtensionHost: records taint for cellChanged events (and passes it to clipboard guard)", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  /** @type {any} */
  let guardArgs = null;

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async setCell() {},
    },
    clipboardApi: {
      readText: async () => "",
      writeText: async () => {},
    },
    clipboardWriteGuard: async (args) => {
      guardArgs = args;
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.cell-event-taint";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "cell-event-taint",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["clipboard"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  extension.active = true;

  host._broadcastEvent("cellChanged", {
    sheetId: "sheet1",
    row: 5,
    col: 6,
    value: "secret",
  });

  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 5, startCol: 6, endRow: 5, endCol: 6 },
  ]);

  await host._executeApi("clipboard", "writeText", ["hello"], extension);

  assert.deepEqual(guardArgs, {
    extensionId,
    taintedRanges: [{ sheetId: "sheet1", startRow: 5, startCol: 6, endRow: 5, endCol: 6 }],
  });
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

test("BrowserExtensionHost: taint persists across worker termination + restart", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  // Two workers: initial load + restart after termination.
  installFakeWorker(t, [{}, {}]);

  /** @type {any} */
  let guardArgs = null;

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      async getCell() {
        return "secret";
      },
      async setCell() {},
    },
    clipboardApi: {
      readText: async () => "",
      writeText: async () => {},
    },
    clipboardWriteGuard: async (args) => {
      guardArgs = args;
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.taint-persist";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "taint-persist",
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
  extension.active = true;

  await host._executeApi("cells", "getCell", [5, 6], extension);
  assert.deepEqual(sortRanges(extension.taintedRanges), [
    { sheetId: "sheet1", startRow: 5, startCol: 6, endRow: 5, endCol: 6 },
  ]);

  host._terminateWorker(extension, { reason: "crash", cause: new Error("boom") });

  // Restart the worker and ensure the clipboard guard still sees the previous taint.
  await host._ensureWorker(extension);
  extension.active = true;

  await host._executeApi("clipboard", "writeText", ["hello"], extension);

  assert.deepEqual(guardArgs, {
    extensionId,
    taintedRanges: [{ sheetId: "sheet1", startRow: 5, startCol: 6, endRow: 5, endCol: 6 }],
  });
});

test("BrowserExtensionHost: taint list is capped to the most recent ranges", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

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
        return null;
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.taint-cap";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "taint-cap",
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

  // Add more than the cap (50) unique single-cell ranges. Use distinct rows/cols so
  // no merging occurs.
  for (let i = 0; i < 60; i += 1) {
    // eslint-disable-next-line no-await-in-loop
    await host._executeApi("cells", "getCell", [i, i], extension);
  }

  assert.equal(extension.taintedRanges.length, 50);
  assert.deepEqual(extension.taintedRanges[0], { sheetId: "sheet1", startRow: 10, startCol: 10, endRow: 10, endCol: 10 });
  assert.deepEqual(extension.taintedRanges[49], { sheetId: "sheet1", startRow: 59, startCol: 59, endRow: 59, endCol: 59 });
});

test("BrowserExtensionHost: cells.getRange without spreadsheetApi.getRange rejects non-active sheet-qualified refs", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  installFakeWorker(t, [{}]);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      getActiveSheet() {
        return { id: "sheet1", name: "Sheet1" };
      },
      listSheets() {
        return [
          { id: "sheet1", name: "Sheet1" },
          { id: "sheet2", name: "Data" },
        ];
      },
      async getSheet(name) {
        const n = String(name);
        if (n === "Sheet1") return { id: "sheet1", name: "Sheet1" };
        if (n === "Data") return { id: "sheet2", name: "Data" };
        return undefined;
      },
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell(_row, _col) {
        return null;
      },
      async setCell() {},
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.range-sheet-qualified";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "range-sheet-qualified",
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

  await assert.rejects(
    () => host._executeApi("cells", "getRange", ["Data!A1:A1"], extension),
    /Sheet-qualified ranges require spreadsheetApi\.getRange/
  );
});
