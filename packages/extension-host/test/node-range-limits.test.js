const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

test("ExtensionHost (node): rejects huge getRange/setRange before calling spreadsheet", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-range-"));

  let getRangeCalls = 0;
  let setRangeCalls = 0;

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    spreadsheet: {
      getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      getRange() {
        getRangeCalls += 1;
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      setRange() {
        setRangeCalls += 1;
      }
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extension = { id: "test" };

  await assert.rejects(() => host._executeApi("cells", "getRange", ["A1:Z10000"], extension), /too large/i);
  await assert.rejects(() => host._executeApi("cells", "setRange", ["A1:Z10000", []], extension), /too large/i);
  assert.equal(getRangeCalls, 0);
  assert.equal(setRangeCalls, 0);
});

test("ExtensionHost (node): truncates huge selectionChanged event payloads", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-selection-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    spreadsheet: {}
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any} */
  let lastMessage = null;
  host._extensions.set("ext", {
    id: "ext",
    active: true,
    worker: { postMessage: (msg) => (lastMessage = msg) }
  });

  // 10,000 rows x 26 cols = 260,000 cells (> 200,000 cap).
  host._broadcastEvent("selectionChanged", {
    selection: { startRow: 0, startCol: 0, endRow: 9999, endCol: 25, values: [[1]] }
  });

  assert.ok(lastMessage);
  assert.deepEqual(lastMessage.data.selection.values, []);
  assert.deepEqual(lastMessage.data.selection.formulas, []);
  assert.equal(lastMessage.data.selection.truncated, true);
});

test("ExtensionHost (node): allows partial matrices when huge selectionChanged payload is marked truncated", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-selection-truncated-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    spreadsheet: {}
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any} */
  let lastMessage = null;
  host._extensions.set("ext", {
    id: "ext",
    active: true,
    worker: { postMessage: (msg) => (lastMessage = msg) }
  });

  host._broadcastEvent("selectionChanged", {
    selection: {
      startRow: 0,
      startCol: 0,
      endRow: 9999,
      endCol: 25,
      values: [[1]],
      formulas: [],
      truncated: true
    }
  });

  assert.ok(lastMessage);
  assert.deepEqual(lastMessage.data.selection.values, [[1]]);
  assert.deepEqual(lastMessage.data.selection.formulas, []);
  assert.equal(lastMessage.data.selection.truncated, true);
});
