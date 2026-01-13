import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { sheetStateFromYjsDoc } from "../packages/versioning/src/yjs/sheetState.js";
import { workbookStateFromYjsDoc } from "../packages/versioning/src/yjs/workbookState.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

test("Yjs state extractors handle cross-instance nested Y.Maps (CJS updates applied into ESM doc)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const sheets = remote.getArray("sheets");
  const sheet = new Ycjs.Map();
  sheet.set("id", "sheet1");
  sheet.set("name", "Sheet1");
  sheet.set("visibility", "hidden");
  sheet.set("tabColor", "#00ff00");
  sheet.set("view", { frozenRows: 2, frozenCols: 1, colWidths: { "0": 120 } });
  sheets.push([sheet]);

  const cells = remote.getMap("cells");
  remote.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", 42);
    cell.set("formula", "=1+1");
    cells.set("sheet1:0:0", cell);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply using the CJS Yjs instance to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update);

  const sheetState = sheetStateFromYjsDoc(doc, { sheetId: "sheet1" });
  assert.equal(sheetState.cells.get("r0c0")?.value, 42);
  assert.equal(sheetState.cells.get("r0c0")?.formula, "=1+1");

  const workbookState = workbookStateFromYjsDoc(doc);
  assert.ok(workbookState.sheets.some((s) => s.id === "sheet1"));
  const meta = workbookState.sheets.find((s) => s.id === "sheet1");
  assert.ok(meta);
  assert.equal(meta.visibility, "hidden");
  assert.equal(meta.tabColor, "FF00FF00");
  assert.deepEqual(meta.view, { frozenRows: 2, frozenCols: 1 });
  assert.equal(workbookState.cellsBySheet.get("sheet1")?.cells.get("r0c0")?.value, 42);
});
