import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { workbookStateFromYjsDoc } from "../packages/versioning/src/yjs/workbookState.js";

test("workbookStateFromYjsDoc groups cells in a single scan of the cells root", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  const sheetCount = 50;
  const cellsPerSheet = 20;

  doc.transact(() => {
    for (let s = 0; s < sheetCount; s++) {
      const sheet = new Y.Map();
      sheet.set("id", `sheet${s}`);
      sheet.set("name", `Sheet ${s}`);
      sheets.push([sheet]);

      for (let r = 0; r < cellsPerSheet; r++) {
        const cell = new Y.Map();
        cell.set("value", `${s}-${r}`);
        cells.set(`sheet${s}:${r}:0`, cell);
      }
    }
  });

  let forEachCalls = 0;
  const originalForEach = cells.forEach;
  // Instrument the specific `cells` root map instance so we can assert the
  // workbook extractor doesn't re-scan it per sheet.
  cells.forEach = function (...args) {
    forEachCalls++;
    return originalForEach.apply(this, args);
  };

  try {
    const state = workbookStateFromYjsDoc(doc);
    assert.equal(forEachCalls, 1);
    assert.equal(state.cellsBySheet.size, sheetCount);
    assert.equal(state.cellsBySheet.get("sheet0")?.cells.get("r0c0")?.value ?? null, "0-0");
    assert.equal(
      state.cellsBySheet.get(`sheet${sheetCount - 1}`)?.cells.get(`r${cellsPerSheet - 1}c0`)?.value ?? null,
      `${sheetCount - 1}-${cellsPerSheet - 1}`,
    );
  } finally {
    cells.forEach = originalForEach;
  }
});

