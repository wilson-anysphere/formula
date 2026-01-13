import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { workbookStateFromYjsDoc } from "../packages/versioning/src/yjs/workbookState.js";

test("workbookStateFromYjsDoc scans the Yjs `cells` map at most once (no O(#sheets * #cells) rescans)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  // Create many sheets to trigger the historical worst-case behavior where we scanned
  // the full `cells` map once per sheet.
  for (let i = 0; i < 50; i++) {
    const sheet = new Y.Map();
    sheet.set("id", `sheet${i}`);
    sheet.set("name", `Sheet ${i}`);
    sheets.push([sheet]);
  }

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cell.set("formula", null);
    cells.set("sheet0:0:0", cell);
  });

  let forEachCalls = 0;
  const originalForEach = cells.forEach.bind(cells);
  cells.forEach = (cb, thisArg) => {
    forEachCalls++;
    return originalForEach(cb, thisArg);
  };

  workbookStateFromYjsDoc(doc);

  // Optimized path should scan `cells` at most once. Allow <= 2 for safety in case we
  // add a second constant pass in the future (e.g. for validation).
  assert.ok(
    forEachCalls <= 2,
    `expected cells.forEach called <= 2 times, got ${forEachCalls} (regression: per-sheet rescans)`,
  );
});

