import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import { DocumentControllerBridge } from "../src/document-controller-bridge.js";

test("DocumentControllerBridge resolves sheet ids by (case-insensitive) sheet name", () => {
  const doc = new DocumentController();
  const alphaId = doc.addSheet({ sheetId: "sheet_alpha", name: "Alpha" });
  const betaId = doc.addSheet({ sheetId: "sheet_beta", name: "Beta" });

  const bridge = new DocumentControllerBridge(doc, { activeSheetId: alphaId });

  assert.equal(bridge.get_sheet_id({ name: "Alpha" }), alphaId);
  assert.equal(bridge.get_sheet_id({ name: "alpha" }), alphaId);
  assert.equal(bridge.get_sheet_id({ name: "Beta" }), betaId);
  assert.equal(bridge.get_sheet_id({ name: "Missing" }), null);

  // Back-compat: allow direct sheet_id lookup.
  assert.equal(bridge.get_sheet_id({ name: betaId }), betaId);
});

test("DocumentControllerBridge get_sheet_name and rename_sheet use DocumentController metadata (stable ids)", () => {
  const doc = new DocumentController();
  const sheetId = doc.addSheet({ sheetId: "sheet_1", name: "First" });

  const bridge = new DocumentControllerBridge(doc, { activeSheetId: sheetId });
  assert.equal(bridge.get_sheet_name({ sheet_id: sheetId }), "First");

  bridge.rename_sheet({ sheet_id: sheetId, name: "Renamed" });
  assert.equal(bridge.get_sheet_name({ sheet_id: sheetId }), "Renamed");
  assert.equal(bridge.get_sheet_id({ name: "Renamed" }), sheetId);
  assert.equal(bridge.get_sheet_id({ name: "First" }), null);
});

test("DocumentControllerBridge create_sheet inserts after active sheet by default", () => {
  const doc = new DocumentController();
  const sheetA = doc.addSheet({ sheetId: "sheet_a", name: "A" });
  const sheetB = doc.addSheet({ sheetId: "sheet_b", name: "B" });

  const bridge = new DocumentControllerBridge(doc, { activeSheetId: sheetA });

  const newId = bridge.create_sheet({ name: "Inserted" });
  const order = doc.getSheetIds();

  assert.deepEqual(order, [sheetA, newId, sheetB]);
  assert.equal(doc.getSheetMeta(newId)?.name, "Inserted");
});

test("DocumentControllerBridge create_sheet materializes the active sheet when the doc is empty", () => {
  const doc = new DocumentController();
  const bridge = new DocumentControllerBridge(doc, { activeSheetId: "Sheet1" });

  const newId = bridge.create_sheet({ name: "Inserted" });
  assert.deepEqual(doc.getSheetIds(), ["Sheet1", newId]);
  assert.equal(doc.getSheetMeta(newId)?.name, "Inserted");
});

test("DocumentControllerBridge create_sheet honors explicit index (0-based) over active sheet", () => {
  const doc = new DocumentController();
  const sheetA = doc.addSheet({ sheetId: "sheet_a", name: "A" });
  const sheetB = doc.addSheet({ sheetId: "sheet_b", name: "B" });
  const sheetC = doc.addSheet({ sheetId: "sheet_c", name: "C" });

  const bridge = new DocumentControllerBridge(doc, { activeSheetId: sheetB });

  const at0 = bridge.create_sheet({ name: "At0", index: 0 });
  assert.deepEqual(doc.getSheetIds(), [at0, sheetA, sheetB, sheetC]);

  const at2 = bridge.create_sheet({ name: "At2", index: 2 });
  assert.deepEqual(doc.getSheetIds(), [at0, sheetA, at2, sheetB, sheetC]);

  const atEnd = bridge.create_sheet({ name: "AtEnd", index: 99 });
  assert.deepEqual(doc.getSheetIds(), [at0, sheetA, at2, sheetB, sheetC, atEnd]);
});

test("DocumentControllerBridge get_range_values rejects huge ranges before scanning cells", () => {
  let scanned = 0;
  const doc = {
    getCell() {
      scanned += 1;
      throw new Error("Should not scan");
    },
  };
  const bridge = new DocumentControllerBridge(doc, { activeSheetId: "Sheet1" });

  assert.throws(
    () =>
      bridge.get_range_values({
        range: { sheet_id: "Sheet1", start_row: 0, end_row: 7999, start_col: 0, end_col: 25 },
      }),
    /get_range_values skipped/i,
  );
  assert.equal(scanned, 0);
});

test("DocumentControllerBridge set_range_values rejects huge scalar fills before materializing a matrix", () => {
  let wrote = 0;
  const doc = {
    setRangeValues() {
      wrote += 1;
      throw new Error("Should not write");
    },
  };
  const bridge = new DocumentControllerBridge(doc, { activeSheetId: "Sheet1" });

  assert.throws(
    () =>
      bridge.set_range_values({
        range: { sheet_id: "Sheet1", start_row: 0, end_row: 7999, start_col: 0, end_col: 25 },
        values: 1,
      }),
    /set_range_values skipped/i,
  );
  assert.equal(wrote, 0);
});

test("DocumentControllerBridge set_range_values spills matrix writes when the destination is a single cell", () => {
  const doc = new DocumentController();
  const sheetId = doc.addSheet({ sheetId: "sheet_1", name: "Sheet1" });
  const bridge = new DocumentControllerBridge(doc, { activeSheetId: sheetId });

  bridge.set_range_values({
    range: { sheet_id: sheetId, start_row: 0, end_row: 0, start_col: 0, end_col: 0 },
    values: [
      [1, 2],
      [3, 4],
    ],
  });

  assert.equal(doc.getCell(sheetId, { row: 0, col: 0 }).value, 1);
  assert.equal(doc.getCell(sheetId, { row: 0, col: 1 }).value, 2);
  assert.equal(doc.getCell(sheetId, { row: 1, col: 0 }).value, 3);
  assert.equal(doc.getCell(sheetId, { row: 1, col: 1 }).value, 4);
});
