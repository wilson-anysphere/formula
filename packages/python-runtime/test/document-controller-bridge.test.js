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

