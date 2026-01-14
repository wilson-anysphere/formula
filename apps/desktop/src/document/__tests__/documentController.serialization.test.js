import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("encodeState/applyState roundtrip restores cell inputs and clears history", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellFormula("Sheet1", "B1", "SUM(A1:A3)");
  doc.setRangeFormat("Sheet1", "A1", { bold: true });
  doc.setFrozen("Sheet1", 2, 1);
  doc.setColWidth("Sheet1", 0, 120);
  doc.setRowHeight("Sheet1", 1, 40);
  assert.equal(doc.canUndo, true);

  const snapshot = doc.encodeState();
  assert.ok(snapshot instanceof Uint8Array);

  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.equal(sheet.defaultFormat, null);
  assert.deepEqual(sheet.rowFormats, []);
  assert.deepEqual(sheet.colFormats, []);

  const restored = new DocumentController();
  let lastChange = null;
  restored.on("change", (payload) => {
    lastChange = payload;
  });
  restored.applyState(snapshot);

  // applyState clears history and marks dirty until the host explicitly marks saved.
  assert.equal(restored.canUndo, false);
  assert.equal(restored.canRedo, false);
  assert.equal(restored.isDirty, true);

  assert.equal(lastChange?.source, "applyState");
  assert.equal(restored.getCell("Sheet1", "A1").value, 1);
  const a1 = restored.getCell("Sheet1", "A1");
  assert.deepEqual(restored.styleTable.get(a1.styleId), { bold: true });
  assert.equal(restored.getCell("Sheet1", "B1").formula, "=SUM(A1:A3)");
  assert.deepEqual(restored.getSheetView("Sheet1"), {
    frozenRows: 2,
    frozenCols: 1,
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
});

test("encodeState/applyState roundtrip preserves layered column formatting", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const snapshot = doc.encodeState();
  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.deepEqual(sheet.defaultFormat, null);
  assert.deepEqual(sheet.rowFormats, []);
  assert.deepEqual(sheet.colFormats, [{ col: 0, format: { font: { bold: true } } }]);

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.deepEqual(restored.getCellFormat("Sheet1", "A1048576"), { font: { bold: true } });
  assert.equal(restored.model.sheets.get("Sheet1").cells.size, 0);
});

test("applyState materializes empty sheets from snapshots", () => {
  const doc = new DocumentController();
  // DocumentController lazily creates sheets on first access. This creates an empty sheet
  // that should still survive encode/apply roundtrips.
  doc.getCell("EmptySheet", "A1");

  const snapshot = doc.encodeState();
  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.deepEqual(restored.getSheetIds(), ["EmptySheet"]);
});

test("applyState removes sheets that are not present in the snapshot", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.getCell("ExtraSheet", "A1"); // empty sheet

  const next = new DocumentController();
  next.setCellValue("OnlySheet", "A1", 2);

  doc.applyState(next.encodeState());

  assert.deepEqual(doc.getSheetIds(), ["OnlySheet"]);
});

test("contentVersion increments when applyState adds/removes sheets (even if empty)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  assert.equal(doc.contentVersion, 1);
  assert.equal(doc.updateVersion, 1);

  // Add an *empty* sheet via applyState: no cell deltas, but sheet structure changes.
  const withExtraSheet = new DocumentController();
  withExtraSheet.setCellValue("Sheet1", "A1", 1);
  withExtraSheet.getCell("Sheet2", "A1"); // materialize Sheet2 without content

  doc.applyState(withExtraSheet.encodeState());
  assert.equal(doc.contentVersion, 2);
  assert.equal(doc.updateVersion, 2);
  assert.deepEqual(doc.getSheetIds().sort(), ["Sheet1", "Sheet2"]);

  // Remove the sheet via applyState: again, structure-only change.
  const withoutExtraSheet = new DocumentController();
  withoutExtraSheet.setCellValue("Sheet1", "A1", 1);

  doc.applyState(withoutExtraSheet.encodeState());
  assert.equal(doc.contentVersion, 3);
  assert.equal(doc.updateVersion, 3);
  assert.deepEqual(doc.getSheetIds().sort(), ["Sheet1"]);
});

test("encodeState/applyState preserves sheet insertion order", () => {
  const doc = new DocumentController();
  // Sheets are created lazily on first access. Create them in a non-sorted order.
  doc.getCell("Sheet2", "A1");
  doc.getCell("Sheet1", "A1");
  assert.deepEqual(doc.getSheetIds(), ["Sheet2", "Sheet1"]);

  const snapshot = doc.encodeState();
  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  assert.deepEqual(parsed.sheetOrder, ["Sheet2", "Sheet1"]);

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.deepEqual(restored.getSheetIds(), ["Sheet2", "Sheet1"]);

  // `sheetOrder` should be authoritative even if the `sheets` array is reordered (defensive).
  // Simulate a consumer that sorts the sheet objects but preserves the explicit order field.
  parsed.sheets.reverse();
  const tamperedSnapshot =
    typeof TextEncoder !== "undefined"
      ? new TextEncoder().encode(JSON.stringify(parsed))
      : // eslint-disable-next-line no-undef
        Buffer.from(JSON.stringify(parsed), "utf8");
  const restoredFromTampered = new DocumentController();
  restoredFromTampered.applyState(tamperedSnapshot);
  assert.deepEqual(restoredFromTampered.getSheetIds(), ["Sheet2", "Sheet1"]);

  // Also ensure applyState can reorder an existing controller to match the snapshot.
  const differentOrder = new DocumentController();
  differentOrder.getCell("Sheet1", "A1");
  differentOrder.getCell("Sheet2", "A1");
  assert.deepEqual(differentOrder.getSheetIds(), ["Sheet1", "Sheet2"]);

  let idsDuringChange = null;
  differentOrder.on("change", () => {
    idsDuringChange = differentOrder.getSheetIds();
  });
  differentOrder.applyState(snapshot);
  assert.deepEqual(idsDuringChange, ["Sheet2", "Sheet1"]);
  assert.deepEqual(differentOrder.getSheetIds(), ["Sheet2", "Sheet1"]);
});

test("update event fires on edits and undo/redo", () => {
  const doc = new DocumentController();
  let updates = 0;
  doc.on("update", () => {
    updates += 1;
  });

  doc.setCellValue("Sheet1", "A1", "x");
  doc.undo();
  doc.redo();

  assert.equal(updates, 3);
});

test("encodeState/applyState roundtrip preserves sheet metadata + order", () => {
  const doc = new DocumentController();
  // Materialize a few empty sheets.
  doc.getCell("S1", "A1");
  doc.getCell("S2", "A1");
  doc.getCell("S3", "A1");

  doc.renameSheet("S1", "Income");
  doc.setSheetVisibility("S2", "hidden");
  doc.setSheetTabColor("S3", { rgb: "FF00FF00", tint: 0.25 });
  doc.reorderSheets(["S3", "S1", "S2"]);

  const snapshot = doc.encodeState();
  const parsed = JSON.parse(new TextDecoder().decode(snapshot));
  assert.deepEqual(parsed.sheetOrder, ["S3", "S1", "S2"]);
  assert.deepEqual(
    parsed.sheets.map((s) => s.id),
    ["S3", "S1", "S2"]
  );

  const restored = new DocumentController();
  let lastChange = null;
  let idsDuringChange = null;
  let metaDuringChange = null;
  restored.on("change", (payload) => {
    lastChange = payload;
    idsDuringChange = restored.getSheetIds();
    metaDuringChange = restored.getSheetMeta("S1");
  });
  restored.applyState(snapshot);

  assert.deepEqual(idsDuringChange, ["S3", "S1", "S2"]);
  assert.deepEqual(metaDuringChange, { name: "Income", visibility: "visible" });

  assert.deepEqual(restored.getSheetIds(), ["S3", "S1", "S2"]);
  assert.deepEqual(restored.getSheetMeta("S1"), { name: "Income", visibility: "visible" });
  assert.deepEqual(restored.getSheetMeta("S2"), { name: "S2", visibility: "hidden" });
  assert.deepEqual(restored.getSheetMeta("S3"), {
    name: "S3",
    visibility: "visible",
    tabColor: { rgb: "FF00FF00", tint: 0.25 },
  });

  assert.equal(lastChange?.source, "applyState");
  assert.ok(Array.isArray(lastChange?.sheetMetaDeltas));
});

test("applyState accepts legacy snapshots without sheet metadata", () => {
  const legacy = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheets: [
        {
          id: "Sheet1",
          frozenRows: 0,
          frozenCols: 0,
          cells: [{ row: 0, col: 0, value: 123, formula: null, format: null }],
        },
      ],
    })
  );

  const doc = new DocumentController();
  doc.applyState(legacy);

  assert.equal(doc.getCell("Sheet1", "A1").value, 123);
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);
  assert.deepEqual(doc.getSheetMeta("Sheet1"), { name: "Sheet1", visibility: "visible" });
});

test("applyState accepts tabColor as an ARGB string (branch/collab snapshots)", () => {
  const snapshot = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheetOrder: ["Sheet1"],
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "visible",
          tabColor: "FF00FF00",
          frozenRows: 0,
          frozenCols: 0,
          cells: [],
        },
      ],
    })
  );

  const doc = new DocumentController();
  doc.applyState(snapshot);
  assert.deepEqual(doc.getSheetMeta("Sheet1"), {
    name: "Sheet1",
    visibility: "visible",
    tabColor: { rgb: "FF00FF00" },
  });
});

test("applyState accepts singleton-wrapped sheets/sheetOrder arrays and nested sheet/view/cells wrappers (interop)", () => {
  const snapshot = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheetOrder: { 0: ["Sheet2", "Sheet1"] },
      sheets: {
        0: [
          {
            0: {
              id: "Sheet2",
              name: { 0: " Sheet Two " },
              visibility: { 0: "hidden" },
              tabColor: { 0: "ff00ff00" },
              // Place view fields only under a wrapped `view` object (no top-level frozenRows/Cols).
              view: { 0: { frozenRows: { 0: 2 }, frozenCols: [1] } },
              // Cells list wrapped as `{0:[...]}`.
              cells: { 0: [{ row: 0, col: 0, value: 123, formula: null, format: null }] },
            },
          },
          { id: "Sheet1", cells: [] },
        ],
      },
    }),
  );

  const doc = new DocumentController();
  doc.applyState(snapshot);

  assert.deepEqual(doc.getSheetIds(), ["Sheet2", "Sheet1"]);
  assert.deepEqual(doc.getSheetMeta("Sheet2"), {
    name: "Sheet Two",
    visibility: "hidden",
    tabColor: { rgb: "FF00FF00" },
  });
  assert.deepEqual(doc.getSheetMeta("Sheet1"), { name: "Sheet1", visibility: "visible" });
  assert.equal(doc.getSheetView("Sheet2").frozenRows, 2);
  assert.equal(doc.getSheetView("Sheet2").frozenCols, 1);
  assert.equal(doc.getCell("Sheet2", "A1").value, 123);
});
