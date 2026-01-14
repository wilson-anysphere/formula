import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("insertCellsShiftRight shifts stored cells (value + styleId) to the right", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "A");
  doc.setCellValue("Sheet1", "B1", "B");
  doc.setCellValue("Sheet1", "C1", "C");

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "B1", { font: { italic: true } });
  doc.setRangeFormat("Sheet1", "C1", { font: { underline: true } });

  const a1 = doc.getCell("Sheet1", "A1");
  const b1 = doc.getCell("Sheet1", "B1");
  const c1 = doc.getCell("Sheet1", "C1");

  doc.insertCellsShiftRight("Sheet1", { startRow: 0, endRow: 0, startCol: 0, endCol: 1 });

  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);
  assert.equal(doc.getCell("Sheet1", "B1").value, null);
  assert.equal(doc.getCell("Sheet1", "B1").styleId, 0);

  assert.equal(doc.getCell("Sheet1", "C1").value, "A");
  assert.equal(doc.getCell("Sheet1", "C1").styleId, a1.styleId);
  assert.equal(doc.getCell("Sheet1", "D1").value, "B");
  assert.equal(doc.getCell("Sheet1", "D1").styleId, b1.styleId);
  assert.equal(doc.getCell("Sheet1", "E1").value, "C");
  assert.equal(doc.getCell("Sheet1", "E1").styleId, c1.styleId);
});

test("insertCellsShiftDown shifts stored cells (value + styleId) down", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "A");
  doc.setCellValue("Sheet1", "A2", "A2");
  doc.setCellValue("Sheet1", "B1", "B");
  doc.setCellValue("Sheet1", "B2", "B2");

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "A2", { font: { italic: true } });
  doc.setRangeFormat("Sheet1", "B1", { font: { underline: true } });
  doc.setRangeFormat("Sheet1", "B2", { font: { strike: true } });

  const a1 = doc.getCell("Sheet1", "A1");
  const a2 = doc.getCell("Sheet1", "A2");
  const b1 = doc.getCell("Sheet1", "B1");
  const b2 = doc.getCell("Sheet1", "B2");

  doc.insertCellsShiftDown("Sheet1", { startRow: 0, endRow: 0, startCol: 0, endCol: 1 });

  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "A1").styleId, 0);
  assert.equal(doc.getCell("Sheet1", "B1").value, null);
  assert.equal(doc.getCell("Sheet1", "B1").styleId, 0);

  assert.equal(doc.getCell("Sheet1", "A2").value, "A");
  assert.equal(doc.getCell("Sheet1", "A2").styleId, a1.styleId);
  assert.equal(doc.getCell("Sheet1", "B2").value, "B");
  assert.equal(doc.getCell("Sheet1", "B2").styleId, b1.styleId);

  assert.equal(doc.getCell("Sheet1", "A3").value, "A2");
  assert.equal(doc.getCell("Sheet1", "A3").styleId, a2.styleId);
  assert.equal(doc.getCell("Sheet1", "B3").value, "B2");
  assert.equal(doc.getCell("Sheet1", "B3").styleId, b2.styleId);
});

test("deleteCellsShiftLeft shifts stored cells (value + styleId) left", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "A");
  doc.setCellValue("Sheet1", "B1", "B");
  doc.setCellValue("Sheet1", "C1", "C");
  doc.setCellValue("Sheet1", "D1", "D");

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "B1", { font: { italic: true } });
  doc.setRangeFormat("Sheet1", "C1", { font: { underline: true } });
  doc.setRangeFormat("Sheet1", "D1", { font: { strike: true } });

  const a1 = doc.getCell("Sheet1", "A1");
  const d1 = doc.getCell("Sheet1", "D1");

  doc.deleteCellsShiftLeft("Sheet1", { startRow: 0, endRow: 0, startCol: 1, endCol: 2 });

  assert.equal(doc.getCell("Sheet1", "A1").value, "A");
  assert.equal(doc.getCell("Sheet1", "A1").styleId, a1.styleId);

  assert.equal(doc.getCell("Sheet1", "B1").value, "D");
  assert.equal(doc.getCell("Sheet1", "B1").styleId, d1.styleId);

  assert.equal(doc.getCell("Sheet1", "C1").value, null);
  assert.equal(doc.getCell("Sheet1", "C1").styleId, 0);
  assert.equal(doc.getCell("Sheet1", "D1").value, null);
  assert.equal(doc.getCell("Sheet1", "D1").styleId, 0);
});

test("deleteCellsShiftUp shifts stored cells (value + styleId) up", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "A1");
  doc.setCellValue("Sheet1", "A2", "A2");
  doc.setCellValue("Sheet1", "A3", "A3");

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "A2", { font: { italic: true } });
  doc.setRangeFormat("Sheet1", "A3", { font: { underline: true } });

  const a3 = doc.getCell("Sheet1", "A3");

  doc.deleteCellsShiftUp("Sheet1", { startRow: 0, endRow: 1, startCol: 0, endCol: 0 });

  assert.equal(doc.getCell("Sheet1", "A1").value, "A3");
  assert.equal(doc.getCell("Sheet1", "A1").styleId, a3.styleId);

  assert.equal(doc.getCell("Sheet1", "A2").value, null);
  assert.equal(doc.getCell("Sheet1", "A2").styleId, 0);
  assert.equal(doc.getCell("Sheet1", "A3").value, null);
  assert.equal(doc.getCell("Sheet1", "A3").styleId, 0);
});

test("insertCellsShiftDown applies formula rewrites (Excel-style) when provided", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 42);
  doc.setCellFormula("Sheet1", "B1", "=A1");

  doc.insertCellsShiftDown(
    "Sheet1",
    { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
    {
      formulaRewrites: [{ sheet: "Sheet1", address: "B1", before: "=A1", after: "=A2" }],
    },
  );

  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "A2").value, 42);
  assert.equal(doc.getCell("Sheet1", "B1").formula, "=A2");
});

test("insert/delete cells shifts are undoable + redoable", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "A");
  doc.setCellValue("Sheet1", "B1", "B");
  doc.setCellValue("Sheet1", "C1", "C");
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  doc.setRangeFormat("Sheet1", "B1", { font: { italic: true } });
  doc.setRangeFormat("Sheet1", "C1", { font: { underline: true } });

  const before = {
    A1: doc.getCell("Sheet1", "A1"),
    B1: doc.getCell("Sheet1", "B1"),
    C1: doc.getCell("Sheet1", "C1"),
  };

  doc.insertCellsShiftRight("Sheet1", "A1:B1");
  const afterInsert = {
    A1: doc.getCell("Sheet1", "A1"),
    B1: doc.getCell("Sheet1", "B1"),
    C1: doc.getCell("Sheet1", "C1"),
    D1: doc.getCell("Sheet1", "D1"),
    E1: doc.getCell("Sheet1", "E1"),
  };

  assert.ok(doc.undo());
  assert.equal(doc.getCell("Sheet1", "A1").value, before.A1.value);
  assert.equal(doc.getCell("Sheet1", "A1").styleId, before.A1.styleId);
  assert.equal(doc.getCell("Sheet1", "B1").value, before.B1.value);
  assert.equal(doc.getCell("Sheet1", "B1").styleId, before.B1.styleId);
  assert.equal(doc.getCell("Sheet1", "C1").value, before.C1.value);
  assert.equal(doc.getCell("Sheet1", "C1").styleId, before.C1.styleId);

  assert.ok(doc.redo());
  assert.equal(doc.getCell("Sheet1", "A1").value, afterInsert.A1.value);
  assert.equal(doc.getCell("Sheet1", "B1").value, afterInsert.B1.value);
  assert.equal(doc.getCell("Sheet1", "C1").value, afterInsert.C1.value);
  assert.equal(doc.getCell("Sheet1", "D1").value, afterInsert.D1.value);
  assert.equal(doc.getCell("Sheet1", "E1").value, afterInsert.E1.value);

  // Now verify delete shift undo/redo on top of the inserted state.
  doc.deleteCellsShiftLeft("Sheet1", "A1:B1");
  const afterDelete = {
    A1: doc.getCell("Sheet1", "A1"),
    B1: doc.getCell("Sheet1", "B1"),
    C1: doc.getCell("Sheet1", "C1"),
  };

  assert.ok(doc.undo());
  assert.equal(doc.getCell("Sheet1", "A1").value, afterInsert.A1.value);
  assert.equal(doc.getCell("Sheet1", "B1").value, afterInsert.B1.value);
  assert.equal(doc.getCell("Sheet1", "C1").value, afterInsert.C1.value);

  assert.ok(doc.redo());
  assert.equal(doc.getCell("Sheet1", "A1").value, afterDelete.A1.value);
  assert.equal(doc.getCell("Sheet1", "B1").value, afterDelete.B1.value);
  assert.equal(doc.getCell("Sheet1", "C1").value, afterDelete.C1.value);
});
