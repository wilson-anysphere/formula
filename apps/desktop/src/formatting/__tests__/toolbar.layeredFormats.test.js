import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toggleBold reads full-column formatting from the column layer (no per-cell scan)", () => {
  const doc = new DocumentController();

  // Apply bold to the entire column A via the layered column format path.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "Sheet1 should exist after formatting");
  assert.equal(sheet.cells.size, 0, "Full-column formatting should not materialize per-cell overrides");

  // Toggling again should flip bold OFF (because all cells are currently bold).
  toggleBold(doc, "Sheet1", "A1:A1048576");
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), false);
  assert.equal(sheet.cells.size, 0, "Toggling full-column formatting should not materialize per-cell overrides");
});

test("toggleBold on a full-column selection treats a single conflicting cell override as mixed state", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  // Introduce a single conflicting cell-level override inside the column.
  doc.setRangeFormat("Sheet1", "A2", { font: { bold: false } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet, "Sheet1 should exist after formatting");
  assert.equal(sheet.cells.size, 1, "A single cell override should not expand to the full column");

  // Selection is now mixed, so toggleBold should choose next=true and make everything bold.
  toggleBold(doc, "Sheet1", "A1:A1048576");

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1").font?.bold), true);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A2").font?.bold), true);
  assert.ok(sheet.cells.size <= 1, "Toggling should not materialize per-cell overrides across the selection");
});

