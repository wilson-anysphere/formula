import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { getEffectiveCellStyle } from "../getEffectiveCellStyle.js";

test("Format Cells dialog helpers read effective (layered) formatting for an inherited bold column", () => {
  const doc = new DocumentController();

  // Full-height column A. This should be stored as a column formatting layer, not a cell styleId.
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.styleId, 0);

  const effective = getEffectiveCellStyle(doc, "Sheet1", "A1");
  assert.equal(effective.font?.bold, true);

  // Clearing inherited bold requires an explicit override (bold: false) at the cell/range.
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: false } });
  const cleared = getEffectiveCellStyle(doc, "Sheet1", "A1");
  assert.equal(cleared.font?.bold, false);
});

