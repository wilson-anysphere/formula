import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { copyRangeToClipboardPayload, pasteClipboardContent } from "../clipboard.js";

test("e2e: copy range -> paste into another area", () => {
  const doc = new DocumentController();

  doc.setRangeValues("Sheet1", "A1", [
    [1, "A"],
    [{ formula: "=A1*2" }, true],
  ]);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1:B2");
  const pasted = pasteClipboardContent(doc, "Sheet1", "C3", payload, { mode: "all" });
  assert.equal(pasted, true);

  assert.equal(doc.getCell("Sheet1", "C3").value, 1);
  assert.equal(doc.getCell("Sheet1", "D3").value, "A");
  assert.equal(doc.getCell("Sheet1", "C4").formula, "=A1*2");
  assert.equal(doc.getCell("Sheet1", "D4").value, true);
});
