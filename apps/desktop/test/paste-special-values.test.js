import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../src/document/documentController.js";
import { pasteClipboardContent } from "../src/clipboard/clipboard.js";

test("pasteClipboardContent: Paste Special Values ignores formulas", () => {
  const doc = new DocumentController();

  // Simulate an Excel-style HTML clipboard payload where the cell text contains the
  // computed value, while `data-formula` carries the underlying formula.
  const content = {
    html: `<!DOCTYPE html><html><body><table>
      <tr>
        <td data-formula="=A1*2">2</td>
        <td>Plain</td>
      </tr>
      <tr>
        <td>3</td>
        <td data-formula="=A1+1">4</td>
      </tr>
    </table></body></html>`,
    text: "2\tPlain\n3\t4",
  };

  const pasted = pasteClipboardContent(doc, "Sheet1", "B2", content, { mode: "values" });
  assert.equal(pasted, true);

  assert.equal(doc.getCell("Sheet1", "B2").value, 2);
  assert.equal(doc.getCell("Sheet1", "C2").value, "Plain");
  assert.equal(doc.getCell("Sheet1", "B3").value, 3);
  assert.equal(doc.getCell("Sheet1", "C3").value, 4);

  assert.equal(doc.getCell("Sheet1", "B2").formula, null);
  assert.equal(doc.getCell("Sheet1", "C2").formula, null);
  assert.equal(doc.getCell("Sheet1", "B3").formula, null);
  assert.equal(doc.getCell("Sheet1", "C3").formula, null);
});

