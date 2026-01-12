import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../src/document/documentController.js";
import { pasteClipboardContent } from "../src/clipboard/clipboard.js";

test("pasteClipboardContent: Paste Special Formats applies formats only", () => {
  const doc = new DocumentController();

  // Destination cell has a value that should remain unchanged.
  doc.setCellValue("Sheet1", "B2", "Keep");

  const before = doc.getCell("Sheet1", "B2");
  assert.equal(before.value, "Keep");
  assert.equal(before.styleId, 0);

  const content = {
    html: `<!DOCTYPE html><html><body><table>
      <tr>
        <td style="font-weight:bold">X</td>
      </tr>
    </table></body></html>`,
    text: "X",
  };

  const pasted = pasteClipboardContent(doc, "Sheet1", "B2", content, { mode: "formats" });
  assert.equal(pasted, true);

  const after = doc.getCell("Sheet1", "B2");
  assert.equal(after.value, "Keep");
  assert.equal(after.formula, null);
  assert.ok(after.styleId > 0);

  const style = doc.styleTable.get(after.styleId);
  assert.equal(style?.font?.bold, true);
});

