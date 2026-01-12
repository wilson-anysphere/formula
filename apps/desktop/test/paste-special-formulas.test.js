import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../src/document/documentController.js";
import { pasteClipboardContent } from "../src/clipboard/clipboard.js";

test("pasteClipboardContent: Paste Special Formulas pastes formulas but not formats", () => {
  const doc = new DocumentController();

  // Seed destination formatting (should remain after pasting formulas-only).
  doc.setCellValue("Sheet1", "B2", "KeepStyle");
  doc.setRangeFormat("Sheet1", "B2", { font: { italic: true } }, { label: "Italic" });
  const before = doc.getCell("Sheet1", "B2");
  const beforeStyle = doc.styleTable.get(before.styleId);
  assert.equal(beforeStyle?.font?.italic, true);

  const content = {
    html: `<!DOCTYPE html><html><body><table>
      <tr>
        <td data-formula="=A1+1" style="font-weight:bold">2</td>
      </tr>
    </table></body></html>`,
    text: "2",
  };

  const pasted = pasteClipboardContent(doc, "Sheet1", "B2", content, { mode: "formulas" });
  assert.equal(pasted, true);

  const after = doc.getCell("Sheet1", "B2");
  assert.equal(after.formula, "=A1+1");
  assert.equal(after.value, null);
  // Formats should not be pasted; the destination formatting should remain.
  assert.equal(after.styleId, before.styleId);
  const afterStyle = doc.styleTable.get(after.styleId);
  assert.equal(afterStyle?.font?.italic, true);
  assert.equal(afterStyle?.font?.bold, undefined);
});

