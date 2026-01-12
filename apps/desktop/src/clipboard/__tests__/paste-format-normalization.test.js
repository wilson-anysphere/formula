import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { pasteClipboardContent } from "../clipboard.js";

test("pasteClipboardContent normalizes flat clipboard formats into canonical DocumentController styles", () => {
  const doc = new DocumentController();

  const html = `<!DOCTYPE html><html><body><table>
    <tr>
      <td style="font-weight:bold;color:#ff0000;background-color:#00ff00">X</td>
    </tr>
  </table></body></html>`;

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", { html });
  assert.equal(pasted, true);

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.value, "X");
  assert.notEqual(cell.styleId, 0);

  const style = doc.styleTable.get(cell.styleId);
  assert.equal(style.font?.bold, true);
  assert.equal(style.font?.color, "#FFFF0000");
  assert.equal(style.fill?.pattern, "solid");
  assert.equal(style.fill?.fgColor, "#FF00FF00");

  // Ensure we didn't intern the clipboard's flat schema directly.
  assert.equal(style.bold, undefined);
  assert.equal(style.italic, undefined);
  assert.equal(style.underline, undefined);
  assert.equal(style.textColor, undefined);
  assert.equal(style.backgroundColor, undefined);
});

