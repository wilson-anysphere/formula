import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../src/document/documentController.js";
import { pasteClipboardContent } from "../src/clipboard/clipboard.js";

test("pasteClipboardContent: Paste Special Values preserves leading '=' strings as text", () => {
  const doc = new DocumentController();

  const content = {
    html: `<!DOCTYPE html><html><body><table>
      <tr>
        <td>=literal</td>
      </tr>
    </table></body></html>`,
    text: "=literal",
  };

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", content, { mode: "values" });
  assert.equal(pasted, true);

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.formula, null);
  assert.equal(cell.value, "=literal");
});

test("pasteClipboardContent: Paste Special Values preserves leading-zero numbers as text", () => {
  const doc = new DocumentController();

  const content = {
    html: `<!DOCTYPE html><html><body><table>
      <tr>
        <td>00123</td>
      </tr>
    </table></body></html>`,
    text: "00123",
  };

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", content, { mode: "values" });
  assert.equal(pasted, true);

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.formula, null);
  // parseScalar treats leading-zero integers as strings, and Paste Values should preserve that.
  assert.equal(cell.value, "00123");
});

