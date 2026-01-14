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

test("pasteClipboardContent converts rgb()/rgba() + #AARRGGBB CSS colors to ARGB", () => {
  const doc = new DocumentController();

  const html = `<!DOCTYPE html><html><body><table>
    <tr>
      <td style="color:rgb(255,0,0);background-color:rgb(0 255 0)">A</td>
      <td style="color:rgba(0,0,255,0.5);background-color:rgba(255,0,0,0.25)">B</td>
      <td style="color:#80FF0000;background-color:#4000FF00">C</td>
      <td style="color:#f00;background-color:#0f08">D</td>
      <td style="color:rgb(100% 0% 0% / 50%);background-color:rgba(0 0 255 / 25%)">E</td>
    </tr>
  </table></body></html>`;

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", { html });
  assert.equal(pasted, true);

  const a1 = doc.getCell("Sheet1", "A1");
  const b1 = doc.getCell("Sheet1", "B1");
  const c1 = doc.getCell("Sheet1", "C1");
  const d1 = doc.getCell("Sheet1", "D1");
  const e1 = doc.getCell("Sheet1", "E1");

  const styleA1 = doc.styleTable.get(a1.styleId);
  assert.equal(styleA1.font?.color, "#FFFF0000");
  assert.equal(styleA1.fill?.fgColor, "#FF00FF00");

  const styleB1 = doc.styleTable.get(b1.styleId);
  assert.equal(styleB1.font?.color, "#800000FF");
  assert.equal(styleB1.fill?.fgColor, "#40FF0000");

  const styleC1 = doc.styleTable.get(c1.styleId);
  assert.equal(styleC1.font?.color, "#80FF0000");
  assert.equal(styleC1.fill?.fgColor, "#4000FF00");

  const styleD1 = doc.styleTable.get(d1.styleId);
  assert.equal(styleD1.font?.color, "#FFFF0000");
  assert.equal(styleD1.fill?.fgColor, "#8800FF00");

  const styleE1 = doc.styleTable.get(e1.styleId);
  assert.equal(styleE1.font?.color, "#80FF0000");
  assert.equal(styleE1.fill?.fgColor, "#400000FF");
});

test("pasteClipboardContent converts common CSS named colors to ARGB (non-DOM fallback)", () => {
  const doc = new DocumentController();

  const html = `<!DOCTYPE html><html><body><table>
    <tr>
      <td style="color:red;background-color:yellow">X</td>
    </tr>
  </table></body></html>`;

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", { html });
  assert.equal(pasted, true);

  const cell = doc.getCell("Sheet1", "A1");
  const style = doc.styleTable.get(cell.styleId);
  assert.equal(style.font?.color, "#FFFF0000");
  assert.equal(style.fill?.fgColor, "#FFFFFF00");
});

test("pasteClipboardContent treats mso-number-format:General as clearing (explicit numberFormat: null override)", () => {
  const doc = new DocumentController();

  const html = `<!DOCTYPE html><html><body><table>
    <tr>
      <td style="mso-number-format:'General'">1</td>
    </tr>
  </table></body></html>`;

  const pasted = pasteClipboardContent(doc, "Sheet1", "A1", { html });
  assert.equal(pasted, true);

  const cell = doc.getCell("Sheet1", "A1");
  assert.equal(cell.value, 1);

  // "General" should map to an explicit `numberFormat: null` override so pasted cells
  // clear any inherited number formats (row/col/sheet defaults).
  assert.notEqual(cell.styleId, 0);
  const style = doc.styleTable.get(cell.styleId);
  assert.equal(style.numberFormat, null);
});
