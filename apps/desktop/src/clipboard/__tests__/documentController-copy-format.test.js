import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { copyRangeToClipboardPayload } from "../clipboard.js";
import { dateToExcelSerial } from "../../shared/valueParsing.js";

test("copyRangeToClipboardPayload maps DocumentController styleId -> clipboard format", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "Hello");
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});

test("copyRangeToClipboardPayload preserves numberFormat for date serials", () => {
  const doc = new DocumentController();

  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));
  doc.setRangeValues("Sheet1", "A1", [[{ value: serial, format: { numberFormat: "yyyy-mm-dd" } }]]);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(payload.text, "2024-01-31");
  assert.ok(payload.html);
  assert.match(payload.html, /data-number-format="yyyy-mm-dd"/);
});

test("copyRangeToClipboardPayload serializes rich text values as plain text", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", { text: "Rich Bold", runs: [{ start: 0, end: 4, style: { bold: true } }] });
  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(payload.text, "Rich Bold");
  assert.ok(payload.html);
  assert.match(payload.html, />Rich Bold</);
});

