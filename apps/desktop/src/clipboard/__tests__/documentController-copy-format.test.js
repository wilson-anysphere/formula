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

test("copyRangeToClipboardPayload uses effective formats (e.g. column defaults) when styleId is 0", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } }, { label: "Bold Column" });

  // The whole-column format should live in a layered style map, not on the cell itself.
  assert.equal(doc.getCell("Sheet1", "A500").styleId, 0);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A500:A500");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});

test("copyRangeToClipboardPayload preserves column-level numberFormat for date serials", () => {
  const doc = new DocumentController();

  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));
  doc.setRangeFormat("Sheet1", "A1:A1048576", { numberFormat: "yyyy-mm-dd" }, { label: "Date Column" });
  doc.setCellValue("Sheet1", "A500", serial);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A500");
  assert.equal(payload.text, "2024-01-31");
  assert.ok(payload.html);
  assert.match(payload.html, /data-number-format="yyyy-mm-dd"/);
});
