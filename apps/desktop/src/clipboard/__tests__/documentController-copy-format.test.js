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

test("copyRangeToClipboardPayload does not fall back to snake_case number_format when numberFormat is cleared", () => {
  const doc = new DocumentController();

  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));
  doc.setCellValue("Sheet1", "A1", serial);
  // Simulate imported formula-model formatting (snake_case).
  doc.setRangeFormat("Sheet1", "A1", { number_format: "yyyy-mm-dd" });
  // User clears back to General via UI (camelCase null override).
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: null });

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  // Should copy as the raw numeric serial when the number format is cleared.
  assert.equal(payload.text, String(serial));
  assert.ok(payload.html);
  assert.doesNotMatch(payload.html, /data-number-format="yyyy-mm-dd"/);
});

test("copyRangeToClipboardPayload formats m/d/yyyy date serials (UI date preset)", () => {
  const doc = new DocumentController();

  const serial = dateToExcelSerial(new Date(Date.UTC(2024, 0, 31)));
  doc.setRangeValues("Sheet1", "A1", [[{ value: serial, format: { numberFormat: "m/d/yyyy" } }]]);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(payload.text, "2024-01-31");
  assert.ok(payload.html);
  assert.match(payload.html, /data-number-format="m\/d\/yyyy"/);
});

test("copyRangeToClipboardPayload formats hh:mm:ss time serials (Time preset)", () => {
  const doc = new DocumentController();

  const serial = (3 * 3600 + 4 * 60 + 5) / 86_400;
  doc.setRangeValues("Sheet1", "A1", [[{ value: serial, format: { numberFormat: "hh:mm:ss" } }]]);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(payload.text, "03:04:05");
  assert.ok(payload.html);
  assert.match(payload.html, /data-number-format="hh:mm:ss"/);
});

test("copyRangeToClipboardPayload serializes rich text values as plain text", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", { text: "Rich Bold", runs: [{ start: 0, end: 4, style: { bold: true } }] });
  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(payload.text, "Rich Bold");
  assert.ok(payload.html);
  assert.match(payload.html, />Rich Bold</);
});

test("copyRangeToClipboardPayload includes RTF output with basic formatting", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", "Hello");
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A1");
  assert.equal(typeof payload.rtf, "string");
  assert.ok(payload.rtf.startsWith("{\\rtf1"));
  assert.match(payload.rtf, /Hello/);
  assert.match(payload.rtf, /\\b(?=\\|\s)/);
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

test("copyRangeToClipboardPayload uses row-level formats from full-width row selections", () => {
  const doc = new DocumentController();

  // Full-width row 1: A1:XFD1.
  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { bold: true } }, { label: "Bold Row" });
  assert.equal(doc.getCell("Sheet1", "B1").styleId, 0);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "B1");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});

test("copyRangeToClipboardPayload uses sheet-level formats from full-sheet selections", () => {
  const doc = new DocumentController();

  // Full-sheet selection: A1:XFD1048576.
  doc.setRangeFormat("Sheet1", "A1:XFD1048576", { font: { bold: true } }, { label: "Bold Sheet" });
  assert.equal(doc.getCell("Sheet1", "C10").styleId, 0);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "C10");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});

test("copyRangeToClipboardPayload uses effective formats from range-run formatting (large rectangles)", () => {
  const doc = new DocumentController();
  // Large rectangle (> 50k cells) should store formatting as per-column row interval runs.
  doc.setRangeFormat("Sheet1", "A1:B30000", { font: { bold: true } }, { label: "Bold Runs" });

  // The formatted cell should not be materialized as a per-cell styleId.
  assert.equal(doc.getCell("Sheet1", "A20000").styleId, 0);

  const [sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId, rangeRunStyleId] = doc.getCellFormatStyleIds(
    "Sheet1",
    "A20000"
  );
  assert.equal(sheetDefaultStyleId, 0);
  assert.equal(rowStyleId, 0);
  assert.equal(colStyleId, 0);
  assert.equal(cellStyleId, 0);
  assert.notEqual(rangeRunStyleId, 0);
  assert.equal(Boolean(doc.styleTable.get(rangeRunStyleId)?.font?.bold), true);

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A20000");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});
