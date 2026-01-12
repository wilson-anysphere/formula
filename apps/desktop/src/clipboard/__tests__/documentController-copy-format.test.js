import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { parseRangeA1 } from "../../document/coords.js";
import { applyStylePatch, StyleTable } from "../../formatting/styleTable.js";
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
  class LayeredDoc {
    constructor() {
      this.styleTable = new StyleTable();
      /** @type {Map<number, number>} */
      this.colStyleIds = new Map();
      /** @type {Map<string, number>} */
      this.cellStyleIds = new Map();
      this.sheetDefaultStyleId = 0;
    }

    /**
     * @param {string} _sheetId
     * @param {{ row: number, col: number }} coord
     */
    getCell(_sheetId, coord) {
      const key = `${coord.row},${coord.col}`;
      const styleId = this.cellStyleIds.get(key) ?? 0;
      return { value: null, formula: null, styleId };
    }

    /**
     * @param {string} _sheetId
     * @param {{ row: number, col: number }} coord
     * @returns {[number, number, number, number]}
     */
    getCellFormatStyleIds(_sheetId, coord) {
      const cellStyleId = this.cellStyleIds.get(`${coord.row},${coord.col}`) ?? 0;
      const colStyleId = this.colStyleIds.get(coord.col) ?? 0;
      return [this.sheetDefaultStyleId, 0, colStyleId, cellStyleId];
    }

    /**
     * @param {string} sheetId
     * @param {{ row: number, col: number }} coord
     */
    getCellFormat(sheetId, coord) {
      const [sheetDefaultStyleId, rowStyleId, colStyleId, cellStyleId] = this.getCellFormatStyleIds(sheetId, coord);
      let merged = {};
      merged = applyStylePatch(merged, this.styleTable.get(sheetDefaultStyleId));
      merged = applyStylePatch(merged, this.styleTable.get(rowStyleId));
      merged = applyStylePatch(merged, this.styleTable.get(colStyleId));
      merged = applyStylePatch(merged, this.styleTable.get(cellStyleId));
      return merged;
    }

    /**
     * @param {string} _sheetId
     * @param {string} range
     * @param {Record<string, any> | null} stylePatch
     */
    setRangeFormat(_sheetId, range, stylePatch) {
      const r = parseRangeA1(range);
      const isFullColumn = r.start.col === r.end.col && r.start.row === 0 && r.end.row === 1048575;
      if (!isFullColumn) {
        throw new Error(`LayeredDoc test helper only supports full-column ranges; got ${range}`);
      }
      const col = r.start.col;
      const beforeId = this.colStyleIds.get(col) ?? 0;
      const baseStyle = this.styleTable.get(beforeId);
      const merged = applyStylePatch(baseStyle, stylePatch);
      const nextId = this.styleTable.intern(merged);
      if (nextId === 0) this.colStyleIds.delete(col);
      else this.colStyleIds.set(col, nextId);
    }
  }

  const doc = new LayeredDoc();
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const payload = copyRangeToClipboardPayload(doc, "Sheet1", "A500:A500");
  assert.ok(payload.html);
  assert.match(payload.html, /font-weight:bold/i);
});
