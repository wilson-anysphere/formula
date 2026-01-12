import assert from "node:assert/strict";
import test from "node:test";

import { StyleTable } from "../apps/desktop/src/formatting/styleTable.js";
import { MacroRecorder } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";

class FakeDocumentController {
  constructor() {
    this.styleTable = new StyleTable();
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
    /** @type {((sheetId: string, coord: { row: number, col: number }) => any) | null} */
    this.getCellFormatImpl = null;
  }

  /**
   * @param {string} event
   * @param {(payload: any) => void} listener
   */
  on(event, listener) {
    let bucket = this.listeners.get(event);
    if (!bucket) {
      bucket = new Set();
      this.listeners.set(event, bucket);
    }
    bucket.add(listener);
    return () => bucket.delete(listener);
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  emit(event, payload) {
    const bucket = this.listeners.get(event);
    if (!bucket) return;
    for (const listener of bucket) listener(payload);
  }

  emitChange(payload) {
    this.emit("change", payload);
  }

  getCell() {
    return { value: null, formula: null, styleId: 0 };
  }

  getCellFormat(sheetId, coord) {
    if (!this.getCellFormatImpl) return null;
    return this.getCellFormatImpl(sheetId, coord);
  }
}

test("Range.getFormat reads effective format via DocumentController.getCellFormat (layered formatting)", () => {
  const doc = new FakeDocumentController();
  doc.getCellFormatImpl = () => ({
    font: { bold: true },
    fill: { fgColor: "#FFFFFF00" },
    numberFormat: "0%",
  });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const sheet = workbook.getActiveSheet();
  assert.deepEqual(sheet.getRange("A1").getFormat(), {
    bold: true,
    backgroundColor: "#FFFFFF00",
    numberFormat: "0%",
  });

  workbook.dispose();
});

test("Macro recorder receives formatChanged for column style deltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const boldId = doc.styleTable.intern({ font: { bold: true } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    colStyleIdDeltas: [{ sheetId: "Sheet1", col: 2, beforeStyleId: 0, afterStyleId: boldId }],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "C1:C1048576", format: { bold: true } },
  ]);

  workbook.dispose();
});

test("Macro recorder receives formatChanged for colFormats changes in sheetViewDeltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const boldId = doc.styleTable.intern({ font: { bold: true } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    sheetViewDeltas: [
      {
        sheetId: "Sheet1",
        before: { frozenRows: 0, frozenCols: 0, colFormats: {} },
        after: { frozenRows: 0, frozenCols: 0, colFormats: { "2": boldId } },
      },
    ],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "C1:C1048576", format: { bold: true } },
  ]);

  workbook.dispose();
});

test("Macro recorder receives formatChanged for row style deltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const italicId = doc.styleTable.intern({ font: { italic: true } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    rowStyleIdDeltas: [{ sheetId: "Sheet1", row: 5, beforeStyleId: 0, afterStyleId: italicId }],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "A6:XFD6", format: { italic: true } },
  ]);

  workbook.dispose();
});

test("Macro recorder receives formatChanged for rowFormats changes in sheetViewDeltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const italicId = doc.styleTable.intern({ font: { italic: true } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    sheetViewDeltas: [
      {
        sheetId: "Sheet1",
        before: { frozenRows: 0, frozenCols: 0, rowFormats: {} },
        after: { frozenRows: 0, frozenCols: 0, rowFormats: { "5": italicId } },
      },
    ],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "A6:XFD6", format: { italic: true } },
  ]);

  workbook.dispose();
});

test("Macro recorder receives formatChanged for sheet default style deltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const fillId = doc.styleTable.intern({ fill: { fgColor: "#FF00FF00" } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    sheetStyleIdDeltas: [{ sheetId: "Sheet1", beforeStyleId: 0, afterStyleId: fillId }],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "A1:XFD1048576", format: { backgroundColor: "#FF00FF00" } },
  ]);

  workbook.dispose();
});

test("Macro recorder receives formatChanged for defaultFormat changes in sheetViewDeltas (layered formatting)", () => {
  const doc = new FakeDocumentController();
  const fillId = doc.styleTable.intern({ fill: { fgColor: "#FF00FF00" } });

  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });
  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.emitChange({
    deltas: [],
    sheetViewDeltas: [
      {
        sheetId: "Sheet1",
        before: { frozenRows: 0, frozenCols: 0, defaultFormat: 0 },
        after: { frozenRows: 0, frozenCols: 0, defaultFormat: fillId },
      },
    ],
  });

  assert.deepEqual(recorder.stop(), [
    { type: "setFormat", sheetName: "Sheet1", address: "A1:XFD1048576", format: { backgroundColor: "#FF00FF00" } },
  ]);

  workbook.dispose();
});
