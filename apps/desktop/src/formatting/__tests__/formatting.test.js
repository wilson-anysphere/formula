import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { renderCellStyle } from "../renderCellStyle.js";
import {
  applyAllBorders,
  applyNumberFormatPreset,
  setFillColor,
  setHorizontalAlign,
  toggleBold,
  toggleStrikethrough,
  toggleWrap,
} from "../toolbar.js";
import { installFormattingShortcuts } from "../shortcuts.js";

class FakeEventTarget {
  constructor() {
    /** @type {Map<string, Set<(event: any) => void>>} */
    this.listeners = new Map();
  }

  addEventListener(type, listener) {
    let set = this.listeners.get(type);
    if (!set) {
      set = new Set();
      this.listeners.set(type, set);
    }
    set.add(listener);
  }

  removeEventListener(type, listener) {
    this.listeners.get(type)?.delete(listener);
  }

  dispatchEvent(type, event) {
    for (const listener of this.listeners.get(type) ?? []) listener(event);
  }
}

test("formatting uses styleIds with deduplication", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [
    ["x", "y"],
    ["z", "w"],
  ]);

  doc.setRangeFormat("Sheet1", "A1:B2", { font: { bold: true } });

  const a1 = doc.getCell("Sheet1", "A1");
  const b2 = doc.getCell("Sheet1", "B2");
  assert.equal(a1.styleId, b2.styleId);
  assert.equal(doc.styleTable.size, 2); // default + bold
});

test("toolbar commands produce expected style rendering (snapshot)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "Styled");

  toggleBold(doc, "Sheet1", "A1");
  setFillColor(doc, "Sheet1", "A1", "#FFFFFF00");
  applyAllBorders(doc, "Sheet1", "A1", { style: "thin", color: "#FF000000" });
  setHorizontalAlign(doc, "Sheet1", "A1", "center");
  toggleWrap(doc, "Sheet1", "A1");
  applyNumberFormatPreset(doc, "Sheet1", "A1", "currency");

  const cell = doc.getCell("Sheet1", "A1");
  const style = doc.styleTable.get(cell.styleId);
  assert.equal(
    renderCellStyle(style),
    "font-weight:bold;background-color:#ffff00;text-align:center;white-space:normal;border-left:thin #000000;border-right:thin #000000;border-top:thin #000000;border-bottom:thin #000000;number-format:$#,##0.00",
  );
});

test("toggle commands accept multiple ranges and apply consistently", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [["x", "y"]]);

  // Seed mixed formatting: A1 bold, B1 not bold.
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

  toggleBold(doc, "Sheet1", ["A1", "B1"]);
  const a1 = doc.styleTable.get(doc.getCell("Sheet1", "A1").styleId);
  const b1 = doc.styleTable.get(doc.getCell("Sheet1", "B1").styleId);
  assert.equal(Boolean(a1.font?.bold), true);
  assert.equal(Boolean(b1.font?.bold), true);

  toggleBold(doc, "Sheet1", ["A1", "B1"]);
  const a1After = doc.styleTable.get(doc.getCell("Sheet1", "A1").styleId);
  const b1After = doc.styleTable.get(doc.getCell("Sheet1", "B1").styleId);
  assert.equal(Boolean(a1After.font?.bold), false);
  assert.equal(Boolean(b1After.font?.bold), false);
});

test("toggle commands can be forced to an explicit state", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [["x", "y"]]);

  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  toggleBold(doc, "Sheet1", ["A1", "B1"], { next: false });

  const a1 = doc.styleTable.get(doc.getCell("Sheet1", "A1").styleId);
  const b1 = doc.styleTable.get(doc.getCell("Sheet1", "B1").styleId);
  assert.equal(Boolean(a1.font?.bold), false);
  assert.equal(Boolean(b1.font?.bold), false);
});

test("Ctrl/Cmd+1 triggers Format Cells shortcut", () => {
  const target = new FakeEventTarget();
  let count = 0;
  installFormattingShortcuts(target, {
    openFormatCells() {
      count += 1;
    },
  });

  let prevented = false;
  target.dispatchEvent("keydown", {
    key: "1",
    ctrlKey: true,
    preventDefault() {
      prevented = true;
    },
  });

  assert.equal(count, 1);
  assert.equal(prevented, true);
});

test("Ctrl/Cmd+Digit1 triggers Format Cells even when event.key differs under keyboard layout", () => {
  const target = new FakeEventTarget();
  let count = 0;
  installFormattingShortcuts(target, {
    openFormatCells() {
      count += 1;
    },
  });

  let prevented = false;
  // Example: on AZERTY layouts, `Digit1` may report `event.key === "&"` without Shift.
  target.dispatchEvent("keydown", {
    key: "&",
    code: "Digit1",
    ctrlKey: true,
    preventDefault() {
      prevented = true;
    },
  });

  assert.equal(count, 1);
  assert.equal(prevented, true);
});

test("toggleBold on a full column selection is fast + uses effective (layered) formats", () => {
  const doc = new DocumentController();

  // Ensure the sheet exists in the model.
  doc.getCell("Sheet1", "A1");

  // Simulate layered formats: column A is bold by default, while individual cells in the
  // column keep `styleId=0` (no cell override).
  const sheet = doc.model.sheets.get("Sheet1");
  const boldStyleId = doc.styleTable.intern({ font: { bold: true } });
  sheet.colStyleIds.set(0, boldStyleId);

  // The regression here is a full-column range in Excel row space. This should not
  // enumerate ~1 million rows just to determine the toggle state.
  const fullColumnA = "A1:A1048576";

  // Guardrail: the previous implementation called `getCell()` once per row; cap it.
  const originalGetCell = doc.getCell.bind(doc);
  let getCellCalls = 0;
  doc.getCell = (...args) => {
    getCellCalls += 1;
    if (getCellCalls > 10_000) {
      throw new Error(`toggleBold performed O(rows) getCell calls (${getCellCalls})`);
    }
    return originalGetCell(...args);
  };

  // Stub the write so the test only exercises the read-path (toggle state), not the
  // implementation details of applying formats to huge ranges.
  let lastPatch = null;
  doc.setRangeFormat = (_sheetId, _range, patch) => {
    lastPatch = patch;
  };

  toggleBold(doc, "Sheet1", fullColumnA);

  // Column A is effectively bold, so toggleBold should toggle *off* (set bold=false).
  assert.deepEqual(lastPatch, { font: { bold: false } });
});

test("toggleStrikethrough produces expected style rendering", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "Styled");

  toggleStrikethrough(doc, "Sheet1", "A1");

  const cell = doc.getCell("Sheet1", "A1");
  const style = doc.styleTable.get(cell.styleId);
  assert.equal(renderCellStyle(style), "text-decoration:line-through");
});

test("toggleStrikethrough on a full column selection is fast + uses effective (layered) formats", () => {
  const doc = new DocumentController();

  // Ensure the sheet exists in the model.
  doc.getCell("Sheet1", "A1");

  // Simulate layered formats: column A is strikethrough by default, while individual cells in the
  // column keep `styleId=0` (no cell override).
  const sheet = doc.model.sheets.get("Sheet1");
  const strikeStyleId = doc.styleTable.intern({ font: { strike: true } });
  sheet.colStyleIds.set(0, strikeStyleId);

  const fullColumnA = "A1:A1048576";

  // Guardrail: the previous naive implementation would have called `getCell()` once per row.
  const originalGetCell = doc.getCell.bind(doc);
  let getCellCalls = 0;
  doc.getCell = (...args) => {
    getCellCalls += 1;
    if (getCellCalls > 10_000) {
      throw new Error(`toggleStrikethrough performed O(rows) getCell calls (${getCellCalls})`);
    }
    return originalGetCell(...args);
  };

  // Stub the write so the test only exercises the read-path (toggle state), not the
  // implementation details of applying formats to huge ranges.
  let lastPatch = null;
  doc.setRangeFormat = (_sheetId, _range, patch) => {
    lastPatch = patch;
  };

  toggleStrikethrough(doc, "Sheet1", fullColumnA);

  // Column A is effectively strikethrough, so toggleStrikethrough should toggle *off* (set strike=false).
  assert.deepEqual(lastPatch, { font: { strike: false } });
});
