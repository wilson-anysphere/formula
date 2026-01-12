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
    "font-weight:bold;background-color:rgb(255 255 0);text-align:center;white-space:normal;border-left:thin rgb(0 0 0);border-right:thin rgb(0 0 0);border-top:thin rgb(0 0 0);border-bottom:thin rgb(0 0 0);number-format:$#,##0.00",
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
