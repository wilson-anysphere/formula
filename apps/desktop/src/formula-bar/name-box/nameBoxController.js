import { formatA1Address, parseGoTo } from "../../../../../packages/search/index.js";

export class NameBoxController {
  constructor({ workbook, getCurrentSheetName, getActiveCell, setActiveCell, selectRange }) {
    if (!workbook) throw new Error("NameBoxController: workbook is required");
    this.workbook = workbook;
    this.getCurrentSheetName = getCurrentSheetName;
    this.getActiveCell = getActiveCell;
    this.setActiveCell = setActiveCell;
    this.selectRange = selectRange;
  }

  getDisplayValue() {
    const active = this.getActiveCell?.();
    if (!active) return "";
    return formatA1Address({ row: active.row, col: active.col });
  }

  submit(text) {
    const currentSheetName = this.getCurrentSheetName?.();
    const parsed = parseGoTo(text, { workbook: this.workbook, currentSheetName });

    if (parsed.type === "range") {
      const { range } = parsed;
      if (range.startRow === range.endRow && range.startCol === range.endCol) {
        this.setActiveCell?.({ sheetName: parsed.sheetName, row: range.startRow, col: range.startCol });
      } else {
        this.selectRange?.({ sheetName: parsed.sheetName, range });
      }
    }
  }
}
