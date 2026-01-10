import { findAll, findNext, replaceAll, replaceNext } from "../../../../../packages/search/index.js";

/**
 * Minimal controller wiring the core search/replace engine to UI/navigation.
 *
 * This is intentionally framework-agnostic (no React dependency) so it can be
 * integrated into whatever desktop shell we end up with.
 */
export class FindReplaceController {
  constructor({
    workbook,
    getCurrentSheetName,
    getActiveCell,
    setActiveCell,
    getSelectionRanges,
    beginBatch,
    endBatch,
  }) {
    if (!workbook) throw new Error("FindReplaceController: workbook is required");
    this.workbook = workbook;

    this.getCurrentSheetName = getCurrentSheetName;
    this.getActiveCell = getActiveCell;
    this.setActiveCell = setActiveCell;
    this.getSelectionRanges = getSelectionRanges;
    this.beginBatch = beginBatch;
    this.endBatch = endBatch;

    this.query = "";
    this.replacement = "";

    this.scope = "sheet"; // "selection" | "sheet" | "workbook"
    this.lookIn = "values"; // "values" | "formulas"
    this.valueMode = "display"; // "display" | "raw"
    this.matchCase = false;
    this.matchEntireCell = false;
    this.useWildcards = true;
    this.searchOrder = "byRows"; // "byRows" | "byColumns"

    this.lastResults = [];
  }

  getSearchOptions(overrides = {}) {
    const currentSheetName = this.getCurrentSheetName?.();
    const selectionRanges = this.getSelectionRanges?.();

    return {
      scope: this.scope,
      currentSheetName,
      selectionRanges,
      lookIn: this.lookIn,
      valueMode: this.valueMode,
      matchCase: this.matchCase,
      matchEntireCell: this.matchEntireCell,
      useWildcards: this.useWildcards,
      searchOrder: this.searchOrder,
      ...overrides,
    };
  }

  async findNext() {
    const from = this.getActiveCell?.();
    const match = await findNext(this.workbook, this.query, this.getSearchOptions(), from);
    if (match && this.setActiveCell) {
      this.setActiveCell({ sheetName: match.sheetName, row: match.row, col: match.col });
    }
    return match;
  }

  async findAll() {
    const results = await findAll(this.workbook, this.query, this.getSearchOptions());
    this.lastResults = results;
    return results;
  }

  async replaceNext() {
    const from = this.getActiveCell?.();
    const res = await replaceNext(
      this.workbook,
      this.query,
      this.replacement,
      this.getSearchOptions(),
      from,
    );
    if (res?.match && this.setActiveCell) {
      this.setActiveCell({
        sheetName: res.match.sheetName,
        row: res.match.row,
        col: res.match.col,
      });
    }
    return res;
  }

  async replaceAll() {
    this.beginBatch?.({ label: "Replace All" });
    try {
      return await replaceAll(this.workbook, this.query, this.replacement, this.getSearchOptions());
    } finally {
      this.endBatch?.();
    }
  }
}
