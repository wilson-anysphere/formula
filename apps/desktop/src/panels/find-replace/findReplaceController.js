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
    canReplace,
    showToast,
  }) {
    if (!workbook) throw new Error("FindReplaceController: workbook is required");
    this.workbook = workbook;

    this.getCurrentSheetName = getCurrentSheetName;
    this.getActiveCell = getActiveCell;
    this.setActiveCell = setActiveCell;
    this.getSelectionRanges = getSelectionRanges;
    this.beginBatch = beginBatch;
    this.endBatch = endBatch;
    this.canReplace = typeof canReplace === "function" ? canReplace : null;
    this.showToast = typeof showToast === "function" ? showToast : null;

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

  /**
   * @returns {{ allowed: boolean, reason?: string | null }}
   */
  getReplacePermission() {
    const guard = this.canReplace;
    if (typeof guard !== "function") return { allowed: true };
    try {
      const result = guard();
      if (typeof result === "boolean") return { allowed: result, reason: null };
      if (result && typeof result === "object" && typeof result.allowed === "boolean") {
        return { allowed: result.allowed, reason: typeof result.reason === "string" ? result.reason : null };
      }
    } catch (err) {
      return { allowed: false, reason: String(err?.message ?? err) };
    }
    return { allowed: true };
  }

  /**
   * @param {string | null | undefined} reason
   */
  toastReplaceBlocked(reason) {
    const showToast = this.showToast;
    if (typeof showToast !== "function") return;
    const message = typeof reason === "string" && reason.trim() ? reason.trim() : "Replacing cells is not allowed.";
    try {
      showToast(message, "warning");
    } catch {
      // ignore
    }
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
    const permission = this.getReplacePermission();
    if (!permission.allowed) {
      this.toastReplaceBlocked(permission.reason);
      return null;
    }
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
    const permission = this.getReplacePermission();
    if (!permission.allowed) {
      this.toastReplaceBlocked(permission.reason);
      return null;
    }
    this.beginBatch?.({ label: "Replace All" });
    try {
      return await replaceAll(this.workbook, this.query, this.replacement, this.getSearchOptions());
    } finally {
      this.endBatch?.();
    }
  }
}
