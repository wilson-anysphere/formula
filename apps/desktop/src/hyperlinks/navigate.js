import { formatA1 } from "../document/coords.js";

/**
 * @typedef {{
 *   activateSheet: (sheetName: string) => void | Promise<void>,
 *   selectCell: (a1: string) => void | Promise<void>,
 * }} WorkbookNavigator
 */

/**
 * Navigate to an internal workbook link (activate sheet + select cell).
 *
 * @param {{ sheet: string, cell: { row: number, col: number } }} target
 * @param {WorkbookNavigator} navigator
 */
export async function navigateInternalHyperlink(target, navigator) {
  if (!navigator) throw new Error("navigateInternalHyperlink requires navigator");
  if (typeof navigator.activateSheet !== "function") {
    throw new Error("navigator.activateSheet must be a function");
  }
  if (typeof navigator.selectCell !== "function") {
    throw new Error("navigator.selectCell must be a function");
  }

  await navigator.activateSheet(target.sheet);
  await navigator.selectCell(formatA1(target.cell));
}

