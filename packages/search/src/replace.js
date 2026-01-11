import { iterateMatches, findNext } from "./search.js";
import { applyReplaceToCell } from "./replaceCore.js";

function getSheetByName(workbook, sheetName) {
  if (typeof workbook.getSheet === "function") return workbook.getSheet(sheetName);
  const sheets = workbook.sheets ?? [];
  const found = sheets.find((s) => s.name === sheetName);
  if (!found) throw new Error(`Unknown sheet: ${sheetName}`);
  return found;
}

export async function replaceAll(workbook, query, replacement, options = {}) {
  let replacedCells = 0;
  let replacedOccurrences = 0;

  for await (const match of iterateMatches(workbook, query, options)) {
    const sheet = getSheetByName(workbook, match.sheetName);
    const cell = sheet.getCell(match.row, match.col);
    const res = applyReplaceToCell(cell, query, replacement, options, { replaceAll: true });
    if (res.replaced) {
      sheet.setCell(match.row, match.col, res.cell);
      replacedCells++;
      replacedOccurrences += res.replacements;
    }
  }

  return { replacedCells, replacedOccurrences };
}

export async function replaceNext(workbook, query, replacement, options = {}, from) {
  const next = await findNext(workbook, query, options, from);
  if (!next) return null;

  const sheet = getSheetByName(workbook, next.sheetName);
  const cell = sheet.getCell(next.row, next.col);
  const res = applyReplaceToCell(cell, query, replacement, options, { replaceAll: false });
  if (res.replaced) {
    sheet.setCell(next.row, next.col, res.cell);
  }

  return { match: next, replaced: res.replaced, replacements: res.replacements };
}

export { applyReplaceToCell };
