import { excelWildcardToRegExp } from "./wildcards.js";
import { iterateMatches, findNext } from "./search.js";

function getSheetByName(workbook, sheetName) {
  if (typeof workbook.getSheet === "function") return workbook.getSheet(sheetName);
  const sheets = workbook.sheets ?? [];
  const found = sheets.find((s) => s.name === sheetName);
  if (!found) throw new Error(`Unknown sheet: ${sheetName}`);
  return found;
}

function looksLikeNumber(text) {
  const s = String(text).trim();
  if (s === "") return false;
  return /^-?(?:\d+\.?\d*|\d*\.?\d+)(?:[eE]-?\d+)?$/.test(s);
}

function coerceReplacementValue(text) {
  if (looksLikeNumber(text)) return Number(text);
  return String(text);
}

function replaceInString(text, re, replacement) {
  let count = 0;
  const out = String(text).replace(re, () => {
    count++;
    return replacement;
  });
  return { out, count };
}

function getValueText(cell, valueMode) {
  if (!cell) return "";
  if (valueMode === "raw") {
    if (cell.value == null) return "";
    return String(cell.value);
  }
  if (cell.display != null) return String(cell.display);
  if (cell.value == null) return "";
  return String(cell.value);
}

function applyReplaceToCell(cell, query, replacement, options = {}, { replaceAll = false } = {}) {
  if (!cell) return { cell: null, replaced: false, replacements: 0 };
  const {
    lookIn = "values",
    valueMode = "display",
    matchCase = false,
    matchEntireCell = false,
    useWildcards = true,
  } = options;

  const re = excelWildcardToRegExp(query, {
    matchCase,
    matchEntireCell,
    useWildcards,
    global: replaceAll,
  });

  // Excel semantics:
  // - Look in: Formulas => replace in the cell "input" (formula string or raw constant)
  // - Look in: Values   => replace in displayed/evaluated value, overwriting formulas with constants
  if (lookIn === "formulas") {
    if (cell.formula != null && cell.formula !== "") {
      const { out, count } = replaceInString(cell.formula, re, replacement);
      if (count === 0) return { cell, replaced: false, replacements: 0 };
      return { cell: { ...cell, formula: out }, replaced: true, replacements: count };
    }

    const original = cell.value == null ? "" : String(cell.value);
    const { out, count } = replaceInString(original, re, replacement);
    if (count === 0) return { cell, replaced: false, replacements: 0 };
    return { cell: { ...cell, value: coerceReplacementValue(out) }, replaced: true, replacements: count };
  }

  // values
  const original = getValueText(cell, valueMode);
  const { out, count } = replaceInString(original, re, replacement);
  if (count === 0) return { cell, replaced: false, replacements: 0 };

  return {
    cell: { ...cell, formula: undefined, value: coerceReplacementValue(out), display: undefined },
    replaced: true,
    replacements: count,
  };
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
