import { formatA1Address } from "./a1.js";
import { excelWildcardToRegExp } from "./wildcards.js";

function yieldToEventLoop() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function getSheetByName(workbook, sheetName) {
  if (typeof workbook.getSheet === "function") return workbook.getSheet(sheetName);
  const sheets = workbook.sheets ?? [];
  const found = sheets.find((s) => s.name === sheetName);
  if (!found) throw new Error(`Unknown sheet: ${sheetName}`);
  return found;
}

function getUsedRange(sheet) {
  if (typeof sheet.getUsedRange === "function") return sheet.getUsedRange();
  return sheet.usedRange ?? null;
}

function getCellText(cell, { lookIn, valueMode }) {
  if (!cell) return "";

  if (lookIn === "formulas") {
    if (cell.formula != null && cell.formula !== "") return String(cell.formula);
    if (cell.value == null) return "";
    return String(cell.value);
  }

  // values
  if (valueMode === "raw") {
    if (cell.value == null) return "";
    return String(cell.value);
  }

  // display
  if (cell.display != null) return String(cell.display);
  if (cell.value == null) return "";
  return String(cell.value);
}

function buildMatcher(query, { matchCase, matchEntireCell, useWildcards }) {
  const pattern = String(query);
  return excelWildcardToRegExp(pattern, {
    matchCase,
    matchEntireCell,
    useWildcards,
  });
}

function buildScopeSegments(workbook, { scope, currentSheetName, selectionRanges }) {
  const segments = [];

  if (scope === "workbook") {
    for (const sheet of workbook.sheets ?? []) {
      const range = getUsedRange(sheet);
      if (!range) continue;
      segments.push({ sheetName: sheet.name, ranges: [range] });
    }
    return segments;
  }

  const sheetName = currentSheetName;
  if (!sheetName) throw new Error("Search scope requires currentSheetName");

  if (scope === "selection") {
    const ranges = selectionRanges ?? [];
    return [{ sheetName, ranges }];
  }

  // sheet (default)
  const range = getUsedRange(getSheetByName(workbook, sheetName));
  if (!range) return [];
  return [{ sheetName, ranges: [range] }];
}

async function* iterateCellsInScope(
  workbook,
  {
    scope = "sheet",
    currentSheetName,
    selectionRanges,
    searchOrder = "byRows",
    yieldEvery = 10_000,
    signal,
  } = {},
) {
  const segments = buildScopeSegments(workbook, { scope, currentSheetName, selectionRanges });

  let scanned = 0;

  for (const segment of segments) {
    const sheet = getSheetByName(workbook, segment.sheetName);

    for (const range of segment.ranges) {
      if (signal?.aborted) return;

      if (typeof sheet.iterateCells === "function") {
        for (const { row, col, cell } of sheet.iterateCells(range, { order: searchOrder })) {
          if (signal?.aborted) return;
          yield { sheetName: segment.sheetName, row, col, cell };
          scanned++;
          if (yieldEvery > 0 && scanned % yieldEvery === 0) {
            await yieldToEventLoop();
          }
        }
        continue;
      }

      // Fallback: scan rectangular range via getCell.
      if (typeof sheet.getCell !== "function") {
        throw new Error(
          `Sheet ${segment.sheetName} does not provide iterateCells(range) or getCell(row,col)`,
        );
      }

      if (searchOrder === "byColumns") {
        for (let col = range.startCol; col <= range.endCol; col++) {
          for (let row = range.startRow; row <= range.endRow; row++) {
            if (signal?.aborted) return;
            yield { sheetName: segment.sheetName, row, col, cell: sheet.getCell(row, col) };
            scanned++;
            if (yieldEvery > 0 && scanned % yieldEvery === 0) {
              await yieldToEventLoop();
            }
          }
        }
      } else {
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            if (signal?.aborted) return;
            yield { sheetName: segment.sheetName, row, col, cell: sheet.getCell(row, col) };
            scanned++;
            if (yieldEvery > 0 && scanned % yieldEvery === 0) {
              await yieldToEventLoop();
            }
          }
        }
      }
    }
  }
}

export async function* iterateMatches(workbook, query, options = {}) {
  if (query == null || String(query) === "") return;

  const {
    lookIn = "values",
    valueMode = "display",
    matchCase = false,
    matchEntireCell = false,
    useWildcards = true,
  } = options;

  const re = buildMatcher(query, { matchCase, matchEntireCell, useWildcards });

  for await (const entry of iterateCellsInScope(workbook, options)) {
    const text = getCellText(entry.cell, { lookIn, valueMode });
    if (re.test(text)) {
      yield {
        sheetName: entry.sheetName,
        row: entry.row,
        col: entry.col,
        address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
        text,
      };
    }
  }
}

export async function findAll(workbook, query, options = {}) {
  const matches = [];
  for await (const m of iterateMatches(workbook, query, options)) {
    matches.push(m);
  }
  return matches;
}

/**
 * Find the next match after `from` (exclusive). If `wrap` is true (default),
 * wraps to the first match before `from` when no match is found after it.
 */
export async function findNext(workbook, query, options = {}, from) {
  if (query == null || String(query) === "") return null;

  const {
    lookIn = "values",
    valueMode = "display",
    matchCase = false,
    matchEntireCell = false,
    useWildcards = true,
    wrap = true,
  } = options;

  const re = buildMatcher(query, { matchCase, matchEntireCell, useWildcards });

  const hasFrom = from && from.sheetName != null && from.row != null && from.col != null;
  let passedFrom = !hasFrom;
  let fromFound = false;
  let firstMatchBeforeFrom = null;
  let matchAtFrom = null;

  for await (const entry of iterateCellsInScope(workbook, options)) {
    const isFromCell =
      hasFrom && entry.sheetName === from.sheetName && entry.row === from.row && entry.col === from.col;
    const text = getCellText(entry.cell, { lookIn, valueMode });

    if (hasFrom && !passedFrom) {
      if (isFromCell) {
        passedFrom = true;
        fromFound = true;
        if (re.test(text)) {
          matchAtFrom = {
            sheetName: entry.sheetName,
            row: entry.row,
            col: entry.col,
            address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
            text,
          };
        }
        continue;
      }

      if (!firstMatchBeforeFrom && re.test(text)) {
        firstMatchBeforeFrom = {
          sheetName: entry.sheetName,
          row: entry.row,
          col: entry.col,
          address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
          text,
        };
      }
      continue;
    }

    // After from (exclusive) or no from at all.
    if (isFromCell) continue;
    if (re.test(text)) {
      return {
        sheetName: entry.sheetName,
        row: entry.row,
        col: entry.col,
        address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
        text,
      };
    }
  }

  if (!hasFrom || !wrap) return null;
  if (!fromFound) return firstMatchBeforeFrom;
  return firstMatchBeforeFrom ?? matchAtFrom;
}

export { getCellText };
