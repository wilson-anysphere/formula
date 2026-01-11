import { formatA1Address } from "./a1.js";
import { excelWildcardToRegExp } from "./wildcards.js";
import { createTimeSlicer } from "./scheduler.js";
import {
  buildScopeSegments,
  expandSelectionRangesForMerges,
  getMergedMasterCell,
  getSheetByName,
  rangesOverlap,
} from "./scope.js";
import { encodeCellId } from "./indexing.js";
import { getCellText } from "./text.js";

function buildMatcher(query, { matchCase, matchEntireCell, useWildcards }) {
  const pattern = String(query);
  return excelWildcardToRegExp(pattern, {
    matchCase,
    matchEntireCell,
    useWildcards,
  });
}

async function* iterateCellsInScope(
  workbook,
  {
    scope = "sheet",
    currentSheetName,
    selectionRanges,
    searchOrder = "byRows",
    timeBudgetMs = 10,
    scheduler,
    checkEvery,
    signal,
  } = {},
) {
  const segments = buildScopeSegments(workbook, { scope, currentSheetName, selectionRanges });
  if (scope === "selection") {
    for (const seg of segments) {
      const sheet = getSheetByName(workbook, seg.sheetName);
      seg.ranges = expandSelectionRangesForMerges(sheet, seg.ranges);
    }
  }

  const slicer = createTimeSlicer({ signal, timeBudgetMs, scheduler, checkEvery });

  for (const segment of segments) {
    const sheet = getSheetByName(workbook, segment.sheetName);
    const visited = segment.ranges.length > 1 && rangesOverlap(segment.ranges) ? new Set() : null;

    for (const range of segment.ranges) {
      if (typeof sheet.iterateCells === "function") {
        for (const { row, col, cell } of sheet.iterateCells(range, { order: searchOrder })) {
          await slicer.checkpoint();
          const master = getMergedMasterCell(sheet, row, col);
          if (master && (master.row !== row || master.col !== col)) continue;
          if (visited) {
            const id = encodeCellId(row, col);
            if (visited.has(id)) continue;
            visited.add(id);
          }
          yield { sheetName: segment.sheetName, row, col, cell };
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
            await slicer.checkpoint();
            const master = getMergedMasterCell(sheet, row, col);
            if (master && (master.row !== row || master.col !== col)) continue;
            if (visited) {
              const id = encodeCellId(row, col);
              if (visited.has(id)) continue;
              visited.add(id);
            }
            yield { sheetName: segment.sheetName, row, col, cell: sheet.getCell(row, col) };
          }
        }
      } else {
        for (let row = range.startRow; row <= range.endRow; row++) {
          for (let col = range.startCol; col <= range.endCol; col++) {
            await slicer.checkpoint();
            const master = getMergedMasterCell(sheet, row, col);
            if (master && (master.row !== row || master.col !== col)) continue;
            if (visited) {
              const id = encodeCellId(row, col);
              if (visited.has(id)) continue;
              visited.add(id);
            }
            yield { sheetName: segment.sheetName, row, col, cell: sheet.getCell(row, col) };
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

  const normalizedFrom = (() => {
    if (!from || from.sheetName == null || from.row == null || from.col == null) return null;
    const sheet = getSheetByName(workbook, from.sheetName);
    const master = getMergedMasterCell(sheet, from.row, from.col);
    if (!master) return from;
    return { ...from, row: master.row, col: master.col };
  })();

  const hasFrom = normalizedFrom && normalizedFrom.sheetName != null;
  let passedFrom = !hasFrom;
  let fromFound = false;
  let firstMatchBeforeFrom = null;
  let matchAtFrom = null;

  for await (const entry of iterateCellsInScope(workbook, options)) {
    const isFromCell =
      hasFrom &&
      entry.sheetName === normalizedFrom.sheetName &&
      entry.row === normalizedFrom.row &&
      entry.col === normalizedFrom.col;
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

/**
 * Find the previous match before `from` (exclusive). If `wrap` is true (default),
 * wraps to the last match after `from` when no match is found before it.
 */
export async function findPrev(workbook, query, options = {}, from) {
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

  const normalizedFrom = (() => {
    if (!from || from.sheetName == null || from.row == null || from.col == null) return null;
    const sheet = getSheetByName(workbook, from.sheetName);
    const master = getMergedMasterCell(sheet, from.row, from.col);
    if (!master) return from;
    return { ...from, row: master.row, col: master.col };
  })();

  const hasFrom = normalizedFrom && normalizedFrom.sheetName != null;
  let passedFrom = !hasFrom;

  let lastMatchBeforeFrom = null;
  let lastMatchOverall = null;
  let matchAtFrom = null;

  for await (const entry of iterateCellsInScope(workbook, options)) {
    const isFromCell =
      hasFrom &&
      entry.sheetName === normalizedFrom.sheetName &&
      entry.row === normalizedFrom.row &&
      entry.col === normalizedFrom.col;
    const text = getCellText(entry.cell, { lookIn, valueMode });

    if (hasFrom && !passedFrom) {
      if (isFromCell) {
        passedFrom = true;
        if (re.test(text)) {
          matchAtFrom = {
            sheetName: entry.sheetName,
            row: entry.row,
            col: entry.col,
            address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
            text,
          };
          lastMatchOverall = matchAtFrom;
        }
        continue;
      }

      if (re.test(text)) {
        lastMatchBeforeFrom = {
          sheetName: entry.sheetName,
          row: entry.row,
          col: entry.col,
          address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
          text,
        };
        lastMatchOverall = lastMatchBeforeFrom;
      }
      continue;
    }

    // After from (exclusive) or no from at all.
    if (isFromCell) {
      if (re.test(text)) {
        matchAtFrom = {
          sheetName: entry.sheetName,
          row: entry.row,
          col: entry.col,
          address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
          text,
        };
        lastMatchOverall = matchAtFrom;
      }
      continue;
    }

    if (re.test(text)) {
      lastMatchOverall = {
        sheetName: entry.sheetName,
        row: entry.row,
        col: entry.col,
        address: `${entry.sheetName}!${formatA1Address({ row: entry.row, col: entry.col })}`,
        text,
      };
    }
  }

  if (!hasFrom) return lastMatchOverall;
  if (lastMatchBeforeFrom) return lastMatchBeforeFrom;
  if (!wrap) return null;
  return lastMatchOverall ?? matchAtFrom;
}

export { getCellText };
