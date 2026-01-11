export function getSheetByName(workbook, sheetName) {
  if (typeof workbook.getSheet === "function") return workbook.getSheet(sheetName);
  const sheets = workbook.sheets ?? [];
  const found = sheets.find((s) => s.name === sheetName);
  if (!found) throw new Error(`Unknown sheet: ${sheetName}`);
  return found;
}

export function getUsedRange(sheet) {
  if (typeof sheet.getUsedRange === "function") return sheet.getUsedRange();
  return sheet.usedRange ?? null;
}

export function getMergedRanges(sheet) {
  if (typeof sheet.getMergedRanges === "function") return sheet.getMergedRanges() ?? [];
  return sheet.mergedRanges ?? [];
}

export function getMergedMasterCell(sheet, row, col) {
  if (typeof sheet.getMergedMasterCell === "function") return sheet.getMergedMasterCell(row, col);
  return null;
}

function rangesIntersect(a, b) {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endCol < b.startCol ||
    a.startCol > b.endCol
  );
}

export function rangesOverlap(ranges) {
  const list = ranges ?? [];
  for (let i = 0; i < list.length; i++) {
    for (let j = i + 1; j < list.length; j++) {
      if (rangesIntersect(list[i], list[j])) return true;
    }
  }
  return false;
}

/**
 * Excel treats merged regions as a single cell with the address of its
 * top-left corner. When the search scope is `selection`, the selection may
 * include any cell within a merged region; Excel still searches the merged
 * cell's value/formula (i.e. the top-left cell).
 *
 * To emulate this, we expand selection ranges to include any merged regions
 * they intersect.
 */
export function expandSelectionRangesForMerges(sheet, ranges) {
  const merges = getMergedRanges(sheet);
  if (!merges || merges.length === 0) return ranges;

  /** @type {Array<{startRow:number,endRow:number,startCol:number,endCol:number}>} */
  const out = ranges.map((r) => ({ ...r }));

  for (const merge of merges) {
    // We model merges as a range with `start*`/`end*` coordinates.
    const mergeRange = {
      startRow: merge.startRow,
      endRow: merge.endRow,
      startCol: merge.startCol,
      endCol: merge.endCol,
    };

    for (const r of out) {
      if (!rangesIntersect(r, mergeRange)) continue;
      r.startRow = Math.min(r.startRow, mergeRange.startRow);
      r.endRow = Math.max(r.endRow, mergeRange.endRow);
      r.startCol = Math.min(r.startCol, mergeRange.startCol);
      r.endCol = Math.max(r.endCol, mergeRange.endCol);
    }
  }

  return out;
}

export function buildScopeSegments(workbook, { scope, currentSheetName, selectionRanges }) {
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

export function rangeContains(range, row, col) {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    col >= range.startCol &&
    col <= range.endCol
  );
}
