import { formatA1Address } from "./a1.js";
import { excelWildcardToRegExp } from "./wildcards.js";
import { createTimeSlicer, throwIfAborted } from "./scheduler.js";
import {
  buildScopeSegments,
  expandSelectionRangesForMerges,
  getMergedMasterCell,
  getSheetByName,
  rangeContains,
} from "./scope.js";
import { decodeCellId } from "./indexing.js";
import { getCellText } from "./text.js";
import { applyReplaceToCell } from "./replaceCore.js";

function buildMatcher(query, { matchCase, matchEntireCell, useWildcards }) {
  return excelWildcardToRegExp(String(query), { matchCase, matchEntireCell, useWildcards });
}

function comparePositions(a, b) {
  if (a.segmentIndex !== b.segmentIndex) return a.segmentIndex - b.segmentIndex;
  if (a.primary !== b.primary) return a.primary - b.primary;
  return a.secondary - b.secondary;
}

function positionKeyFor({ row, col }, searchOrder) {
  // Compare keys as a tuple (primary,secondary) to avoid reliance on sheet size.
  if (searchOrder === "byColumns") return { primary: col, secondary: row };
  return { primary: row, secondary: col };
}

function normalizeCursorToMergeMaster(workbook, cursor) {
  if (!cursor?.sheetName || cursor.row == null || cursor.col == null) return cursor;
  const sheet = getSheetByName(workbook, cursor.sheetName);
  const master = getMergedMasterCell(sheet, cursor.row, cursor.col);
  if (!master) return cursor;
  return { sheetName: cursor.sheetName, row: master.row, col: master.col };
}

function matchToResult(sheetName, row, col, text, { wrapped = false } = {}) {
  return {
    sheetName,
    row,
    col,
    address: `${sheetName}!${formatA1Address({ row, col })}`,
    text,
    wrapped,
  };
}

/**
 * Stateful Excel-style Find/Replace session.
 *
 * The session caches matches so repeated `findNext` / `findPrev` calls are fast.
 * Use a shared `WorkbookSearchIndex` to avoid rescanning large workbooks across
 * sessions.
 */
export class SearchSession {
  constructor(workbook, query, options = {}) {
    if (!workbook) throw new Error("SearchSession: workbook is required");
    this.workbook = workbook;
    this.query = query;
    this.options = { ...options };

    this.cursor = normalizeCursorToMergeMaster(workbook, options.from ?? null);

    this.stats = {
      cellsScanned: 0,
      indexCellsVisited: 0,
    };

    /** @type {null | Array<ReturnType<typeof matchToResult> & { _pos: {segmentIndex:number,primary:number,secondary:number} }>} */
    this._matches = null;
    this._building = null;
  }

  setCursor(pos) {
    this.cursor = normalizeCursorToMergeMaster(this.workbook, pos);
  }

  _effectiveSegments() {
    const segments = buildScopeSegments(this.workbook, this.options);
    // Merge-aware selection expansion.
    if (this.options.scope === "selection") {
      for (const seg of segments) {
        const sheet = getSheetByName(this.workbook, seg.sheetName);
        seg.ranges = expandSelectionRangesForMerges(sheet, seg.ranges);
      }
    }
    return segments;
  }

  async _buildMatches({ signal } = {}) {
    if (this._matches) return this._matches;
    if (this._building) return this._building;

    this._building = (async () => {
      try {
        const query = this.query;
        if (query == null || String(query) === "") {
          this._matches = [];
          return this._matches;
        }

        const {
          lookIn = "values",
          valueMode = "display",
          matchCase = false,
          matchEntireCell = false,
          useWildcards = true,
          searchOrder = "byRows",
          // Scheduling / cancellation.
          timeBudgetMs = 10,
          scheduler,
          checkEvery,
          // Indexing.
          index = null,
          indexStrategy = "auto", // "auto" | "always" | "never"
        } = this.options;

        const segments = this._effectiveSegments();
        const re = buildMatcher(query, { matchCase, matchEntireCell, useWildcards });
        const slicer = createTimeSlicer({ signal, timeBudgetMs, scheduler, checkEvery });

        /** @type {Array<ReturnType<typeof matchToResult> & { _pos: {segmentIndex:number,primary:number,secondary:number} }>} */
        const matches = [];

        // Decide whether to build/use index in this session.
        const canUseIndex = index && indexStrategy !== "never";

        for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex++) {
          throwIfAborted(signal);
          const segment = segments[segmentIndex];
          const sheet = getSheetByName(this.workbook, segment.sheetName);

          let candidates = null;
          if (canUseIndex) {
            const modeKey = lookIn === "formulas" ? "formulas" : `values:${valueMode}`;
            if (indexStrategy === "always") {
              const before = index.getSheetModeIndex(segment.sheetName, modeKey)?.stats.cellsVisited ?? 0;
              const built = await index.ensureSheetModeBuilt(segment.sheetName, modeKey, { lookIn, valueMode }, {
                signal,
                timeBudgetMs,
                scheduler,
                checkEvery,
              });
              this.stats.indexCellsVisited += Math.max(0, built.stats.cellsVisited - before);
              candidates = index.queryCandidates(segment.sheetName, query, { lookIn, valueMode, useWildcards });
            } else if (indexStrategy === "auto") {
              const before = index.getSheetModeIndex(segment.sheetName, modeKey)?.stats.cellsVisited ?? 0;
              const res = await index.ensureIndexForQuery(
                query,
                { ...this.options, scope: "sheet", currentSheetName: segment.sheetName },
                { signal, timeBudgetMs, scheduler, checkEvery },
              );
              if (res) {
                const after = index.getSheetModeIndex(segment.sheetName, modeKey)?.stats.cellsVisited ?? before;
                this.stats.indexCellsVisited += Math.max(0, after - before);
                candidates = index.queryCandidates(segment.sheetName, query, { lookIn, valueMode, useWildcards });
              }
            }
          }

          if (Array.isArray(candidates) && typeof sheet.getCell !== "function") {
            // The index can narrow candidates, but we still need `getCell` to
            // validate wildcard/substring semantics.
            candidates = null;
          }

          // Process each range in-order (important for multi-area selections).
          const seenCandidateIds = Array.isArray(candidates) ? new Set() : null;
          for (const range of segment.ranges) {
            throwIfAborted(signal);

            if (Array.isArray(candidates)) {
              // Filter candidates to the current range.
              for (const id of candidates) {
                const { row, col } = decodeCellId(id);
                if (!rangeContains(range, row, col)) continue;
                if (seenCandidateIds.has(id)) continue;
                seenCandidateIds.add(id);

                await slicer.checkpoint();
                this.stats.cellsScanned++;

                const cell = sheet.getCell(row, col);
                const text = getCellText(cell, { lookIn, valueMode });
                if (!re.test(text)) continue;

                const posKey = positionKeyFor({ row, col }, searchOrder);
                matches.push({
                  ...matchToResult(segment.sheetName, row, col, text),
                  _pos: { segmentIndex, ...posKey },
                });
              }
              continue;
            }

            // Fallback: scan the range.
            if (typeof sheet.iterateCells === "function") {
              for (const { row, col, cell } of sheet.iterateCells(range, { order: searchOrder })) {
                await slicer.checkpoint();
                this.stats.cellsScanned++;

                const master = getMergedMasterCell(sheet, row, col);
                if (master && (master.row !== row || master.col !== col)) continue;

                const text = getCellText(cell, { lookIn, valueMode });
                if (!re.test(text)) continue;

                const posKey = positionKeyFor({ row, col }, searchOrder);
                matches.push({
                  ...matchToResult(segment.sheetName, row, col, text),
                  _pos: { segmentIndex, ...posKey },
                });
              }
              continue;
            }

            if (typeof sheet.getCell !== "function") {
              throw new Error(
                `Sheet ${segment.sheetName} does not provide iterateCells(range) or getCell(row,col)`,
              );
            }

            if (searchOrder === "byColumns") {
              for (let col = range.startCol; col <= range.endCol; col++) {
                for (let row = range.startRow; row <= range.endRow; row++) {
                  await slicer.checkpoint();
                  this.stats.cellsScanned++;

                  const master = getMergedMasterCell(sheet, row, col);
                  if (master && (master.row !== row || master.col !== col)) continue;

                  const cell = sheet.getCell(row, col);
                  const text = getCellText(cell, { lookIn, valueMode });
                  if (!re.test(text)) continue;

                  const posKey = positionKeyFor({ row, col }, searchOrder);
                  matches.push({
                    ...matchToResult(segment.sheetName, row, col, text),
                    _pos: { segmentIndex, ...posKey },
                  });
                }
              }
            } else {
              for (let row = range.startRow; row <= range.endRow; row++) {
                for (let col = range.startCol; col <= range.endCol; col++) {
                  await slicer.checkpoint();
                  this.stats.cellsScanned++;

                  const master = getMergedMasterCell(sheet, row, col);
                  if (master && (master.row !== row || master.col !== col)) continue;

                  const cell = sheet.getCell(row, col);
                  const text = getCellText(cell, { lookIn, valueMode });
                  if (!re.test(text)) continue;

                  const posKey = positionKeyFor({ row, col }, searchOrder);
                  matches.push({
                    ...matchToResult(segment.sheetName, row, col, text),
                    _pos: { segmentIndex, ...posKey },
                  });
                }
              }
            }
          }
        }

        matches.sort((a, b) => comparePositions(a._pos, b._pos));

        this._matches = matches;
        return matches;
      } finally {
        // Allow retry after abort/error.
        this._building = null;
      }
    })();

    return this._building;
  }

  _cursorPosKey() {
    const { searchOrder = "byRows" } = this.options;
    const segments = this._effectiveSegments();
    const cursor = normalizeCursorToMergeMaster(this.workbook, this.cursor);
    if (!cursor?.sheetName || cursor.row == null || cursor.col == null) return null;

    const segmentIndex = segments.findIndex((s) => s.sheetName === cursor.sheetName);
    if (segmentIndex === -1) return null;
    const key = positionKeyFor({ row: cursor.row, col: cursor.col }, searchOrder);
    return { segmentIndex, ...key };
  }

  _pickNext(matches, cursorKey) {
    const wrap = this.options.wrap ?? true;

    // Find the first match strictly after the cursor.
    let lo = 0;
    let hi = matches.length;
    if (cursorKey) {
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        const cmp = comparePositions(matches[mid]._pos, cursorKey);
        if (cmp <= 0) lo = mid + 1;
        else hi = mid;
      }
    }

    if (lo < matches.length) return { picked: matches[lo], wrapped: false };
    if (!wrap) return { picked: null, wrapped: false };
    return { picked: matches[0], wrapped: true };
  }

  _pickPrev(matches, cursorKey) {
    const wrap = this.options.wrap ?? true;

    let lo = 0;
    let hi = matches.length;
    if (cursorKey) {
      // Find first element >= cursorKey, then step back for strict prev.
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        const cmp = comparePositions(matches[mid]._pos, cursorKey);
        if (cmp < 0) lo = mid + 1;
        else hi = mid;
      }
    }

    const idx = cursorKey ? lo - 1 : matches.length - 1;
    if (idx >= 0) return { picked: matches[idx], wrapped: false };
    if (!wrap) return { picked: null, wrapped: false };
    return { picked: matches[matches.length - 1], wrapped: true };
  }

  async findNext({ signal } = {}) {
    const matches = await this._buildMatches({ signal });
    if (matches.length === 0) return null;

    const cursorKey = this._cursorPosKey();
    const { picked, wrapped } = this._pickNext(matches, cursorKey);
    if (!picked) return null;
    this.setCursor({ sheetName: picked.sheetName, row: picked.row, col: picked.col });

    const { _pos, ...publicMatch } = picked;
    return { ...publicMatch, wrapped };
  }

  async findPrev({ signal } = {}) {
    const matches = await this._buildMatches({ signal });
    if (matches.length === 0) return null;

    const cursorKey = this._cursorPosKey();
    const { picked, wrapped } = this._pickPrev(matches, cursorKey);
    if (!picked) return null;
    this.setCursor({ sheetName: picked.sheetName, row: picked.row, col: picked.col });

    const { _pos, ...publicMatch } = picked;
    return { ...publicMatch, wrapped };
  }

  async replaceNext(replacement, { signal } = {}) {
    const matches = await this._buildMatches({ signal });
    if (matches.length === 0) return null;

    const cursorKey = this._cursorPosKey();
    const { picked, wrapped } = this._pickNext(matches, cursorKey);
    if (!picked) return null;

    const sheet = getSheetByName(this.workbook, picked.sheetName);
    const oldCell = sheet.getCell(picked.row, picked.col);
    const res = applyReplaceToCell(oldCell, this.query, replacement, this.options, { replaceAll: false });
    if (res.replaced) {
      sheet.setCell(picked.row, picked.col, res.cell);
      if (this.options.index) {
        this.options.index.updateCell(picked.sheetName, picked.row, picked.col, { oldCell, newCell: res.cell });
      }
    }

    // Update cached matches for this cell so subsequent findNext calls don't
    // return stale results.
    if (this._matches) {
      const { lookIn = "values", valueMode = "display", matchCase = false, matchEntireCell = false, useWildcards = true } =
        this.options;
      const re = buildMatcher(this.query, { matchCase, matchEntireCell, useWildcards });
      const newText = getCellText(res.cell, { lookIn, valueMode });

      // Locate the match entry by position.
      const key = picked._pos;
      const idx = this._matches.findIndex((m) => comparePositions(m._pos, key) === 0);
      const stillMatches = re.test(newText);
      if (idx !== -1) {
        if (stillMatches) {
          this._matches[idx] = { ...this._matches[idx], text: newText };
        } else {
          this._matches.splice(idx, 1);
        }
      } else if (stillMatches) {
        // The cell didn't previously match (possible when replacing with
        // wildcards disabled). Insert into cache for correctness.
        const posKey = positionKeyFor(
          { row: picked.row, col: picked.col },
          this.options.searchOrder ?? "byRows",
        );
        this._matches.push({
          ...matchToResult(picked.sheetName, picked.row, picked.col, newText),
          _pos: { segmentIndex: key.segmentIndex, ...posKey },
        });
        this._matches.sort((a, b) => comparePositions(a._pos, b._pos));
      }
    }

    this.setCursor({ sheetName: picked.sheetName, row: picked.row, col: picked.col });

    const { _pos, ...publicMatch } = picked;
    return { match: { ...publicMatch, wrapped }, replaced: res.replaced, replacements: res.replacements };
  }
}
