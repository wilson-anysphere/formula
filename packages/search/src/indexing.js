import { createTimeSlicer } from "./scheduler.js";
import { getSheetByName, getUsedRange, getMergedMasterCell, rangeContains } from "./scope.js";
import { getCellText } from "./text.js";

// Excel's hard max is 1,048,576 rows and 16,384 columns. Using 2^20 as the
// stride gives us a stable, sortable cell id for row-major ordering while
// staying within JS safe integers.
export const CELL_ID_STRIDE = 1_048_576;

export function encodeCellId(row, col) {
  return row * CELL_ID_STRIDE + col;
}

export function decodeCellId(id) {
  const row = Math.floor(id / CELL_ID_STRIDE);
  const col = id - row * CELL_ID_STRIDE;
  return { row, col };
}

function toModeKey({ lookIn, valueMode }) {
  if (lookIn === "formulas") return "formulas";
  return `values:${valueMode ?? "display"}`;
}

function normalizeForIndex(text) {
  return String(text).toLowerCase();
}

function computeGrams(text, gramLength) {
  const s = String(text);
  if (s.length < gramLength) return [];
  /** @type {string[]} */
  const grams = [];
  for (let i = 0; i <= s.length - gramLength; i++) {
    grams.push(s.slice(i, i + gramLength));
  }
  // De-dup within a single cell so the per-gram set doesn't see repeated adds.
  return Array.from(new Set(grams));
}

function parseLiteralRuns(pattern, { useWildcards = true } = {}) {
  const input = String(pattern);
  /** @type {string[]} */
  const runs = [];
  let current = "";

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];

    if (useWildcards && ch === "~") {
      const next = input[i + 1];
      if (next == null) {
        current += "~";
      } else {
        current += next;
        i++;
      }
      continue;
    }

    if (useWildcards && (ch === "*" || ch === "?")) {
      if (current) runs.push(current);
      current = "";
      continue;
    }

    current += ch;
  }

  if (current) runs.push(current);
  return runs;
}

function pickBestLiteral(pattern, { useWildcards = true } = {}) {
  const runs = parseLiteralRuns(pattern, { useWildcards });
  if (runs.length === 0) return "";
  runs.sort((a, b) => b.length - a.length);
  return runs[0];
}

export function isQueryIndexable(query, { useWildcards = true } = {}, { gramLength = 3 } = {}) {
  const literal = pickBestLiteral(query, { useWildcards });
  return normalizeForIndex(literal).length >= gramLength;
}

class SheetModeIndex {
  constructor({ gramLength }) {
    this.gramLength = gramLength;
    /** @type {Map<string, Set<number>>} */
    this.grams = new Map();
    this.built = false;
    this.building = null;

    this.stats = {
      cellsVisited: 0,
      cellsIndexed: 0,
      yields: 0,
    };
  }

  _addCell(id, textLower) {
    const grams = computeGrams(textLower, this.gramLength);
    if (grams.length === 0) return;
    for (const gram of grams) {
      let set = this.grams.get(gram);
      if (!set) {
        set = new Set();
        this.grams.set(gram, set);
      }
      set.add(id);
    }
    this.stats.cellsIndexed++;
  }

  _removeCell(id, textLower) {
    const grams = computeGrams(textLower, this.gramLength);
    if (grams.length === 0) return;
    for (const gram of grams) {
      const set = this.grams.get(gram);
      if (!set) continue;
      set.delete(id);
      if (set.size === 0) this.grams.delete(gram);
    }
  }

  /**
   * @returns {number[] | null} Candidate cell ids or null when the index can't help.
   */
  queryCandidates(query, { useWildcards = true } = {}) {
    const literal = pickBestLiteral(query, { useWildcards });
    const literalLower = normalizeForIndex(literal);
    if (literalLower.length < this.gramLength) return null;

    const grams = computeGrams(literalLower, this.gramLength);
    if (grams.length === 0) return null;

    /** @type {Array<Set<number>>} */
    const sets = [];
    for (const gram of grams) {
      const set = this.grams.get(gram);
      if (!set) return [];
      sets.push(set);
    }

    // Intersect starting from the smallest set to minimize work.
    sets.sort((a, b) => a.size - b.size);

    const [smallest, ...rest] = sets;
    /** @type {number[]} */
    const out = [];
    for (const id of smallest) {
      let ok = true;
      for (const set of rest) {
        if (!set.has(id)) {
          ok = false;
          break;
        }
      }
      if (ok) out.push(id);
    }
    return out;
  }
}

/**
 * Workbook-level index that can be shared across multiple SearchSessions.
 *
 * The index is built lazily, per sheet and per "look in" mode.
 */
export class WorkbookSearchIndex {
  constructor(workbook, { gramLength = 3, autoThresholdCells = 50_000 } = {}) {
    this.workbook = workbook;
    this.gramLength = gramLength;
    this.autoThresholdCells = autoThresholdCells;

    /** @type {Map<string, Map<string, SheetModeIndex>>} sheetName -> modeKey -> index */
    this._sheetIndexes = new Map();
  }

  getSheetModeIndex(sheetName, modeKey) {
    const sheetMap = this._sheetIndexes.get(sheetName);
    return sheetMap?.get(modeKey) ?? null;
  }

  _getOrCreateSheetModeIndex(sheetName, modeKey) {
    let sheetMap = this._sheetIndexes.get(sheetName);
    if (!sheetMap) {
      sheetMap = new Map();
      this._sheetIndexes.set(sheetName, sheetMap);
    }

    let idx = sheetMap.get(modeKey);
    if (!idx) {
      idx = new SheetModeIndex({ gramLength: this.gramLength });
      sheetMap.set(modeKey, idx);
    }
    return idx;
  }

  _estimateUsedRangeCellCount(sheet) {
    const range = getUsedRange(sheet);
    if (!range) return 0;
    const rows = range.endRow - range.startRow + 1;
    const cols = range.endCol - range.startCol + 1;
    if (rows <= 0 || cols <= 0) return 0;
    return rows * cols;
  }

  /**
   * Ensure the index exists for a given query/options. Returns `null` when the
   * caller should fall back to scanning (e.g. small sheet or unsupported pattern).
   */
  async ensureIndexForQuery(query, options = {}, { signal, timeBudgetMs, scheduler, checkEvery } = {}) {
    const { scope = "sheet", currentSheetName } = options;
    if (scope !== "sheet" && scope !== "workbook" && scope !== "selection") {
      return null;
    }

    // If we can't extract a useful literal substring (e.g. query is too short),
    // building the index is usually a net loss. Fall back to scanning.
    if (
      !isQueryIndexable(query, { useWildcards: options.useWildcards ?? true }, { gramLength: this.gramLength })
    ) {
      return null;
    }

    // Only build indexes on sheet/workbook scopes. For selection searches, the
    // sheet-level index can still help, but we rely on the same heuristics.
    const modeKey = toModeKey(options);

    // For workbook scope, the SearchSession will call `ensureSheetModeBuilt` per sheet.
    if (scope === "workbook") return { modeKey };

    if (!currentSheetName) throw new Error("Search scope requires currentSheetName");
    const sheet = getSheetByName(this.workbook, currentSheetName);
    const estimate = this._estimateUsedRangeCellCount(sheet);
    if (estimate < this.autoThresholdCells) return null;

    await this.ensureSheetModeBuilt(currentSheetName, modeKey, options, {
      signal,
      timeBudgetMs,
      scheduler,
      checkEvery,
    });

    return { modeKey };
  }

  async ensureSheetModeBuilt(
    sheetName,
    modeKey,
    options = {},
    { signal, timeBudgetMs = 10, scheduler, checkEvery } = {},
  ) {
    const sheet = getSheetByName(this.workbook, sheetName);
    const idx = this._getOrCreateSheetModeIndex(sheetName, modeKey);
    if (idx.built) return idx;
    if (idx.building) return idx.building;

    // Reset any partial state from a previous aborted build.
    idx.grams.clear();
    idx.stats.cellsVisited = 0;
    idx.stats.cellsIndexed = 0;
    idx.stats.yields = 0;

    const range = getUsedRange(sheet);
    if (!range) {
      idx.built = true;
      return idx;
    }

    idx.building = (async () => {
      try {
        const slicer = createTimeSlicer({ signal, timeBudgetMs, scheduler, checkEvery });
        const gramLength = idx.gramLength;

        const iterate =
          typeof sheet.iterateCells === "function"
            ? sheet.iterateCells(range, { order: "byRows" })
            : null;

        if (iterate) {
          for (const { row, col, cell } of iterate) {
            await slicer.checkpoint();
            idx.stats.cellsVisited++;

            const master = getMergedMasterCell(sheet, row, col);
            if (master && (master.row !== row || master.col !== col)) continue;

            const text = getCellText(cell, options);
            if (text == null || text === "") continue;
            const textStr = String(text);
            if (textStr.length < gramLength) continue;
            idx._addCell(encodeCellId(row, col), textStr.toLowerCase());
          }
        } else if (typeof sheet.getCell === "function") {
          // Scan full used range via getCell. This is O(area); callers should
          // provide iterateCells for sparsity.
          for (let r = range.startRow; r <= range.endRow; r++) {
            for (let c = range.startCol; c <= range.endCol; c++) {
              await slicer.checkpoint();
              idx.stats.cellsVisited++;

              const master = getMergedMasterCell(sheet, r, c);
              if (master && (master.row !== r || master.col !== c)) continue;

              const cell = sheet.getCell(r, c);
              const text = getCellText(cell, options);
              if (text == null || text === "") continue;
              const textStr = String(text);
              if (textStr.length < gramLength) continue;
              idx._addCell(encodeCellId(r, c), textStr.toLowerCase());
            }
          }
        } else {
          throw new Error(`Sheet ${sheetName} does not provide iterateCells(range) or getCell(row,col)`);
        }

        idx.built = true;
        return idx;
      } finally {
        // Allow retry after abort/error.
        idx.building = null;
      }
    })();

    return idx.building;
  }

  queryCandidates(sheetName, query, options = {}) {
    const modeKey = toModeKey(options);
    const sheetMap = this._sheetIndexes.get(sheetName);
    const idx = sheetMap?.get(modeKey);
    if (!idx || !idx.built) return null;
    return idx.queryCandidates(query, { useWildcards: options.useWildcards ?? true });
  }

  /**
   * Incrementally update the index after a cell edit.
   *
   * Callers are expected to provide both `oldCell` and `newCell` so we can
   * remove old grams without storing per-cell state in the index.
   */
  updateCell(sheetName, row, col, { oldCell, newCell }) {
    const sheetMap = this._sheetIndexes.get(sheetName);
    if (!sheetMap) return;

    const id = encodeCellId(row, col);
    const modes = sheetMap.entries();
    for (const [modeKey, idx] of modes) {
      if (!idx.built) continue;

      const modeOptions =
        modeKey === "formulas"
          ? { lookIn: "formulas" }
          : { lookIn: "values", valueMode: modeKey.split(":")[1] ?? "display" };

      const oldTextLower = normalizeForIndex(getCellText(oldCell, modeOptions));
      const newTextLower = normalizeForIndex(getCellText(newCell, modeOptions));

      if (oldTextLower !== "") idx._removeCell(id, oldTextLower);
      if (newTextLower !== "") idx._addCell(id, newTextLower);
    }
  }

  /**
   * Helper for filtering candidates to a rectangular range.
   */
  filterCandidatesToRange(candidates, range) {
    /** @type {number[]} */
    const out = [];
    for (const id of candidates) {
      const { row, col } = decodeCellId(id);
      if (rangeContains(range, row, col)) out.push(id);
    }
    return out;
  }
}
