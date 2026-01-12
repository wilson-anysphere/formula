import { extractCells } from "./extractCells.js";
import { rectIntersectionArea, rectSize, rectToA1 } from "./rect.js";
import { getSheetCellMap, getSheetMatrix, normalizeCell } from "./normalizeCell.js";

const DEFAULT_EXTRACT_MAX_ROWS = 50;
const DEFAULT_EXTRACT_MAX_COLS = 50;
// Region detection for matrix-backed sheets can allocate large visited grids.
// Cap the number of cells we consider to avoid catastrophic allocations on
// Excel-scale sheets.
const DEFAULT_DETECT_REGIONS_CELL_LIMIT = 200000;

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

function isNonEmptyCell(cell) {
  if (!cell) return false;
  if (cell.f != null && String(cell.f).trim() !== "") return true;
  const v = cell.v;
  if (v == null) return false;
  if (typeof v === "string") return v.trim() !== "";
  return true;
}

function isFormulaCell(cell) {
  return !!(cell && cell.f != null && String(cell.f).trim() !== "");
}

/**
 * @param {import('./workbookTypes').Workbook} workbook
 * @returns {Map<string, import('./workbookTypes').Sheet>}
 */
function sheetMap(workbook) {
  const map = new Map();
  for (const s of workbook.sheets || []) map.set(s.name, s);
  return map;
}

/**
 * Detect connected regions (4-neighbor) for a predicate over sheet cells.
 *
 * @param {import('./workbookTypes').Sheet} sheet
 * @param {(cell: any) => boolean} predicate
 * @param {AbortSignal | undefined} [signal]
 * @returns {{ r0: number, c0: number, r1: number, c1: number }[]}
 */
function detectRegions(sheet, predicate, signal) {
  throwIfAborted(signal);
  const matrix = getSheetMatrix(sheet);
  if (matrix) {
    /** @type {Map<string, { row: number, col: number }>} */
    const coords = new Map();
    let truncated = false;

    // Treat matrix-backed sheets as sparse: use `for..in` to iterate only defined
    // rows/cols (avoids scanning/allocating for large sparse arrays).
    try {
      for (const rKey in matrix) {
        throwIfAborted(signal);
        const r = Number(rKey);
        if (!Number.isInteger(r) || r < 0) continue;
        const row = matrix[r];
        if (!Array.isArray(row)) continue;
        for (const cKey in row) {
          throwIfAborted(signal);
          const c = Number(cKey);
          if (!Number.isInteger(c) || c < 0) continue;
          const cell = normalizeCell(row[c]);
          if (!predicate(cell)) continue;
          coords.set(`${r},${c}`, { row: r, col: c });
          if (coords.size > DEFAULT_DETECT_REGIONS_CELL_LIMIT) {
            truncated = true;
            break;
          }
        }
        if (truncated) break;
      }
    } catch {
      // Fall back to no regions on unexpected enumerable shapes.
      return [];
    }

    if (coords.size === 0) return [];

    /** @type {Set<string>} */
    const visited = new Set();
    /** @type {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }[]} */
    const components = [];

    const entries = Array.from(coords.values()).sort((a, b) => a.row - b.row || a.col - b.col);
    for (const start of entries) {
      throwIfAborted(signal);
      const startKey = `${start.row},${start.col}`;
      if (visited.has(startKey)) continue;
      visited.add(startKey);
      const stack = [start];
      let r0 = start.row;
      let r1 = start.row;
      let c0 = start.col;
      let c1 = start.col;
      let count = 0;

      while (stack.length) {
        throwIfAborted(signal);
        const cur = stack.pop();
        if (!cur) continue;
        count += 1;
        r0 = Math.min(r0, cur.row);
        r1 = Math.max(r1, cur.row);
        c0 = Math.min(c0, cur.col);
        c1 = Math.max(c1, cur.col);

        const neighbors = [
          { row: cur.row - 1, col: cur.col },
          { row: cur.row + 1, col: cur.col },
          { row: cur.row, col: cur.col - 1 },
          { row: cur.row, col: cur.col + 1 },
        ];
        for (const n of neighbors) {
          const nk = `${n.row},${n.col}`;
          if (!coords.has(nk)) continue;
          if (visited.has(nk)) continue;
          visited.add(nk);
          const entry = coords.get(nk);
          if (entry) stack.push(entry);
        }
      }

      components.push({ rect: { r0, c0, r1, c1 }, count });
    }

    components.sort(
      (a, b) =>
        a.rect.r0 - b.rect.r0 ||
        a.rect.c0 - b.rect.c0 ||
        a.rect.r1 - b.rect.r1 ||
        a.rect.c1 - b.rect.c1
    );

    // Drop trivial single-cell regions (often incidental labels).
    return components.filter((c) => c.count >= 2).map((c) => c.rect);
  }

  const map = getSheetCellMap(sheet);
  if (map) {
    /**
     * @param {string} key
     */
    function parseRowColKey(key) {
      const raw = String(key);
      const delimiter = raw.includes(",") ? "," : raw.includes(":") ? ":" : null;
      if (!delimiter) return null;
      const parts = raw.split(delimiter);
      if (parts.length !== 2) return null;
      const row = Number(parts[0]);
      const col = Number(parts[1]);
      if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) return null;
      return { row, col };
    }

    /** @type {Map<string, { row: number, col: number }>} */
    const coords = new Map();
    for (const [key, raw] of map.entries()) {
      throwIfAborted(signal);
      const parsed = parseRowColKey(key);
      if (!parsed) continue;
      const cell = normalizeCell(raw);
      if (!predicate(cell)) continue;
      coords.set(`${parsed.row},${parsed.col}`, parsed);
    }

    /** @type {Set<string>} */
    const visited = new Set();
    /** @type {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }[]} */
    const components = [];

    const entries = Array.from(coords.values()).sort((a, b) => a.row - b.row || a.col - b.col);
    for (const start of entries) {
      throwIfAborted(signal);
      const startKey = `${start.row},${start.col}`;
      if (visited.has(startKey)) continue;
      visited.add(startKey);
      const stack = [start];
      let r0 = start.row;
      let r1 = start.row;
      let c0 = start.col;
      let c1 = start.col;
      let count = 0;

      while (stack.length) {
        throwIfAborted(signal);
        const cur = stack.pop();
        if (!cur) continue;
        count += 1;
        r0 = Math.min(r0, cur.row);
        r1 = Math.max(r1, cur.row);
        c0 = Math.min(c0, cur.col);
        c1 = Math.max(c1, cur.col);

        const neighbors = [
          { row: cur.row - 1, col: cur.col },
          { row: cur.row + 1, col: cur.col },
          { row: cur.row, col: cur.col - 1 },
          { row: cur.row, col: cur.col + 1 },
        ];
        for (const n of neighbors) {
          const nk = `${n.row},${n.col}`;
          if (!coords.has(nk)) continue;
          if (visited.has(nk)) continue;
          visited.add(nk);
          const entry = coords.get(nk);
          if (entry) stack.push(entry);
        }
      }

      components.push({ rect: { r0, c0, r1, c1 }, count });
    }

    components.sort(
      (a, b) =>
        a.rect.r0 - b.rect.r0 ||
        a.rect.c0 - b.rect.c0 ||
        a.rect.r1 - b.rect.r1 ||
        a.rect.c1 - b.rect.c1
    );

    return components.filter((c) => c.count >= 2).map((c) => c.rect);
  }

  return [];
}

/**
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {{ r0: number, c0: number, r1: number, c1: number }[]} existing
 */
function overlapsExisting(rect, existing) {
  for (const ex of existing) {
    const inter = rectIntersectionArea(rect, ex);
    if (inter === 0) continue;
    const ratio = inter / Math.min(rectSize(rect), rectSize(ex));
    if (ratio > 0.8) return true;
  }
  return false;
}

/**
 * Chunk workbook into semantic regions.
 *
 * Strategy:
 * - Use explicit tables & named ranges first (stable, user-authored).
 * - Detect remaining data regions by connected non-empty cell blocks.
 * - Detect formula-heavy regions by connected formula blocks.
 *
 * @param {import('./workbookTypes').Workbook} workbook
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {import('./workbookTypes').WorkbookChunk[]}
 */
function chunkWorkbook(workbook, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const sheets = sheetMap(workbook);
  /** @type {import('./workbookTypes').WorkbookChunk[]} */
  const chunks = [];

  /** @type {{ sheetName: string, rect: any }[]} */
  const occupied = [];

  for (const table of workbook.tables || []) {
    throwIfAborted(signal);
    const sheet = sheets.get(table.sheetName);
    if (!sheet) continue;
    const id = `${workbook.id}::${table.sheetName}::table::${table.name}`;
    chunks.push({
      id,
      workbookId: workbook.id,
      sheetName: table.sheetName,
      kind: "table",
      title: table.name,
      rect: table.rect,
      cells: extractCells(sheet, table.rect, {
        maxRows: DEFAULT_EXTRACT_MAX_ROWS,
        maxCols: DEFAULT_EXTRACT_MAX_COLS,
      }),
      meta: { tableName: table.name },
    });
    occupied.push({ sheetName: table.sheetName, rect: table.rect });
  }

  for (const nr of workbook.namedRanges || []) {
    throwIfAborted(signal);
    const sheet = sheets.get(nr.sheetName);
    if (!sheet) continue;
    const id = `${workbook.id}::${nr.sheetName}::namedRange::${nr.name}`;
    chunks.push({
      id,
      workbookId: workbook.id,
      sheetName: nr.sheetName,
      kind: "namedRange",
      title: nr.name,
      rect: nr.rect,
      cells: extractCells(sheet, nr.rect, {
        maxRows: DEFAULT_EXTRACT_MAX_ROWS,
        maxCols: DEFAULT_EXTRACT_MAX_COLS,
      }),
      meta: { namedRange: nr.name },
    });
    occupied.push({ sheetName: nr.sheetName, rect: nr.rect });
  }

  for (const sheet of workbook.sheets || []) {
    throwIfAborted(signal);
    const existingRects = occupied
      .filter((o) => o.sheetName === sheet.name)
      .map((o) => o.rect);

    const dataRegions = detectRegions(sheet, isNonEmptyCell, signal).filter(
      (rect) => !overlapsExisting(rect, existingRects)
    );
    for (const rect of dataRegions) {
      throwIfAborted(signal);
      const coordKey = `${rect.r0},${rect.c0},${rect.r1},${rect.c1}`;
      const id = `${workbook.id}::${sheet.name}::dataRegion::${coordKey}`;
      chunks.push({
        id,
        workbookId: workbook.id,
        sheetName: sheet.name,
        kind: "dataRegion",
        title: `Data region ${rectToA1(rect)}`,
        rect,
        cells: extractCells(sheet, rect, {
          maxRows: DEFAULT_EXTRACT_MAX_ROWS,
          maxCols: DEFAULT_EXTRACT_MAX_COLS,
        }),
      });
    }

    const formulaRegions = detectRegions(sheet, isFormulaCell, signal).filter(
      (rect) => !overlapsExisting(rect, existingRects)
    );
    for (const rect of formulaRegions) {
      throwIfAborted(signal);
      const coordKey = `${rect.r0},${rect.c0},${rect.r1},${rect.c1}`;
      const id = `${workbook.id}::${sheet.name}::formulaRegion::${coordKey}`;
      chunks.push({
        id,
        workbookId: workbook.id,
        sheetName: sheet.name,
        kind: "formulaRegion",
        title: `Formula region ${rectToA1(rect)}`,
        rect,
        cells: extractCells(sheet, rect, {
          maxRows: DEFAULT_EXTRACT_MAX_ROWS,
          maxCols: DEFAULT_EXTRACT_MAX_COLS,
        }),
      });
    }
  }

  return chunks;
}

export { chunkWorkbook };
