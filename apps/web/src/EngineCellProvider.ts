import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";

type EngineScalar = string | number | boolean | null;

export interface EngineClientLike {
  getRange(range: string, sheet?: string): Promise<Array<Array<{ value: EngineScalar }>>>;
}

function columnIndexToLetters(index: number): string {
  if (!Number.isFinite(index) || index < 0) return "";
  let n = Math.floor(index);
  let out = "";
  while (n >= 0) {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  }
  return out;
}

function toA1Address(row: number, col: number): string {
  return `${columnIndexToLetters(col)}${row + 1}`;
}

function toGridValue(value: EngineScalar): string | number | null {
  if (value == null) return null;
  if (typeof value === "number" || typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function cacheKey(row: number, col: number): string {
  return `${row},${col}`;
}

export class EngineCellProvider implements CellProvider {
  private readonly engine: EngineClientLike;
  private readonly sheet?: string;
  private readonly rowCount: number;
  private readonly colCount: number;

  private readonly cache = new Map<string, string | number | null>();
  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();

  constructor(options: { engine: EngineClientLike; rowCount: number; colCount: number; sheet?: string }) {
    this.engine = options.engine;
    this.sheet = options.sheet;
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  private emit(update: CellProviderUpdate): void {
    for (const listener of this.listeners) listener(update);
  }

  getCell(row: number, col: number): CellData | null {
    if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) return null;

    const headerStyle: CellStyle = { fill: "#f5f5f5", fontWeight: "600", textAlign: "center" };
    const rowHeaderStyle: CellStyle = { ...headerStyle, textAlign: "end" };

    if (row === 0 && col === 0) return { row, col, value: null, style: headerStyle };
    if (row === 0) return { row, col, value: columnIndexToLetters(col - 1), style: headerStyle };
    if (col === 0) return { row, col, value: row, style: rowHeaderStyle };

    const key = cacheKey(row, col);
    const cached = this.cache.get(key) ?? null;

    const fill = row % 2 === 0 ? "#ffffff" : "#fcfcfc";
    return { row, col, value: cached, style: { fill } };
  }

  async prefetch(range: CellRange): Promise<void> {
    const startRow = Math.max(1, range.startRow);
    const endRow = Math.max(1, range.endRow);
    const startCol = Math.max(1, range.startCol);
    const endCol = Math.max(1, range.endCol);

    // Only data cells (row>=1, col>=1) map to the engine.
    if (startRow >= endRow || startCol >= endCol) return;

    const engineStartRow = startRow - 1;
    const engineStartCol = startCol - 1;
    const engineEndRow = endRow - 2; // inclusive
    const engineEndCol = endCol - 2; // inclusive

    const engineRange = `${toA1Address(engineStartRow, engineStartCol)}:${toA1Address(engineEndRow, engineEndCol)}`;

    let gridStartRow = Number.POSITIVE_INFINITY;
    let gridStartCol = Number.POSITIVE_INFINITY;
    let gridEndRow = -1;
    let gridEndCol = -1;

    try {
      const rows = await this.engine.getRange(engineRange, this.sheet);
      const expectedRows = endRow - startRow;
      const expectedCols = endCol - startCol;

      const rowCount = Math.min(expectedRows, rows.length);
      for (let r = 0; r < rowCount; r++) {
        const cols = rows[r] ?? [];
        const colCount = Math.min(expectedCols, cols.length);
        for (let c = 0; c < colCount; c++) {
          const value = toGridValue(cols[c]?.value ?? null);

          const gridRow = startRow + r;
          const gridCol = startCol + c;
          const key = cacheKey(gridRow, gridCol);

          const previous = this.cache.get(key);
          if (previous === value) continue;

          this.cache.set(key, value);

          gridStartRow = Math.min(gridStartRow, gridRow);
          gridStartCol = Math.min(gridStartCol, gridCol);
          gridEndRow = Math.max(gridEndRow, gridRow);
          gridEndCol = Math.max(gridEndCol, gridCol);
        }
      }
    } catch {
      // Prefetch failures should not crash rendering; missing cells will remain blank.
      return;
    }

    if (gridEndRow === -1 || gridEndCol === -1) return;

    this.emit({
      type: "cells",
      range: {
        startRow: gridStartRow,
        endRow: gridEndRow + 1,
        startCol: gridStartCol,
        endCol: gridEndCol + 1
      }
    });
  }
}

