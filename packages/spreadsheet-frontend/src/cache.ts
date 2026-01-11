import type { CellChange, CellData as EngineCellData, CellScalar, EngineClient } from "@formula/engine";
import { fromA1, range0ToA1, type Range0 } from "./a1";

function defaultSheetName(sheet?: string): string {
  return sheet ?? "Sheet1";
}

function cacheKey(sheet: string, row0: number, col0: number): string {
  return `${sheet}\n${row0}\n${col0}`;
}

function normalizeCellValue(value: CellScalar): string | number | null {
  if (value === null) return null;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number" || typeof value === "string") return value;
  // Future-proof in case the engine widens its scalar type.
  return String(value);
}

export class EngineCellCache {
  readonly engine: EngineClient;

  private readonly values = new Map<string, string | number | null>();
  private readonly inflight = new Map<string, Promise<void>>();

  constructor(engine: EngineClient) {
    this.engine = engine;
  }

  getValue(row0: number, col0: number, sheet?: string): string | number | null {
    const sheetName = defaultSheetName(sheet);
    return this.values.get(cacheKey(sheetName, row0, col0)) ?? null;
  }

  hasValue(row0: number, col0: number, sheet?: string): boolean {
    const sheetName = defaultSheetName(sheet);
    return this.values.has(cacheKey(sheetName, row0, col0));
  }

  isRangeCached(range: Range0, sheet?: string): boolean {
    const sheetName = defaultSheetName(sheet);

    const rowCount = range.endRow0Exclusive - range.startRow0;
    const colCount = range.endCol0Exclusive - range.startCol0;
    if (rowCount <= 0 || colCount <= 0) return true;

    const cellCount = rowCount * colCount;
    if (this.values.size < cellCount) return false;

    for (let row0 = range.startRow0; row0 < range.endRow0Exclusive; row0++) {
      for (let col0 = range.startCol0; col0 < range.endCol0Exclusive; col0++) {
        if (!this.values.has(cacheKey(sheetName, row0, col0))) return false;
      }
    }
    return true;
  }

  async prefetch(range: Range0, sheet?: string): Promise<void> {
    const sheetName = defaultSheetName(sheet);
    if (this.isRangeCached(range, sheetName)) {
      return;
    }

    const rangeA1 = range0ToA1(range);
    const key = `${sheetName}\n${rangeA1}`;
    const existing = this.inflight.get(key);
    if (existing) return existing;

    const task = (async () => {
      const rows = await this.engine.getRange(rangeA1, sheetName);
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] ?? [];
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          const value = normalizeCellValue(cell.value);
          const cellRow0 = range.startRow0 + r;
          const cellCol0 = range.startCol0 + c;
          this.values.set(cacheKey(sheetName, cellRow0, cellCol0), value);
        }
      }
    })();

    this.inflight.set(key, task);
    try {
      await task;
    } finally {
      this.inflight.delete(key);
    }
  }

  applyRecalcChanges(changes: CellChange[]): void {
    for (const change of changes) {
      const sheet = defaultSheetName(change.sheet);
      const { row0, col0 } = fromA1(change.address);
      this.values.set(cacheKey(sheet, row0, col0), normalizeCellValue(change.value));
    }
  }

  async recalculate(sheet?: string): Promise<CellChange[]> {
    const changes = await this.engine.recalculate(sheet);
    this.applyRecalcChanges(changes);
    return changes;
  }
}
