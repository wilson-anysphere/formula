import type { CellChange, CellData as EngineCellData, CellScalar, EngineClient } from "@formula/engine";
import { fromA1, range0ToA1, type Range0 } from "./a1.ts";

function defaultSheetName(sheet?: string): string {
  return sheet ?? "Sheet1";
}

function cacheKey(sheet: string, row0: number, col0: number): string {
  return `${sheet}\n${row0}\n${col0}`;
}

function normalizeCellValue(value: CellScalar): string | number | boolean | null {
  if (value === null) return null;
  if (typeof value === "boolean" || typeof value === "number" || typeof value === "string") return value;
  // Future-proof in case the engine widens its scalar type.
  return String(value);
}

export interface EngineCellCacheOptions {
  /**
   * Maximum number of cached cell entries (including cached empty cells).
   *
   * Prefetching a large sparse sheet can otherwise grow the cache without bound as
   * the user scrolls around. When the limit is exceeded, the oldest entries are
   * evicted (insertion order).
   */
  maxEntries?: number;
}

export class EngineCellCache {
  readonly engine: EngineClient;

  private readonly values = new Map<string, string | number | boolean | null>();
  private readonly inflight = new Map<string, Promise<void>>();
  private readonly maxEntries: number;
  private generation = 0;

  constructor(engine: EngineClient, options?: EngineCellCacheOptions) {
    this.engine = engine;
    const maxEntries = options?.maxEntries ?? 200_000;
    if (!Number.isSafeInteger(maxEntries) || maxEntries <= 0) {
      throw new Error(`EngineCellCache: maxEntries must be a positive safe integer, got ${maxEntries}`);
    }
    this.maxEntries = maxEntries;
  }

  clear(): void {
    this.generation += 1;
    this.values.clear();
    this.inflight.clear();
  }

  getValue(row0: number, col0: number, sheet?: string): string | number | boolean | null {
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

    const generation = this.generation;
    const task = (async () => {
      const rows = await this.engine.getRange(rangeA1, sheetName);
      if (generation !== this.generation) {
        return;
      }
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] ?? [];
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          const value = normalizeCellValue(cell.value);
          const cellRow0 = range.startRow0 + r;
          const cellCol0 = range.startCol0 + c;
          this.setValue(cacheKey(sheetName, cellRow0, cellCol0), value);
        }
      }
      this.trim();
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
      this.setValue(cacheKey(sheet, row0, col0), normalizeCellValue(change.value));
    }
    this.trim();
  }

  async recalculate(sheet?: string): Promise<CellChange[]> {
    const changes = await this.engine.recalculate(sheet);
    this.applyRecalcChanges(changes);
    return changes;
  }

  private setValue(key: string, value: string | number | boolean | null): void {
    // Refresh insertion order when updating an existing key so eviction behaves
    // like an LRU-by-prefetch cache.
    if (this.values.has(key)) {
      this.values.delete(key);
    }
    this.values.set(key, value);
  }

  private trim(): void {
    while (this.values.size > this.maxEntries) {
      const firstKey = this.values.keys().next().value as string | undefined;
      if (!firstKey) break;
      this.values.delete(firstKey);
    }
  }
}
