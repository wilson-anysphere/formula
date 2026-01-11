import type { CellChange } from "@formula/engine";
import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";
import { colToName, fromA1, type Range0 } from "./a1";
import { EngineCellCache } from "./cache";

function defaultSheetName(sheet?: string): string {
  return sheet ?? "Sheet1";
}

export interface EngineGridProviderOptions {
  cache: EngineCellCache;
  rowCount: number;
  colCount: number;
  sheet?: string;
  /** If true, reserve row 0 / col 0 for spreadsheet headers and offset engine coordinates by 1. */
  headers?: boolean;
}

export class EngineGridProvider implements CellProvider {
  private readonly cache: EngineCellCache;
  private readonly rowCount: number;
  private readonly colCount: number;
  private sheet: string;
  private readonly headers: boolean;

  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private pendingInvalidations: CellRange[] = [];
  private flushScheduled = false;

  private pendingPrefetchRanges: Range0[] = [];
  private pendingPrefetchResolvers: Array<() => void> = [];
  private prefetchScheduled = false;
  private prefetchInFlight: Promise<void> | null = null;

  constructor(options: EngineGridProviderOptions) {
    this.cache = options.cache;
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
    this.sheet = defaultSheetName(options.sheet);
    this.headers = options.headers ?? false;
  }

  setSheet(sheet: string): void {
    const nextSheet = defaultSheetName(sheet);
    if (nextSheet === this.sheet) return;

    this.sheet = nextSheet;
    // Drop any cached values so we don't briefly show stale cells while the
    // renderer prefetches the new sheet's visible range.
    this.cache.clear();

    if (this.listeners.size === 0) return;
    const update: CellProviderUpdate = { type: "invalidateAll" };
    for (const listener of this.listeners) listener(update);
  }

  getCell(row: number, col: number): CellData | null {
    if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) return null;

    if (this.headers) {
      const headerStyle: CellStyle = { fill: "#f5f5f5", fontWeight: "600", textAlign: "center" };
      const rowHeaderStyle: CellStyle = { ...headerStyle, textAlign: "end" };
      if (row === 0 && col === 0) return { row, col, value: null, style: headerStyle };
      if (row === 0) return { row, col, value: colToName(col - 1), style: headerStyle };
      if (col === 0) return { row, col, value: row, style: rowHeaderStyle };
    }

    const row0 = this.headers ? row - 1 : row;
    const col0 = this.headers ? col - 1 : col;
    if (row0 < 0 || col0 < 0) return null;

    const value = this.cache.getValue(row0, col0, this.sheet);
    const fill = row % 2 === 0 ? "#ffffff" : "#fcfcfc";
    return { row, col, value, style: { fill } };
  }

  prefetch(range: CellRange): void {
    void this.prefetchAsync(range);
  }

  async prefetchAsync(range: CellRange): Promise<void> {
    const offset = this.headers ? 1 : 0;
    const startRow0 = range.startRow - offset;
    const endRow0Exclusive = range.endRow - offset;
    const startCol0 = range.startCol - offset;
    const endCol0Exclusive = range.endCol - offset;

    const clamped: Range0 = {
      startRow0: Math.max(0, startRow0),
      endRow0Exclusive: Math.max(0, endRow0Exclusive),
      startCol0: Math.max(0, startCol0),
      endCol0Exclusive: Math.max(0, endCol0Exclusive)
    };

    if (clamped.endRow0Exclusive <= clamped.startRow0 || clamped.endCol0Exclusive <= clamped.startCol0) return;

    const result = new Promise<void>((resolve) => {
      this.pendingPrefetchRanges.push(clamped);
      this.pendingPrefetchResolvers.push(resolve);
    });

    if (!this.prefetchScheduled) {
      this.prefetchScheduled = true;
      queueMicrotask(() => {
        this.prefetchScheduled = false;
        void this.flushPrefetches();
      });
    }

    return result;
  }

  applyRecalcChanges(changes: CellChange[]): void {
    this.cache.applyRecalcChanges(changes);

    for (const change of changes) {
      if (defaultSheetName(change.sheet) !== this.sheet) continue;
      const { row0, col0 } = fromA1(change.address);
      const row = this.headers ? row0 + 1 : row0;
      const col = this.headers ? col0 + 1 : col0;
      this.queueInvalidation({ startRow: row, endRow: row + 1, startCol: col, endCol: col + 1 });
    }
  }

  async recalculate(sheet?: string): Promise<CellChange[]> {
    const changes = await this.cache.engine.recalculate(sheet ?? this.sheet);
    this.applyRecalcChanges(changes);
    return changes;
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private queueInvalidation(range: CellRange): void {
    if (this.listeners.size === 0) return;

    this.pendingInvalidations.push(range);
    if (this.flushScheduled) return;
    this.flushScheduled = true;

    queueMicrotask(() => this.flushInvalidations());
  }

  private flushInvalidations(): void {
    this.flushScheduled = false;
    const ranges = coalesceRanges(this.pendingInvalidations);
    this.pendingInvalidations = [];

    for (const range of ranges) {
      const update: CellProviderUpdate = { type: "cells", range };
      for (const listener of this.listeners) listener(update);
    }
  }

  private async flushPrefetches(): Promise<void> {
    if (this.prefetchInFlight) return this.prefetchInFlight;

    this.prefetchInFlight = (async () => {
      while (this.pendingPrefetchRanges.length > 0) {
        const ranges = this.pendingPrefetchRanges;
        const resolvers = this.pendingPrefetchResolvers;
        this.pendingPrefetchRanges = [];
        this.pendingPrefetchResolvers = [];

        const merged = mergeRange0(ranges);

        const isCached = this.cache.isRangeCached(merged, this.sheet);
        if (!isCached) {
          try {
            await this.cache.prefetch(merged, this.sheet);
          } catch {
            // Engine fetch failures should not crash the grid; the next scroll/prefetch will retry.
          }

          const offset = this.headers ? 1 : 0;
          this.queueInvalidation({
            startRow: merged.startRow0 + offset,
            endRow: merged.endRow0Exclusive + offset,
            startCol: merged.startCol0 + offset,
            endCol: merged.endCol0Exclusive + offset
          });
        }

        for (const resolve of resolvers) resolve();
      }
    })().finally(() => {
      this.prefetchInFlight = null;
    });

    return this.prefetchInFlight;
  }
}

function mergeRange0(ranges: Range0[]): Range0 {
  let startRow0 = Number.POSITIVE_INFINITY;
  let endRow0Exclusive = 0;
  let startCol0 = Number.POSITIVE_INFINITY;
  let endCol0Exclusive = 0;

  for (const range of ranges) {
    startRow0 = Math.min(startRow0, range.startRow0);
    endRow0Exclusive = Math.max(endRow0Exclusive, range.endRow0Exclusive);
    startCol0 = Math.min(startCol0, range.startCol0);
    endCol0Exclusive = Math.max(endCol0Exclusive, range.endCol0Exclusive);
  }

  if (!Number.isFinite(startRow0) || !Number.isFinite(startCol0)) {
    throw new Error("mergeRange0: expected at least one range");
  }

  return { startRow0, endRow0Exclusive, startCol0, endCol0Exclusive };
}

function coalesceRanges(ranges: CellRange[]): CellRange[] {
  const normalized = ranges
    .filter((r) => r.endRow > r.startRow && r.endCol > r.startCol)
    .map((r) => ({
      startRow: Math.min(r.startRow, r.endRow),
      endRow: Math.max(r.startRow, r.endRow),
      startCol: Math.min(r.startCol, r.endCol),
      endCol: Math.max(r.startCol, r.endCol)
    }));

  let pending = normalized;
  let changed = true;

  while (changed) {
    changed = false;
    const next: CellRange[] = [];

    for (const range of pending) {
      let merged = range;

      for (let i = 0; i < next.length; ) {
        const other = next[i];
        if (canMergeRanges(merged, other)) {
          merged = {
            startRow: Math.min(merged.startRow, other.startRow),
            endRow: Math.max(merged.endRow, other.endRow),
            startCol: Math.min(merged.startCol, other.startCol),
            endCol: Math.max(merged.endCol, other.endCol)
          };
          next.splice(i, 1);
          changed = true;
          continue;
        }
        i++;
      }

      next.push(merged);
    }

    pending = next;
  }

  return pending;
}

function canMergeRanges(a: CellRange, b: CellRange): boolean {
  const rowOverlap = a.startRow < b.endRow && b.startRow < a.endRow;
  const colOverlap = a.startCol < b.endCol && b.startCol < a.endCol;
  const rowNoGap = a.startRow <= b.endRow && b.startRow <= a.endRow;
  const colNoGap = a.startCol <= b.endCol && b.startCol <= a.endCol;

  return (rowOverlap && colNoGap) || (colOverlap && rowNoGap);
}
