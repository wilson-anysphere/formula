import type { CellChange } from "@formula/engine";
import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";
import { colToName, fromA1, type Range0 } from "./a1.ts";
import { EngineCellCache } from "./cache.ts";

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
      const headerStyle: CellStyle = { fontWeight: "600", textAlign: "center" };
      const rowHeaderStyle: CellStyle = { ...headerStyle, textAlign: "end" };
      if (row === 0 && col === 0) return { row, col, value: null, style: headerStyle };
      if (row === 0) return { row, col, value: colToName(col - 1), style: headerStyle };
      if (col === 0) return { row, col, value: row, style: rowHeaderStyle };
    }

    const row0 = this.headers ? row - 1 : row;
    const col0 = this.headers ? col - 1 : col;
    if (row0 < 0 || col0 < 0) return null;

    const value = this.cache.getValue(row0, col0, this.sheet);
    return { row, col, value };
  }

  prefetch(range: CellRange): void {
    const promise = this.prefetchAsync(range);
    // Prefetch is best-effort and often fire-and-forgotten by renderers. Attach a no-op rejection
    // handler so unexpected errors never surface as unhandled rejections.
    void promise.catch(() => {});
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

    const relevant = changes.filter((change) => defaultSheetName(change.sheet) === this.sheet);
    if (relevant.length === 0) return;

    const offset = this.headers ? 1 : 0;

    // For very large recalc batches, invalidate a single bounding box rather than
    // enqueueing thousands of 1x1 ranges and running an O(n^2) coalescer.
    if (relevant.length > 256) {
      let minRow0 = Number.POSITIVE_INFINITY;
      let maxRow0 = -1;
      let minCol0 = Number.POSITIVE_INFINITY;
      let maxCol0 = -1;

      for (const change of relevant) {
        const { row0, col0 } = fromA1(change.address);
        minRow0 = Math.min(minRow0, row0);
        maxRow0 = Math.max(maxRow0, row0);
        minCol0 = Math.min(minCol0, col0);
        maxCol0 = Math.max(maxCol0, col0);
      }

      if (minRow0 !== Number.POSITIVE_INFINITY && minCol0 !== Number.POSITIVE_INFINITY && maxRow0 >= 0 && maxCol0 >= 0) {
        this.queueInvalidation({
          startRow: minRow0 + offset,
          endRow: maxRow0 + offset + 1,
          startCol: minCol0 + offset,
          endCol: maxCol0 + offset + 1
        });
      }
      return;
    }

    for (const change of relevant) {
      const { row0, col0 } = fromA1(change.address);
      this.queueInvalidation({
        startRow: row0 + offset,
        endRow: row0 + offset + 1,
        startCol: col0 + offset,
        endCol: col0 + offset + 1
      });
    }
  }

  recalculate(sheet?: string): Promise<CellChange[]> {
    const promise = (async () => {
      const changes = await this.cache.engine.recalculate(sheet ?? this.sheet);
      this.applyRecalcChanges(changes);
      return changes;
    })();

    // Recalc is sometimes fired-and-forgotten (e.g. demo UI sheet switch). Attach a no-op
    // rejection handler so a failing recalc doesn't surface as an unhandled rejection when
    // callers don't await.
    //
    // Awaiting the returned promise still observes the rejection.
    void promise.catch(() => {});

    return promise;
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

        try {
          const coalesced = coalesceRange0(ranges);
          const offset = this.headers ? 1 : 0;

          for (const range of coalesced) {
            const isCached = this.cache.isRangeCached(range, this.sheet);
            if (isCached) continue;

            try {
              await this.cache.prefetch(range, this.sheet);
            } catch {
              // Engine fetch failures should not crash the grid; the next scroll/prefetch will retry.
            }

            this.queueInvalidation({
              startRow: range.startRow0 + offset,
              endRow: range.endRow0Exclusive + offset,
              startCol: range.startCol0 + offset,
              endCol: range.endCol0Exclusive + offset
            });
          }
        } catch {
          // Best-effort: prefetch is an optimization. Never allow unexpected errors to wedge awaiting
          // callers or surface as unhandled rejections from a fire-and-forget prefetch path.
        } finally {
          for (const resolve of resolvers) {
            try {
              resolve();
            } catch {
              // ignore
            }
          }
        }
      }
    })().finally(() => {
      this.prefetchInFlight = null;
    });

    return this.prefetchInFlight;
  }
}

function coalesceRange0(ranges: Range0[]): Range0[] {
  const normalized = ranges
    .filter((r) => r.endRow0Exclusive > r.startRow0 && r.endCol0Exclusive > r.startCol0)
    .map((r) => ({
      startRow0: Math.min(r.startRow0, r.endRow0Exclusive),
      endRow0Exclusive: Math.max(r.startRow0, r.endRow0Exclusive),
      startCol0: Math.min(r.startCol0, r.endCol0Exclusive),
      endCol0Exclusive: Math.max(r.startCol0, r.endCol0Exclusive)
    }));

  let pending = normalized;
  let changed = true;

  while (changed) {
    changed = false;
    const next: Range0[] = [];

    for (const range of pending) {
      let merged = range;

      for (let i = 0; i < next.length; ) {
        const other = next[i];
        if (canMergeRange0(merged, other)) {
          merged = {
            startRow0: Math.min(merged.startRow0, other.startRow0),
            endRow0Exclusive: Math.max(merged.endRow0Exclusive, other.endRow0Exclusive),
            startCol0: Math.min(merged.startCol0, other.startCol0),
            endCol0Exclusive: Math.max(merged.endCol0Exclusive, other.endCol0Exclusive)
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

function canMergeRange0(a: Range0, b: Range0): boolean {
  const rowOverlap = a.startRow0 < b.endRow0Exclusive && b.startRow0 < a.endRow0Exclusive;
  const colOverlap = a.startCol0 < b.endCol0Exclusive && b.startCol0 < a.endCol0Exclusive;
  const rowNoGap = a.startRow0 <= b.endRow0Exclusive && b.startRow0 <= a.endRow0Exclusive;
  const colNoGap = a.startCol0 <= b.endCol0Exclusive && b.startCol0 <= a.endCol0Exclusive;

  return (rowOverlap && colNoGap) || (colOverlap && rowNoGap);
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
