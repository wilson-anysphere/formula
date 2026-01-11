import type { CellChange, CellData as EngineCellData, CellScalar, EngineClient } from "@formula/engine";
import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";

export interface Range0 {
  startRow0: number;
  endRow0Exclusive: number;
  startCol0: number;
  endCol0Exclusive: number;
}

export function colToName(col0: number): string {
  if (!Number.isSafeInteger(col0) || col0 < 0) {
    throw new Error(`colToName: col0 must be a non-negative safe integer, got ${col0}`);
  }

  let col = col0 + 1;
  let name = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    col = Math.floor((col - 1) / 26);
  }
  return name;
}

export function toA1(row0: number, col0: number): string {
  if (!Number.isSafeInteger(row0) || row0 < 0) {
    throw new Error(`toA1: row0 must be a non-negative safe integer, got ${row0}`);
  }
  return `${colToName(col0)}${row0 + 1}`;
}

export function fromA1(address: string): { row0: number; col0: number } {
  const trimmed = address.trim();
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)$/.exec(trimmed);
  if (!match) {
    throw new Error(`Invalid A1 address: "${address}"`);
  }

  const colLabel = match[1].toUpperCase();
  let col1 = 0;
  for (const ch of colLabel) {
    col1 = col1 * 26 + (ch.charCodeAt(0) - 64);
  }
  const row1 = Number.parseInt(match[2], 10);

  if (!Number.isSafeInteger(row1) || row1 <= 0) {
    throw new Error(`Invalid A1 address row: "${address}"`);
  }
  return { row0: row1 - 1, col0: col1 - 1 };
}

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

function range0ToA1(range: Range0): string {
  if (range.endRow0Exclusive <= range.startRow0 || range.endCol0Exclusive <= range.startCol0) {
    throw new Error(`Invalid range0 (empty): ${JSON.stringify(range)}`);
  }
  const start = toA1(range.startRow0, range.startCol0);
  const end = toA1(range.endRow0Exclusive - 1, range.endCol0Exclusive - 1);
  return start === end ? start : `${start}:${end}`;
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

  async prefetch(range: Range0, sheet?: string): Promise<void> {
    const sheetName = defaultSheetName(sheet);
    const rangeA1 = range0ToA1(range);
    const key = `${sheetName}\n${rangeA1}`;
    const existing = this.inflight.get(key);
    if (existing) return existing;

    const task = (async () => {
      const rows = (await this.engine.getRange(rangeA1, sheetName)) as EngineCellData[][];
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r] ?? [];
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          this.values.set(cacheKey(sheetName, range.startRow0 + r, range.startCol0 + c), normalizeCellValue(cell.value));
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
  private readonly sheet: string;
  private readonly headers: boolean;

  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private pendingInvalidations: CellRange[] = [];
  private flushScheduled = false;

  constructor(options: EngineGridProviderOptions) {
    this.cache = options.cache;
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
    this.sheet = defaultSheetName(options.sheet);
    this.headers = options.headers ?? false;
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

  async prefetch(range: CellRange): Promise<void> {
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

    try {
      await this.cache.prefetch(clamped, this.sheet);
    } catch {
      // Engine fetch failures should not crash the grid; the next scroll/prefetch will retry.
      return;
    }

    const gridRange: CellRange = {
      startRow: clamped.startRow0 + offset,
      endRow: clamped.endRow0Exclusive + offset,
      startCol: clamped.startCol0 + offset,
      endCol: clamped.endCol0Exclusive + offset
    };
    this.queueInvalidation(gridRange);
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
