import type { CellAddress, CollabSession } from "@formula/collab-session";
import { getYArray } from "@formula/collab-yjs-utils";
import * as Y from "yjs";

export const ENCRYPTED_RANGES_METADATA_KEY = "encryptedRanges";

export type EncryptedRange = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  keyId: string;
};

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeNonNegativeInt(value: unknown): number | null {
  const num = Number(value);
  if (!Number.isFinite(num) || !Number.isInteger(num) || num < 0) return null;
  return num;
}

function normalizeEncryptedRange(raw: unknown): EncryptedRange | null {
  // We store ranges as plain objects in Yjs, but tolerate Y.Map-like entries too.
  const obj: any = (() => {
    if (isPlainObject(raw)) return raw;
    if (raw && typeof raw === "object" && typeof (raw as any).get === "function") {
      return {
        sheetId: (raw as any).get("sheetId") ?? (raw as any).get("sheetName"),
        startRow: (raw as any).get("startRow"),
        startCol: (raw as any).get("startCol"),
        endRow: (raw as any).get("endRow"),
        endCol: (raw as any).get("endCol"),
        keyId: (raw as any).get("keyId"),
      };
    }
    return null;
  })();
  if (!obj) return null;

  const sheetId = String(obj.sheetId ?? obj.sheetName ?? "").trim();
  const keyId = String(obj.keyId ?? "").trim();
  const startRow = normalizeNonNegativeInt(obj.startRow);
  const startCol = normalizeNonNegativeInt(obj.startCol);
  const endRow = normalizeNonNegativeInt(obj.endRow);
  const endCol = normalizeNonNegativeInt(obj.endCol);
  if (!sheetId || !keyId) return null;
  if (startRow == null || startCol == null || endRow == null || endCol == null) return null;

  return {
    sheetId,
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endCol: Math.max(startCol, endCol),
    keyId,
  };
}

function cellInRange(cell: CellAddress, range: EncryptedRange): boolean {
  if (cell.sheetId !== range.sheetId) return false;
  if (cell.row < range.startRow || cell.row > range.endRow) return false;
  if (cell.col < range.startCol || cell.col > range.endCol) return false;
  return true;
}

function ensureEncryptedRangesArray(metadata: Y.Map<unknown>): Y.Array<unknown> {
  const existing = metadata.get(ENCRYPTED_RANGES_METADATA_KEY);
  const yarr = getYArray(existing);
  if (yarr) return yarr;

  const next = new Y.Array<unknown>();
  // Best-effort migration from legacy/plain array storage.
  if (Array.isArray(existing)) {
    next.push(existing.map((item) => structuredClone(item)));
  }
  metadata.set(ENCRYPTED_RANGES_METADATA_KEY, next);
  return next;
}

function readEncryptedRanges(metadata: Y.Map<unknown>): EncryptedRange[] {
  const raw = metadata.get(ENCRYPTED_RANGES_METADATA_KEY);
  const yArr = getYArray(raw);
  const entries = yArr ? yArr.toArray() : Array.isArray(raw) ? raw : [];
  const out: EncryptedRange[] = [];
  for (const entry of entries) {
    const range = normalizeEncryptedRange(entry);
    if (range) out.push(range);
  }
  return out;
}

export class EncryptedRangeManager {
  private readonly session: CollabSession;
  private readonly metadata: Y.Map<unknown>;
  private cachedRanges: EncryptedRange[] = [];
  private dirty = true;
  private readonly observer: (events: any[]) => void;

  constructor(opts: { session: CollabSession }) {
    this.session = opts.session;
    this.metadata = opts.session.metadata as Y.Map<unknown>;

    this.observer = (events: any[]) => {
      // Invalidate when the encrypted ranges subtree changes.
      for (const event of events ?? []) {
        const path = event?.path;
        if (Array.isArray(path) && path[0] === ENCRYPTED_RANGES_METADATA_KEY) {
          this.dirty = true;
          return;
        }
        const changes = event?.changes?.keys;
        if (changes && typeof changes.has === "function" && changes.has(ENCRYPTED_RANGES_METADATA_KEY)) {
          this.dirty = true;
          return;
        }
      }
    };

    // observeDeep so nested array edits (push/delete) invalidate the cache.
    this.metadata.observeDeep(this.observer);
  }

  destroy(): void {
    this.metadata.unobserveDeep(this.observer);
  }

  private ensureCached(): void {
    if (!this.dirty) return;
    this.cachedRanges = readEncryptedRanges(this.metadata);
    this.dirty = false;
  }

  listEncryptedRanges(): EncryptedRange[] {
    this.ensureCached();
    // Defensive copy so callers don't mutate internal cache.
    return this.cachedRanges.map((r) => ({ ...r }));
  }

  addEncryptedRange(range: EncryptedRange): void {
    const normalized = normalizeEncryptedRange(range);
    if (!normalized) throw new Error("Invalid encrypted range");

    this.session.transactLocal(() => {
      const arr = ensureEncryptedRangesArray(this.metadata);
      arr.push([structuredClone(normalized)]);
    });

    this.dirty = true;
  }

  findRangeForCell(cell: CellAddress): EncryptedRange | null {
    this.ensureCached();
    // Prefer the most recently added range when overlaps exist.
    for (let i = this.cachedRanges.length - 1; i >= 0; i -= 1) {
      const range = this.cachedRanges[i]!;
      if (cellInRange(cell, range)) return { ...range };
    }
    return null;
  }

  getKeyIdForCell(cell: CellAddress): string | null {
    const range = this.findRangeForCell(cell);
    return range ? range.keyId : null;
  }

  shouldEncryptCell(cell: CellAddress): boolean {
    return Boolean(this.getKeyIdForCell(cell));
  }
}
