import * as Y from "yjs";
import { getWorkbookRoots } from "@formula/collab-workbook";
import { getYArray, getYMap, getYText } from "@formula/collab-yjs-utils";

export type EncryptedRange = {
  id: string;
  sheetId: string;
  /**
   * Start row (0-based, inclusive).
   */
  startRow: number;
  /**
   * Start column (0-based, inclusive).
   */
  startCol: number;
  /**
   * End row (0-based, inclusive).
   */
  endRow: number;
  /**
   * End column (0-based, inclusive).
   */
  endCol: number;
  /**
   * Identifier for the encryption key to use for cells in this range.
   *
   * Important: this is metadata only; no secret key material is stored in the Yjs doc.
   */
  keyId: string;
  createdAt?: number;
  createdBy?: string;
};

export type EncryptedRangeAddInput = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  keyId: string;
  createdAt?: number;
  createdBy?: string;
};

export type EncryptedRangeUpdatePatch = Partial<EncryptedRangeAddInput>;

export type WorkbookTransact = (fn: () => void) => void;

const METADATA_KEY = "encryptedRanges";

function coerceString(value: unknown): string | null {
  const text = getYText(value);
  if (text) return text.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

function parseNonNegativeInt(value: unknown, field: string): number {
  const n = typeof value === "number" ? value : typeof value === "string" && value.trim() ? Number(value) : NaN;
  if (!Number.isFinite(n) || !Number.isSafeInteger(n) || Math.floor(n) !== n) {
    throw new Error(`Invalid ${field} (expected non-negative integer): ${String(value)}`);
  }
  if (n < 0) {
    throw new Error(`Invalid ${field} (expected non-negative integer): ${String(value)}`);
  }
  return n;
}

function normalizeId(id: unknown): string {
  const str = String(id ?? "").trim();
  if (!str) throw new Error("Invalid encrypted range id");
  return str;
}

function normalizeSheetId(sheetId: unknown): string {
  const str = String(sheetId ?? "").trim();
  if (!str) throw new Error("Invalid sheetId (expected non-empty string)");
  return str;
}

function normalizeKeyId(keyId: unknown): string {
  const str = String(keyId ?? "").trim();
  if (!str) throw new Error("Invalid keyId (expected non-empty string)");
  return str;
}

function validateRangeOrder(range: {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}): void {
  if (range.startRow > range.endRow) {
    throw new Error(`Invalid encrypted range: startRow (${range.startRow}) > endRow (${range.endRow})`);
  }
  if (range.startCol > range.endCol) {
    throw new Error(`Invalid encrypted range: startCol (${range.startCol}) > endCol (${range.endCol})`);
  }
}

function canonicalizeAddInput(input: EncryptedRangeAddInput): EncryptedRangeAddInput {
  const sheetId = normalizeSheetId(input.sheetId);
  const keyId = normalizeKeyId(input.keyId);
  const startRow = parseNonNegativeInt(input.startRow, "startRow");
  const startCol = parseNonNegativeInt(input.startCol, "startCol");
  const endRow = parseNonNegativeInt(input.endRow, "endRow");
  const endCol = parseNonNegativeInt(input.endCol, "endCol");
  validateRangeOrder({ startRow, startCol, endRow, endCol });

  const createdAt = input.createdAt == null ? undefined : Number(input.createdAt);
  if (createdAt != null && (!Number.isFinite(createdAt) || createdAt < 0)) {
    throw new Error(`Invalid createdAt (expected non-negative number): ${String(input.createdAt)}`);
  }

  const createdByRaw = input.createdBy;
  const createdBy = createdByRaw == null ? undefined : String(createdByRaw).trim() || undefined;

  return { sheetId, startRow, startCol, endRow, endCol, keyId, createdAt, createdBy };
}

function createId(): string {
  const globalCrypto = (globalThis as any).crypto as Crypto | undefined;
  if (globalCrypto?.randomUUID) {
    return globalCrypto.randomUUID();
  }
  return `er_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

function yRangeToEncryptedRange(value: unknown, fallbackId?: string): EncryptedRange | null {
  const map = getYMap(value);
  const obj = map ? null : value && typeof value === "object" ? (value as any) : null;
  const get = (k: string): unknown => (map ? map.get(k) : obj ? obj[k] : undefined);

  const sheetIdRaw = coerceString(get("sheetId")) ?? coerceString(get("sheetName")) ?? coerceString(get("sheet"));
  const keyIdRaw = coerceString(get("keyId"));
  const sheetId = sheetIdRaw?.trim() ?? "";
  const keyId = keyIdRaw?.trim() ?? "";
  if (!sheetId || !keyId) return null;

  const startRow = typeof get("startRow") === "number" ? (get("startRow") as number) : Number(get("startRow"));
  const startCol = typeof get("startCol") === "number" ? (get("startCol") as number) : Number(get("startCol"));
  const endRow = typeof get("endRow") === "number" ? (get("endRow") as number) : Number(get("endRow"));
  const endCol = typeof get("endCol") === "number" ? (get("endCol") as number) : Number(get("endCol"));

  if (![startRow, startCol, endRow, endCol].every((n) => Number.isFinite(n) && Math.floor(n) === n && n >= 0)) {
    return null;
  }
  if (startRow > endRow || startCol > endCol) return null;

  // Legacy support: older clients stored encryptedRanges entries without an `id` field
  // (plain objects in a Y.Array). Derive a deterministic id from the range fields so:
  // - policy helpers can still find these ranges, and
  // - migrations that rewrite encryptedRanges into canonical Y.Maps don't drop them.
  //
  // This also intentionally dedupes identical legacy ranges.
  const idRaw = coerceString(get("id"))?.trim() ?? "";
  const idFromKey = String(fallbackId ?? "").trim();
  const id = idRaw || idFromKey || `legacy:${sheetId}:${startRow}:${startCol}:${endRow}:${endCol}:${keyId}`;
  if (!id) return null;

  const createdAtRaw = get("createdAt");
  const createdAtNum =
    typeof createdAtRaw === "number"
      ? createdAtRaw
      : typeof createdAtRaw === "string" && createdAtRaw.trim()
        ? Number(createdAtRaw)
        : undefined;
  const createdAt =
    createdAtNum != null && Number.isFinite(createdAtNum) && createdAtNum >= 0 ? createdAtNum : undefined;

  const createdByRaw = get("createdBy");
  const createdBy = typeof createdByRaw === "string" ? createdByRaw : createdByRaw != null ? String(createdByRaw) : undefined;
  const createdByTrimmed = createdBy?.trim() || undefined;

  return {
    id,
    sheetId,
    startRow,
    startCol,
    endRow,
    endCol,
    keyId,
    ...(createdAt != null ? { createdAt } : {}),
    ...(createdByTrimmed ? { createdBy: createdByTrimmed } : {}),
  };
}

function isSameRange(a: EncryptedRange, b: EncryptedRangeAddInput): boolean {
  return (
    a.sheetId === b.sheetId &&
    a.startRow === b.startRow &&
    a.startCol === b.startCol &&
    a.endRow === b.endRow &&
    a.endCol === b.endCol &&
    a.keyId === b.keyId
  );
}

export class EncryptedRangeManager {
  private readonly doc: Y.Doc;
  private readonly metadata: Y.Map<unknown>;
  private readonly transact: WorkbookTransact;

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    if (!opts?.doc) throw new Error("EncryptedRangeManager requires { doc }");
    this.doc = opts.doc;
    // Use workbook root helpers so we tolerate mixed-module docs (ESM/CJS Yjs).
    this.metadata = getWorkbookRoots(opts.doc).metadata;
    this.transact = opts.transact ?? ((fn) => opts.doc.transact(fn));
  }

  list(): EncryptedRange[] {
    const raw = this.metadata.get(METADATA_KEY);

    const rangesById = new Map<string, EncryptedRange>();

    const addValue = (value: unknown, fallbackId?: string) => {
      const parsed = yRangeToEncryptedRange(value, fallbackId);
      if (!parsed) return;
      rangesById.set(parsed.id, parsed);
    };

    const arr = getYArray(raw);
    if (arr) {
      for (const item of arr.toArray()) addValue(item);
    } else {
      const map = getYMap(raw);
      if (map) {
        map.forEach((value, key) => addValue(value, String(key)));
      } else if (Array.isArray(raw)) {
        for (const item of raw) addValue(item);
      }
    }

    const out = Array.from(rangesById.values());
    // Deterministic ordering across clients regardless of insertion order / concurrency.
    out.sort((a, b) => a.id.localeCompare(b.id));
    return out;
  }

  add(range: EncryptedRangeAddInput): string {
    // Normalize foreign nested Yjs types (ESM/CJS) before we start an undo-tracked
    // transaction so collaborative undo only captures the user's change.
    this.normalizeEncryptedRangesForUndoScope();
    const canonical = canonicalizeAddInput(range);

    let outId: string | null = null;
    this.transact(() => {
      const arr = this.ensureEncryptedRangesArrayForWrite();

      // Deduplicate identical ranges.
      for (const existing of this.list()) {
        if (isSameRange(existing, canonical)) {
          outId = existing.id;
          return;
        }
      }

      const id = createId();
      const yRange = new Y.Map<unknown>();
      yRange.set("id", id);
      yRange.set("sheetId", canonical.sheetId);
      yRange.set("startRow", canonical.startRow);
      yRange.set("startCol", canonical.startCol);
      yRange.set("endRow", canonical.endRow);
      yRange.set("endCol", canonical.endCol);
      yRange.set("keyId", canonical.keyId);
      if (canonical.createdAt != null) yRange.set("createdAt", canonical.createdAt);
      if (canonical.createdBy != null) yRange.set("createdBy", canonical.createdBy);

      arr.push([yRange]);
      outId = id;
    });

    if (!outId) throw new Error("Failed to add encrypted range");
    return outId;
  }

  remove(id: string): void {
    const normalizedId = normalizeId(id);

    this.normalizeEncryptedRangesForUndoScope();
    this.transact(() => {
      const arr = getYArray(this.metadata.get(METADATA_KEY));
      if (!arr) return;

      // Delete back-to-front so indices remain stable when multiple duplicates exist.
      const items = arr.toArray();
      for (let i = items.length - 1; i >= 0; i -= 1) {
        const entry = yRangeToEncryptedRange(items[i]);
        if (entry?.id === normalizedId) {
          arr.delete(i, 1);
        }
      }
    });
  }

  update(id: string, patch: EncryptedRangeUpdatePatch): void {
    const normalizedId = normalizeId(id);
    const patchSheetId = patch.sheetId == null ? undefined : normalizeSheetId(patch.sheetId);
    const patchKeyId = patch.keyId == null ? undefined : normalizeKeyId(patch.keyId);

    const patchStartRow = patch.startRow == null ? undefined : parseNonNegativeInt(patch.startRow, "startRow");
    const patchStartCol = patch.startCol == null ? undefined : parseNonNegativeInt(patch.startCol, "startCol");
    const patchEndRow = patch.endRow == null ? undefined : parseNonNegativeInt(patch.endRow, "endRow");
    const patchEndCol = patch.endCol == null ? undefined : parseNonNegativeInt(patch.endCol, "endCol");

    const patchCreatedAt = patch.createdAt == null ? undefined : Number(patch.createdAt);
    if (patchCreatedAt != null && (!Number.isFinite(patchCreatedAt) || patchCreatedAt < 0)) {
      throw new Error(`Invalid createdAt (expected non-negative number): ${String(patch.createdAt)}`);
    }

    const patchCreatedBy = patch.createdBy == null ? undefined : String(patch.createdBy).trim() || undefined;

    this.normalizeEncryptedRangesForUndoScope();
    this.transact(() => {
      const arr = getYArray(this.metadata.get(METADATA_KEY));
      if (!arr) return;

      const items = arr.toArray();
      for (let i = 0; i < items.length; i += 1) {
        const yMap = getYMap(items[i]);
        if (!yMap) continue;
        const entryIdRaw = coerceString(yMap.get("id"))?.trim() ?? "";
        // Fast-path: if the range has an id and it doesn't match, skip without parsing.
        if (entryIdRaw && entryIdRaw !== normalizedId) continue;

        const existing = yRangeToEncryptedRange(yMap);
        if (!existing) {
          // If a row has an explicit id but is missing required fields, treat it as corrupt
          // rather than silently ignoring the update.
          if (entryIdRaw === normalizedId) {
            throw new Error(`Encrypted range is missing required fields: ${normalizedId}`);
          }
          continue;
        }

        // Legacy support: tolerate entries without an explicit `id` field by matching against
        // the derived id returned by `yRangeToEncryptedRange` (which may be `legacy:...`).
        if (existing.id !== normalizedId) continue;

        // Persist the derived id so future updates/removes can reference a stable identifier
        // even if the range is later resized (which would change the legacy derived id).
        if (!entryIdRaw) {
          yMap.set("id", normalizedId);
        }

        const next: EncryptedRangeAddInput = {
          sheetId: patchSheetId ?? existing.sheetId,
          startRow: patchStartRow ?? existing.startRow,
          startCol: patchStartCol ?? existing.startCol,
          endRow: patchEndRow ?? existing.endRow,
          endCol: patchEndCol ?? existing.endCol,
          keyId: patchKeyId ?? existing.keyId,
          createdAt: patchCreatedAt ?? existing.createdAt,
          createdBy: patchCreatedBy ?? existing.createdBy,
        };
        // Throws if invalid.
        const canonical = canonicalizeAddInput(next);

        if (patchSheetId != null) yMap.set("sheetId", canonical.sheetId);
        if (patchStartRow != null) yMap.set("startRow", canonical.startRow);
        if (patchStartCol != null) yMap.set("startCol", canonical.startCol);
        if (patchEndRow != null) yMap.set("endRow", canonical.endRow);
        if (patchEndCol != null) yMap.set("endCol", canonical.endCol);
        if (patchKeyId != null) yMap.set("keyId", canonical.keyId);
        if (patchCreatedAt != null) yMap.set("createdAt", canonical.createdAt);
        if (patchCreatedBy != null) yMap.set("createdBy", canonical.createdBy);
      }
    });
  }

  /**
   * Normalize `metadata.encryptedRanges` to the canonical schema (local `Y.Array`
   * containing local `Y.Map` entries).
   *
   * This is needed when a doc was hydrated using a different `yjs` module
   * instance (ESM vs CJS), which can leave nested types with foreign constructors.
   * UndoManager relies on `instanceof` checks, so we normalize in an *untracked*
   * transaction before applying user edits.
   */
  private normalizeEncryptedRangesForUndoScope(): void {
    const existing = this.metadata.get(METADATA_KEY);
    if (existing == null) return;

    // Fast-path: already the canonical local schema.
    const existingArr = getYArray(existing);
    if (existingArr && existingArr instanceof Y.Array) {
      const items = existingArr.toArray();
      let allLocal = true;
      for (const item of items) {
        const map = getYMap(item);
        if (!map || !(map instanceof Y.Map)) {
          allLocal = false;
          break;
        }
        // Canonical entries must include a stable `id` field. Older/buggy clients may
        // omit it even when using a local Y.Map, so treat missing ids as non-canonical
        // and normalize below.
        const id = coerceString(map.get("id"))?.trim();
        if (!id) {
          allLocal = false;
          break;
        }
      }
      if (allLocal) return;
    }

    const cloneEntryToLocal = (value: unknown, fallbackId?: string): Y.Map<unknown> | null => {
      const parsed = yRangeToEncryptedRange(value, fallbackId);
      if (!parsed) return null;

      const out = new Y.Map<unknown>();
      out.set("id", parsed.id);
      out.set("sheetId", parsed.sheetId);
      out.set("startRow", parsed.startRow);
      out.set("startCol", parsed.startCol);
      out.set("endRow", parsed.endRow);
      out.set("endCol", parsed.endCol);
      out.set("keyId", parsed.keyId);
      if (parsed.createdAt != null) out.set("createdAt", parsed.createdAt);
      if (parsed.createdBy != null) out.set("createdBy", parsed.createdBy);
      return out;
    };

    // Normalize in an untracked transaction (origin is undefined) so collaborative
    // undo only captures the user's explicit edits.
    this.doc.transact(() => {
      const current = this.metadata.get(METADATA_KEY);
      if (current == null) return;

      const next = new Y.Array<Y.Map<unknown>>();

      const pushFrom = (value: unknown, fallbackId?: string) => {
        const cloned = cloneEntryToLocal(value, fallbackId);
        if (!cloned) return;
        next.push([cloned]);
      };

      const arr = getYArray(current);
      if (arr) {
        for (const item of arr.toArray()) pushFrom(item);
      } else {
        const map = getYMap(current);
        if (map) {
          map.forEach((value, key) => {
            pushFrom(value, String(key));
          });
        } else if (Array.isArray(current)) {
          for (const item of current) pushFrom(item);
        } else {
          // Unknown schema; do not clobber.
          return;
        }
      }

      this.metadata.set(METADATA_KEY, next);
    });
  }

  private ensureEncryptedRangesArrayForWrite(): Y.Array<Y.Map<unknown>> {
    const existing = this.metadata.get(METADATA_KEY);
    const arr = getYArray(existing);

    // Already the canonical schema. Prefer keeping it.
    if (arr && arr instanceof Y.Array) return arr as Y.Array<Y.Map<unknown>>;

    // If the doc is hydrated by a different Yjs build (ESM vs CJS), nested arrays
    // can fail `instanceof` checks. Rather than mixing constructors (which can
    // throw inside Yjs), migrate the array to local types.
    const next = new Y.Array<Y.Map<unknown>>();

    const cloneEntry = (value: unknown, fallbackId?: string) => {
      const parsed = yRangeToEncryptedRange(value, fallbackId);
      if (!parsed) return;
      const yRange = new Y.Map<unknown>();
      yRange.set("id", parsed.id);
      yRange.set("sheetId", parsed.sheetId);
      yRange.set("startRow", parsed.startRow);
      yRange.set("startCol", parsed.startCol);
      yRange.set("endRow", parsed.endRow);
      yRange.set("endCol", parsed.endCol);
      yRange.set("keyId", parsed.keyId);
      if (parsed.createdAt != null) yRange.set("createdAt", parsed.createdAt);
      if (parsed.createdBy != null) yRange.set("createdBy", parsed.createdBy);
      next.push([yRange]);
    };

    if (arr) {
      for (const item of arr.toArray()) cloneEntry(item);
    } else {
      const map = getYMap(existing);
      if (map) {
        map.forEach((value, key) => cloneEntry(value, String(key)));
      } else if (Array.isArray(existing)) {
        for (const item of existing) cloneEntry(item);
      }
    }

    this.metadata.set(METADATA_KEY, next);
    return next;
  }
}

export function createEncryptedRangeManagerForSession(session: {
  doc: Y.Doc;
  transactLocal: (fn: () => void) => void;
}): EncryptedRangeManager {
  // Be careful to preserve the caller's `this` binding (transactLocal is usually a method).
  return new EncryptedRangeManager({ doc: session.doc, transact: (fn) => session.transactLocal(fn) });
}

export function createEncryptionPolicyFromDoc(doc: Y.Doc): {
  shouldEncryptCell(cell: { sheetId: string; row: number; col: number }): boolean;
  keyIdForCell(cell: { sheetId: string; row: number; col: number }): string | null;
} {
  const mgr = new EncryptedRangeManager({ doc });
  const sheetsRoot = getWorkbookRoots(doc).sheets;

  function normalizeCell(cell: { sheetId: string; row: number; col: number }): { sheetId: string; row: number; col: number } | null {
    const sheetId = String(cell?.sheetId ?? "").trim();
    const row = Number(cell?.row);
    const col = Number(cell?.col);
    if (!sheetId) return null;
    if (!Number.isFinite(row) || Math.floor(row) !== row || row < 0) return null;
    if (!Number.isFinite(col) || Math.floor(col) !== col || col < 0) return null;
    return { sheetId, row, col };
  }

  function resolveSheetName(sheetId: string): string | null {
    try {
      const entries = typeof (sheetsRoot as any)?.toArray === "function" ? (sheetsRoot as any).toArray() : [];
      for (const entry of entries) {
        const map = getYMap(entry);
        const obj = map ? null : entry && typeof entry === "object" ? (entry as any) : null;
        const get = (k: string): unknown => (map ? map.get(k) : obj ? obj[k] : undefined);
        const entryId = coerceString(get("id"))?.trim() ?? "";
        if (!entryId || entryId !== sheetId) continue;
        const name = coerceString(get("name"))?.trim() ?? "";
        return name || null;
      }
    } catch {
      // Best-effort; fall back to matching by id only.
    }
    return null;
  }

  function findMatch(cell: { sheetId: string; row: number; col: number }): EncryptedRange | null {
    const normalized = normalizeCell(cell);
    if (!normalized) return null;

    const { sheetId, row, col } = normalized;
    let sheetName: string | null = null;
    for (const range of mgr.list()) {
      if (range.sheetId !== sheetId) {
        // Legacy support: older clients stored `sheetName` instead of the stable
        // workbook sheet id. Match those entries against the current sheet name.
        sheetName ??= resolveSheetName(sheetId);
        if (!sheetName || range.sheetId !== sheetName) continue;
      }
      if (row < range.startRow || row > range.endRow) continue;
      if (col < range.startCol || col > range.endCol) continue;
      return range;
    }
    return null;
  }

  return {
    shouldEncryptCell(cell): boolean {
      return findMatch(cell) != null;
    },
    keyIdForCell(cell): string | null {
      return findMatch(cell)?.keyId ?? null;
    },
  };
}
