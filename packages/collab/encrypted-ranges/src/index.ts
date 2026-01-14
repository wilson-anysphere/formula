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

function normalizeSheetNameForCaseInsensitiveCompare(name: string): string {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics used elsewhere in the codebase (desktop SheetNameResolver).
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

function createSheetIdResolverFromWorkbook(doc: Y.Doc): (ref: string) => string | null {
  const sheets = getWorkbookRoots(doc).sheets;
  const idsByCi = new Map<string, string>();
  const idsByNameCi = new Map<string, string>();

  const entries = typeof (sheets as any)?.toArray === "function" ? (sheets as any).toArray() : [];
  for (const entry of entries) {
    const map = getYMap(entry);
    const obj = map ? null : entry && typeof entry === "object" ? (entry as any) : null;
    const get = (k: string): unknown => (map ? map.get(k) : obj ? obj[k] : undefined);
    const id = coerceString(get("id"))?.trim() ?? "";
    if (!id) continue;
    idsByCi.set(id.toLowerCase(), id);
    const name = coerceString(get("name"))?.trim() ?? "";
    if (name) {
      idsByNameCi.set(normalizeSheetNameForCaseInsensitiveCompare(name), id);
    }
  }

  return (ref: string): string | null => {
    const trimmed = String(ref ?? "").trim();
    if (!trimmed) return null;
    const direct = idsByCi.get(trimmed.toLowerCase());
    if (direct) return direct;
    return idsByNameCi.get(normalizeSheetNameForCaseInsensitiveCompare(trimmed)) ?? null;
  };
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

function rangeSignature(range: Pick<EncryptedRange, "sheetId" | "startRow" | "startCol" | "endRow" | "endCol" | "keyId">): string {
  return rangeSignatureWithSheetId(range.sheetId, range);
}

function rangeSignatureWithSheetId(
  sheetId: string,
  range: Pick<EncryptedRange, "startRow" | "startCol" | "endRow" | "endCol" | "keyId">
): string {
  // Use a delimiter that cannot appear in numbers to avoid ambiguous concatenations.
  return `${sheetId}\n${range.startRow},${range.startCol},${range.endRow},${range.endCol}\n${range.keyId}`;
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

    // If the doc contains duplicate encryptedRanges entries (e.g. concurrent inserts),
    // normalize before applying the tracked mutation. Prefer keeping the entry
    // referenced by `normalizedId` so `remove()` reliably removes the intended range.
    this.normalizeEncryptedRangesForUndoScope(normalizedId);
    this.transact(() => {
      const arr = getYArray(this.metadata.get(METADATA_KEY));
      if (!arr) return;

      // Delete back-to-front so indices remain stable when multiple duplicates exist.
      for (let i = arr.length - 1; i >= 0; i -= 1) {
        const entry = yRangeToEncryptedRange(arr.get(i));
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

    // If the doc contains duplicate encryptedRanges entries (e.g. concurrent inserts),
    // normalize before applying the tracked mutation. Prefer keeping the entry
    // referenced by `normalizedId` so updates are never silently dropped.
    this.normalizeEncryptedRangesForUndoScope(normalizedId);
    this.transact(() => {
      const arr = getYArray(this.metadata.get(METADATA_KEY));
      if (!arr) return;

      const len = arr.length;
      for (let i = 0; i < len; i += 1) {
        const yMap = getYMap(arr.get(i));
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
  private normalizeEncryptedRangesForUndoScope(preferId?: string): void {
    const existing = this.metadata.get(METADATA_KEY);
    if (existing == null) return;

    const resolveSheetId = createSheetIdResolverFromWorkbook(this.doc);

    // Fast-path: already the canonical local schema *and* does not require cleanup.
    const existingArr = getYArray(existing);
    if (existingArr && existingArr instanceof Y.Array) {
      const items = existingArr.toArray();
      /** @type {Set<string>} */
      const ids = new Set();
      /** @type {Set<string>} */
      const signatures = new Set();
      let needsNormalize = false;

      for (const item of items) {
        const map = getYMap(item);
        // Foreign constructors need normalization for UndoManager.
        if (!map || !(map instanceof Y.Map)) {
          needsNormalize = true;
          break;
        }

        const parsed = yRangeToEncryptedRange(map);
        if (!parsed) {
          // Malformed entries should be dropped during normalization.
          needsNormalize = true;
          break;
        }

        // Ensure ids are unique.
        if (ids.has(parsed.id)) {
          needsNormalize = true;
          break;
        }
        ids.add(parsed.id);

        // Dedupe identical ranges (can happen after concurrent inserts).
        const canonicalSheetId = resolveSheetId(parsed.sheetId) ?? parsed.sheetId;
        const sig = rangeSignatureWithSheetId(canonicalSheetId, parsed);
        if (signatures.has(sig)) {
          needsNormalize = true;
          break;
        }
        signatures.add(sig);

        // Canonicalize storage types + trims (e.g. avoid Y.Text and stringified numbers).
        // We intentionally only verify required keys; any unknown keys will be dropped during normalization.
        const idVal = map.get("id");
        if (typeof idVal !== "string") needsNormalize = true;
        else if (idVal.trim() !== parsed.id) needsNormalize = true;

        const sheetIdVal = map.get("sheetId");
        if (typeof sheetIdVal !== "string") needsNormalize = true;
        else if (sheetIdVal.trim() !== canonicalSheetId) needsNormalize = true;

        const keyIdVal = map.get("keyId");
        if (typeof keyIdVal !== "string") needsNormalize = true;
        else if (keyIdVal.trim() !== parsed.keyId) needsNormalize = true;

        const startRowVal = map.get("startRow");
        if (typeof startRowVal !== "number" || startRowVal !== parsed.startRow) needsNormalize = true;
        const startColVal = map.get("startCol");
        if (typeof startColVal !== "number" || startColVal !== parsed.startCol) needsNormalize = true;
        const endRowVal = map.get("endRow");
        if (typeof endRowVal !== "number" || endRowVal !== parsed.endRow) needsNormalize = true;
        const endColVal = map.get("endCol");
        if (typeof endColVal !== "number" || endColVal !== parsed.endCol) needsNormalize = true;

        const createdAtVal = map.get("createdAt");
        if (createdAtVal !== undefined) {
          if (typeof createdAtVal !== "number") needsNormalize = true;
          else if (!Number.isFinite(createdAtVal) || createdAtVal < 0) needsNormalize = true;
        }

        const createdByVal = map.get("createdBy");
        if (createdByVal !== undefined && typeof createdByVal !== "string") needsNormalize = true;

        if (needsNormalize) break;
      }

      if (!needsNormalize) return;
    }

    const cloneRangeToLocal = (parsed: EncryptedRange, canonicalSheetId?: string): Y.Map<unknown> => {
      const out = new Y.Map<unknown>();
      out.set("id", parsed.id);
      const resolvedSheetId = canonicalSheetId ?? resolveSheetId(parsed.sheetId) ?? parsed.sheetId;
      out.set("sheetId", resolvedSheetId);
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

      /** @type {Set<string>} */
      const ids = new Set<string>();
      /** @type {Map<string, string>} */
      const idBySignature = new Map<string, string>();
      /** @type {Map<string, number>} */
      const indexBySignature = new Map<string, number>();

      // Build the normalized array in JS first. Some Yjs APIs warn when reading
      // data from unintegrated types (e.g. `yarray.length`). Building in JS avoids
      // that and keeps normalization noise-free.
      /** @type {Array<Y.Map<unknown>>} */
      const out: Array<Y.Map<unknown>> = [];

      const tryPushParsed = (parsed: EncryptedRange) => {
        // Enforce unique ids in the canonical schema.
        if (ids.has(parsed.id)) return;
        const canonicalSheetId = resolveSheetId(parsed.sheetId) ?? parsed.sheetId;
        const sig = rangeSignatureWithSheetId(canonicalSheetId, parsed);

        // Dedupe identical ranges (can happen after concurrent inserts). When
        // `preferId` is provided (e.g. update/remove), prefer keeping that id so
        // the caller's mutation is applied to a surviving entry.
        const existingId = idBySignature.get(sig);
        if (existingId) {
          const preferred = preferId ? String(preferId).trim() : "";
          if (!preferred) return;
          if (existingId === preferred) return;
          if (parsed.id !== preferred) return;

          const idx = indexBySignature.get(sig);
          if (idx == null) return;
          // Replace the previously-kept entry with the preferred id.
          ids.delete(existingId);
          ids.add(parsed.id);
          out[idx] = cloneRangeToLocal(parsed, canonicalSheetId);
          idBySignature.set(sig, parsed.id);
          return;
        }

        ids.add(parsed.id);
        const idx = out.length;
        out.push(cloneRangeToLocal(parsed, canonicalSheetId));
        indexBySignature.set(sig, idx);
        idBySignature.set(sig, parsed.id);
      };

      const pushFrom = (value: unknown, fallbackId?: string) => {
        const parsed = yRangeToEncryptedRange(value, fallbackId);
        if (!parsed) return;
        tryPushParsed(parsed);
      };

      const arr = getYArray(current);
      if (arr) {
        for (const item of arr.toArray()) pushFrom(item);
      } else {
        const map = getYMap(current);
        if (map) {
          const keys = Array.from(map.keys())
            .map((k) => String(k))
            .sort();
          for (const key of keys) {
            pushFrom(map.get(key), key);
          }
        } else if (Array.isArray(current)) {
          for (const item of current) pushFrom(item);
        } else {
          // Unknown schema; do not clobber.
          return;
        }
      }

      const next = new Y.Array<Y.Map<unknown>>();
      if (out.length > 0) next.push(out);
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
  const roots = getWorkbookRoots(doc);
  const metadata = roots.metadata;
  const sheetsRoot = roots.sheets;

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

    const matchesSheet = (rangeSheetId: string): boolean => {
      // Stable sheet id match (case-insensitive for resilience to legacy/case-mismatched ids).
      if (rangeSheetId === sheetId) return true;
      if (rangeSheetId.toLowerCase() === sheetId.toLowerCase()) return true;
      // Legacy support: older clients stored `sheetName` instead of the stable
      // workbook sheet id. Match those entries against the current sheet name.
      sheetName ??= resolveSheetName(sheetId);
      if (!sheetName) return false;
      return (
        normalizeSheetNameForCaseInsensitiveCompare(rangeSheetId) ===
        normalizeSheetNameForCaseInsensitiveCompare(sheetName)
      );
    };

    const raw = metadata.get(METADATA_KEY);

    // Policy precedence: if multiple encrypted ranges overlap, prefer the most recently
    // added entry when the doc stores ranges in an array (canonical schema).
    //
    // This mirrors typical user expectations: adding a new range is treated as "override"
    // for subsequent writes, and the array ordering is deterministic across collaborators
    // in Yjs.
    const scanValue = (value: unknown, fallbackId?: string): EncryptedRange | null => {
      const range = yRangeToEncryptedRange(value, fallbackId);
      if (!range) return null;
      if (!matchesSheet(range.sheetId)) return null;
      if (row < range.startRow || row > range.endRow) return null;
      if (col < range.startCol || col > range.endCol) return null;
      return range;
    };

    const arr = getYArray(raw);
    if (arr) {
      const len = typeof (arr as any).length === "number" ? (arr as any).length : arr.toArray().length;
      for (let i = len - 1; i >= 0; i -= 1) {
        const match = scanValue(arr.get(i));
        if (match) return match;
      }
      return null;
    }

    const map = getYMap(raw);
    if (map) {
      // Legacy schema: encryptedRanges stored as a map keyed by a range id.
      //
      // Deterministic precedence when overlaps exist: prefer the lexicographically greatest key.
      // This is equivalent to sorting keys and scanning in reverse, but avoids allocations and
      // O(n log n) work on each lookup.
      let bestKey: string | null = null;
      let best: EncryptedRange | null = null;
      map.forEach((value, key) => {
        const keyStr = String(key);
        const match = scanValue(value, keyStr);
        if (!match) return;
        if (bestKey == null || keyStr > bestKey) {
          bestKey = keyStr;
          best = match;
        }
      });
      return best;
    }

    if (Array.isArray(raw)) {
      for (let i = raw.length - 1; i >= 0; i -= 1) {
        const match = scanValue(raw[i]);
        if (match) return match;
      }
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
