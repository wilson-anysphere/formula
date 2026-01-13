import * as Y from "yjs";
import { getWorkbookRoots } from "@formula/collab-workbook";
import { getYArray, getYMap, getYText } from "@formula/collab-yjs-utils";

import { normalizeCellKey } from "../cell-key.js";

export type CellKeyMigrationConflictStrategy = "prefer-canonical" | "prefer-legacy" | "merge";

export type MigrateLegacyCellKeysOptions = {
  /**
   * Sheet id to use when migrating `r{row}c{col}` keys.
   *
   * Defaults to "Sheet1" to match CollabSession schema defaults.
   */
  defaultSheetId?: string;
  /**
   * Yjs origin token to associate with the transaction when applying the migration.
   */
  origin?: unknown;
  /**
   * Collision strategy when both canonical and legacy keys exist and both are plaintext.
   *
   * Defaults to "prefer-canonical".
   */
  conflict?: CellKeyMigrationConflictStrategy;
  /**
   * When true, do not mutate the document; only compute the migration stats.
   */
  dryRun?: boolean;
};

export type MigrateLegacyCellKeysResult = { migrated: number; removed: number; collisions: number };

type CellsMapRead = {
  keys: () => IterableIterator<unknown>;
  get: (key: string) => unknown;
  has: (key: string) => boolean;
};

type YMapLike = {
  get: (key: string) => unknown;
  set: (key: string, value: unknown) => void;
  delete: (key: string) => void;
  forEach: (cb: (value: unknown, key: string) => void) => void;
};

function getYMapLike(value: unknown): YMapLike | null {
  if (!value || typeof value !== "object") return null;
  // Keep the check small; we only rely on the Map-like APIs for cell payloads.
  const maybe = value as any;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.forEach !== "function") return null;
  return maybe as YMapLike;
}

function isEncryptedCellValue(value: unknown): boolean {
  const map = getYMapLike(value);
  if (map) return map.get("enc") !== undefined;
  if (!value || typeof value !== "object") return false;
  return (value as any).enc !== undefined;
}

function deletePlaintextFields(value: unknown): void {
  const map = getYMapLike(value);
  if (map) {
    // For encrypted cells, ensure there is no plaintext payload accessible via
    // `value`/`formula` even if an older client left them behind.
    map.delete("value");
    map.delete("formula");
    return;
  }
  if (!value || typeof value !== "object") return;
  // Best-effort for non-Yjs payloads.
  delete (value as any).value;
  delete (value as any).formula;
}

function cloneYjsValue(value: unknown): unknown {
  const map = getYMap(value);
  if (map) {
    // Always clone to *local* constructors (this module's Yjs instance). Docs can
    // contain nested types created by a different Yjs module instance (CJS vs
    // ESM, or duplicate dependency trees). Those foreign types may not pass
    // `instanceof Y.AbstractType` checks during integration, so copying their
    // constructor would re-introduce "Unexpected content type" crashes.
    const out = new Y.Map();
    map.forEach((v: unknown, k: string) => {
      out.set(k, cloneYjsValue(v));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new Y.Array();
    const items = typeof array.toArray === "function" ? array.toArray() : [];
    for (const item of items) {
      out.push([cloneYjsValue(item)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const out = new Y.Text();
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  if (value && typeof value === "object") return structuredClone(value);
  return value;
}

function cloneCellValue(value: unknown, MapCtor: new () => Y.Map<unknown>): unknown {
  const map = getYMapLike(value);
  if (!map) return cloneYjsValue(value);

  const out = new MapCtor();
  map.forEach((v, k) => {
    out.set(k, cloneYjsValue(v));
  });
  return out;
}

function getCellsMapForDryRun(doc: Y.Doc): CellsMapRead | null {
  const existing = doc.share.get("cells");
  if (!existing) return null;

  // Fast path: already a Map-like root (including foreign module instances).
  const map = getYMap(existing);
  if (map) {
    return map as unknown as CellsMapRead;
  }

  // Slow path: root is a generic `AbstractType` placeholder. Avoid calling
  // `doc.getMap("cells")` since that would mutate the document (and may throw
  // "different constructor" for foreign placeholders). Instead, inspect the
  // internal `_map` structure to read keys/values without instantiating the root.
  const placeholder = existing as any;
  const internalMap = placeholder?._map;
  if (!(internalMap instanceof Map)) return null;

  // If this looks like an Array root, don't try to treat it as a Map.
  const hasStart = placeholder?._start != null;
  if (hasStart && internalMap.size === 0) return null;

  function getValue(key: string): unknown {
    const item = internalMap.get(key);
    if (!item || item.deleted) return undefined;
    const content = item.content?.getContent?.() ?? [];
    return content[content.length - 1];
  }

  function hasKey(key: string): boolean {
    const item = internalMap.get(key);
    return Boolean(item && !item.deleted);
  }

  function* iterKeys(): IterableIterator<string> {
    for (const [key, item] of internalMap.entries()) {
      if (!item || item.deleted) continue;
      yield String(key);
    }
  }

  return {
    keys: () => iterKeys(),
    get: (key) => getValue(key),
    has: (key) => hasKey(key),
  };
}

function runMigrationDry(params: {
  cells: CellsMapRead;
  defaultSheetId: string;
  conflict: CellKeyMigrationConflictStrategy;
}): MigrateLegacyCellKeysResult {
  const { cells, defaultSheetId, conflict } = params;

  /** @type {Map<string, string[]>} */
  const legacyKeysByCanonical = new Map<string, string[]>();
  for (const rawKey of cells.keys()) {
    const key = String(rawKey);
    const canonical = normalizeCellKey(key, { defaultSheetId });
    if (!canonical || canonical === key) continue;
    const list = legacyKeysByCanonical.get(canonical);
    if (list) list.push(key);
    else legacyKeysByCanonical.set(canonical, [key]);
  }

  let migrated = 0;
  let removed = 0;
  let collisions = 0;

  if (legacyKeysByCanonical.size === 0) return { migrated, removed, collisions };

  for (const [canonicalKey, legacyKeysRaw] of legacyKeysByCanonical.entries()) {
    const legacyKeys = legacyKeysRaw.slice().sort();
    const canonicalExists = cells.has(canonicalKey);
    const canonicalValue = canonicalExists ? cells.get(canonicalKey) : undefined;

    const candidateCount = legacyKeys.length + (canonicalExists ? 1 : 0);
    if (candidateCount > 1) collisions += candidateCount - 1;

    const canonicalEncrypted = canonicalExists && isEncryptedCellValue(canonicalValue);
    let legacyEncryptedKey: string | null = null;
    if (!canonicalEncrypted) {
      for (const k of legacyKeys) {
        if (isEncryptedCellValue(cells.get(k))) {
          legacyEncryptedKey = k;
          break;
        }
      }
    }

    const hasEncrypted = canonicalEncrypted || legacyEncryptedKey != null;
    const shouldWriteCanonical = (() => {
      if (hasEncrypted) {
        return !canonicalEncrypted && legacyEncryptedKey != null;
      }
      if (!canonicalExists) return true;
      if (conflict === "prefer-canonical") return false;
      return true; // prefer-legacy or merge
    })();

    if (shouldWriteCanonical) migrated += 1;
    removed += legacyKeys.length;
  }

  return { migrated, removed, collisions };
}

function mergeCellValues(params: {
  canonical: unknown;
  legacies: unknown[];
  MapCtor: new () => Y.Map<unknown>;
}): Y.Map<unknown> {
  const { canonical, legacies, MapCtor } = params;
  const out = new MapCtor();
  const canonicalMap = getYMapLike(canonical);
  if (canonicalMap) {
    canonicalMap.forEach((v, k) => {
      out.set(k, cloneYjsValue(v));
    });
  }

  for (const legacy of legacies) {
    const legacyMap = getYMapLike(legacy);
    if (!legacyMap) continue;
    legacyMap.forEach((v, k) => {
      // "merge" is intentionally conservative: keep canonical values when a field
      // exists, but salvage missing fields from legacy payloads.
      if (out.get(k) === undefined) out.set(k, cloneYjsValue(v));
    });
  }
  return out;
}

function runMigration(params: {
  cells: Y.Map<unknown>;
  defaultSheetId: string;
  conflict: CellKeyMigrationConflictStrategy;
  dryRun: boolean;
}): MigrateLegacyCellKeysResult {
  const { cells, defaultSheetId, conflict, dryRun } = params;
  const MapCtor = cells.constructor as unknown as new () => Y.Map<unknown>;

  /** @type {Map<string, string[]>} */
  const legacyKeysByCanonical = new Map<string, string[]>();
  for (const rawKey of cells.keys()) {
    const key = String(rawKey);
    const canonical = normalizeCellKey(key, { defaultSheetId });
    if (!canonical || canonical === key) continue;
    const list = legacyKeysByCanonical.get(canonical);
    if (list) list.push(key);
    else legacyKeysByCanonical.set(canonical, [key]);
  }

  let migrated = 0;
  let removed = 0;
  let collisions = 0;

  if (legacyKeysByCanonical.size === 0) return { migrated, removed, collisions };

  for (const [canonicalKey, legacyKeysRaw] of legacyKeysByCanonical.entries()) {
    const legacyKeys = legacyKeysRaw.slice().sort();
    const canonicalExists = cells.has(canonicalKey);
    const canonicalValue = canonicalExists ? cells.get(canonicalKey) : undefined;

    const candidateCount = legacyKeys.length + (canonicalExists ? 1 : 0);
    if (candidateCount > 1) collisions += candidateCount - 1;

    // Determine whether any candidate is encrypted; encryption always wins.
    const canonicalEncrypted = canonicalExists && isEncryptedCellValue(canonicalValue);
    let legacyEncryptedKey: string | null = null;
    if (!canonicalEncrypted) {
      for (const k of legacyKeys) {
        if (isEncryptedCellValue(cells.get(k))) {
          legacyEncryptedKey = k;
          break;
        }
      }
    }

    const hasEncrypted = canonicalEncrypted || legacyEncryptedKey != null;

    let nextCanonicalValue: unknown = undefined;
    let shouldWriteCanonical = false;

    if (hasEncrypted) {
      if (canonicalEncrypted) {
        // Canonical is already encrypted; keep it, but ensure it does not retain
        // plaintext fields (defense-in-depth).
        if (!dryRun) deletePlaintextFields(canonicalValue);
      } else if (legacyEncryptedKey) {
        const legacyValue = cells.get(legacyEncryptedKey);
        nextCanonicalValue = cloneCellValue(legacyValue, MapCtor);
        deletePlaintextFields(nextCanonicalValue);
        shouldWriteCanonical = true;
      }
    } else {
      // All plaintext; resolve based on conflict strategy.
      if (!canonicalExists) {
        // No canonical entry exists yet; deterministically migrate the first legacy key.
        const winnerKey = legacyKeys[0];
        nextCanonicalValue = cloneCellValue(cells.get(winnerKey), MapCtor);
        shouldWriteCanonical = true;
      } else if (conflict === "prefer-canonical") {
        // Keep canonical as-is.
      } else if (conflict === "prefer-legacy") {
        const winnerKey = legacyKeys[0];
        nextCanonicalValue = cloneCellValue(cells.get(winnerKey), MapCtor);
        shouldWriteCanonical = true;
      } else {
        // merge
        nextCanonicalValue = mergeCellValues({
          canonical: canonicalValue,
          legacies: legacyKeys.map((k) => cells.get(k)),
          MapCtor,
        });
        shouldWriteCanonical = true;
      }
    }

    if (!dryRun) {
      if (shouldWriteCanonical) {
        cells.set(canonicalKey, nextCanonicalValue);
        migrated += 1;
      }

      for (const legacyKey of legacyKeys) {
        if (!cells.has(legacyKey)) continue;
        cells.delete(legacyKey);
        removed += 1;
      }
    } else {
      if (shouldWriteCanonical) migrated += 1;
      removed += legacyKeys.length;
    }
  }

  return { migrated, removed, collisions };
}

/**
 * Rewrite historical cell key encodings into the canonical `${sheetId}:${row}:${col}` format.
 *
 * This migration is safe to run multiple times (idempotent) and defends against
 * legacy plaintext/encrypted duplication by preferring encrypted payloads when
 * any candidate cell entry has an `enc` field.
 */
export function migrateLegacyCellKeys(doc: Y.Doc, opts: MigrateLegacyCellKeysOptions = {}): MigrateLegacyCellKeysResult {
  const defaultSheetId = opts.defaultSheetId ?? "Sheet1";
  const conflict = opts.conflict ?? "prefer-canonical";
  const dryRun = Boolean(opts.dryRun);

  if (dryRun) {
    const cells = getCellsMapForDryRun(doc);
    if (!cells) return { migrated: 0, removed: 0, collisions: 0 };
    return runMigrationDry({ cells, defaultSheetId, conflict });
  }

  // Avoid instantiating workbook roots when the document doesn't contain the
  // expected schema yet (e.g. a brand new doc). If there's no `cells` root, there
  // are no cell keys to migrate, and creating the root would be an unexpected
  // mutation.
  if (!doc.share.has("cells")) {
    return { migrated: 0, removed: 0, collisions: 0 };
  }

  let result: MigrateLegacyCellKeysResult = { migrated: 0, removed: 0, collisions: 0 };
  doc.transact(
    () => {
      const cells = getWorkbookRoots(doc).cells;
      result = runMigration({ cells, defaultSheetId, conflict, dryRun: false });
    },
    opts.origin ?? "collab-session:migrateLegacyCellKeys",
  );
  return result;
}
