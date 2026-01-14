import { throwIfAborted } from "./abort.js";

/**
 * Return the unambiguous id prefix used for sheet-level region chunks.
 *
 * The legacy `${sheetName}-region-...` ids can collide when sheet names share a prefix
 * (e.g. "Sales" vs "Sales-region-2024"). Length-prefixing prevents this, even when
 * sheet names contain punctuation or substrings like "-region-".
 *
 * @param {string} sheetName
 */
export function sheetChunkIdPrefix(sheetName) {
  const name = typeof sheetName === "string" ? sheetName : String(sheetName ?? "");
  return `sheet:${name.length}:${name}:region:`;
}

/**
 * Legacy id prefix used by older releases.
 * @param {string} sheetName
 */
export function legacySheetChunkIdPrefix(sheetName) {
  const name = typeof sheetName === "string" ? sheetName : String(sheetName ?? "");
  return `${name}-region-`;
}

/**
 * Delete all region chunks for a given sheet name, using the current id scheme.
 *
 * Falls back to scanning `store.items` when the store does not implement
 * `deleteByPrefix` (common in simple in-memory store shapes).
 *
 * Also performs a best-effort cleanup of legacy ids.
 *
 * @param {any} store
 * @param {string} sheetName
 * @param {{ signal?: AbortSignal }} [options]
 */
export async function deleteSheetRegionChunks(store, sheetName, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);

  if (!store || typeof store !== "object") return;

  const prefix = sheetChunkIdPrefix(sheetName);

  if (typeof store.deleteByPrefix === "function") {
    await store.deleteByPrefix(prefix, { signal });
  } else {
    const items = store.items;
    if (items && typeof items.keys === "function" && typeof items.delete === "function") {
      for (const id of Array.from(items.keys())) {
        throwIfAborted(signal);
        if (typeof id === "string" && id.startsWith(prefix)) items.delete(id);
      }
    }
  }

  deleteLegacySheetRegionChunks(store, sheetName, { signal });
}

/**
 * Best-effort cleanup for legacy `${sheetName}-region-...` ids.
 *
 * We avoid `deleteByPrefix(legacySheetChunkIdPrefix(sheetName))` because that prefix is
 * ambiguous (and is the root cause of this bug). Instead, when the vector store exposes
 * a mutable `items: Map<string, ...>` (the in-memory store shape used by ai-context),
 * we delete only ids that match the legacy *chunk id* pattern for this exact sheet name.
 *
 * @param {any} store
 * @param {string} sheetName
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {number} number of deleted items
 */
export function deleteLegacySheetRegionChunks(store, sheetName, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);

  if (!store || typeof store !== "object") return 0;
  const items = store.items;
  if (!items || typeof items.keys !== "function" || typeof items.delete !== "function") return 0;

  const escaped = String(sheetName ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`^${escaped}-region-(\\d+)(?:-o\\d+x\\d+)?(?:-rows-\\d+)?$`);

  let deleted = 0;
  // `Map.prototype.keys()` iterators are tolerant of deletes, but iterating over a
  // snapshot keeps behavior consistent across store implementations.
  for (const id of Array.from(items.keys())) {
    throwIfAborted(signal);
    if (re.test(id)) {
      items.delete(id);
      deleted += 1;
    }
  }
  return deleted;
}
