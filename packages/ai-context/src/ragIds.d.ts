/**
 * Return the unambiguous id prefix used for sheet-level region chunks.
 *
 * Format: `sheet:<len>:<name>:region:`
 */
export function sheetChunkIdPrefix(sheetName: string): string;

/**
 * Legacy id prefix used by older releases (`${sheetName}-region-`).
 */
export function legacySheetChunkIdPrefix(sheetName: string): string;

export interface LegacySheetRegionChunkStoreLike {
  /**
   * In-memory store shape used by ai-context's sheet-level RAG store.
   */
  items?: {
    keys(): IterableIterator<string> | Iterable<string>;
    delete(id: string): boolean;
    [key: string]: unknown;
  } | null;
  [key: string]: unknown;
}

/**
 * Delete all region chunks for a given sheet name (current id scheme), plus best-effort
 * cleanup of legacy `${sheetName}-region-*` ids when supported by the store.
 */
export function deleteSheetRegionChunks(
  store: any,
  sheetName: string,
  options?: { signal?: AbortSignal },
): Promise<void>;

/**
 * Best-effort cleanup for legacy `${sheetName}-region-*` ids.
 *
 * This helper avoids using the legacy prefix for `deleteByPrefix()` because that
 * prefix is ambiguous and can delete chunks for other sheets with similar names.
 *
 * Returns the number of deleted items when supported by the store.
 */
export function deleteLegacySheetRegionChunks(
  store: LegacySheetRegionChunkStoreLike,
  sheetName: string,
  options?: { signal?: AbortSignal },
): number;
