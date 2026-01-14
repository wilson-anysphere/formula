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

/**
 * Best-effort cleanup for legacy `${sheetName}-region-*` ids.
 *
 * This helper avoids using the legacy prefix for `deleteByPrefix()` because that
 * prefix is ambiguous and can delete chunks for other sheets with similar names.
 *
 * Returns the number of deleted items when supported by the store.
 */
export function deleteLegacySheetRegionChunks(
  store: any,
  sheetName: string,
  options?: { signal?: AbortSignal },
): number;

