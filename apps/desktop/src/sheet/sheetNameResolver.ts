export type SheetNameResolver = {
  /**
   * Resolve a stable sheet id to the user-facing display name.
   *
   * Return null when unknown.
   */
  getSheetNameById(id: string): string | null;
  /**
   * Resolve a user-facing display name to a stable sheet id.
   *
   * Implementations must treat lookup as case-insensitive.
   * Return null when unknown.
   */
  getSheetIdByName(name: string): string | null;
};

function normalizeSheetNameForCaseInsensitiveCompare(name: string): string {
  // Excel compares sheet names case-insensitively with Unicode NFKC normalization.
  // Match the semantics used by `@formula/workbook-backend` and the Rust backend.
  try {
    return String(name ?? "").normalize("NFKC").toUpperCase();
  } catch {
    return String(name ?? "").toUpperCase();
  }
}

/**
 * Creates a {@link SheetNameResolver} backed by a live `Map<sheetId, sheetName>`.
 *
 * This helper intentionally reads from the map on each lookup, so callers can
 * mutate the map (e.g. on workbook load / rename) without rebuilding the resolver.
 *
 * Both id and name lookups are case-insensitive and return the canonical id/name
 * stored in the map. Sheet name comparisons use Unicode NFKC normalization to
 * match Excel-like semantics.
 */
export function createSheetNameResolverFromIdToNameMap(sheetIdToName: Map<string, string>): SheetNameResolver {
  return {
    getSheetIdByName(name: string): string | null {
      const needle = String(name ?? "").trim();
      if (!needle) return null;
      const needleIdCi = needle.toLowerCase();
      const needleNameCi = normalizeSheetNameForCaseInsensitiveCompare(needle);

      for (const [id, displayName] of sheetIdToName.entries()) {
        // Accept ids directly too so callers can canonicalize case ("sheet2" -> "Sheet2").
        if (String(id).toLowerCase() === needleIdCi) return id;
        if (normalizeSheetNameForCaseInsensitiveCompare(displayName) === needleNameCi) return id;
      }

      return null;
    },

    getSheetNameById(id: string): string | null {
      const needle = String(id ?? "").trim();
      if (!needle) return null;
      const needleCi = needle.toLowerCase();

      for (const [sheetId, displayName] of sheetIdToName.entries()) {
        if (String(sheetId).toLowerCase() === needleCi) return displayName;
      }

      return null;
    },
  };
}
