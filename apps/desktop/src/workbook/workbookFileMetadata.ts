import type { WorkbookInfo } from "@formula/workbook-backend";

export type WorkbookFileMetadata = {
  /**
   * Workbook directory (typically includes a trailing path separator).
   *
   * `null` indicates the workbook is unsaved / unknown (Excel `CELL("filename")` should return `""`).
   */
  directory: string | null;
  /**
   * Workbook filename (including extension).
   *
   * `null` indicates the workbook is unsaved / unknown.
   */
  filename: string | null;
};

export function splitWorkbookPath(path: string): WorkbookFileMetadata | null {
  const trimmed = String(path ?? "").trim();
  if (!trimmed) return null;

  const lastSlash = trimmed.lastIndexOf("/");
  const lastBackslash = trimmed.lastIndexOf("\\");
  const lastSep = Math.max(lastSlash, lastBackslash);

  if (lastSep < 0) {
    // Filename-only (web/file-picker scenarios).
    return { directory: "", filename: trimmed };
  }

  const filename = trimmed.slice(lastSep + 1);
  if (!filename) return null;

  // Include the separator in the directory, matching Excel's display form.
  const directory = trimmed.slice(0, lastSep + 1);
  return { directory, filename };
}

/**
 * Compute the workbook file metadata that should be injected into the formula engine.
 *
 * Excel semantics:
 * - When the workbook is unsaved, metadata should be cleared so `CELL("filename")` returns `""`.
 * - When a workbook has an associated file identity, we provide directory + filename.
 */
export function getWorkbookFileMetadataFromWorkbookInfo(
  info: Pick<WorkbookInfo, "path" | "origin_path"> | null,
): WorkbookFileMetadata {
  const rawPath = typeof info?.path === "string" && info.path.trim() !== "" ? info.path : null;
  const rawOrigin = typeof info?.origin_path === "string" && info.origin_path.trim() !== "" ? info.origin_path : null;

  const best = rawPath ?? rawOrigin;
  if (!best) return { directory: null, filename: null };

  const parsed = splitWorkbookPath(best);
  if (!parsed) return { directory: null, filename: null };

  return parsed;
}

