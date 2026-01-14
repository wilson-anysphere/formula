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
  const rawPath = typeof info?.path === "string" ? info.path.trim() : "";
  const rawOrigin = typeof info?.origin_path === "string" ? info.origin_path.trim() : "";

  const best = rawPath || rawOrigin;
  if (!best) return { directory: null, filename: null };

  const parsed = splitWorkbookPath(best);
  if (!parsed) return { directory: null, filename: null };

  return parsed;
}

const WORKBOOK_SAVE_EXTENSIONS = new Set(["xlsx", "xlsm", "xltx", "xltm", "xlam", "xlsb"]);

/**
 * Mirror the desktop backend's save-path coercion behavior.
 *
 * The Tauri backend may rewrite saves for non-workbook origins (CSV/Parquet/etc) and legacy
 * formats (XLS) by changing the file extension to `.xlsx` when the user uses plain "Save"
 * instead of "Save As".
 *
 * Keep this logic in sync with `apps/desktop/src-tauri/src/commands.rs:coerce_save_path_to_xlsx`.
 */
export function coerceSavePathToXlsx(path: string): string {
  const raw = String(path ?? "");
  if (!raw) return raw;

  const lastSlash = raw.lastIndexOf("/");
  const lastBackslash = raw.lastIndexOf("\\");
  const lastSep = Math.max(lastSlash, lastBackslash);
  const basename = raw.slice(lastSep + 1);

  const lastDot = basename.lastIndexOf(".");
  if (lastDot <= 0) return raw; // no extension or hidden file like `.foo`
  if (lastDot >= basename.length - 1) return raw; // empty extension (e.g. `foo.`)

  const ext = basename.slice(lastDot + 1);
  if (WORKBOOK_SAVE_EXTENSIONS.has(ext.toLowerCase())) return raw;

  const coercedBasename = `${basename.slice(0, lastDot)}.xlsx`;
  const prefix = lastSep >= 0 ? raw.slice(0, lastSep + 1) : "";
  return `${prefix}${coercedBasename}`;
}
