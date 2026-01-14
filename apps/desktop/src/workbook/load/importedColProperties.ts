import { excelColWidthCharsToPixels } from "@formula/engine";

export type ImportedColPropertiesEntry = {
  width?: number | null;
  hidden?: boolean;
};

export type ImportedSheetColPropertiesPayload = {
  schemaVersion?: number;
  defaultColWidth?: number | null;
  colProperties?: Record<string, ImportedColPropertiesEntry>;
};

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Resolve the initial `SheetViewState.colWidths` value to apply when opening a workbook.
 *
 * The persisted sheet view state (when present) stores **user overrides** (CSS px at zoom=1).
 *
 * The imported workbook model (`Worksheet.col_properties.width`) stores Excel/OOXML widths
 * in "character" units.
 *
 * On open, merge the two so:
 * - persisted (user) widths win for any overlapping column indices
 * - imported widths provide the baseline for columns that have no persisted override
 */
export function sheetColWidthsFromViewOrImportedColProperties(
  view: unknown,
  importedColProperties: unknown,
): unknown | null {
  const normalizeColWidths = (raw: unknown): Record<string, number> | null => {
    if (!raw) return null;
    const out: Record<string, number> = {};

    if (Array.isArray(raw)) {
      // Backwards compatibility with older encodings that used arrays of { index, size } or [index,size].
      for (const entry of raw) {
        const index = Array.isArray(entry) ? entry[0] : (entry as any)?.col ?? (entry as any)?.index;
        const size = Array.isArray(entry) ? entry[1] : (entry as any)?.width ?? (entry as any)?.size;
        const idx = Number(index);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const value = Number(size);
        if (!Number.isFinite(value) || value <= 0) continue;
        out[String(idx)] = value;
      }
    } else if (isPlainObject(raw)) {
      for (const [key, valueRaw] of Object.entries(raw)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const value = Number(valueRaw);
        if (!Number.isFinite(value) || value <= 0) continue;
        out[String(idx)] = value;
      }
    }

    return Object.keys(out).length > 0 ? out : null;
  };

  const persistedColWidthsRaw = isPlainObject(view) ? (view as any).colWidths : (view as any)?.colWidths;
  const persistedColWidths = normalizeColWidths(persistedColWidthsRaw);

  const importedColWidths = docColWidthsFromImportedColProperties(importedColProperties);

  if (persistedColWidths && importedColWidths) {
    return { ...importedColWidths, ...persistedColWidths };
  }
  return persistedColWidths ?? importedColWidths;
}

/**
 * Convert a backend `get_sheet_imported_col_properties` payload into a DocumentController
 * `SheetViewState.colWidths` map (CSS px at zoom=1).
 */
export function docColWidthsFromImportedColProperties(raw: unknown): Record<string, number> | null {
  const payload = isPlainObject(raw) ? (raw as ImportedSheetColPropertiesPayload) : null;
  const colProperties = payload && isPlainObject(payload.colProperties) ? payload.colProperties : null;
  if (!colProperties) return null;

  const out: Record<string, number> = {};

  for (const [key, entryRaw] of Object.entries(colProperties)) {
    const col = Number(key);
    if (!Number.isInteger(col) || col < 0) continue;

    const entry = isPlainObject(entryRaw) ? (entryRaw as ImportedColPropertiesEntry) : null;
    const widthChars = Number(entry?.width);
    if (!Number.isFinite(widthChars) || widthChars <= 0) continue;

    const px = excelColWidthCharsToPixels(widthChars);
    if (!Number.isFinite(px) || px <= 0) continue;

    out[String(col)] = px;
  }

  return Object.keys(out).length > 0 ? out : null;
}

/**
 * Extract 0-based hidden column indices from a backend `get_sheet_imported_col_properties` payload.
 */
export function hiddenColsFromImportedColProperties(raw: unknown): number[] {
  const payload = isPlainObject(raw) ? (raw as ImportedSheetColPropertiesPayload) : null;
  const colProperties = payload && isPlainObject(payload.colProperties) ? payload.colProperties : null;
  if (!colProperties) return [];

  const out: number[] = [];
  for (const [key, entryRaw] of Object.entries(colProperties)) {
    const col = Number(key);
    if (!Number.isInteger(col) || col < 0) continue;

    const entry = isPlainObject(entryRaw) ? (entryRaw as ImportedColPropertiesEntry) : null;
    if (entry?.hidden === true) out.push(col);
  }

  out.sort((a, b) => a - b);
  return out;
}

/**
 * Extract the sheet's default column width (OOXML `<sheetFormatPr defaultColWidth="...">`).
 *
 * This is expressed in Excel "character" units (the same units used by `col/@width`).
 */
export function defaultColWidthCharsFromImportedColProperties(raw: unknown): number | null {
  const payload = isPlainObject(raw) ? (raw as ImportedSheetColPropertiesPayload) : null;
  const width = Number(payload?.defaultColWidth);
  if (!Number.isFinite(width) || width <= 0) return null;
  return width;
}
