import { excelColWidthCharsToPixels } from "@formula/engine";

export type ImportedColPropertiesEntry = {
  width?: number | null;
  hidden?: boolean;
};

export type ImportedSheetColPropertiesPayload = {
  schemaVersion?: number;
  colProperties?: Record<string, ImportedColPropertiesEntry>;
};

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
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

