import type { SheetMeta } from "./workbookSheetStore";

/**
 * Pick the visible sheet to activate after deleting/hiding the current visible sheet.
 *
 * Mirrors Excel behavior:
 * - Prefer the next visible sheet to the right
 * - Otherwise fall back to the previous visible sheet
 *
 * Returns `null` when the reference sheet isn't found in the visible list.
 */
export function pickAdjacentVisibleSheetId(
  sheets: ReadonlyArray<Pick<SheetMeta, "id" | "visibility">>,
  referenceSheetId: string,
): string | null {
  const id = String(referenceSheetId ?? "").trim();
  if (!id) return null;

  const visibleSheets = sheets.filter((s) => s.visibility === "visible");
  const idx = visibleSheets.findIndex((s) => s.id === id);
  if (idx === -1) return null;
  return visibleSheets[idx + 1]?.id ?? visibleSheets[idx - 1]?.id ?? null;
}

