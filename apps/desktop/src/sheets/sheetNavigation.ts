import type { SheetMeta } from "./workbookSheetStore";

/**
 * Pick the visible sheet to activate after deleting/hiding a sheet.
 *
 * Mirrors Excel behavior:
 * - Prefer the next visible sheet to the right
 * - Otherwise fall back to the previous visible sheet
 *
 * Returns `null` when the reference sheet isn't found or there is no visible
 * neighbor to activate.
 */
export function pickAdjacentVisibleSheetId(
  sheets: ReadonlyArray<Pick<SheetMeta, "id" | "visibility">>,
  referenceSheetId: string,
): string | null {
  const id = String(referenceSheetId ?? "").trim();
  if (!id) return null;

  const idx = sheets.findIndex((s) => s.id === id);
  if (idx === -1) return null;

  for (let i = idx + 1; i < sheets.length; i += 1) {
    const sheet = sheets[i];
    if (sheet?.visibility === "visible") return sheet.id;
  }

  for (let i = idx - 1; i >= 0; i -= 1) {
    const sheet = sheets[i];
    if (sheet?.visibility === "visible") return sheet.id;
  }

  return null;
}
