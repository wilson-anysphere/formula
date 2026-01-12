import type { SheetUsedRange } from "@formula/workbook-backend";

import type { WorkbookLoadLimits } from "./clampUsedRange.js";

export type WorkbookLoadTruncation = {
  sheetId: string;
  sheetName: string;
  originalRange: SheetUsedRange;
  loadedRange: { startRow: number; endRow: number; startCol: number; endCol: number };
  truncatedRows: boolean;
  truncatedCols: boolean;
};

function formatInt(value: number): string {
  const raw = String(Math.trunc(value));
  return raw.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function rangeEndToExtent(end: number): number {
  if (!Number.isFinite(end)) return 0;
  return Math.max(0, Math.floor(end)) + 1;
}

export function createWorkbookLoadTruncationWarning(
  truncations: WorkbookLoadTruncation[],
  limits: WorkbookLoadLimits,
  options?: { maxSheetsToShow?: number },
): string | null {
  if (!Array.isArray(truncations) || truncations.length === 0) return null;

  const maxSheetsToShow = Math.max(1, options?.maxSheetsToShow ?? 3);
  const capText = `${formatInt(limits.maxRows)} rows × ${formatInt(limits.maxCols)} cols`;

  const sheetSummaries = truncations.slice(0, maxSheetsToShow).map((t) => {
    const sheetLabel = String(t.sheetName ?? "").trim() || String(t.sheetId ?? "").trim() || "Sheet";

    const usedRows = rangeEndToExtent(t.originalRange.end_row);
    const usedCols = rangeEndToExtent(t.originalRange.end_col);

    const hasIntersection = t.loadedRange.startRow <= t.loadedRange.endRow && t.loadedRange.startCol <= t.loadedRange.endCol;
    const loadedRows = hasIntersection ? rangeEndToExtent(t.loadedRange.endRow) : 0;
    const loadedCols = hasIntersection ? rangeEndToExtent(t.loadedRange.endCol) : 0;

    const loadedText = hasIntersection ? `${formatInt(loadedRows)}×${formatInt(loadedCols)}` : "no cells";
    const usedText = `${formatInt(usedRows)}×${formatInt(usedCols)}`;

    return `${sheetLabel} (loaded ${loadedText}, used ${usedText})`;
  });

  const remaining = truncations.length - sheetSummaries.length;
  const sheetsText = remaining > 0 ? `${sheetSummaries.join("; ")}; +${remaining} more` : sheetSummaries.join("; ");

  const hint =
    "To load more, increase limits (?loadMaxRows=…&loadMaxCols=… in dev, or env DESKTOP_LOAD_MAX_ROWS/DESKTOP_LOAD_MAX_COLS).";

  return `Workbook partially loaded (limited to ${capText}). Sheets: ${sheetsText}. ${hint}`;
}

export function warnIfWorkbookLoadTruncated(
  truncations: WorkbookLoadTruncation[],
  limits: WorkbookLoadLimits,
  showToast: (message: string, type: "info" | "warning" | "error", options?: { timeoutMs?: number }) => void,
): void {
  const message = createWorkbookLoadTruncationWarning(truncations, limits);
  if (!message) return;

  try {
    console.warn(`[formula][desktop] ${message}`);
  } catch {
    // Ignore console formatting failures.
  }

  showToast(message, "warning", { timeoutMs: 15_000 });
}

