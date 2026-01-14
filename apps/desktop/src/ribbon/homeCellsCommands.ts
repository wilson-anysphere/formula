import type { SpreadsheetApp } from "../app/spreadsheetApp";
import { normalizeSelectionRange } from "../formatting/selectionSizeGuard.js";
import type { Range } from "../selection/types";

export type QuickPickItem<T> = { label: string; value: T; description?: string; detail?: string };

export type ShowQuickPick = <T>(items: QuickPickItem<T>[], options?: { placeHolder?: string }) => Promise<T | null>;
export type ShowToast = (message: string, type?: "info" | "warning" | "error") => void;

const MAX_STRUCTURAL_CELL_SELECTION = 50_000;

function selectionCellCount(range: Range): number {
  const r = normalizeSelectionRange(range);
  const rows = r.endRow - r.startRow + 1;
  const cols = r.endCol - r.startCol + 1;
  if (rows <= 0 || cols <= 0) return 0;
  return rows * cols;
}

export async function handleHomeCellsInsertDeleteCommand(params: {
  app: SpreadsheetApp;
  commandId: string;
  showQuickPick: ShowQuickPick;
  showToast: ShowToast;
}): Promise<boolean> {
  const { app, commandId, showQuickPick, showToast } = params;

  if (commandId !== "home.cells.insert.insertCells" && commandId !== "home.cells.delete.deleteCells") {
    return false;
  }

  if (app.isEditing()) return true;

  // Always restore focus to the grid after the command completes/cancels (ribbon commands
  // otherwise leave focus on the trigger button).
  try {
    const ranges = app.getSelectionRanges();
    const range: Range | null =
      ranges.length === 0
        ? (() => {
            const active = app.getActiveCell();
            return { startRow: active.row, endRow: active.row, startCol: active.col, endCol: active.col };
          })()
        : ranges.length === 1
          ? normalizeSelectionRange(ranges[0]!)
          : null;

    if (!range) {
      showToast("Insert/Delete Cells currently supports only a single selection range.", "warning");
      return true;
    }

    const cellCount = selectionCellCount(range);
    if (cellCount > MAX_STRUCTURAL_CELL_SELECTION) {
      showToast(
        `Selection too large (>${MAX_STRUCTURAL_CELL_SELECTION.toLocaleString()} cells). Select fewer cells and try again.`,
        "warning",
      );
      return true;
    }

    if (commandId === "home.cells.insert.insertCells") {
      const choice = await showQuickPick(
        [
          { label: "Shift cells right", value: "right" as const },
          { label: "Shift cells down", value: "down" as const },
        ],
        { placeHolder: "Insert Cells" },
      );
      if (!choice) return true;

      await app.insertCells(range, choice);
      return true;
    }

    const choice = await showQuickPick(
      [
        { label: "Shift cells left", value: "left" as const },
        { label: "Shift cells up", value: "up" as const },
      ],
      { placeHolder: "Delete Cells" },
    );
    if (!choice) return true;

    await app.deleteCells(range, choice);
    return true;
  } finally {
    app.focus();
  }
}
