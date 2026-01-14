import type { SpreadsheetApp } from "../app/spreadsheetApp";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast.js";
import { normalizeSelectionRange } from "../formatting/selectionSizeGuard.js";
import type { Range } from "../selection/types";

export type QuickPickItem<T> = { label: string; value: T; description?: string; detail?: string };

export type ShowQuickPick = <T>(items: QuickPickItem<T>[], options?: { placeHolder?: string }) => Promise<T | null>;
export type ShowToast = (message: string, type?: "info" | "warning" | "error") => void;

const MAX_STRUCTURAL_CELL_SELECTION = 50_000;
type InsertCellsChoice = "shiftRight" | "shiftDown" | "entireRow" | "entireColumn";
type DeleteCellsChoice = "shiftLeft" | "shiftUp" | "entireRow" | "entireColumn";

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

  const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
  if (app.isEditing() || globalEditing === true) return true;

  // Always restore focus to the grid after the command completes/cancels (ribbon commands
  // otherwise leave focus on the trigger button).
  try {
    if (typeof (app as any).isReadOnly === "function" && (app as any).isReadOnly()) {
      showCollabEditRejectedToast([{ rejectionKind: "editCells", rejectionReason: "permission" }]);
      return true;
    }
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
    const tooLargeForShift = cellCount > MAX_STRUCTURAL_CELL_SELECTION;

    const rowCount = range.endRow - range.startRow + 1;
    const colCount = range.endCol - range.startCol + 1;

    if (commandId === "home.cells.insert.insertCells") {
      const choice = await showQuickPick(
        [
          {
            label: "Shift cells right",
            value: "shiftRight" as const satisfies InsertCellsChoice,
            description: tooLargeForShift ? "Selection too large" : undefined,
          },
          {
            label: "Shift cells down",
            value: "shiftDown" as const satisfies InsertCellsChoice,
            description: tooLargeForShift ? "Selection too large" : undefined,
          },
          { label: "Entire row", value: "entireRow" as const satisfies InsertCellsChoice },
          { label: "Entire column", value: "entireColumn" as const satisfies InsertCellsChoice },
        ],
        { placeHolder: "Insert Cells" },
      );
      if (!choice) return true;

      switch (choice) {
        case "shiftRight":
        case "shiftDown": {
          if (tooLargeForShift) {
            showToast(
              `Selection too large (>${MAX_STRUCTURAL_CELL_SELECTION.toLocaleString()} cells). Select fewer cells and try again.`,
              "warning",
            );
            return true;
          }
          await app.insertCells(range, choice === "shiftRight" ? "right" : "down");
          return true;
        }
        case "entireRow": {
          if (typeof (app as any).insertRows === "function") {
            await (app as any).insertRows(range.startRow, rowCount);
            return true;
          }
          const doc = app.getDocument();
          const sheetId = app.getCurrentSheetId();
          try {
            doc.insertRows(sheetId, range.startRow, rowCount, { label: "Insert Rows", source: "ribbon" });
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            showToast(`Failed to insert rows: ${message}`, "error");
          }
          return true;
        }
        case "entireColumn": {
          if (typeof (app as any).insertCols === "function") {
            await (app as any).insertCols(range.startCol, colCount);
            return true;
          }
          const doc = app.getDocument();
          const sheetId = app.getCurrentSheetId();
          try {
            doc.insertCols(sheetId, range.startCol, colCount, { label: "Insert Columns", source: "ribbon" });
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            showToast(`Failed to insert columns: ${message}`, "error");
          }
          return true;
        }
        default:
          return true;
      }
    }

    const choice = await showQuickPick(
      [
        {
          label: "Shift cells left",
          value: "shiftLeft" as const satisfies DeleteCellsChoice,
          description: tooLargeForShift ? "Selection too large" : undefined,
        },
        {
          label: "Shift cells up",
          value: "shiftUp" as const satisfies DeleteCellsChoice,
          description: tooLargeForShift ? "Selection too large" : undefined,
        },
        { label: "Entire row", value: "entireRow" as const satisfies DeleteCellsChoice },
        { label: "Entire column", value: "entireColumn" as const satisfies DeleteCellsChoice },
      ],
      { placeHolder: "Delete Cells" },
    );
    if (!choice) return true;

    switch (choice) {
      case "shiftLeft":
      case "shiftUp": {
        if (tooLargeForShift) {
          showToast(
            `Selection too large (>${MAX_STRUCTURAL_CELL_SELECTION.toLocaleString()} cells). Select fewer cells and try again.`,
            "warning",
          );
          return true;
        }
        await app.deleteCells(range, choice === "shiftLeft" ? "left" : "up");
        return true;
      }
      case "entireRow": {
        if (typeof (app as any).deleteRows === "function") {
          await (app as any).deleteRows(range.startRow, rowCount);
          return true;
        }
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();
        try {
          doc.deleteRows(sheetId, range.startRow, rowCount, { label: "Delete Rows", source: "ribbon" });
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          showToast(`Failed to delete rows: ${message}`, "error");
        }
        return true;
      }
      case "entireColumn": {
        if (typeof (app as any).deleteCols === "function") {
          await (app as any).deleteCols(range.startCol, colCount);
          return true;
        }
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();
        try {
          doc.deleteCols(sheetId, range.startCol, colCount, { label: "Delete Columns", source: "ribbon" });
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          showToast(`Failed to delete columns: ${message}`, "error");
        }
        return true;
      }
      default:
        return true;
    }
  } finally {
    app.focus();
  }
}
