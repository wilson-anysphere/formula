import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { promptAndApplyAxisSizing } from "../ribbon/axisSizing.js";

export const AXIS_SIZING_COMMAND_IDS = {
  rowHeight: "home.cells.format.rowHeight",
  columnWidth: "home.cells.format.columnWidth",
} as const;

export function registerAxisSizingCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  /**
   * Optional override for determining whether the spreadsheet is in "editing" mode.
   * The desktop uses a custom guard (`isSpreadsheetEditing`) that includes split-view
   * secondary editor state; callers can pass that in here.
   */
  isEditing?: (() => boolean) | null;
  category?: string | null;
}): void {
  const { commandRegistry, app, isEditing = null, category = null } = params;
  const isEditingFn =
    isEditing ??
    (() => {
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      const appAny = app as any;
      const primaryEditing = typeof appAny?.isEditing === "function" && appAny.isEditing() === true;
      return primaryEditing || globalEditing === true;
    });

  commandRegistry.registerBuiltinCommand(
    AXIS_SIZING_COMMAND_IDS.rowHeight,
    "Row Height…",
    () => promptAndApplyAxisSizing(app, "rowHeight", { isEditing: isEditingFn }),
    {
      category,
      icon: null,
      keywords: ["row", "height", "resize row", "row height"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    AXIS_SIZING_COMMAND_IDS.columnWidth,
    "Column Width…",
    () => promptAndApplyAxisSizing(app, "colWidth", { isEditing: isEditingFn }),
    {
      category,
      icon: null,
      keywords: ["column", "width", "resize column", "column width"],
    },
  );
}
