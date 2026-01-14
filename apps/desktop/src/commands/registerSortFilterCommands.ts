import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";
import { openCustomSortDialog } from "../sort-filter/openCustomSortDialog.js";
import { sortSelection } from "../sort-filter/sortSelection.js";

export const SORT_FILTER_RIBBON_COMMANDS = {
  // Canonical sort command ids (used by both Home and Data tab ribbon controls).
  sortAtoZ: "data.sortFilter.sortAtoZ",
  sortZtoA: "data.sortFilter.sortZtoA",

  // Custom Sort ribbon ids are still schema-scoped (`home.*` vs `data.*`). Register them so the
  // ribbon doesn't auto-disable and so other surfaces (command palette/keybindings) can invoke
  // the same behavior.
  homeCustomSort: "home.editing.sortFilter.customSort",
  dataCustomSort: "data.sortFilter.sort.customSort",
} as const;

export function registerSortFilterCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  /**
   * Optional spreadsheet edit-state predicate. When omitted, falls back to `app.isEditing()`.
   *
   * The desktop shell passes a custom predicate that includes split-view secondary editing state.
   */
  isEditing?: (() => boolean) | null;
}): void {
  const { commandRegistry, app, isEditing: isEditingParam = null } = params;

  const isEditingActive = (): boolean => {
    if (typeof isEditingParam === "function") return isEditingParam();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const appAny = app as any;
    if (typeof appAny?.isEditing === "function") return Boolean(appAny.isEditing());
    return false;
  };

  const category = t("commandCategory.data");

  const registerSortCommand = (commandId: string, title: string, order: "ascending" | "descending"): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () => {
        if (isEditingActive()) return;
        // `sortSelection` already restores focus to the grid in all supported code paths.
        sortSelection(app, { order });
      },
      { category, icon: null, keywords: ["sort", order === "ascending" ? "a to z" : "z to a"] },
    );
  };

  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.sortAtoZ, "Sort A to Z", "ascending");
  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.sortZtoA, "Sort Z to A", "descending");

  const registerCustomSortCommand = (commandId: string): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      "Custom Sort…",
      () => {
        if (isEditingActive()) return;
        openCustomSortDialog({
          isEditing: isEditingActive,
          getDocument: () => app.getDocument(),
          getSheetId: () => app.getCurrentSheetId(),
          getSelectionRanges: () => app.getSelectionRanges(),
          getCellValue: (sheetId, cell) => app.getCellComputedValueForSheet(sheetId, cell),
          focusGrid: () => app.focus(),
        });
      },
      { category, icon: null, keywords: ["sort", "custom sort"] },
    );
  };

  registerCustomSortCommand(SORT_FILTER_RIBBON_COMMANDS.homeCustomSort);
  registerCustomSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataCustomSort);
}

