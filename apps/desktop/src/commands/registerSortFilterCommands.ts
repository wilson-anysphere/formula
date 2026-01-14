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
   * Optional spreadsheet edit-state predicate.
   *
   * When omitted, falls back to `app.isEditing()` and the desktop-shell-owned
   * `globalThis.__formulaSpreadsheetIsEditing` flag (when present).
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
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    const primaryEditing = typeof appAny?.isEditing === "function" && appAny.isEditing() === true;
    return primaryEditing || globalEditing === true;
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

  const registerCustomSortCommand = (
    commandId: string,
    options: { when?: string | null; delegateTo?: string | null } = {},
  ): void => {
    const delegateTo = options.delegateTo ?? null;
    commandRegistry.registerBuiltinCommand(
      commandId,
      "Custom Sort…",
      delegateTo
        ? () => commandRegistry.executeCommand(delegateTo)
        : () => {
            if (isEditingActive()) return;
            openCustomSortDialog({
              isEditing: isEditingActive,
              isReadOnly: () => {
                try {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const appAny = app as any;
                  return typeof appAny?.isReadOnly === "function" && appAny.isReadOnly() === true;
                } catch {
                  return false;
                }
              },
              getDocument: () => app.getDocument(),
              getSheetId: () => app.getCurrentSheetId(),
              getSelectionRanges: () => app.getSelectionRanges(),
              getCellValue: (sheetId, cell) => app.getCellComputedValueForSheet(sheetId, cell),
              focusGrid: () => app.focus(),
            });
          },
      { category, icon: null, keywords: ["sort", "custom sort"], when: options.when ?? null },
    );
  };

  // Home uses a ribbon-scoped id for UI parity; hide it from the command palette to avoid
  // duplicate "Custom Sort…" entries (Data tab id is treated as canonical).
  registerCustomSortCommand(SORT_FILTER_RIBBON_COMMANDS.homeCustomSort, {
    when: "false",
    // Ensure command-palette recents tracking lands on the canonical command id.
    delegateTo: SORT_FILTER_RIBBON_COMMANDS.dataCustomSort,
  });
  registerCustomSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataCustomSort);
}
