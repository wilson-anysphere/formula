import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { t } from "../i18n/index.js";
import { sortSelection } from "../sort-filter/sortSelection.js";

export const SORT_FILTER_RIBBON_COMMANDS = {
  homeSortAtoZ: "home.editing.sortFilter.sortAtoZ",
  homeSortZtoA: "home.editing.sortFilter.sortZtoA",

  dataSortAtoZ: "data.sortFilter.sortAtoZ",
  dataSortZtoA: "data.sortFilter.sortZtoA",

  // Data tab "Sort" dropdown menu items.
  dataDropdownSortAtoZ: "data.sortFilter.sort.sortAtoZ",
  dataDropdownSortZtoA: "data.sortFilter.sort.sortZtoA",
} as const;

export function registerSortFilterCommands(params: { commandRegistry: CommandRegistry; app: SpreadsheetApp }): void {
  const { commandRegistry, app } = params;

  const category = t("commandCategory.data");

  const registerSortCommand = (commandId: string, title: string, order: "ascending" | "descending"): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () => {
        // `sortSelection` already restores focus to the grid in all supported code paths.
        sortSelection(app, { order });
      },
      { category, icon: null, keywords: ["sort", order === "ascending" ? "a to z" : "z to a"] },
    );
  };

  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.homeSortAtoZ, "Sort A to Z", "ascending");
  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.homeSortZtoA, "Sort Z to A", "descending");

  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataSortAtoZ, "Sort A to Z", "ascending");
  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataSortZtoA, "Sort Z to A", "descending");

  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataDropdownSortAtoZ, "Sort A to Z", "ascending");
  registerSortCommand(SORT_FILTER_RIBBON_COMMANDS.dataDropdownSortZtoA, "Sort Z to A", "descending");
}

