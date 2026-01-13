import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { QuickPickItem } from "../extensions/ui.js";
import type { LayoutController } from "../layout/layoutController.js";
import { t } from "../i18n/index.js";
import type { ThemeController } from "../theme/themeController.js";

import {
  NUMBER_FORMATS,
  toggleStrikethrough,
  type CellRange,
} from "../formatting/toolbar.js";

import { registerBuiltinCommands } from "./registerBuiltinCommands.js";
import { registerBuiltinFormatFontCommands } from "./registerBuiltinFormatFontCommands.js";
import { registerFormatAlignmentCommands } from "./registerFormatAlignmentCommands.js";
import { registerNumberFormatCommands } from "./registerNumberFormatCommands.js";
import { registerPageLayoutCommands, type PageLayoutCommandHandlers } from "./registerPageLayoutCommands.js";
import { registerWorkbenchFileCommands, type WorkbenchFileCommandHandlers } from "./registerWorkbenchFileCommands.js";

export type ApplyFormattingToSelection = (
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
  options?: { forceBatch?: boolean },
) => void;

export type FindReplaceCommandHandlers = {
  openFind: () => void;
  openReplace: () => void;
  openGoTo: () => void;
};

export function registerDesktopCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController | null;
  focusAfterSheetNavigation?: (() => void) | null;
  getVisibleSheetIds?: (() => string[]) | null;
  ensureExtensionsLoaded?: (() => Promise<void>) | null;
  onExtensionsLoaded?: (() => void) | null;
  themeController?: Pick<ThemeController, "setThemePreference"> | null;
  refreshRibbonUiState?: (() => void) | null;
  applyFormattingToSelection: ApplyFormattingToSelection;
  getActiveCellNumberFormat: () => string | null;
  getActiveCellIndentLevel: () => number;
  openFormatCells: () => void | Promise<void>;
  showQuickPick: <T>(items: QuickPickItem<T>[], options?: { placeHolder?: string }) => Promise<T | null>;
  findReplace: FindReplaceCommandHandlers;
  workbenchFileHandlers: WorkbenchFileCommandHandlers;
  pageLayoutHandlers?: PageLayoutCommandHandlers | null;
  /**
   * Optional command palette opener. When provided, `workbench.showCommandPalette` will be
   * overridden to invoke this handler (instead of the built-in no-op registration).
   */
  openCommandPalette?: (() => void) | null;
}): void {
  const {
    commandRegistry,
    app,
    layoutController,
    focusAfterSheetNavigation = null,
    getVisibleSheetIds = null,
    ensureExtensionsLoaded = null,
    onExtensionsLoaded = null,
    themeController = null,
    refreshRibbonUiState = null,
    applyFormattingToSelection,
    getActiveCellNumberFormat,
    getActiveCellIndentLevel,
    openFormatCells,
    showQuickPick,
    findReplace,
    workbenchFileHandlers,
    pageLayoutHandlers = null,
    openCommandPalette = null,
  } = params;

  const commandCategoryFormat = t("commandCategory.format");

  commandRegistry.registerBuiltinCommand(
    "format.toggleStrikethrough",
    t("command.format.toggleStrikethrough"),
    (next?: boolean) =>
      applyFormattingToSelection(
        t("command.format.toggleStrikethrough"),
        (doc, sheetId, ranges) => toggleStrikethrough(doc, sheetId, ranges, { next }),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  registerNumberFormatCommands({
    commandRegistry,
    applyFormattingToSelection,
    getActiveCellNumberFormat,
    t,
    category: commandCategoryFormat,
  });

  if (layoutController) {
    registerBuiltinCommands({
      commandRegistry,
      app,
      layoutController,
      focusAfterSheetNavigation,
      getVisibleSheetIds,
      ensureExtensionsLoaded,
      onExtensionsLoaded,
      themeController,
      refreshRibbonUiState,
    });
  }

  registerWorkbenchFileCommands({ commandRegistry, handlers: workbenchFileHandlers });

  registerBuiltinFormatFontCommands({
    commandRegistry,
    app,
    applyFormattingToSelection,
  });

  registerFormatAlignmentCommands({
    commandRegistry,
    applyFormattingToSelection,
    activeCellIndentLevel: getActiveCellIndentLevel,
    openAlignmentDialog: () => void openFormatCells(),
  });

  if (pageLayoutHandlers) {
    registerPageLayoutCommands({ commandRegistry, handlers: pageLayoutHandlers });
  }
  commandRegistry.registerBuiltinCommand("format.openFormatCells", t("command.format.openFormatCells"), () => openFormatCells(), {
    category: commandCategoryFormat,
    icon: null,
    keywords: ["format cells", "number format", "font"],
  });

  // Quick-pick variant for applying common number formats without opening the full dialog.
  commandRegistry.registerBuiltinCommand(
    "format.applyNumberFormatPresetQuickPick",
    t("command.format.applyNumberFormatPresetQuickPick"),
    async () => {
      type Choice = "general" | "currency" | "percent" | "date";
      const choice = await showQuickPick<Choice>(
        [
          { label: "General", description: "Clear number format", value: "general" },
          { label: "Currency", description: NUMBER_FORMATS.currency, value: "currency" },
          { label: "Percent", description: NUMBER_FORMATS.percent, value: "percent" },
          { label: "Date", description: NUMBER_FORMATS.date, value: "date" },
        ],
        { placeHolder: "Number format" },
      );
      if (!choice) return;

      const patch = choice === "general" ? { numberFormat: null } : { numberFormat: NUMBER_FORMATS[choice] };

      applyFormattingToSelection(
        "Number format",
        (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, patch, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        },
        { forceBatch: true },
      );
    },
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand("edit.find", t("command.edit.find"), () => findReplace.openFind(), {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.find"),
    keywords: ["find", "search"],
  });

  commandRegistry.registerBuiltinCommand("edit.replace", t("command.edit.replace"), () => findReplace.openReplace(), {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.replace"),
    keywords: ["replace", "find"],
  });

  commandRegistry.registerBuiltinCommand(
    "navigation.goTo",
    t("command.navigation.goTo"),
    () => findReplace.openGoTo(),
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.navigation.goTo"),
      keywords: ["go to", "goto", "reference", "name box"],
    },
  );

  if (layoutController && openCommandPalette) {
    // `registerBuiltinCommands(...)` wires this as a no-op so the desktop shell can own
    // opening the palette. Override it in the desktop UI so keybinding dispatch through
    // `CommandRegistry.executeCommand(...)` works as well.
    commandRegistry.registerBuiltinCommand(
      "workbench.showCommandPalette",
      t("command.workbench.showCommandPalette"),
      () => openCommandPalette(),
      {
        category: t("commandCategory.navigation"),
        icon: null,
        description: t("commandDescription.workbench.showCommandPalette"),
        keywords: ["command palette", "commands"],
      },
    );
  }
}
