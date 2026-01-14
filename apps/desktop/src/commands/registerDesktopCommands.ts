import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { QuickPickItem } from "../extensions/ui.js";
import type { LayoutController } from "../layout/layoutController.js";
import { t } from "../i18n/index.js";
import type { ThemeController } from "../theme/themeController.js";

import { NUMBER_FORMATS, toggleStrikethrough, toggleSubscript, toggleSuperscript, type CellRange } from "../formatting/toolbar.js";

import { registerBuiltinCommands } from "./registerBuiltinCommands.js";
import { registerAxisSizingCommands } from "./registerAxisSizingCommands.js";
import { registerBuiltinFormatFontCommands } from "./registerBuiltinFormatFontCommands.js";
import { registerFormatPainterCommand } from "./formatPainterCommand.js";
import { registerFormatAlignmentCommands } from "./registerFormatAlignmentCommands.js";
import { registerFormatFontDropdownCommands } from "./registerFormatFontDropdownCommands.js";
import { registerDataQueriesCommands } from "./registerDataQueriesCommands.js";
import { registerNumberFormatCommands } from "./registerNumberFormatCommands.js";
import { registerPageLayoutCommands, type PageLayoutCommandHandlers } from "./registerPageLayoutCommands.js";
import { registerRibbonMacroCommands, type RibbonMacroCommandHandlers } from "./registerRibbonMacroCommands.js";
import { registerWorkbenchFileCommands, type WorkbenchFileCommandHandlers } from "./registerWorkbenchFileCommands.js";
import { registerSortFilterCommands } from "./registerSortFilterCommands.js";
import { registerHomeStylesCommands } from "./registerHomeStylesCommands.js";

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

export type FormatPainterCommandHandlers = {
  isArmed: () => boolean;
  arm: () => void;
  disarm: () => void;
  onCancel?: (() => void) | null;
};

export type DataQueriesCommandHandlers = Pick<
  Parameters<typeof registerDataQueriesCommands>[0],
  "getPowerQueryService" | "showToast" | "notify" | "now" | "focusAfterExecute"
>;

export function registerDesktopCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController | null;
  /**
   * Optional "editing mode" guard for commands that should not run while editing.
   *
   * The desktop shell uses a custom guard (`isSpreadsheetEditing`) that includes
   * split-view secondary editor state, so callers should pass that in when available.
   */
  isEditing?: (() => boolean) | null;
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
  formatPainter?: FormatPainterCommandHandlers | null;
  ribbonMacroHandlers?: RibbonMacroCommandHandlers | null;
  dataQueriesHandlers?: DataQueriesCommandHandlers | null;
  pageLayoutHandlers?: PageLayoutCommandHandlers | null;
  /**
   * Optional command palette opener. When provided, `workbench.showCommandPalette` will be
   * overridden to invoke this handler (instead of the built-in no-op registration).
   */
  openCommandPalette?: (() => void) | null;
  /**
   * Optional host hook to open the Goal Seek dialog (What-If Analysis).
   */
  openGoalSeekDialog?: (() => void) | null;
}): void {
  const {
    commandRegistry,
    app,
    layoutController,
    isEditing = null,
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
    formatPainter = null,
    ribbonMacroHandlers = null,
    dataQueriesHandlers = null,
    pageLayoutHandlers = null,
    openCommandPalette = null,
    openGoalSeekDialog = null,
  } = params;

  const commandCategoryFormat = t("commandCategory.format");

  registerAxisSizingCommands({ commandRegistry, app, isEditing, category: commandCategoryFormat });

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

  // Ribbon-only formatting toggles that are not yet part of the canonical `format.*` command namespace.
  // These are still registered in the CommandRegistry so the ribbon does not auto-disable them and
  // so other UI surfaces (command palette/keybindings) can invoke them consistently.
  commandRegistry.registerBuiltinCommand(
    "home.font.subscript",
    "Subscript",
    (next?: boolean) =>
      applyFormattingToSelection("Subscript", (doc, sheetId, ranges) => toggleSubscript(doc, sheetId, ranges, { next }), {
        forceBatch: true,
      }),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "home.font.superscript",
    "Superscript",
    (next?: boolean) =>
      applyFormattingToSelection(
        "Superscript",
        (doc, sheetId, ranges) => toggleSuperscript(doc, sheetId, ranges, { next }),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );
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
      openGoalSeekDialog,
    });

    // Home → Font dropdown actions are registered as canonical `format.*` commands so
    // ribbon actions and the command palette share a single command surface.
    registerFormatFontDropdownCommands({
      commandRegistry,
      category: commandCategoryFormat,
      applyFormattingToSelection,
    });
  }

  // Register number formats after `registerBuiltinCommands` so the desktop shell can override
  // any builtin wiring with the host-provided `applyFormattingToSelection` + `getActiveCellNumberFormat`.
  registerNumberFormatCommands({
    commandRegistry,
    applyFormattingToSelection,
    getActiveCellNumberFormat,
    t,
    category: commandCategoryFormat,
  });

  registerWorkbenchFileCommands({ commandRegistry, handlers: workbenchFileHandlers });

  if (formatPainter) {
    registerFormatPainterCommand({ commandRegistry, ...formatPainter });
  }

  if (ribbonMacroHandlers) {
    registerRibbonMacroCommands({ commandRegistry, handlers: ribbonMacroHandlers });
  }

  if (dataQueriesHandlers) {
    const focusAfterExecute =
      dataQueriesHandlers.focusAfterExecute === undefined
        ? typeof (app as any)?.focus === "function"
          ? () => (app as any).focus()
          : null
        : dataQueriesHandlers.focusAfterExecute;

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      refreshRibbonUiState,
      ...dataQueriesHandlers,
      focusAfterExecute,
    });
  }

  registerBuiltinFormatFontCommands({
    commandRegistry,
    applyFormattingToSelection,
  });

  registerSortFilterCommands({ commandRegistry, app });

  registerFormatAlignmentCommands({
    commandRegistry,
    applyFormattingToSelection,
    activeCellIndentLevel: getActiveCellIndentLevel,
    openAlignmentDialog: () => {
      const focusAlignmentSection = () => {
        if (typeof document === "undefined") return;
        const input = document.querySelector<HTMLElement>('[data-testid="format-cells-horizontal"]');
        if (!input) return;
        try {
          input.scrollIntoView({ block: "center" });
        } catch {
          // ignore (non-DOM contexts/tests)
        }
        try {
          input.focus();
        } catch {
          // ignore
        }
      };

      const result = openFormatCells();
      // Support async openers (even though desktop currently uses a sync dialog).
      if (result && typeof (result as any)?.then === "function") {
        void (result as Promise<void>).then(() => focusAlignmentSection());
      } else {
        focusAlignmentSection();
      }
    },
  });

  registerHomeStylesCommands({
    commandRegistry,
    app,
    category: commandCategoryFormat,
    applyFormattingToSelection,
    showQuickPick,
    isEditing,
  });

  // Page Layout → Arrange drawing order commands. These are desktop-only (drawing overlay)
  // but still registered in CommandRegistry so the ribbon does not auto-disable them and
  // so other UI surfaces (command palette/keybindings) can invoke them consistently.
  const commandCategoryPageLayout = "Page Layout";
  commandRegistry.registerBuiltinCommand(
    "pageLayout.arrange.bringForward",
    "Bring Forward",
    () => {
      app.bringSelectedDrawingForward();
      app.focus();
    },
    {
      category: commandCategoryPageLayout,
      icon: null,
      description: "Bring the selected drawing forward",
      keywords: ["arrange", "drawing", "bring forward", "z order", "layer"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "pageLayout.arrange.sendBackward",
    "Send Backward",
    () => {
      app.sendSelectedDrawingBackward();
      app.focus();
    },
    {
      category: commandCategoryPageLayout,
      icon: null,
      description: "Send the selected drawing backward",
      keywords: ["arrange", "drawing", "send backward", "z order", "layer"],
    },
  );

  if (pageLayoutHandlers) {
    registerPageLayoutCommands({ commandRegistry, handlers: pageLayoutHandlers });
  }

  // Override the builtin quick-pick implementation to open the full Format Cells UI.
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

  commandRegistry.registerBuiltinCommand("navigation.goTo", t("command.navigation.goTo"), () => findReplace.openGoTo(), {
    category: t("commandCategory.navigation"),
    icon: null,
    description: t("commandDescription.navigation.goTo"),
    keywords: ["go to", "goto", "reference", "name box"],
  });

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
