import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import { mergeAcross, mergeCells, mergeCenter, unmergeCells } from "../document/mergedCells.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { showInputBox, showToast, type QuickPickItem } from "../extensions/ui.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";
import type { ThemeController } from "../theme/themeController.js";

import { NUMBER_FORMATS, toggleStrikethrough, toggleSubscript, toggleSuperscript, type CellRange } from "../formatting/toolbar.js";
import { promptAndApplyCustomNumberFormat } from "../formatting/promptCustomNumberFormat.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT } from "../formatting/selectionSizeGuard.js";
import { handleHomeCellsInsertDeleteCommand } from "../ribbon/homeCellsCommands.js";

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
  const commandCategoryEditing = t("commandCategory.editing");
  const isEditingFn =
    isEditing ?? (() => (typeof (app as any)?.isEditing === "function" ? (app as any).isEditing() : false));
  const focusGrid = (): void => {
    try {
      (app as any).focus?.();
    } catch {
      // ignore (tests/headless)
    }
  };
  const safeShowToast = (message: string, type: Parameters<typeof showToast>[1] = "info"): void => {
    try {
      showToast(message, type);
    } catch {
      // `showToast` requires a DOM #toast-root; ignore in tests/headless.
    }
  };

  registerAxisSizingCommands({ commandRegistry, app, isEditing, category: commandCategoryFormat });

  // Ribbon-only editing command (Home → Editing → Fill → Series…).
  // Register it in CommandRegistry so the ribbon does not need an exemption and so the command palette
  // can invoke it consistently.
  commandRegistry.registerBuiltinCommand(
    "home.editing.fill.series",
    "Series…",
    async () => {
      if (isEditingFn()) {
        app.focus();
        return;
      }

      const selectionRanges = app.getSelectionRanges();
      if (!Array.isArray(selectionRanges) || selectionRanges.length === 0) {
        app.focus();
        return;
      }

      let minRow = Number.POSITIVE_INFINITY;
      let maxRow = Number.NEGATIVE_INFINITY;
      let minCol = Number.POSITIVE_INFINITY;
      let maxCol = Number.NEGATIVE_INFINITY;
      for (const range of selectionRanges) {
        const startRow = Math.min(range.startRow, range.endRow);
        const endRow = Math.max(range.startRow, range.endRow);
        const startCol = Math.min(range.startCol, range.endCol);
        const endCol = Math.max(range.startCol, range.endCol);
        minRow = Math.min(minRow, startRow);
        maxRow = Math.max(maxRow, endRow);
        minCol = Math.min(minCol, startCol);
        maxCol = Math.max(maxCol, endCol);
      }

      const height = Number.isFinite(minRow) && Number.isFinite(maxRow) ? Math.max(0, maxRow - minRow + 1) : 0;
      const width = Number.isFinite(minCol) && Number.isFinite(maxCol) ? Math.max(0, maxCol - minCol + 1) : 0;
      const suggestVertical = height > width;

      type FillDirection = "down" | "right" | "up" | "left";
      const ordered: FillDirection[] = suggestVertical ? ["down", "up", "right", "left"] : ["right", "left", "down", "up"];

      const picked = await showQuickPick<FillDirection>(
        ordered.map((dir) => ({ label: `Series ${dir[0]!.toUpperCase()}${dir.slice(1)}`, value: dir })),
        { placeHolder: "Series direction" },
      );

      if (!picked) {
        app.focus();
        return;
      }

      app.fillSeries(picked);
      app.focus();
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Fill the selection with a series",
      keywords: ["fill", "series", "auto fill", "autofill", "excel"],
    },
  );

  // Ribbon schema still uses `home.alignment.mergeCenter.*` ids for Merge & Center menu items.
  // Register them here so ribbon enable/disable logic can rely on the CommandRegistry baseline.
  const normalizeSelectionRect = (range: { startRow: number; endRow: number; startCol: number; endCol: number }) => ({
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endCol: Math.max(range.startCol, range.endCol),
  });
  const getSingleSelectionRect = (): { rect: { startRow: number; endRow: number; startCol: number; endCol: number }; multiple: boolean } => {
    const selectionRanges = typeof (app as any)?.getSelectionRanges === "function" ? (app as any).getSelectionRanges() : [];
    const ranges = Array.isArray(selectionRanges) ? selectionRanges : [];
    if (ranges.length > 1) return { rect: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, multiple: true };
    if (ranges.length === 1) return { rect: normalizeSelectionRect(ranges[0]!), multiple: false };
    const cell = typeof (app as any)?.getActiveCell === "function" ? (app as any).getActiveCell() : { row: 0, col: 0 };
    return { rect: { startRow: cell.row, endRow: cell.row, startCol: cell.col, endCol: cell.col }, multiple: false };
  };
  const registerMergeCommand = (args: {
    id: string;
    title: string;
    kind: "mergeCenter" | "mergeAcross" | "mergeCells" | "unmergeCells";
  }): void => {
    commandRegistry.registerBuiltinCommand(
      args.id,
      args.title,
      () => {
        if (isEditingFn()) return;

        if (typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true) {
          safeShowToast("Read-only: cannot merge cells.", "warning");
          focusGrid();
          return;
        }

        const { rect, multiple } = getSingleSelectionRect();
        if (multiple) {
          safeShowToast("Merge commands only support a single selection range.", "warning");
          focusGrid();
          return;
        }

        const rows = rect.endRow - rect.startRow + 1;
        const cols = rect.endCol - rect.startCol + 1;
        const totalCells = rows * cols;
        if (totalCells > DEFAULT_FORMATTING_APPLY_CELL_LIMIT) {
          safeShowToast(
            `Selection too large to merge (>${DEFAULT_FORMATTING_APPLY_CELL_LIMIT.toLocaleString()} cells). Select fewer cells and try again.`,
            "warning",
          );
          focusGrid();
          return;
        }

        // Merge Across is only meaningful for multi-column selections.
        if (args.kind === "mergeAcross" && cols <= 1) {
          focusGrid();
          return;
        }

        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

        const beginBatch = typeof (doc as any)?.beginBatch === "function";
        if (beginBatch) {
          (doc as any).beginBatch({ label: args.title });
        }
        let committed = false;
        try {
          switch (args.kind) {
            case "mergeCenter":
              mergeCenter(doc, sheetId, rect, { label: args.title });
              break;
            case "mergeAcross":
              mergeAcross(doc, sheetId, rect, { label: args.title });
              break;
            case "mergeCells":
              mergeCells(doc, sheetId, rect, { label: args.title });
              break;
            case "unmergeCells":
              unmergeCells(doc, sheetId, rect, { label: args.title });
              break;
            default:
              break;
          }
          committed = true;
        } finally {
          if (beginBatch) {
            if (committed) {
              (doc as any).endBatch?.();
            } else {
              (doc as any).cancelBatch?.();
            }
          }
        }
        focusGrid();
      },
      { category: commandCategoryFormat },
    );
  };

  registerMergeCommand({ id: "home.alignment.mergeCenter.mergeCenter", title: "Merge & Center", kind: "mergeCenter" });
  registerMergeCommand({ id: "home.alignment.mergeCenter.mergeAcross", title: "Merge Across", kind: "mergeAcross" });
  registerMergeCommand({ id: "home.alignment.mergeCenter.mergeCells", title: "Merge Cells", kind: "mergeCells" });
  registerMergeCommand({ id: "home.alignment.mergeCenter.unmergeCells", title: "Unmerge Cells", kind: "unmergeCells" });

  // Ribbon schema uses `home.number.moreFormats.custom` for "Custom…" number formats.
  // Register it so it can be enabled/disabled via CommandRegistry (and used outside the ribbon).
  commandRegistry.registerBuiltinCommand(
    "home.number.moreFormats.custom",
    "Custom number format…",
    async () => {
      try {
        await promptAndApplyCustomNumberFormat({
          isEditing: isEditingFn,
          showInputBox,
          getActiveCellNumberFormat,
          applyFormattingToSelection: (label, fn) => applyFormattingToSelection(label, fn),
        });
      } finally {
        focusGrid();
      }
    },
    { category: commandCategoryFormat },
  );

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

  // Formatting toggles that are wired to the ribbon but are not registered in `registerBuiltinCommands`.
  commandRegistry.registerBuiltinCommand(
    "format.toggleSubscript",
    "Subscript",
    (next?: boolean) =>
      applyFormattingToSelection("Subscript", (doc, sheetId, ranges) => toggleSubscript(doc, sheetId, ranges, { next }), {
        forceBatch: true,
      }),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleSuperscript",
    "Superscript",
    (next?: boolean) =>
      applyFormattingToSelection(
        "Superscript",
        (doc, sheetId, ranges) => toggleSuperscript(doc, sheetId, ranges, { next }),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  // Home → Cells → Insert/Delete Cells… (shift structural edits).
  //
  // These were originally ribbon-only ids handled in `main.ts`. Register them as commands so:
  // - keybindings can invoke them (Excel: Ctrl+Shift+= / Ctrl+-),
  // - they appear in the command palette, and
  // - ribbon disabling can rely on CommandRegistry as the source of truth.
  const showToastForStructuralCellsCommand = (message: string, type?: "info" | "warning" | "error") => showToast(message, type);
  commandRegistry.registerBuiltinCommand(
    "home.cells.insert.insertCells",
    "Insert Cells…",
    async () => {
      await handleHomeCellsInsertDeleteCommand({
        app,
        commandId: "home.cells.insert.insertCells",
        showQuickPick,
        showToast: showToastForStructuralCellsCommand,
      });
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Insert cells and shift existing cells right or down",
      keywords: ["insert", "cells", "shift", "excel"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.cells.delete.deleteCells",
    "Delete Cells…",
    async () => {
      await handleHomeCellsInsertDeleteCommand({
        app,
        commandId: "home.cells.delete.deleteCells",
        showQuickPick,
        showToast: showToastForStructuralCellsCommand,
      });
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Delete cells and shift remaining cells left or up",
      keywords: ["delete", "cells", "shift", "excel"],
    },
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

    // Register ribbon dropdown trigger ids that have a default action when invoked.
    //
    // These ids exist in the ribbon schema (`home.font.*`) for Excel-style UI parity, but the
    // underlying behavior is owned by canonical `format.*` commands. Register them as hidden
    // aliases so ribbon wiring coverage can treat them as implemented without surfacing
    // duplicate entries in the command palette.
    commandRegistry.registerBuiltinCommand("home.font.borders", "Borders", () => commandRegistry.executeCommand("format.borders.all"), {
      category: commandCategoryFormat,
      when: "false",
    });
    commandRegistry.registerBuiltinCommand("home.font.fillColor", "Fill Color", () => commandRegistry.executeCommand("format.fillColor"), {
      category: commandCategoryFormat,
      when: "false",
    });
    commandRegistry.registerBuiltinCommand("home.font.fontColor", "Font Color", () => commandRegistry.executeCommand("format.fontColor"), {
      category: commandCategoryFormat,
      when: "false",
    });
    commandRegistry.registerBuiltinCommand("home.font.fontSize", "Font Size", () => commandRegistry.executeCommand("format.fontSize.set"), {
      category: commandCategoryFormat,
      when: "false",
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

  commandRegistry.registerBuiltinCommand(
    "pageLayout.arrange.selectionPane",
    "Selection Pane",
    () => {
      if (!layoutController) {
        try {
          app.focus();
        } catch {
          // ignore (tests/minimal harnesses)
        }
        return;
      }

      // Excel-style: "Selection Pane" should be idempotent. If the panel is already open,
      // activate/focus it instead of toggling it closed.
      const panelId = PanelIds.SELECTION_PANE;
      const placement = getPanelPlacement(layoutController.layout, panelId);
      layoutController.openPanel(panelId);

      // Floating panels can be minimized; opening should restore them.
      if (placement.kind === "floating" && (layoutController.layout as any)?.floating?.[panelId]?.minimized) {
        layoutController.setFloatingPanelMinimized(panelId, false);
      }

      // The panel is a React mount; wait a frame (or two) so DOM nodes exist before focusing.
      if (typeof document !== "undefined" && typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() =>
          requestAnimationFrame(() => {
            const el = document.querySelector<HTMLElement>('[data-testid="selection-pane"]');
            try {
              el?.focus();
            } catch {
              // Best-effort.
            }
          }),
        );
      }
    },
    {
      category: commandCategoryPageLayout,
      icon: null,
      description: "Open the Selection Pane panel",
      keywords: ["selection pane", "arrange", "drawing", "objects"],
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
