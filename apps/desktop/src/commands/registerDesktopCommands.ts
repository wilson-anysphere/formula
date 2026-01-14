import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { DocumentController } from "../document/documentController.js";
import { mergeAcross, mergeCells, mergeCenter, unmergeCells } from "../document/mergedCells.js";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { showInputBox, showToast, type QuickPickItem } from "../extensions/ui.js";
import type { LayoutController } from "../layout/layoutController.js";
import { t } from "../i18n/index.js";
import type { ThemeController } from "../theme/themeController.js";

import { NUMBER_FORMATS, toggleStrikethrough, toggleSubscript, toggleSuperscript, type CellRange } from "../formatting/toolbar.js";
import { promptAndApplyCustomNumberFormat } from "../formatting/promptCustomNumberFormat.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize } from "../formatting/selectionSizeGuard.js";
import { DEFAULT_GRID_LIMITS } from "../selection/selection.js";
import { computeSelectionFormatState } from "../ribbon/selectionFormatState.js";
import { executeCellsStructuralRibbonCommand } from "../ribbon/cellsStructuralCommands.js";
import { handleHomeCellsInsertDeleteCommand } from "../ribbon/homeCellsCommands.js";
import { handleInsertPicturesRibbonCommand } from "../main.insertPicturesRibbonCommand.js";
import { exportDocumentRangeToCsv } from "../import-export/csv/export.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast.js";

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
import { registerRibbonAutoFilterCommands, type RibbonAutoFilterCommandHandlers } from "./registerRibbonAutoFilterCommands.js";
import { registerWorkbenchFileCommands, type WorkbenchFileCommandHandlers } from "./registerWorkbenchFileCommands.js";
import { registerSortFilterCommands } from "./registerSortFilterCommands.js";
import { registerHomeStylesCommands } from "./registerHomeStylesCommands.js";

export type { RibbonAutoFilterCommandHandlers } from "./registerRibbonAutoFilterCommands.js";

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

export type SheetStructureCommandHandlers = {
  /**
   * Insert a new sheet and activate it (Excel-style "Insert Sheet").
   *
   * The desktop shell owns workbook sheet metadata and confirmations, so the handler is
   * provided by `main.ts` (typically via `createAddSheetCommand(...)`).
   */
  insertSheet: () => void | Promise<void>;
  /**
   * Delete the active sheet (with confirmation) and activate a remaining sheet if needed.
   */
  deleteActiveSheet: () => void | Promise<void>;
  /**
   * Open the Organize Sheets dialog (reorder/rename/hide/etc).
   */
  openOrganizeSheets?: (() => void | Promise<void>) | null;
};

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
  /**
   * Optional hook to commit any in-progress edits before running commands that need the latest
   * document state (e.g. exports).
   *
   * Defaults to calling `app.commitPendingEditsForCommand()` when available. `main.ts` should pass
   * its `commitAllPendingEditsForCommand()` helper so split-view secondary pane edits are also
   * included.
   */
  commitPendingEditsForCommand?: (() => void) | null;
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
  /**
   * Optional handler for File → Info → Protect Workbook → "Encrypt with Password…".
   *
   * The desktop shell owns file dialogs and save flows, so the implementation lives in `main.ts`.
   */
  encryptWithPassword?: (() => void | Promise<void>) | null;
  formatPainter?: FormatPainterCommandHandlers | null;
  ribbonMacroHandlers?: RibbonMacroCommandHandlers | null;
  dataQueriesHandlers?: DataQueriesCommandHandlers | null;
  pageLayoutHandlers?: PageLayoutCommandHandlers | null;
  /**
   * Optional handlers for sheet structure commands (Insert/Delete Sheet).
   *
   * These ids come from the ribbon schema (`home.cells.*`) but the implementation depends on
   * workbook sheet store state that lives in the desktop shell (`main.ts`).
   */
  sheetStructureHandlers?: SheetStructureCommandHandlers | null;
  /**
   * Optional handlers for the MVP ribbon AutoFilter commands (`data.sortFilter.*`).
   *
   * AutoFilter logic currently lives in `main.ts` (local-only UI + store). Register the commands in
   * CommandRegistry so baseline ribbon enable/disable can rely on registration (no exemptions needed).
   */
  autoFilterHandlers?: RibbonAutoFilterCommandHandlers | null;
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
    commitPendingEditsForCommand: commitPendingEditsForCommandHandler = null,
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
    encryptWithPassword = null,
    formatPainter = null,
    ribbonMacroHandlers = null,
    dataQueriesHandlers = null,
    pageLayoutHandlers = null,
    sheetStructureHandlers = null,
    autoFilterHandlers = null,
    openCommandPalette = null,
    openGoalSeekDialog = null,
  } = params;

  const commandCategoryFormat = t("commandCategory.format");
  const commandCategoryEditing = t("commandCategory.editing");
  const commandCategoryData = t("commandCategory.data");
  const commandCategoryFile = t("menu.file");
  const isEditingFn =
    isEditing ??
    (() => {
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      const appAny = app as any;
      const primaryEditing = typeof appAny?.isEditing === "function" && appAny.isEditing() === true;
      return primaryEditing || globalEditing === true;
    });
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
  const commitEditsForCommand = (): void => {
    try {
      if (commitPendingEditsForCommandHandler) {
        commitPendingEditsForCommandHandler();
        return;
      }
      if (typeof (app as any)?.commitPendingEditsForCommand === "function") {
        (app as any).commitPendingEditsForCommand();
      }
    } catch {
      // ignore (tests/headless)
    }
  };
  const sanitizeFilename = (raw: string): string => {
    const cleaned = String(raw ?? "")
      // Windows-reserved + generally-illegal filename characters.
      .replace(/[\\/:*?"<>|]+/g, "_")
      .trim();
    return cleaned || "export";
  };
  const canDownloadText = (): boolean => {
    if (typeof document === "undefined") return false;
    if (typeof Blob === "undefined") return false;
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return false;
    return true;
  };
  const downloadText = (text: string, filename: string, mime: string): void => {
    // Download behavior is browser-specific. Make this best-effort so unit tests/non-browser
    // contexts can still register commands without crashing.
    if (typeof document === "undefined") return;
    if (typeof Blob === "undefined") return;
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return;

    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    try {
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.rel = "noopener";
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } finally {
      URL.revokeObjectURL(url);
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
      if (typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true) {
        showCollabEditRejectedToast([{ rejectionKind: "fillCells", rejectionReason: "permission" }]);
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
          showCollabEditRejectedToast([{ rejectionKind: "mergeCells", rejectionReason: "permission" }]);
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
          let applied = true;
          switch (args.kind) {
            case "mergeCenter":
              applied = mergeCenter(doc, sheetId, rect, { label: args.title });
              break;
            case "mergeAcross":
              applied = mergeAcross(doc, sheetId, rect, { label: args.title });
              break;
            case "mergeCells":
              applied = mergeCells(doc, sheetId, rect, { label: args.title });
              break;
            case "unmergeCells":
              unmergeCells(doc, sheetId, rect, { label: args.title });
              break;
            default:
              break;
          }

          // `mergeCells` helpers return `false` when a merge is blocked by `DocumentController.canEditCell`
          // (permissions, missing encryption key, etc). Without a UX signal this looks like a silent no-op.
          // Only toast for multi-cell selections; single-cell "merges" are a no-op in Excel semantics.
          if (!applied && (rows > 1 || cols > 1)) {
            showCollabEditRejectedToast([
              { sheetId, row: rect.startRow, col: rect.startCol, rejectionKind: "cell", rejectionReason: "permission" },
            ]);
            focusGrid();
            return;
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
    t("command.home.number.moreFormats.custom"),
    async () => {
      try {
        // Mirror `applyFormattingToSelection` selection guards before prompting so users don't
        // enter a format code only to hit selection size caps (or read-only role restrictions)
        // when applying.
        if (isEditingFn()) return;

        const selectionRanges =
          typeof (app as any)?.getSelectionRanges === "function" ? ((app as any).getSelectionRanges() as any[]) : [];
        const ranges = Array.isArray(selectionRanges) ? selectionRanges : [];

        const rawLimits = typeof (app as any)?.getGridLimits === "function" ? (app as any).getGridLimits() : null;
        const limits = {
          maxRows:
            Number.isInteger((rawLimits as any)?.maxRows) && (rawLimits as any).maxRows > 0
              ? (rawLimits as any).maxRows
              : DEFAULT_GRID_LIMITS.maxRows,
          maxCols:
            Number.isInteger((rawLimits as any)?.maxCols) && (rawLimits as any).maxCols > 0
              ? (rawLimits as any).maxCols
              : DEFAULT_GRID_LIMITS.maxCols,
        };

        const decision = evaluateFormattingSelectionSize(ranges as any, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });
        if (!decision.allowed) {
          safeShowToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
          return;
        }

        if (typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true && !decision.allRangesBand) {
          showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
          return;
        }

        const getSelectionNumberFormat = (): string | null => {
          // If the app isn't exposing selection ranges, fall back to the active cell (consistent with
          // other formatting commands that treat "no selection" as "active cell").
          if (!Array.isArray(ranges) || ranges.length === 0) {
            return getActiveCellNumberFormat();
          }
          const sheetId = (app as any).getCurrentSheetId?.() ?? null;
          if (typeof sheetId !== "string" || sheetId.trim() === "") {
            return getActiveCellNumberFormat();
          }
          // Be precise when the selection is within our normal formatting cell cap. When the selection is
          // a band (full row/col/sheet), the cell count can be enormous; fall back to sampling so the UI
          // remains responsive.
          const maxInspectCells = decision.allRangesBand ? 256 : decision.totalCells;
          const state = computeSelectionFormatState(app.getDocument(), sheetId, ranges as any, { maxInspectCells });
          return typeof state.numberFormat === "string" ? state.numberFormat : null;
        };

        await promptAndApplyCustomNumberFormat({
          isEditing: isEditingFn,
          showInputBox,
          getSelectionNumberFormat,
          applyFormattingToSelection: (label, fn) => applyFormattingToSelection(label, fn),
          showToast: safeShowToast,
        });
      } finally {
        focusGrid();
      }
    },
    { category: commandCategoryFormat },
  );

  // Ribbon schema uses `insert.illustrations.pictures.*` ids for Insert → Pictures.
  // Register them so the ribbon does not need to exempt them from CommandRegistry disabling.
  const commandCategoryInsert = "Insert";
  const registerInsertPicturesCommand = (
    commandId: string,
    title: string,
    options: { when?: string | null; run?: (() => void | Promise<void>) | null } = {},
  ): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      async () => {
        // Match other insert commands: don't open a file picker while the spreadsheet is in edit mode
        // (including split-view secondary editor state when the host provides `isEditing`).
        if (isEditingFn()) return;
        if (typeof options.run === "function") {
          await options.run();
          return;
        }
        await handleInsertPicturesRibbonCommand(commandId, app);
      },
      { category: commandCategoryInsert, when: options.when ?? null },
    );
  };
  registerInsertPicturesCommand("insert.illustrations.pictures", "Pictures…");
  registerInsertPicturesCommand("insert.illustrations.pictures.thisDevice", "Pictures: This Device…", {
    // `insert.illustrations.pictures` is treated as the canonical command (it maps to
    // "This Device" today). Keep the ribbon menu id registered as an alias so ribbon + recents
    // tracking still land on the canonical command id.
    when: "false",
    run: () => commandRegistry.executeCommand("insert.illustrations.pictures"),
  });
  // Stock Images is not implemented yet; keep it registered for ribbon coverage, but hide it from
  // context-aware UI surfaces (command palette, etc) until it is functional.
  registerInsertPicturesCommand("insert.illustrations.pictures.stockImages", "Pictures: Stock Images…", { when: "false" });
  registerInsertPicturesCommand("insert.illustrations.pictures.onlinePictures", "Pictures: Online Pictures…", {
    // Alias of the standalone Online Pictures command (both ids route through the same handler today).
    when: "false",
    run: () => commandRegistry.executeCommand("insert.illustrations.onlinePictures"),
  });
  // Online Pictures is not implemented yet; keep it registered for ribbon coverage, but hide it from
  // context-aware UI surfaces (command palette, etc) until it is functional.
  registerInsertPicturesCommand("insert.illustrations.onlinePictures", "Online Pictures…", { when: "false" });

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
  // - keybindings can invoke them (Excel: Ctrl+Shift+Plus / Ctrl+-),
  // - they appear in the command palette, and
  // - ribbon disabling can rely on CommandRegistry as the source of truth.
  const showToastForStructuralCellsCommand = (message: string, type?: "info" | "warning" | "error") => showToast(message, type);
  commandRegistry.registerBuiltinCommand(
    "home.cells.insert.insertCells",
    "Insert Cells…",
    async () => {
      // `handleHomeCellsInsertDeleteCommand` blocks while `app.isEditing()` is true and also
      // consults the desktop-shell-owned `__formulaSpreadsheetIsEditing` flag when present.
      // Still respect the caller-provided `isEditing` predicate (e.g. split view secondary editor
      // state) so we don't attempt structural edits while any editor is active (even if the global
      // flag is stale/unavailable).
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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

  // Home → Cells → Insert/Delete Sheet Rows/Columns.
  //
  // This logic lives in `ribbon/cellsStructuralCommands.ts` so it can be shared by both:
  // - the desktop ribbon fallback handlers (main.ts), and
  // - CommandRegistry registrations (so the ribbon can rely on baseline enable/disable).
  const registerCellsStructuralCommand = (commandId: string, title: string, keywords: string[]): void => {
    commandRegistry.registerBuiltinCommand(
      commandId,
      title,
      () => {
        // `executeCellsStructuralRibbonCommand` blocks while `app.isEditing()` is true and also
        // consults the desktop-shell-owned `__formulaSpreadsheetIsEditing` flag when present.
        // Still respect the caller-provided `isEditing` override (e.g. split view secondary editor
        // state) here so we don't attempt structural edits while any editor is active (even if the
        // global flag is stale/unavailable).
        if (isEditingFn()) return;
        executeCellsStructuralRibbonCommand(app, commandId);
      },
      { category: commandCategoryEditing, icon: null, keywords },
    );
  };
  registerCellsStructuralCommand("home.cells.insert.insertSheetRows", "Insert Sheet Rows", ["insert", "rows", "sheet rows", "excel"]);
  registerCellsStructuralCommand("home.cells.insert.insertSheetColumns", "Insert Sheet Columns", [
    "insert",
    "columns",
    "sheet columns",
    "excel",
  ]);
  registerCellsStructuralCommand("home.cells.delete.deleteSheetRows", "Delete Sheet Rows", ["delete", "rows", "sheet rows", "excel"]);
  registerCellsStructuralCommand("home.cells.delete.deleteSheetColumns", "Delete Sheet Columns", [
    "delete",
    "columns",
    "sheet columns",
    "excel",
  ]);

  // Home → Cells → Insert/Delete Sheet.
  //
  // These ribbon ids are registered here so ribbon enable/disable can rely on CommandRegistry.
  // The actual workbook mutation logic lives in `main.ts` (sheet store + confirmations), so
  // we delegate to optional `sheetStructureHandlers`.
  const isReadOnly = (): boolean => {
    try {
      return typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true;
    } catch {
      return false;
    }
  };
  commandRegistry.registerBuiltinCommand(
    "home.cells.insert.insertSheet",
    "Insert Sheet",
    async () => {
      if (isEditingFn()) return;
      if (isReadOnly()) {
        safeShowToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        focusGrid();
        return;
      }
      const handler = sheetStructureHandlers?.insertSheet;
      if (!handler) {
        safeShowToast("Insert Sheet is not available in this environment.", "warning");
        return;
      }
      await handler();
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Insert a new sheet after the active sheet",
      keywords: ["insert", "sheet", "worksheet", "tab"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.cells.delete.deleteSheet",
    "Delete Sheet",
    async () => {
      if (isEditingFn()) return;
      if (isReadOnly()) {
        safeShowToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        focusGrid();
        return;
      }
      const handler = sheetStructureHandlers?.deleteActiveSheet;
      if (!handler) {
        safeShowToast("Delete Sheet is not available in this environment.", "warning");
        return;
      }
      await handler();
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Delete the active sheet",
      keywords: ["delete", "sheet", "worksheet", "tab"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.cells.format.organizeSheets",
    "Organize Sheets…",
    async () => {
      if (isEditingFn()) return;
      if (isReadOnly()) {
        safeShowToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        focusGrid();
        return;
      }
      const handler = sheetStructureHandlers?.openOrganizeSheets;
      if (!handler) {
        safeShowToast("Organize Sheets is not available in this environment.", "warning");
        return;
      }
      await handler();
    },
    {
      category: commandCategoryEditing,
      icon: null,
      description: "Open the Organize Sheets dialog",
      keywords: ["organize sheets", "sheets", "worksheets", "tabs", "reorder", "rename", "hide"],
    },
  );
  if (layoutController) {
    registerBuiltinCommands({
      commandRegistry,
      app,
      layoutController,
      isEditing: isEditingFn,
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

  // Ribbon schema uses Excel-like `file.*` ids, but the canonical desktop commands are the
  // `workbench.*` file commands (plus a few desktop-only export helpers). Register the ribbon ids
  // as hidden aliases so ribbon baseline enable/disable can rely on CommandRegistry registration
  // (without maintaining a growing exemption list).

  const selectionBoundingBox0Based = (): CellRange => {
    const active = app.getActiveCell();
    const selectionRanges = app.getSelectionRanges();
    if (!Array.isArray(selectionRanges) || selectionRanges.length === 0) {
      return { start: { row: active.row, col: active.col }, end: { row: active.row, col: active.col } };
    }
 
    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;
 
    for (const r of selectionRanges) {
      const startRow0 = Math.min(r.startRow, r.endRow);
      const endRow0 = Math.max(r.startRow, r.endRow);
      const startCol0 = Math.min(r.startCol, r.endCol);
      const endCol0 = Math.max(r.startCol, r.endCol);
      minRow = Math.min(minRow, startRow0);
      maxRow = Math.max(maxRow, endRow0);
      minCol = Math.min(minCol, startCol0);
      maxCol = Math.max(maxCol, endCol0);
    }
 
    if (!Number.isFinite(minRow) || !Number.isFinite(minCol) || !Number.isFinite(maxRow) || !Number.isFinite(maxCol)) {
      return { start: { row: active.row, col: active.col }, end: { row: active.row, col: active.col } };
    }
 
    return { start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } };
  };
 
  const clipBandSelectionToUsedRange = (sheetId: string, range0: CellRange): CellRange => {
    const normalized: CellRange = {
      start: { row: Math.min(range0.start.row, range0.end.row), col: Math.min(range0.start.col, range0.end.col) },
      end: { row: Math.max(range0.start.row, range0.end.row), col: Math.max(range0.start.col, range0.end.col) },
    };
 
    const rawLimits = typeof (app as any)?.getGridLimits === "function" ? (app as any).getGridLimits() : null;
    const limits = {
      maxRows:
        Number.isInteger((rawLimits as any)?.maxRows) && (rawLimits as any).maxRows > 0
          ? (rawLimits as any).maxRows
          : DEFAULT_GRID_LIMITS.maxRows,
      maxCols:
        Number.isInteger((rawLimits as any)?.maxCols) && (rawLimits as any).maxCols > 0
          ? (rawLimits as any).maxCols
          : DEFAULT_GRID_LIMITS.maxCols,
    };
 
    const active = app.getActiveCell();
    const activeCellFallback0: CellRange = {
      start: { row: active.row, col: active.col },
      end: { row: active.row, col: active.col },
    };
 
    const isFullHeight = normalized.start.row === 0 && normalized.end.row === limits.maxRows - 1;
    const isFullWidth = normalized.start.col === 0 && normalized.end.col === limits.maxCols - 1;
    if (!isFullHeight && !isFullWidth) return normalized;
 
    const used = app.getDocument().getUsedRange(sheetId);
    if (!used) return activeCellFallback0;
 
    const startRow = Math.max(normalized.start.row, used.startRow);
    const endRow = Math.min(normalized.end.row, used.endRow);
    const startCol = Math.max(normalized.start.col, used.startCol);
    const endCol = Math.min(normalized.end.col, used.endCol);
    const clipped =
      startRow <= endRow && startCol <= endCol
        ? { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } }
        : null;
    return clipped ?? activeCellFallback0;
  };
 
  const exportDelimitedText = (args: { delimiter: string; extension: string; mime: string; label: string }): void => {
    try {
      commitEditsForCommand();
      if (isEditingFn()) return;
      if (!canDownloadText()) return;
 
      const sheetId = app.getCurrentSheetId();
      const sheetName = typeof (app as any)?.getCurrentSheetDisplayName === "function" ? (app as any).getCurrentSheetDisplayName() : sheetId;
      const doc = app.getDocument();
 
      // Use the selection by default, but clip full-row/full-column/full-sheet selections to the
      // used range to avoid attempting to export millions of empty cells.
      const exportRange0 = clipBandSelectionToUsedRange(sheetId, selectionBoundingBox0Based());
 
      const csv = exportDocumentRangeToCsv(doc, sheetId, exportRange0 as any, { delimiter: args.delimiter });
      downloadText(csv, `${sanitizeFilename(sheetName)}.${args.extension}`, args.mime);
    } catch (err) {
      console.error(`Failed to export ${args.label}:`, err);
      safeShowToast(`Failed to export ${args.label}: ${String(err)}`, "error");
    } finally {
      focusGrid();
    }
  };
 
  commandRegistry.registerBuiltinCommand("file.new.new", "New", () => commandRegistry.executeCommand("workbench.newWorkbook"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand(
    "file.new.blankWorkbook",
    "Blank workbook",
    () => commandRegistry.executeCommand("workbench.newWorkbook"),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand("file.open.open", "Open…", () => commandRegistry.executeCommand("workbench.openWorkbook"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand("file.save.save", "Save", () => commandRegistry.executeCommand("workbench.saveWorkbook"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand("file.save.saveAs", "Save As…", () => commandRegistry.executeCommand("workbench.saveWorkbookAs"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand("file.save.saveAs.copy", "Save a Copy…", () => commandRegistry.executeCommand("workbench.saveWorkbookAs"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand(
    "file.save.saveAs.download",
    "Download a Copy",
    () => commandRegistry.executeCommand("workbench.saveWorkbookAs"),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand(
    "file.save.autoSave",
    "AutoSave",
    async (enabled?: boolean) => {
      try {
        // Allow both toggle-style invocation (no args) and ribbon toggle invocation (boolean arg).
        await commandRegistry.executeCommand("workbench.setAutoSaveEnabled", enabled);
      } finally {
        focusGrid();
      }
    },
    {
      category: commandCategoryFile,
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "file.info.protectWorkbook.encryptWithPassword",
    "Encrypt with Password…",
    async () => {
      if (encryptWithPassword) {
        await encryptWithPassword();
      } else {
        safeShowToast("Encrypt with Password is not available in this environment.", "warning");
      }
    },
    {
      category: commandCategoryFile,
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand("file.print.print", "Print…", () => commandRegistry.executeCommand("workbench.print"), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand(
    "file.print.printPreview",
    "Print Preview",
    () => commandRegistry.executeCommand("workbench.printPreview"),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand("file.options.close", "Close", () => commandRegistry.executeCommand("workbench.closeWorkbook"), {
    category: commandCategoryFile,
    when: "false",
  });
 
  commandRegistry.registerBuiltinCommand(
    "file.export.export.xlsx",
    "Excel Workbook",
    () => commandRegistry.executeCommand("workbench.saveWorkbookAs"),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand(
    "file.export.changeFileType.xlsx",
    "Change File Type: Excel Workbook",
    () => commandRegistry.executeCommand("workbench.saveWorkbookAs"),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand("file.export.export.csv", "Export: CSV", () => exportDelimitedText({ delimiter: ",", extension: "csv", mime: "text/csv", label: "CSV" }), {
    category: commandCategoryFile,
    when: "false",
  });
  commandRegistry.registerBuiltinCommand(
    "file.export.changeFileType.csv",
    "Change File Type: CSV",
    () => exportDelimitedText({ delimiter: ",", extension: "csv", mime: "text/csv", label: "CSV" }),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
  commandRegistry.registerBuiltinCommand(
    "file.export.changeFileType.tsv",
    "Change File Type: TSV",
    () =>
      exportDelimitedText({
        delimiter: "\t",
        extension: "tsv",
        mime: "text/tab-separated-values",
        label: "TSV",
      }),
    {
      category: commandCategoryFile,
      when: "false",
    },
  );
 
  if (pageLayoutHandlers) {
    commandRegistry.registerBuiltinCommand(
      "file.print.pageSetup",
      "Page Setup…",
      () => commandRegistry.executeCommand("pageLayout.pageSetup.pageSetupDialog"),
      {
        category: commandCategoryFile,
        when: "false",
      },
    );
    commandRegistry.registerBuiltinCommand(
      "file.print.pageSetup.printTitles",
      "Print Titles…",
      () => commandRegistry.executeCommand("pageLayout.pageSetup.pageSetupDialog"),
      {
        category: commandCategoryFile,
        when: "false",
      },
    );
    commandRegistry.registerBuiltinCommand(
      "file.print.pageSetup.margins",
      "Margins",
      () => commandRegistry.executeCommand("pageLayout.pageSetup.pageSetupDialog"),
      {
        category: commandCategoryFile,
        when: "false",
      },
    );
  
    commandRegistry.registerBuiltinCommand(
      "file.export.createPdf",
      "Create PDF/XPS",
      () => commandRegistry.executeCommand("pageLayout.export.exportPdf"),
      {
        category: commandCategoryFile,
        when: "false",
      },
    );
    commandRegistry.registerBuiltinCommand("file.export.export.pdf", "Export: PDF", () => commandRegistry.executeCommand("pageLayout.export.exportPdf"), {
      category: commandCategoryFile,
      when: "false",
    });
    commandRegistry.registerBuiltinCommand(
      "file.export.changeFileType.pdf",
      "Change File Type: PDF",
      () => commandRegistry.executeCommand("pageLayout.export.exportPdf"),
      {
        category: commandCategoryFile,
        when: "false",
      },
    );
  }
 
  if (layoutController) {
    const registerPanelAlias = (id: string, title: string, targetId: string) => {
      commandRegistry.registerBuiltinCommand(
        id,
        title,
        async () => {
          if (commandRegistry.getCommand(targetId)) {
            await commandRegistry.executeCommand(targetId);
          } else {
            safeShowToast("This panel is not available in this environment.", "warning");
          }
        },
        { category: commandCategoryFile, when: "false" },
      );
    };
    registerPanelAlias("file.info.manageWorkbook.versions", t("panels.versionHistory.title"), "view.togglePanel.versionHistory");
    registerPanelAlias("file.info.manageWorkbook.branches", t("branchManager.title"), "view.togglePanel.branchManager");
  }
 
  if (formatPainter) {
    registerFormatPainterCommand({
      commandRegistry,
      ...formatPainter,
      isEditing: isEditingFn,
      isReadOnly,
    });
  }

  if (ribbonMacroHandlers) {
    registerRibbonMacroCommands({ commandRegistry, handlers: ribbonMacroHandlers, isEditing: isEditingFn, isReadOnly });
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
      isEditing: isEditingFn,
      isReadOnly,
      ...dataQueriesHandlers,
      focusAfterExecute,
    });
  }

  registerBuiltinFormatFontCommands({
    commandRegistry,
    applyFormattingToSelection,
  });

  registerSortFilterCommands({ commandRegistry, app, isEditing });

  // Data → Sort & Filter: MVP AutoFilter commands.
  //
  // These ids exist in `defaultRibbonSchema` but the MVP implementation lives in `main.ts`
  // (local-only UI + store). Register them in CommandRegistry so baseline ribbon disabling
  // can rely on registration instead of exemptions.
  const getAutoFilterHandlers = (): RibbonAutoFilterCommandHandlers | null => {
    const handlers = autoFilterHandlers ?? null;
    if (!handlers) {
      safeShowToast("Filtering is not available in this environment.", "warning");
      focusGrid();
      return null;
    }
    return handlers;
  };
  registerRibbonAutoFilterCommands({
    commandRegistry,
    getHandlers: getAutoFilterHandlers,
    isEditing: isEditingFn,
    category: commandCategoryData,
  });

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
        void (result as Promise<void>)
          .then(() => focusAlignmentSection())
          .catch(() => {
            // Best-effort: focus follow-up should never surface as an unhandled rejection.
          });
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
  const commandCategoryPageLayout = t("commandCategory.pageLayout");
  commandRegistry.registerBuiltinCommand(
    "pageLayout.arrange.bringForward",
    "Bring Forward",
    () => {
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      const labelByChoice: Record<Choice, string> = {
        general: t("command.format.numberFormat.general"),
        currency: t("command.format.numberFormat.currency"),
        percent: t("command.format.numberFormat.percent"),
        date: t("command.format.numberFormat.date"),
      };
      const choice = await showQuickPick<Choice>(
        [
          { label: labelByChoice.general, description: t("quickPick.numberFormat.general.description"), value: "general" },
          { label: labelByChoice.currency, description: NUMBER_FORMATS.currency, value: "currency" },
          { label: labelByChoice.percent, description: NUMBER_FORMATS.percent, value: "percent" },
          { label: labelByChoice.date, description: NUMBER_FORMATS.date, value: "date" },
        ],
        { placeHolder: t("quickPick.numberFormat.placeholder") },
      );
      if (!choice) return;

      const patch = choice === "general" ? { numberFormat: null } : { numberFormat: NUMBER_FORMATS[choice] };

      applyFormattingToSelection(
        labelByChoice[choice],
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

  commandRegistry.registerBuiltinCommand("edit.find", t("command.edit.find"), () => {
    if (isEditingFn()) return;
    findReplace.openFind();
  }, {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.find"),
    keywords: ["find", "search"],
  });

  commandRegistry.registerBuiltinCommand("edit.replace", t("command.edit.replace"), () => {
    if (isEditingFn()) return;
    findReplace.openReplace();
  }, {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.replace"),
    keywords: ["replace", "find"],
  });

  commandRegistry.registerBuiltinCommand("navigation.goTo", t("command.navigation.goTo"), () => {
    if (isEditingFn()) return;
    findReplace.openGoTo();
  }, {
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
