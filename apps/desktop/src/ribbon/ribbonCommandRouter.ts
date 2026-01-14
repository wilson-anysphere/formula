import type { SpreadsheetApp } from "../app/spreadsheetApp.js";
import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import { isSpreadsheetEditingCommandBlockedError } from "../commands/spreadsheetEditingCommandBlockedError.js";
import { PAGE_LAYOUT_COMMANDS } from "../commands/registerPageLayoutCommands.js";
import { WORKBENCH_FILE_COMMANDS } from "../commands/registerWorkbenchFileCommands.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";
import { promptAndApplyCustomNumberFormat } from "../formatting/promptCustomNumberFormat.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize } from "../formatting/selectionSizeGuard.js";
import { DEFAULT_DESKTOP_LOAD_MAX_COLS, DEFAULT_DESKTOP_LOAD_MAX_ROWS } from "../workbook/load/clampUsedRange.js";
import { handleInsertPicturesRibbonCommand } from "../main.insertPicturesRibbonCommand.js";

import { createRibbonActionsFromCommands, createRibbonFileActionsFromCommands } from "./createRibbonActionsFromCommands.js";
import { executeCellsStructuralRibbonCommand } from "./cellsStructuralCommands.js";
import { handleRibbonCommand as handleRibbonFormattingCommand, handleRibbonToggle as handleRibbonFormattingToggle } from "./commandHandlers.js";
import { handleHomeCellsInsertDeleteCommand } from "./homeCellsCommands.js";
import { computeSelectionFormatState } from "./selectionFormatState.js";
import type { RibbonActions } from "./ribbonSchema.js";

export type RibbonToastType = "info" | "success" | "warning" | "error";

export type RibbonQuickPickItem<T> = {
  label: string;
  value: T;
  description?: string;
  detail?: string;
};

export interface RibbonCommandRouterDeps {
  app: SpreadsheetApp;
  commandRegistry: CommandRegistry;
  /**
   * Spreadsheet edit state guard.
   *
   * Desktop uses a custom predicate that includes split-view secondary editor state; callers
   * should pass that in so ribbon-only handlers match the rest of the shell behavior.
   */
  isSpreadsheetEditing: () => boolean;

  // UI helpers.
  showToast: (message: string, type?: RibbonToastType, options?: { timeoutMs?: number }) => void;
  showQuickPick: <T>(items: RibbonQuickPickItem<T>[], options?: { placeHolder?: string }) => Promise<T | null>;
  showInputBox: (options: { prompt?: string; value?: string; placeHolder?: string }) => Promise<string | null>;

  // Sheet/dialog helpers.
  openOrganizeSheets: () => void;
  handleAddSheet: () => Promise<void>;
  handleDeleteActiveSheet: () => Promise<void>;
  openCustomSortDialog: (commandId: string) => void;

  // AutoFilter MVP wiring.
  toggleAutoFilter: () => void;
  clearAutoFilter: () => void;
  reapplyAutoFilter: () => void;

  // Formatting helpers used by ribbon-only handlers.
  applyFormattingToSelection: (
    label: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    fn: (doc: any, sheetId: string, ranges: any[]) => void | boolean,
    options?: { forceBatch?: boolean; allowReadOnlyBandSelection?: boolean },
  ) => void;
  getActiveCellNumberFormat: () => string | null;

  // Extension lazy-load refs for extension-contributed commands.
  getEnsureExtensionsLoadedRef: () => (() => Promise<void>) | null;
  getSyncContributedCommandsRef: () => (() => void) | null;
}

function safeFocusGrid(app: { focus: () => void }): void {
  try {
    app.focus();
  } catch {
    // Best-effort focus restore.
  }
}

function reportRibbonCommandError(deps: RibbonCommandRouterDeps, commandId: string, err: unknown): void {
  // DLP policy violations are already surfaced via a dedicated toast (e.g. clipboard copy blocked).
  // Avoid double-toasting "Command failed" for expected policy restrictions.
  if ((err as any)?.name === "DlpViolationError") return;
  // Edit-mode command guards already surface a dedicated warning toast.
  if (isSpreadsheetEditingCommandBlockedError(err)) return;
  deps.showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
}

function unimplementedToast(deps: RibbonCommandRouterDeps, commandId: string): void {
  if (commandId.startsWith("file.")) {
    deps.showToast(`File command not implemented: ${commandId}`);
  } else {
    deps.showToast(`Ribbon: ${commandId}`);
  }
  safeFocusGrid(deps.app);
}

export function createRibbonActions(deps: RibbonCommandRouterDeps): RibbonActions {
  const executeCommand = (commandId: string, ...args: any[]): void => {
    void deps.commandRegistry.executeCommand(commandId, ...args).catch((err) => reportRibbonCommandError(deps, commandId, err));
  };

  const getGridLimitsForFormatting = (): { maxRows: number; maxCols: number } => {
    const raw = deps.app.getGridLimits();
    const maxRows =
      Number.isInteger(raw?.maxRows) && (raw.maxRows as number) > 0 ? (raw.maxRows as number) : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
    const maxCols =
      Number.isInteger(raw?.maxCols) && (raw.maxCols as number) > 0 ? (raw.maxCols as number) : DEFAULT_DESKTOP_LOAD_MAX_COLS;
    return { maxRows, maxCols };
  };

  const promptCustomNumberFormat = (): void => {
    if (deps.isSpreadsheetEditing()) return;

    // Guard before prompting so users don't enter a format code only to hit selection size caps on apply.
    // (Matches `applyFormattingToSelection` behavior.)
    const selection = deps.app.getSelectionRanges();
    const limits = getGridLimitsForFormatting();
    const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });
    if (!decision.allowed) {
      deps.showToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
      safeFocusGrid(deps.app);
      return;
    }

    if (deps.app.isReadOnly?.() === true && !decision.allRangesBand) {
      showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
      safeFocusGrid(deps.app);
      return;
    }

    const getSelectionNumberFormat = (): string | null => {
      const ranges = selection;
      if (!Array.isArray(ranges) || ranges.length === 0) return deps.getActiveCellNumberFormat();
      // Be exhaustive when selections are under the standard formatting cap, but sample for large band selections.
      const maxInspectCells = decision.allRangesBand ? 256 : decision.totalCells;
      const state = computeSelectionFormatState(deps.app.getDocument(), deps.app.getCurrentSheetId(), ranges, { maxInspectCells });
      return typeof state.numberFormat === "string" ? state.numberFormat : null;
    };

    void promptAndApplyCustomNumberFormat({
      isEditing: () => deps.isSpreadsheetEditing(),
      showInputBox: deps.showInputBox,
      getSelectionNumberFormat,
      applyFormattingToSelection: deps.applyFormattingToSelection,
      showToast: (message, type) => deps.showToast(message, type),
    })
      .finally(() => safeFocusGrid(deps.app))
      .catch(() => {
        // Best-effort: avoid unhandled rejections if the `.finally` bookkeeping throws.
      });
  };

  // Context for ribbon formatting helpers (`ribbon/commandHandlers.ts`).
  const ribbonCommandHandlersCtx = {
    app: deps.app,
    isEditing: deps.isSpreadsheetEditing,
    applyFormattingToSelection: deps.applyFormattingToSelection,
    showToast: deps.showToast,
    executeCommand: (commandId: string, ...args: any[]) => executeCommand(commandId, ...args),
    openCustomSort: deps.openCustomSortDialog,
    promptCustomNumberFormat,
    toggleAutoFilter: deps.toggleAutoFilter,
    clearAutoFilter: deps.clearAutoFilter,
    reapplyAutoFilter: deps.reapplyAutoFilter,
  };

  const handleRibbonCommand = (commandId: string): void => {
    if (handleRibbonFormattingCommand(ribbonCommandHandlersCtx, commandId)) {
      return;
    }

    switch (commandId) {
      case "home.cells.insert":
      case "home.cells.delete":
      case "home.cells.format":
        // These ids are dropdown triggers (menu containers). They should not normally fire as commands,
        // but some ribbon interactions can surface them via `onCommand`. Treat them as no-ops to avoid
        // spurious "Ribbon: ..." toasts.
        return;
      case "insert.illustrations.pictures":
      case "insert.illustrations.pictures.thisDevice":
      case "insert.illustrations.pictures.stockImages":
      case "insert.illustrations.pictures.onlinePictures":
      case "insert.illustrations.onlinePictures":
        if (deps.isSpreadsheetEditing()) return;
        void handleInsertPicturesRibbonCommand(commandId, deps.app);
        return;

      case "home.cells.format.organizeSheets":
        if (deps.isSpreadsheetEditing()) return;
        if (deps.app.isReadOnly()) {
          deps.showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
          safeFocusGrid(deps.app);
          return;
        }
        deps.openOrganizeSheets();
        return;

      case "home.cells.insert.insertCells":
      case "home.cells.delete.deleteCells":
        if (deps.isSpreadsheetEditing()) return;
        void handleHomeCellsInsertDeleteCommand({
          app: deps.app,
          commandId,
          showQuickPick: deps.showQuickPick,
          showToast: deps.showToast,
        }).catch((err) => reportRibbonCommandError(deps, commandId, err));
        return;

      case "home.cells.insert.insertSheetRows":
      case "home.cells.insert.insertSheetColumns":
      case "home.cells.delete.deleteSheetRows":
      case "home.cells.delete.deleteSheetColumns":
        if (deps.isSpreadsheetEditing()) return;
        executeCellsStructuralRibbonCommand(deps.app, commandId);
        return;

      case "home.cells.insert.insertSheet":
        if (deps.isSpreadsheetEditing()) return;
        if (deps.app.isReadOnly()) {
          deps.showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
          safeFocusGrid(deps.app);
          return;
        }
        void deps.handleAddSheet().catch((err) => reportRibbonCommandError(deps, commandId, err));
        return;

      case "home.cells.delete.deleteSheet":
        if (deps.isSpreadsheetEditing()) return;
        if (deps.app.isReadOnly()) {
          deps.showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
          safeFocusGrid(deps.app);
          return;
        }
        void deps.handleDeleteActiveSheet().catch((err) => reportRibbonCommandError(deps, commandId, err));
        return;

      default:
        unimplementedToast(deps, commandId);
        return;
    }
  };

  const ribbonActions = createRibbonActionsFromCommands({
    commandRegistry: deps.commandRegistry,
    onCommandError: (commandId, err) => reportRibbonCommandError(deps, commandId, err),
    onUnknownToggle: (commandId, pressed) => {
      const handled = handleRibbonFormattingToggle(ribbonCommandHandlersCtx, commandId, pressed);
      if (handled) return true;

      // Toggle buttons invoke `onToggle` (not `onCommand`) in `Ribbon`. Route unknown toggles
      // through the same fallback handler so unimplemented buttons still surface a toast.
      handleRibbonCommand(commandId);
      return true;
    },
    onBeforeExecuteCommand: async (_commandId, source: CommandContribution["source"]) => {
      if (source.kind !== "extension") return;
      // Match keybinding/command palette behavior: executing an extension command should
      // lazy-load the extension runtime first.
      await deps.getEnsureExtensionsLoadedRef()?.();
      deps.getSyncContributedCommandsRef()?.();
    },
    onUnknownCommand: handleRibbonCommand,
  });

  const fileActions = createRibbonFileActionsFromCommands({
    commandRegistry: deps.commandRegistry,
    onCommandError: (commandId, err) => reportRibbonCommandError(deps, commandId, err),
    commandIds: {
      newWorkbook: WORKBENCH_FILE_COMMANDS.newWorkbook,
      openWorkbook: WORKBENCH_FILE_COMMANDS.openWorkbook,
      saveWorkbook: WORKBENCH_FILE_COMMANDS.saveWorkbook,
      saveWorkbookAs: WORKBENCH_FILE_COMMANDS.saveWorkbookAs,
      toggleAutoSave: WORKBENCH_FILE_COMMANDS.setAutoSaveEnabled,
      versionHistory: "view.togglePanel.versionHistory",
      branchManager: "view.togglePanel.branchManager",
      pageSetup: PAGE_LAYOUT_COMMANDS.pageSetupDialog,
      printPreview: WORKBENCH_FILE_COMMANDS.printPreview,
      print: WORKBENCH_FILE_COMMANDS.print,
      closeWindow: WORKBENCH_FILE_COMMANDS.closeWorkbook,
      quit: WORKBENCH_FILE_COMMANDS.quit,
    },
  });

  return { ...ribbonActions, fileActions };
}

/**
 * Ribbon schema coverage: every schema id should be either "handled" (has a real implementation,
 * whether via CommandRegistry or in this router) or explicitly allowlisted as intentionally
 * unimplemented.
 */
export const handledRibbonCommandIds = new Set<string>([
  "ai.inlineEdit",
  "audit.toggleDependents",
  "audit.togglePrecedents",
  "audit.toggleTransitive",
  "clipboard.copy",
  "clipboard.cut",
  "clipboard.paste",
  "clipboard.pasteSpecial",
  "clipboard.pasteSpecial.formats",
  "clipboard.pasteSpecial.formulas",
  "clipboard.pasteSpecial.transpose",
  "clipboard.pasteSpecial.values",
  "comments.addComment",
  "comments.togglePanel",
  "data.forecast.whatIfAnalysis.goalSeek",
  "data.forecast.whatIfAnalysis.monteCarlo",
  "data.forecast.whatIfAnalysis.scenarioManager",
  "data.queriesConnections.queriesConnections",
  "data.queriesConnections.refreshAll",
  "data.queriesConnections.refreshAll.refresh",
  "data.queriesConnections.refreshAll.refreshAllConnections",
  "data.queriesConnections.refreshAll.refreshAllQueries",
  "data.sortFilter.advanced.clearFilter",
  "data.sortFilter.clear",
  "data.sortFilter.filter",
  "data.sortFilter.reapply",
  "data.sortFilter.sort.customSort",
  "data.sortFilter.sortAtoZ",
  "data.sortFilter.sortZtoA",
  "developer.code.macroSecurity",
  "developer.code.macroSecurity.trustCenter",
  "developer.code.macros",
  "developer.code.macros.edit",
  "developer.code.macros.run",
  "developer.code.recordMacro",
  "developer.code.recordMacro.stop",
  "developer.code.useRelativeReferences",
  "developer.code.visualBasic",
  "edit.autoSum",
  "edit.clearContents",
  "edit.fillDown",
  "edit.fillLeft",
  "edit.fillRight",
  "edit.fillUp",
  "edit.find",
  "edit.replace",
  "file.export.changeFileType.csv",
  "file.export.changeFileType.pdf",
  "file.export.changeFileType.tsv",
  "file.export.changeFileType.xlsx",
  "file.export.createPdf",
  "file.export.export.csv",
  "file.export.export.pdf",
  "file.export.export.xlsx",
  "file.info.manageWorkbook.branches",
  "file.info.manageWorkbook.versions",
  "file.info.protectWorkbook.encryptWithPassword",
  "file.new.blankWorkbook",
  "file.new.new",
  "file.open.open",
  "file.options.close",
  "file.print.pageSetup",
  "file.print.pageSetup.margins",
  "file.print.pageSetup.printTitles",
  "file.print.print",
  "file.print.printPreview",
  "file.save.autoSave",
  "file.save.save",
  "file.save.saveAs",
  "file.save.saveAs.copy",
  "file.save.saveAs.download",
  "format.alignBottom",
  "format.alignCenter",
  "format.alignLeft",
  "format.alignMiddle",
  "format.alignRight",
  "format.alignTop",
  "format.borders.all",
  "format.borders.bottom",
  "format.borders.left",
  "format.borders.none",
  "format.borders.outside",
  "format.borders.right",
  "format.borders.thickBox",
  "format.borders.top",
  "format.clearAll",
  "format.clearFormats",
  "format.fontSize.decrease",
  "format.decreaseIndent",
  "format.fillColor.blue",
  "format.fillColor.green",
  "format.fillColor.lightGray",
  "format.fillColor.moreColors",
  "format.fillColor.none",
  "format.fillColor.red",
  "format.fillColor.yellow",
  "format.fontColor.automatic",
  "format.fontColor.black",
  "format.fontColor.blue",
  "format.fontColor.green",
  "format.fontColor.moreColors",
  "format.fontColor.red",
  "format.fontName.arial",
  "format.fontName.calibri",
  "format.fontName.courier",
  "format.fontName.times",
  "format.fontSize.10",
  "format.fontSize.11",
  "format.fontSize.12",
  "format.fontSize.14",
  "format.fontSize.16",
  "format.fontSize.18",
  "format.fontSize.20",
  "format.fontSize.24",
  "format.fontSize.28",
  "format.fontSize.36",
  "format.fontSize.48",
  "format.fontSize.72",
  "format.fontSize.8",
  "format.fontSize.9",
  "format.fontSize.increase",
  "format.increaseIndent",
  "format.numberFormat.accounting",
  "format.numberFormat.accounting.eur",
  "format.numberFormat.accounting.gbp",
  "format.numberFormat.accounting.jpy",
  "format.numberFormat.accounting.usd",
  "format.numberFormat.commaStyle",
  "format.numberFormat.currency",
  "format.numberFormat.decreaseDecimal",
  "format.numberFormat.fraction",
  "format.numberFormat.general",
  "format.numberFormat.increaseDecimal",
  "format.numberFormat.longDate",
  "format.numberFormat.number",
  "format.numberFormat.percent",
  "format.numberFormat.scientific",
  "format.numberFormat.shortDate",
  "format.numberFormat.text",
  "format.numberFormat.time",
  "format.openAlignmentDialog",
  "format.openFormatCells",
  "format.textRotation.angleClockwise",
  "format.textRotation.angleCounterclockwise",
  "format.textRotation.rotateDown",
  "format.textRotation.rotateUp",
  "format.textRotation.verticalText",
  "format.toggleBold",
  "format.toggleFormatPainter",
  "format.toggleItalic",
  "format.toggleStrikethrough",
  "format.toggleUnderline",
  "format.toggleWrapText",
  "formulas.formulaAuditing.removeArrows",
  "formulas.formulaAuditing.traceDependents",
  "formulas.formulaAuditing.tracePrecedents",
  "formulas.solutions.solver",
  "home.alignment.mergeCenter",
  "home.alignment.mergeCenter.mergeAcross",
  "home.alignment.mergeCenter.mergeCells",
  "home.alignment.mergeCenter.mergeCenter",
  "home.alignment.mergeCenter.unmergeCells",
  "home.alignment.orientation",
  "home.cells.delete.deleteCells",
  "home.cells.delete.deleteSheet",
  "home.cells.delete.deleteSheetColumns",
  "home.cells.delete.deleteSheetRows",
  "home.cells.format",
  "home.cells.format.columnWidth",
  "home.cells.format.organizeSheets",
  "home.cells.format.rowHeight",
  "home.cells.insert.insertCells",
  "home.cells.insert.insertSheet",
  "home.cells.insert.insertSheetColumns",
  "home.cells.insert.insertSheetRows",
  "home.editing.autoSum.average",
  "home.editing.autoSum.countNumbers",
  "home.editing.autoSum.max",
  "home.editing.autoSum.min",
  "home.editing.fill.series",
  "home.editing.sortFilter.customSort",
  "home.font.borders",
  "home.font.clearFormatting",
  "home.font.fillColor",
  "home.font.fontColor",
  "home.font.fontName",
  "home.font.fontSize",
  "format.toggleSubscript",
  "format.toggleSuperscript",
  "home.number.moreFormats",
  "home.number.moreFormats.custom",
  "home.number.numberFormat",
  "home.styles.cellStyles.goodBadNeutral",
  "home.styles.formatAsTable.dark",
  "home.styles.formatAsTable.light",
  "home.styles.formatAsTable.medium",
  "insert.illustrations.onlinePictures",
  "insert.illustrations.pictures",
  "insert.illustrations.pictures.onlinePictures",
  "insert.illustrations.pictures.stockImages",
  "insert.illustrations.pictures.thisDevice",
  "insert.tables.pivotTable.fromTableRange",
  "navigation.goTo",
  "pageLayout.arrange.bringForward",
  "pageLayout.arrange.sendBackward",
  "pageLayout.export.exportPdf",
  "pageLayout.pageSetup.margins.custom",
  "pageLayout.pageSetup.margins.narrow",
  "pageLayout.pageSetup.margins.normal",
  "pageLayout.pageSetup.margins.wide",
  "pageLayout.pageSetup.orientation.landscape",
  "pageLayout.pageSetup.orientation.portrait",
  "pageLayout.pageSetup.pageSetupDialog",
  "pageLayout.pageSetup.printArea.addTo",
  "pageLayout.pageSetup.printArea.clear",
  "pageLayout.pageSetup.printArea.set",
  "pageLayout.pageSetup.size.a4",
  "pageLayout.pageSetup.size.letter",
  "pageLayout.pageSetup.size.more",
  "pageLayout.printArea.clearPrintArea",
  "pageLayout.printArea.setPrintArea",
  "view.appearance.theme",
  "view.appearance.theme.dark",
  "view.appearance.theme.highContrast",
  "view.appearance.theme.light",
  "view.appearance.theme.system",
  "view.freezeFirstColumn",
  "view.freezePanes",
  "view.freezeTopRow",
  "view.insertPivotTable",
  "view.macros.recordMacro",
  "view.macros.recordMacro.stop",
  "view.macros.useRelativeReferences",
  "view.macros.viewMacros",
  "view.macros.viewMacros.delete",
  "view.macros.viewMacros.edit",
  "view.macros.viewMacros.run",
  "view.splitHorizontal",
  "view.splitNone",
  "view.splitVertical",
  "view.togglePanel.aiAudit",
  "view.togglePanel.aiChat",
  "view.togglePanel.branchManager",
  "view.togglePanel.dataQueries",
  "view.togglePanel.extensions",
  "view.togglePanel.macros",
  "view.togglePanel.marketplace",
  "view.togglePanel.python",
  "view.togglePanel.scriptEditor",
  "view.togglePanel.vbaMigrate",
  "view.togglePanel.versionHistory",
  "view.togglePerformanceStats",
  "view.toggleShowFormulas",
  "view.toggleSplitView",
  "view.unfreezePanes",
  "view.zoom.openPicker",
  "view.zoom.zoom",
  "view.zoom.zoom100",
  "view.zoom.zoom150",
  "view.zoom.zoom200",
  "view.zoom.zoom25",
  "view.zoom.zoom400",
  "view.zoom.zoom50",
  "view.zoom.zoom75",
  "view.zoom.zoomToSelection",
]);

export const unimplementedRibbonCommandIds = new Set<string>([
  "data.dataTools.consolidate",
  "data.dataTools.dataValidation",
  "data.dataTools.dataValidation.circleInvalid",
  "data.dataTools.dataValidation.clearCircles",
  "data.dataTools.flashFill",
  "data.dataTools.manageDataModel",
  "data.dataTools.manageDataModel.addToDataModel",
  "data.dataTools.relationships",
  "data.dataTools.relationships.manage",
  "data.dataTools.removeDuplicates",
  "data.dataTools.removeDuplicates.advanced",
  "data.dataTools.textToColumns",
  "data.dataTools.textToColumns.reapply",
  "data.dataTypes.geography",
  "data.dataTypes.stocks",
  "data.forecast.forecastSheet",
  "data.forecast.forecastSheet.options",
  "data.forecast.whatIfAnalysis",
  "data.forecast.whatIfAnalysis.dataTable",
  "data.getTransform.existingConnections",
  "data.getTransform.getData",
  "data.getTransform.getData.fromAzure",
  "data.getTransform.getData.fromDatabase",
  "data.getTransform.getData.fromFile",
  "data.getTransform.getData.fromOnlineServices",
  "data.getTransform.getData.fromOtherSources",
  "data.getTransform.recentSources",
  "data.outline.group",
  "data.outline.group.group",
  "data.outline.group.groupSelection",
  "data.outline.hideDetail",
  "data.outline.showDetail",
  "data.outline.subtotal",
  "data.outline.ungroup",
  "data.outline.ungroup.clearOutline",
  "data.outline.ungroup.ungroup",
  "data.queriesConnections.properties",
  "data.sortFilter.advanced",
  "data.sortFilter.advanced.advancedFilter",
  "data.sortFilter.sort",
  "developer.addins.addins",
  "developer.addins.addins.browse",
  "developer.addins.addins.excelAddins",
  "developer.addins.addins.manage",
  "developer.addins.comAddins",
  "developer.controls.designMode",
  "developer.controls.insert",
  "developer.controls.insert.button",
  "developer.controls.insert.checkbox",
  "developer.controls.insert.combobox",
  "developer.controls.insert.listbox",
  "developer.controls.insert.scrollbar",
  "developer.controls.insert.spinButton",
  "developer.controls.properties",
  "developer.controls.properties.viewProperties",
  "developer.controls.runDialog",
  "developer.controls.viewCode",
  "developer.xml.export",
  "developer.xml.import",
  "developer.xml.mapProperties",
  "developer.xml.refreshData",
  "developer.xml.source",
  "developer.xml.source.refresh",
  "file.export.changeFileType",
  "file.export.export",
  "file.info.inspectWorkbook",
  "file.info.inspectWorkbook.checkAccessibility",
  "file.info.inspectWorkbook.checkCompatibility",
  "file.info.inspectWorkbook.documentInspector",
  "file.info.manageWorkbook",
  "file.info.manageWorkbook.properties",
  "file.info.manageWorkbook.recoverUnsaved",
  "file.info.protectWorkbook",
  "file.info.protectWorkbook.protectCurrentSheet",
  "file.info.protectWorkbook.protectWorkbookStructure",
  "file.new.fromExisting",
  "file.new.templates",
  "file.new.templates.budget",
  "file.new.templates.calendar",
  "file.new.templates.invoice",
  "file.new.templates.more",
  "file.open.pinned",
  "file.open.pinned.kpis",
  "file.open.pinned.q4",
  "file.open.recent",
  "file.open.recent.book1",
  "file.open.recent.budget",
  "file.open.recent.forecast",
  "file.open.recent.more",
  "file.options.account",
  "file.options.options",
  "file.share.email",
  "file.share.email.attachment",
  "file.share.email.link",
  "file.share.presentOnline",
  "file.share.share",
  "formulas.calculation.calculateNow",
  "formulas.calculation.calculateSheet",
  "formulas.calculation.calculationOptions",
  "formulas.definedNames.createFromSelection",
  "formulas.definedNames.defineName",
  "formulas.definedNames.nameManager",
  "formulas.definedNames.useInFormula",
  "formulas.formulaAuditing.errorChecking",
  "formulas.formulaAuditing.evaluateFormula",
  "formulas.formulaAuditing.watchWindow",
  "formulas.functionLibrary.autoSum",
  "formulas.functionLibrary.autoSum.average",
  "formulas.functionLibrary.autoSum.countNumbers",
  "formulas.functionLibrary.autoSum.max",
  "formulas.functionLibrary.autoSum.min",
  "formulas.functionLibrary.autoSum.moreFunctions",
  "formulas.functionLibrary.autoSum.sum",
  "formulas.functionLibrary.dateTime",
  "formulas.functionLibrary.financial",
  "formulas.functionLibrary.insertFunction",
  "formulas.functionLibrary.logical",
  "formulas.functionLibrary.lookupReference",
  "formulas.functionLibrary.mathTrig",
  "formulas.functionLibrary.moreFunctions",
  "formulas.functionLibrary.recentlyUsed",
  "formulas.functionLibrary.text",
  "formulas.solutions.analysisToolPak",
  "help.support.contactSupport",
  "help.support.feedback",
  "help.support.help",
  "help.support.training",
  "home.cells.delete",
  "home.cells.insert",
  "home.clipboard.clipboardPane",
  "home.clipboard.clipboardPane.clearAll",
  "home.clipboard.clipboardPane.open",
  "home.clipboard.clipboardPane.options",
  "home.editing.autoSum.moreFunctions",
  "home.editing.clear",
  "home.editing.clear.clearComments",
  "home.editing.clear.clearHyperlinks",
  "home.editing.fill",
  "home.editing.findSelect",
  "home.editing.sortFilter",
  "home.styles.cellStyles",
  "home.styles.cellStyles.dataModel",
  "home.styles.cellStyles.newStyle",
  "home.styles.cellStyles.numberFormat",
  "home.styles.cellStyles.titlesHeadings",
  "home.styles.conditionalFormatting",
  "home.styles.conditionalFormatting.clearRules",
  "home.styles.conditionalFormatting.colorScales",
  "home.styles.conditionalFormatting.dataBars",
  "home.styles.conditionalFormatting.highlightCellsRules",
  "home.styles.conditionalFormatting.iconSets",
  "home.styles.conditionalFormatting.manageRules",
  "home.styles.conditionalFormatting.topBottomRules",
  "home.styles.formatAsTable",
  "home.styles.formatAsTable.newStyle",
  "insert.addins.getAddins",
  "insert.addins.myAddins",
  "insert.charts.area",
  "insert.charts.area.area",
  "insert.charts.area.more",
  "insert.charts.area.stackedArea",
  "insert.charts.bar",
  "insert.charts.bar.clusteredBar",
  "insert.charts.bar.more",
  "insert.charts.bar.stackedBar",
  "insert.charts.boxWhisker",
  "insert.charts.column",
  "insert.charts.column.clusteredColumn",
  "insert.charts.column.more",
  "insert.charts.column.stackedColumn",
  "insert.charts.column.stackedColumn100",
  "insert.charts.combo",
  "insert.charts.funnel",
  "insert.charts.histogram",
  "insert.charts.line",
  "insert.charts.line.line",
  "insert.charts.line.lineWithMarkers",
  "insert.charts.line.more",
  "insert.charts.line.stackedArea",
  "insert.charts.map",
  "insert.charts.map.filledMap",
  "insert.charts.map.more",
  "insert.charts.pie",
  "insert.charts.pie.doughnut",
  "insert.charts.pie.more",
  "insert.charts.pie.pie",
  "insert.charts.pivotChart",
  "insert.charts.radar",
  "insert.charts.recommendedCharts",
  "insert.charts.recommendedCharts.column",
  "insert.charts.recommendedCharts.line",
  "insert.charts.recommendedCharts.more",
  "insert.charts.recommendedCharts.pie",
  "insert.charts.scatter",
  "insert.charts.scatter.more",
  "insert.charts.scatter.scatter",
  "insert.charts.scatter.smoothLines",
  "insert.charts.stock",
  "insert.charts.sunburst",
  "insert.charts.surface",
  "insert.charts.treemap",
  "insert.charts.waterfall",
  "insert.comments.comment",
  "insert.comments.note",
  "insert.equations.equation",
  "insert.equations.inkEquation",
  "insert.filters.slicer",
  "insert.filters.slicer.reportConnections",
  "insert.filters.timeline",
  "insert.illustrations.icons",
  "insert.illustrations.screenshot",
  "insert.illustrations.shapes",
  "insert.illustrations.shapes.arrows",
  "insert.illustrations.shapes.basicShapes",
  "insert.illustrations.shapes.callouts",
  "insert.illustrations.shapes.flowchart",
  "insert.illustrations.shapes.lines",
  "insert.illustrations.shapes.rectangles",
  "insert.illustrations.smartArt",
  "insert.links.link",
  "insert.pivotcharts.pivotChart",
  "insert.pivotcharts.recommendedPivotCharts",
  "insert.sparklines.column",
  "insert.sparklines.line",
  "insert.sparklines.winLoss",
  "insert.symbols.equation",
  "insert.symbols.symbol",
  "insert.tables.pivotTable.fromDataModel",
  "insert.tables.pivotTable.fromExternal",
  "insert.tables.recommendedPivotTables",
  "insert.tables.table",
  "insert.text.headerFooter",
  "insert.text.object",
  "insert.text.signatureLine",
  "insert.text.textBox",
  "insert.text.wordArt",
  "insert.tours.3dMap",
  "insert.tours.launchTour",
  "pageLayout.arrange.align",
  "pageLayout.arrange.align.alignBottom",
  "pageLayout.arrange.align.alignCenter",
  "pageLayout.arrange.align.alignLeft",
  "pageLayout.arrange.align.alignMiddle",
  "pageLayout.arrange.align.alignRight",
  "pageLayout.arrange.align.alignTop",
  "pageLayout.arrange.group",
  "pageLayout.arrange.group.group",
  "pageLayout.arrange.group.regroup",
  "pageLayout.arrange.group.ungroup",
  "pageLayout.arrange.rotate",
  "pageLayout.arrange.rotate.flipHorizontal",
  "pageLayout.arrange.rotate.flipVertical",
  "pageLayout.arrange.rotate.rotateLeft90",
  "pageLayout.arrange.rotate.rotateRight90",
  "pageLayout.arrange.selectionPane",
  "pageLayout.pageSetup.background",
  "pageLayout.pageSetup.background.background",
  "pageLayout.pageSetup.background.delete",
  "pageLayout.pageSetup.breaks",
  "pageLayout.pageSetup.breaks.insertPageBreak",
  "pageLayout.pageSetup.breaks.removePageBreak",
  "pageLayout.pageSetup.breaks.resetAll",
  "pageLayout.pageSetup.margins",
  "pageLayout.pageSetup.orientation",
  "pageLayout.pageSetup.printArea",
  "pageLayout.pageSetup.printTitles",
  "pageLayout.pageSetup.printTitles.printTitles",
  "pageLayout.pageSetup.size",
  "pageLayout.scaleToFit.height",
  "pageLayout.scaleToFit.height.1page",
  "pageLayout.scaleToFit.height.2pages",
  "pageLayout.scaleToFit.height.automatic",
  "pageLayout.scaleToFit.scale",
  "pageLayout.scaleToFit.scale.100",
  "pageLayout.scaleToFit.scale.70",
  "pageLayout.scaleToFit.scale.80",
  "pageLayout.scaleToFit.scale.90",
  "pageLayout.scaleToFit.scale.more",
  "pageLayout.scaleToFit.width",
  "pageLayout.scaleToFit.width.1page",
  "pageLayout.scaleToFit.width.2pages",
  "pageLayout.scaleToFit.width.automatic",
  "pageLayout.sheetOptions.gridlinesPrint",
  "pageLayout.sheetOptions.gridlinesView",
  "pageLayout.sheetOptions.headingsPrint",
  "pageLayout.sheetOptions.headingsView",
  "pageLayout.themes.colors",
  "pageLayout.themes.colors.colorful",
  "pageLayout.themes.colors.customize",
  "pageLayout.themes.colors.office",
  "pageLayout.themes.effects",
  "pageLayout.themes.effects.intense",
  "pageLayout.themes.effects.moderate",
  "pageLayout.themes.effects.subtle",
  "pageLayout.themes.fonts",
  "pageLayout.themes.fonts.aptos",
  "pageLayout.themes.fonts.customize",
  "pageLayout.themes.fonts.office",
  "pageLayout.themes.themes",
  "pageLayout.themes.themes.customize",
  "pageLayout.themes.themes.facet",
  "pageLayout.themes.themes.integral",
  "pageLayout.themes.themes.office",
  "review.changes.protectShareWorkbook",
  "review.changes.protectShareWorkbook.protectWorkbook",
  "review.changes.shareWorkbook",
  "review.changes.shareWorkbook.shareNow",
  "review.changes.trackChanges",
  "review.changes.trackChanges.highlight",
  "review.comments.deleteComment",
  "review.comments.deleteComment.deleteAll",
  "review.comments.deleteComment.deleteThread",
  "review.comments.next",
  "review.comments.previous",
  "review.ink.startInking",
  "review.language.language",
  "review.language.language.setProofing",
  "review.language.language.translate",
  "review.language.translate",
  "review.language.translate.translateSelection",
  "review.language.translate.translateSheet",
  "review.notes.editNote",
  "review.notes.newNote",
  "review.notes.showAllNotes",
  "review.notes.showHideNote",
  "review.proofing.accessibility",
  "review.proofing.smartLookup",
  "review.proofing.spelling",
  "review.proofing.spelling.thesaurus",
  "review.proofing.spelling.wordCount",
  "review.protect.allowEditRanges",
  "review.protect.allowEditRanges.new",
  "review.protect.protectSheet",
  "review.protect.protectWorkbook",
  "review.protect.unprotectSheet",
  "review.protect.unprotectWorkbook",
  "view.show.formulaBar",
  "view.show.gridlines",
  "view.show.headings",
  "view.show.ruler",
  "view.window.arrangeAll",
  "view.window.arrangeAll.cascade",
  "view.window.arrangeAll.horizontal",
  "view.window.arrangeAll.tiled",
  "view.window.arrangeAll.vertical",
  "view.window.freezePanes",
  "view.window.hide",
  "view.window.newWindow",
  "view.window.newWindow.newWindowForActiveSheet",
  "view.window.resetWindowPosition",
  "view.window.switchWindows",
  "view.window.switchWindows.window1",
  "view.window.switchWindows.window2",
  "view.window.synchronousScrolling",
  "view.window.unhide",
  "view.window.viewSideBySide",
  "view.workbookViews.customViews",
  "view.workbookViews.customViews.manage",
  "view.workbookViews.normal",
  "view.workbookViews.pageBreakPreview",
  "view.workbookViews.pageLayout",
]);
