import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";
import { showQuickPick } from "../extensions/ui.js";
import { getPasteSpecialMenuItems } from "../clipboard/pasteSpecial.js";
import type { ThemeController } from "../theme/themeController.js";

export function registerBuiltinCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController;
  /**
   * Optional focus restoration hook after sheet navigation actions.
   *
   * The desktop shell can use this to avoid stealing focus while the user is in
   * inline rename mode or while menus are open, while still ensuring normal sheet
   * switching leaves the grid ready for typing/shortcuts.
   */
  focusAfterSheetNavigation?: (() => void) | null;
  /**
   * Optional source of truth for the current visible sheet order (e.g. the UI's sheet store).
   * When provided, sheet navigation commands (Ctrl/Cmd+PgUp/PgDn) use this list so they match
   * the order the user sees in the tab strip.
   */
  getVisibleSheetIds?: (() => string[]) | null;
  ensureExtensionsLoaded?: (() => Promise<void>) | null;
  onExtensionsLoaded?: (() => void) | null;
  themeController?: Pick<ThemeController, "setThemePreference"> | null;
  /**
   * Optional callback to refresh ribbon UI-state overrides (e.g. label overrides).
   *
   * Theme preference commands call this so the ribbon's "Theme" dropdown label
   * updates immediately after executing a theme command from the command palette
   * or extensions.
   */
  refreshRibbonUiState?: (() => void) | null;
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
  } = params;

  const toggleDockPanel = (panelId: string) => {
    const placement = getPanelPlacement(layoutController.layout, panelId);
    if (placement.kind === "closed") layoutController.openPanel(panelId);
    else layoutController.closePanel(panelId);
  };

  const listVisibleSheetIds = (): string[] => {
    if (getVisibleSheetIds) {
      try {
        const ids = getVisibleSheetIds();
        if (Array.isArray(ids) && ids.length > 0) return ids;
      } catch {
        // Best-effort: fall back to DocumentController sheet ids.
      }
    }

    let ids: string[] = [];
    try {
      const doc = app.getDocument();
      ids = doc.getVisibleSheetIds();
    } catch {
      try {
        ids = app.getDocument().getSheetIds();
      } catch {
        ids = [];
      }
    }
    // DocumentController materializes sheets lazily; mimic the UI fallback behavior so
    // navigation commands are stable even before any edits occur.
    return ids.length > 0 ? ids : ["Sheet1"];
  };

  const activateRelativeSheet = (delta: -1 | 1): void => {
    // Excel-like behavior: generally do not allow sheet navigation while editing.
    //
    // Exception: while the formula bar is actively editing a *formula* (range selection mode),
    // allow switching sheets so users can build cross-sheet references.
    if (app.isEditing() && !app.isFormulaBarFormulaEditing()) return;
    const sheetIds = listVisibleSheetIds();
    if (sheetIds.length <= 1) return;
    const active = app.getCurrentSheetId();
    const activeIndex = sheetIds.indexOf(active);
    if (activeIndex === -1) {
      // Current sheet is no longer visible (should be rare; typically we auto-fallback elsewhere).
      // Treat this as a "jump to first visible sheet" so navigation is deterministic.
      const first = sheetIds[0];
      if (!first || first === active) return;
      app.activateSheet(first);
      if (focusAfterSheetNavigation) {
        focusAfterSheetNavigation();
      } else {
        app.focusAfterSheetNavigation();
      }
      return;
    }

    const idx = activeIndex;
    const nextIndex = (idx + delta + sheetIds.length) % sheetIds.length;
    const next = sheetIds[nextIndex];
    if (!next || next === active) return;
    app.activateSheet(next);
    if (focusAfterSheetNavigation) {
      focusAfterSheetNavigation();
    } else {
      app.focusAfterSheetNavigation();
    }
  };

  const getTextEditingTarget = (): HTMLElement | null => {
    if (typeof document === "undefined") return null;
    const target = document.activeElement as HTMLElement | null;
    if (!target) return null;
    const tag = target.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return target;
    return null;
  };

  const tryExecCommand = (command: string): boolean => {
    if (typeof document === "undefined") return false;
    try {
      return document.execCommand(command, false);
    } catch {
      return false;
    }
  };

  commandRegistry.registerBuiltinCommand(
    "edit.undo",
    t("command.edit.undo"),
    () => {
      // Excel-like behavior: when focus is in a text editing surface, undo/redo should
      // apply to that surface instead of spreadsheet history.
      if (getTextEditingTarget()) {
        tryExecCommand("undo");
        return;
      }

      // Formula bar range selection mode can temporarily move focus back to the grid while the
      // formula bar is still actively editing. In that case, treat undo/redo as text editing
      // operations (Excel behavior) rather than workbook history.
      if ((app as any).isFormulaBarEditing?.()) {
        (app as any).focusFormulaBar?.();
        tryExecCommand("undo");
        return;
      }
      app.undo();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.undo"),
      keywords: ["undo", "history"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.redo",
    t("command.edit.redo"),
    () => {
      if (getTextEditingTarget()) {
        tryExecCommand("redo");
        return;
      }
      if ((app as any).isFormulaBarEditing?.()) {
        (app as any).focusFormulaBar?.();
        tryExecCommand("redo");
        return;
      }
      app.redo();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.redo"),
      keywords: ["redo", "history"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.toggleShowFormulas",
    t("command.view.toggleShowFormulas"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.toggleShowFormulas();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.toggleShowFormulas"),
      keywords: ["show formulas", "formulas", "values", "display"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePerformanceStats",
    t("command.view.togglePerformanceStats"),
    (next?: boolean) => {
      const perfStats = app.getGridPerfStats() as any;
      const current = Boolean(perfStats?.enabled);
      const enabled = typeof next === "boolean" ? next : !current;
      app.setGridPerfStatsEnabled(enabled);
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePerformanceStats"),
      keywords: ["performance", "perf", "stats", "overlay", "fps"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.toggleSplitView",
    t("command.view.toggleSplitView"),
    (next?: boolean) => {
      const currentDirection = layoutController.layout.splitView.direction;
      const shouldEnable = typeof next === "boolean" ? next : currentDirection === "none";

      if (!shouldEnable) {
        layoutController.setSplitDirection("none");
      } else if (currentDirection === "none") {
        // Match ribbon toggle behavior: default to a 50/50 vertical split the first
        // time split view is enabled.
        layoutController.setSplitDirection("vertical", 0.5);
      }

      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.toggleSplitView"),
      keywords: ["split", "split view", "pane", "panes"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.togglePrecedents",
    t("command.audit.togglePrecedents"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.toggleAuditingPrecedents();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.togglePrecedents"),
      keywords: ["audit", "precedents", "trace", "toggle"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.toggleDependents",
    t("command.audit.toggleDependents"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.toggleAuditingDependents();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.toggleDependents"),
      keywords: ["audit", "dependents", "trace", "toggle"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbench.showCommandPalette",
    t("command.workbench.showCommandPalette"),
    () => {
      // Intentionally a no-op: the desktop shell owns opening the palette, but we still
      // register the id so keybinding and menu systems can reference it.
    },
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbench.showCommandPalette"),
      keywords: ["command palette", "commands"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbook.previousSheet",
    t("command.workbook.previousSheet"),
    () => activateRelativeSheet(-1),
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbook.previousSheet"),
      keywords: ["sheet", "previous", "navigation", "pageup", "pgup"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbook.nextSheet",
    t("command.workbook.nextSheet"),
    () => activateRelativeSheet(1),
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbook.nextSheet"),
      keywords: ["sheet", "next", "navigation", "pagedown", "pgdn"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "ai.inlineEdit",
    t("command.ai.inlineEdit"),
    () => app.openInlineAiEdit(),
    {
      category: t("commandCategory.ai"),
      icon: null,
      description: t("commandDescription.ai.inlineEdit"),
      keywords: ["ai", "inline edit", "transform"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.insertPivotTable",
    t("command.view.insertPivotTable"),
    () => {
      layoutController.openPanel(PanelIds.PIVOT_BUILDER);
      // If the panel is already open, we still want to refresh its source range from
      // the latest selection.
      window.dispatchEvent(new CustomEvent("pivot-builder:use-selection"));
    },
    {
      category: t("commandCategory.data"),
      icon: null,
      description: t("commandDescription.view.insertPivotTable"),
      keywords: ["pivot", "pivot table", "pivotbuilder"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.aiChat",
    t("command.view.togglePanel.aiChat"),
    () => toggleDockPanel(PanelIds.AI_CHAT),
    {
      category: t("commandCategory.ai"),
      icon: null,
      description: t("commandDescription.view.togglePanel.aiChat"),
      keywords: ["ai", "chat", "assistant", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.aiAudit",
    t("command.view.togglePanel.aiAudit"),
    () => toggleDockPanel(PanelIds.AI_AUDIT),
    {
      category: t("commandCategory.ai"),
      icon: null,
      description: t("commandDescription.view.togglePanel.aiAudit"),
      keywords: ["ai", "audit", "log", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.extensions",
    t("command.view.togglePanel.extensions"),
    () => {
      if (ensureExtensionsLoaded) {
        void ensureExtensionsLoaded()
          .then(() => onExtensionsLoaded?.())
          .catch(() => {
            // ignore; panel open/close should still work
          });
      }
      toggleDockPanel(PanelIds.EXTENSIONS);
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.extensions"),
      keywords: ["extensions", "plugins", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.macros",
    t("command.view.togglePanel.macros"),
    () => toggleDockPanel(PanelIds.MACROS),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.macros"),
      keywords: ["macros", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.dataQueries",
    t("command.view.togglePanel.dataQueries"),
    () => toggleDockPanel(PanelIds.DATA_QUERIES),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.dataQueries"),
      keywords: ["data", "queries", "power query", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.scriptEditor",
    t("command.view.togglePanel.scriptEditor"),
    () => toggleDockPanel(PanelIds.SCRIPT_EDITOR),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.scriptEditor"),
      keywords: ["script", "editor", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.python",
    t("command.view.togglePanel.python"),
    () => toggleDockPanel(PanelIds.PYTHON),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.python"),
      keywords: ["python", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.marketplace",
    t("command.view.togglePanel.marketplace"),
    () => {
      if (ensureExtensionsLoaded) {
        void ensureExtensionsLoaded()
          .then(() => onExtensionsLoaded?.())
          .catch(() => {
            // ignore; panel open/close should still work
          });
      }
      toggleDockPanel(PanelIds.MARKETPLACE);
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.marketplace"),
      keywords: ["marketplace", "extensions", "plugins", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.versionHistory",
    t("command.view.togglePanel.versionHistory"),
    () => toggleDockPanel(PanelIds.VERSION_HISTORY),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.versionHistory"),
      keywords: ["version", "versions", "history", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.branchManager",
    t("command.view.togglePanel.branchManager"),
    () => toggleDockPanel(PanelIds.BRANCH_MANAGER),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.branchManager"),
      keywords: ["branch", "branches", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "comments.togglePanel",
    t("command.comments.togglePanel"),
    () => app.toggleCommentsPanel(),
    {
      category: t("commandCategory.comments"),
      icon: null,
      description: t("commandDescription.comments.togglePanel"),
      keywords: ["comments", "notes", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "comments.addComment",
    t("command.comments.addComment"),
    () => {
      // Match spreadsheet shortcut behavior: don't trigger comment UX while the user is
      // actively editing a cell/formula (Excel-style).
      if (app.isEditing()) return;
      app.openCommentsPanel();
      app.focusNewCommentInput();
    },
    {
      category: t("commandCategory.comments"),
      icon: null,
      description: t("commandDescription.comments.addComment"),
      keywords: ["comment", "add comment", "new comment"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezePanes",
    t("command.view.freezePanes"),
    () => {
      app.freezePanes();
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.freezePanes"),
      keywords: ["freeze", "panes"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezeTopRow",
    t("command.view.freezeTopRow"),
    () => {
      app.freezeTopRow();
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.freezeTopRow"),
      keywords: ["freeze", "top row"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezeFirstColumn",
    t("command.view.freezeFirstColumn"),
    () => {
      app.freezeFirstColumn();
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.freezeFirstColumn"),
      keywords: ["freeze", "first column"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.unfreezePanes",
    t("command.view.unfreezePanes"),
    () => {
      app.unfreezePanes();
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.unfreezePanes"),
      keywords: ["unfreeze", "panes"],
    },
  );

  // Ribbon's View → Window → Freeze Panes dropdown uses a stable trigger id in a canonical
  // namespace (`view.window.freezePanes`). Register it as an alias so it can be executed via
  // the CommandRegistry (e.g. command palette) and so Ribbon↔CommandRegistry coverage tests
  // can enforce that canonical ribbon ids stay wired.
  commandRegistry.registerBuiltinCommand(
    "view.window.freezePanes",
    t("command.view.freezePanes"),
    () => commandRegistry.executeCommand("view.freezePanes"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.freezePanes"),
      keywords: ["freeze", "panes"],
    },
  );

  const setZoomPercent = (percent: number): void => {
    if (!app.supportsZoom()) return;
    const value = typeof percent === "number" ? percent : Number(percent);
    if (!Number.isFinite(value) || value <= 0) return;
    app.setZoom(value / 100);
  };

  const registerZoomPreset = (percent: number): void => {
    const value = Math.round(percent);
    const id = `view.zoom.zoom${value}`;
    commandRegistry.registerBuiltinCommand(
      id,
      t(`command.${id}`),
      () => {
        setZoomPercent(value);
        if (!app.supportsZoom()) return;
        app.focus();
      },
      {
        category: t("commandCategory.view"),
        icon: null,
        description: t(`commandDescription.${id}`),
        keywords: ["zoom", `${value}%`, "view", "scale"],
      },
    );
  };

  for (const percent of [25, 50, 75, 100, 150, 200, 400]) {
    registerZoomPreset(percent);
  }

  commandRegistry.registerBuiltinCommand(
    "view.zoom.set",
    t("command.view.zoom.set"),
    async (...args: any[]) => {
      if (!app.supportsZoom()) return;
      if (args.length === 0) {
        await commandRegistry.executeCommand("view.zoom.openPicker");
        return;
      }
      const percent = Number(args[0]);
      if (!Number.isFinite(percent) || percent <= 0) return;
      setZoomPercent(percent);
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.zoom.set"),
      keywords: ["zoom", "view", "scale", "percent"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.zoom.zoomToSelection",
    t("command.view.zoom.zoomToSelection"),
    () => {
      if (!app.supportsZoom()) return;
      app.zoomToSelection();
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.zoom.zoomToSelection"),
      keywords: ["zoom", "selection", "fit", "view"],
    },
  );

  const zoomPickerTitle = t("command.view.zoom.openPicker");
  commandRegistry.registerBuiltinCommand(
    "view.zoom.openPicker",
    zoomPickerTitle,
    async () => {
      if (!app.supportsZoom()) return;
      // Keep the custom zoom picker aligned with the shared-grid zoom clamp
      // (currently 25%–400%, Excel-style).
      const baseOptions = [25, 50, 75, 100, 125, 150, 200, 400];
      const current = Math.round(app.getZoom() * 100);
      const options = baseOptions.includes(current) ? baseOptions : [current, ...baseOptions];
      const picked = await showQuickPick(
        options.map((value) => ({ label: `${value}%`, value })),
        { placeHolder: zoomPickerTitle },
      );
      if (picked == null) return;
      setZoomPercent(picked);
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.zoom.openPicker"),
      keywords: ["zoom", "custom zoom", "view", "scale"],
    },
  );

  // Ribbon's View → Zoom dropdown uses `view.zoom.zoom` as its stable trigger id.
  // Provide it as an alias for the picker command so it can be executed directly and
  // so Ribbon↔CommandRegistry coverage can treat the id as "real".
  commandRegistry.registerBuiltinCommand(
    "view.zoom.zoom",
    zoomPickerTitle,
    () => commandRegistry.executeCommand("view.zoom.openPicker"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.zoom.openPicker"),
      keywords: ["zoom", "custom zoom", "view", "scale"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.tracePrecedents",
    t("command.audit.tracePrecedents"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.clearAuditing();
      app.toggleAuditingPrecedents();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.tracePrecedents"),
      keywords: ["audit", "precedents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceDependents",
    t("command.audit.traceDependents"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.clearAuditing();
      app.toggleAuditingDependents();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.traceDependents"),
      keywords: ["audit", "dependents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceBoth",
    t("command.audit.traceBoth"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.clearAuditing();
      app.toggleAuditingPrecedents();
      app.toggleAuditingDependents();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.traceBoth"),
      keywords: ["audit", "precedents", "dependents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.clearAuditing",
    t("command.audit.clearAuditing"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.clearAuditing();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.clearAuditing"),
      keywords: ["audit", "clear"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.toggleTransitive",
    t("command.audit.toggleTransitive"),
    () => {
      if (app.isEditing()) return;
      if (getTextEditingTarget()) return;
      app.toggleAuditingTransitive();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.toggleTransitive"),
      keywords: ["audit", "transitive", "toggle"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.copy",
    t("command.clipboard.copy"),
    () => {
      // Excel-like behavior: when focus is in a text editing surface, copy/cut/paste should apply to
      // that surface instead of the spreadsheet selection.
      if (getTextEditingTarget()) {
        tryExecCommand("copy");
        return;
      }
      // Formula bar range selection mode can temporarily move focus back to the grid while the formula
      // bar is still actively editing. In that case, treat copy/cut/paste as text editing operations.
      if ((app as any).isFormulaBarEditing?.()) {
        (app as any).focusFormulaBar?.();
        tryExecCommand("copy");
        return;
      }
      return app.copyToClipboard();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.copy"),
      keywords: ["copy", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.cut",
    t("command.clipboard.cut"),
    () => {
      if (getTextEditingTarget()) {
        tryExecCommand("cut");
        return;
      }
      if ((app as any).isFormulaBarEditing?.()) {
        (app as any).focusFormulaBar?.();
        tryExecCommand("cut");
        return;
      }
      return app.cutToClipboard();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.cut"),
      keywords: ["cut", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.paste",
    t("command.clipboard.paste"),
    () => {
      if (getTextEditingTarget()) {
        tryExecCommand("paste");
        return;
      }
      if ((app as any).isFormulaBarEditing?.()) {
        (app as any).focusFormulaBar?.();
        tryExecCommand("paste");
        return;
      }
      return app.pasteFromClipboard();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.paste"),
      keywords: ["paste", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.all",
    t("clipboard.pasteSpecial.paste"),
    () => app.pasteFromClipboard(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.all"),
      keywords: ["paste", "clipboard", "all"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.values",
    t("clipboard.pasteSpecial.pasteValues"),
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      try {
        await app.clipboardPasteSpecial("values");
      } finally {
        app.focus();
      }
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.values"),
      keywords: ["paste", "clipboard", "values"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.formulas",
    t("clipboard.pasteSpecial.pasteFormulas"),
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      try {
        await app.clipboardPasteSpecial("formulas");
      } finally {
        app.focus();
      }
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.formulas"),
      keywords: ["paste", "clipboard", "formulas"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.formats",
    t("clipboard.pasteSpecial.pasteFormats"),
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      try {
        await app.clipboardPasteSpecial("formats");
      } finally {
        app.focus();
      }
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.formats"),
      keywords: ["paste", "clipboard", "formats"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.transpose",
    t("clipboard.pasteSpecial.transpose"),
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      try {
        await app.clipboardPasteSpecial("all", { transpose: true });
      } finally {
        app.focus();
      }
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.transpose"),
      keywords: ["paste", "clipboard", "transpose"],
    },
  );

  const pasteSpecialTitle = t("clipboard.pasteSpecial.title");
  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial",
    pasteSpecialTitle,
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      const items = getPasteSpecialMenuItems();
      const picked = await showQuickPick(
        items.map((item) => ({ label: item.label, value: item.mode })),
        { placeHolder: pasteSpecialTitle },
      );
      if (!picked) {
        app.focus();
        return;
      }
      try {
        await commandRegistry.executeCommand(`clipboard.pasteSpecial.${picked === "all" ? "all" : picked}`);
      } finally {
        app.focus();
      }
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial"),
      keywords: ["paste", "paste special", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.editCell",
    t("command.edit.editCell"),
    () => {
      if (app.isEditing()) return;
      app.openCellEditorAtActiveCell();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.editCell"),
      keywords: ["edit", "cell", "f2"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.clearContents",
    t("command.edit.clearContents"),
    () => app.clearSelectionContents(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.clearContents"),
      keywords: ["clear", "contents", "delete"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillDown",
    t("command.edit.fillDown"),
    () => app.fillDown(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.fillDown"),
      keywords: ["fill", "fill down", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillRight",
    t("command.edit.fillRight"),
    () => app.fillRight(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.fillRight"),
      keywords: ["fill", "fill right", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.selectCurrentRegion",
    t("command.edit.selectCurrentRegion"),
    () => {
      if (app.isEditing()) return;
      app.selectCurrentRegion();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.selectCurrentRegion"),
      keywords: ["select", "current region", "region", "excel", "ctrl+shift+8", "ctrl+shift+*", "ctrl+*", "numpad"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.insertDate",
    t("command.edit.insertDate"),
    () => app.insertDate(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.insertDate"),
      keywords: ["date", "insert date", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.insertTime",
    t("command.edit.insertTime"),
    () => app.insertTime(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.insertTime"),
      keywords: ["time", "insert time", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.autoSum",
    t("command.edit.autoSum"),
    () => app.autoSum(),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.autoSum"),
      keywords: ["autosum", "sum", "excel"],
    },
  );

  if (themeController) {
    const categoryView = t("commandCategory.view");
    const refresh = () => {
      try {
        refreshRibbonUiState?.();
      } catch {
        // ignore
      }
    };

    const focusApp = () => {
      try {
        (app as any)?.focus?.();
      } catch {
        // ignore
      }
    };

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.system",
      "Theme: System",
      () => {
        themeController.setThemePreference("system");
        refresh();
        focusApp();
      },
      {
        category: categoryView,
        icon: null,
        description: "Use the system theme",
        keywords: ["theme", "appearance", "system", "auto", "os", "dark mode", "light mode"],
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.light",
      "Theme: Light",
      () => {
        themeController.setThemePreference("light");
        refresh();
        focusApp();
      },
      {
        category: categoryView,
        icon: null,
        description: "Use the light theme",
        keywords: ["theme", "appearance", "light", "light mode"],
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.dark",
      "Theme: Dark",
      () => {
        themeController.setThemePreference("dark");
        refresh();
        focusApp();
      },
      {
        category: categoryView,
        icon: null,
        description: "Use the dark theme",
        keywords: ["theme", "appearance", "dark", "dark mode"],
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.highContrast",
      "Theme: High Contrast",
      () => {
        themeController.setThemePreference("high-contrast");
        refresh();
        focusApp();
      },
      {
        category: categoryView,
        icon: null,
        description: "Use the high contrast theme",
        keywords: ["theme", "appearance", "high contrast", "contrast", "accessibility", "a11y"],
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme",
      "Theme…",
      async () => {
        const picked = await showQuickPick(
          [
            { label: "System", value: "view.appearance.theme.system" },
            { label: "Light", value: "view.appearance.theme.light" },
            { label: "Dark", value: "view.appearance.theme.dark" },
            { label: "High Contrast", value: "view.appearance.theme.highContrast" },
          ],
          { placeHolder: "Theme" },
        );
        if (!picked) return;
        await commandRegistry.executeCommand(picked);
      },
      {
        category: categoryView,
        icon: null,
        description: "Choose an application theme",
        keywords: ["theme", "appearance", "dark mode", "light mode"],
      },
    );
  }
}
