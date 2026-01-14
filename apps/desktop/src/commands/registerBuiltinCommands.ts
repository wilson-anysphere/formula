import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";
import { showQuickPick, showToast } from "../extensions/ui.js";
import { getPasteSpecialMenuItems } from "../clipboard/pasteSpecial.js";
import type { ThemeController } from "../theme/themeController.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";
import { cycleWorkbenchFocusRegion, type WorkbenchFocusCycleDeps } from "./workbenchFocusCycle.js";
import { registerNumberFormatCommands } from "./registerNumberFormatCommands.js";
import { DEFAULT_GRID_LIMITS } from "../selection/selection.js";
import type { GridLimits, Range } from "../selection/types";
import { DEFAULT_DESKTOP_LOAD_MAX_COLS, DEFAULT_DESKTOP_LOAD_MAX_ROWS } from "../workbook/load/clampUsedRange.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize, normalizeSelectionRange } from "../formatting/selectionSizeGuard.js";
import {
  setFillColor,
  setFontColor,
  setFontSize,
  toggleBold,
  toggleItalic,
  toggleUnderline,
  toggleWrap,
  type CellRange,
} from "../formatting/toolbar.js";

export function registerBuiltinCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController;
  /**
   * Optional spreadsheet edit-state predicate. When omitted, falls back to `app.isEditing()`.
   *
   * The desktop shell passes a custom predicate that includes split-view secondary editor state.
   */
  isEditing?: (() => boolean) | null;
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
  /**
   * Optional hook to open the Goal Seek dialog (What-If Analysis).
   *
   * The desktop shell owns modal dialog mounting in `main.ts`, so the command
   * catalog accepts a host callback rather than embedding portal/DOM wiring here.
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
    openGoalSeekDialog = null,
  } = params;

  const refresh = (): void => {
    if (!refreshRibbonUiState) return;
    try {
      refreshRibbonUiState();
    } catch {
      // Best-effort only.
    }
  };

  const focusApp = (): void => {
    const focusFn = (app as any)?.focus;
    if (typeof focusFn !== "function") return;
    try {
      focusFn.call(app);
    } catch {
      // Best-effort only.
    }
  };

  const isEditingFn =
    isEditing ?? (() => (typeof (app as any)?.isEditing === "function" ? Boolean((app as any).isEditing()) : false));
  const isReadOnlyFn = (): boolean => {
    try {
      return typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true;
    } catch {
      return false;
    }
  };

  const commandCategoryFormat = t("commandCategory.format");
  const commandCategoryData = t("commandCategory.data");
  const commandCategoryPageLayout = t("commandCategory.pageLayout");

  const toggleDockPanel = (panelId: string) => {
    const placement = getPanelPlacement(layoutController.layout, panelId);
    if (placement.kind === "closed") {
      layoutController.openPanel(panelId);
      return;
    }

    // Dock zones can be collapsed. Treat a collapsed docked panel as "closed" for toggle purposes
    // so toggle commands restore the dock (instead of removing the panel).
    if (placement.kind === "docked" && (layoutController.layout as any)?.docks?.[placement.side]?.collapsed) {
      layoutController.openPanel(panelId);
      return;
    }

    // Floating panels can be minimized. Treat a minimized floating panel as "closed" for toggle
    // purposes so toggle commands restore the panel instead of closing it.
    if (placement.kind === "floating" && (layoutController.layout as any)?.floating?.[panelId]?.minimized) {
      layoutController.setFloatingPanelMinimized(panelId, false);
      return;
    }

    layoutController.closePanel(panelId);
  };

  const openDockPanel = (panelId: string) => {
    const placement = getPanelPlacement(layoutController.layout, panelId);
    // Always call openPanel so we activate docked panels and also trigger a layout
    // re-render even when the panel is already open (important for tabbed docks:
    // inactive tabpanels are empty until activated).
    layoutController.openPanel(panelId);

    // Floating panels can be minimized; opening should restore them.
    if (placement.kind === "floating" && (layoutController.layout as any)?.floating?.[panelId]?.minimized) {
      layoutController.setFloatingPanelMinimized(panelId, false);
    }
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
    if (isEditingFn()) {
      // Only allow the formula-bar exception when the primary editor is the one that is editing.
      // When a split-view secondary editor is active, sheet navigation should still be blocked.
      const appAny = app as any;
      const primaryEditing = typeof appAny?.isEditing === "function" ? Boolean(appAny.isEditing()) : false;
      const formulaEditing = typeof appAny?.isFormulaBarFormulaEditing === "function" ? Boolean(appAny.isFormulaBarFormulaEditing()) : false;
      if (!(primaryEditing && formulaEditing)) return;
    }
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

  const getWorkbenchFocusCycleDeps = (): WorkbenchFocusCycleDeps | null => {
    if (typeof document === "undefined") return null;
    const ribbonRootEl = document.getElementById("ribbon") as HTMLElement | null;
    const formulaBarRootEl = document.getElementById("formula-bar") as HTMLElement | null;
    const gridRootEl = document.getElementById("grid") as HTMLElement | null;
    const statusBarRootEl = document.querySelector<HTMLElement>(".statusbar");
    if (!ribbonRootEl || !formulaBarRootEl || !gridRootEl || !statusBarRootEl) return null;
    return {
      ribbonRootEl,
      formulaBarRootEl,
      gridRootEl,
      statusBarRootEl,
      focusGrid: () => {
        try {
          app.focus();
        } catch {
          // ignore (tests/minimal harnesses)
        }
      },
      getSecondaryGridRoot: () => document.getElementById("grid-secondary") as HTMLElement | null,
      getSheetTabsRoot: () => document.getElementById("sheet-tabs") as HTMLElement | null,
    };
  };

  const getGridLimitsForFormatting = (): GridLimits => {
    const raw = typeof (app as any)?.getGridLimits === "function" ? (app as any).getGridLimits() : null;
    const maxRows =
      Number.isInteger(raw?.maxRows) && raw.maxRows > 0 ? raw.maxRows : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
    const maxCols =
      Number.isInteger(raw?.maxCols) && raw.maxCols > 0 ? raw.maxCols : DEFAULT_DESKTOP_LOAD_MAX_COLS;
    return { maxRows, maxCols };
  };

  const selectionRangesForFormatting = (): CellRange[] => {
    const limits = getGridLimitsForFormatting();
    const selection = typeof (app as any)?.getSelectionRanges === "function" ? (app as any).getSelectionRanges() : [];
    if (!Array.isArray(selection) || selection.length === 0) {
      const cell = typeof (app as any)?.getActiveCell === "function" ? (app as any).getActiveCell() : { row: 0, col: 0 };
      return [{ start: { row: cell.row, col: cell.col }, end: { row: cell.row, col: cell.col } }];
    }

    return selection.map((range: Range) => {
      const r = normalizeSelectionRange(range);
      const isFullColBand = r.startRow === 0 && r.endRow === limits.maxRows - 1;
      const isFullRowBand = r.startCol === 0 && r.endCol === limits.maxCols - 1;

      return {
        start: { row: r.startRow, col: r.startCol },
        end: {
          row: isFullColBand ? DEFAULT_GRID_LIMITS.maxRows - 1 : r.endRow,
          col: isFullRowBand ? DEFAULT_GRID_LIMITS.maxCols - 1 : r.endCol,
        },
      };
    });
  };

  const applyFormattingToSelection = (
    label: string,
    fn: (doc: any, sheetId: string, ranges: CellRange[]) => void | boolean,
    options: { forceBatch?: boolean } = {},
  ): void => {
    // Match SpreadsheetApp guards: formatting commands should never mutate the sheet while the user is
    // actively editing (cell editor / formula bar / inline edit).
    if (isEditingFn()) return;

    const doc = typeof (app as any)?.getDocument === "function" ? (app as any).getDocument() : null;
    const sheetId = typeof (app as any)?.getCurrentSheetId === "function" ? (app as any).getCurrentSheetId() : null;
    if (!doc || !sheetId) return;

    const selection = typeof (app as any)?.getSelectionRanges === "function" ? (app as any).getSelectionRanges() : [];
    const limits = getGridLimitsForFormatting();
    const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });

    if (!decision.allowed) {
      try {
        showToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return;
    }

    const isReadOnly = typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true;
    if (isReadOnly && !decision.allRangesBand) {
      showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
      return;
    }

    const ranges = selectionRangesForFormatting();
    const shouldBatch = Boolean(options.forceBatch) || ranges.length > 1;

    if (shouldBatch) doc.beginBatch?.({ label });
    let committed = false;
    let applied = true;
    try {
      const result = fn(doc, sheetId, ranges);
      if (result === false) applied = false;
      committed = true;
    } finally {
      if (!shouldBatch) {
        // no-op
      } else if (committed) {
        doc.endBatch?.();
      } else {
        doc.cancelBatch?.();
      }
    }
    if (!applied) {
      try {
        showToast("Formatting could not be applied to the full selection. Try selecting fewer cells/rows.", "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
    }
    (app as any).focus?.();
  };

  const shouldOpenFormattingPrompt = (): boolean => {
    // These commands are disabled while editing (ribbon state + Excel parity). Guard here so
    // command palette / keybindings can't open a prompt (color picker, font size picker, etc)
    // only to no-op later when applying the formatting.
    if (isEditingFn()) return false;
    const selection = typeof (app as any)?.getSelectionRanges === "function" ? (app as any).getSelectionRanges() : [];
    const limits = getGridLimitsForFormatting();
    const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });
    if (!decision.allowed) {
      try {
        showToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return false;
    }

    const isReadOnly = typeof (app as any)?.isReadOnly === "function" && (app as any).isReadOnly() === true;
    if (isReadOnly && !decision.allRangesBand) {
      showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
      return false;
    }

    return true;
  };

  const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

  const activeCellFontSizePt = (): number => {
    try {
      const sheetId = (app as any).getCurrentSheetId?.();
      const cell = (app as any).getActiveCell?.();
      const docAny = (app as any).getDocument?.();
      if (!sheetId || !cell || !docAny) return 11;
      const effectiveSize = docAny.getCellFormat?.(sheetId, cell)?.font?.size;
      const state = docAny.getCell?.(sheetId, cell);
      const style = docAny.styleTable?.get?.(state?.styleId ?? 0) ?? {};
      const size = typeof effectiveSize === "number" ? effectiveSize : style.font?.size;
      return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : 11;
    } catch {
      return 11;
    }
  };

  const activeCellNumberFormat = (): string | null => {
    try {
      const sheetId = (app as any).getCurrentSheetId?.();
      const cell = (app as any).getActiveCell?.();
      const docAny = (app as any).getDocument?.();
      if (!sheetId || !cell || !docAny) return null;
      const format = docAny.getCellFormat?.(sheetId, cell)?.numberFormat;
      return typeof format === "string" && format.trim() ? format : null;
    } catch {
      return null;
    }
  };

  const stepFontSize = (current: number, direction: "increase" | "decrease"): number => {
    const value = Number(current);
    const resolved = Number.isFinite(value) && value > 0 ? value : 11;
    if (direction === "increase") {
      for (const step of FONT_SIZE_STEPS) {
        if (step > resolved + 1e-6) return step;
      }
      return resolved;
    }

    for (let i = FONT_SIZE_STEPS.length - 1; i >= 0; i -= 1) {
      const step = FONT_SIZE_STEPS[i]!;
      if (step < resolved - 1e-6) return step;
    }
    return resolved;
  };

  const rgbHexToArgb = (rgb: string): string | null => {
    if (typeof rgb !== "string") return null;
    if (!/^#[0-9A-Fa-f]{6}$/.test(rgb)) return null;
    // DocumentController formatting expects #AARRGGBB.
    return `#FF${rgb.slice(1).toUpperCase()}`;
  };

  const normalizeArgb = (value: unknown): string | null => {
    if (typeof value !== "string") return null;
    const trimmed = value.trim();
    if (!trimmed) return null;
    // #RRGGBB
    if (/^#?[0-9a-f]{6}$/i.test(trimmed)) {
      const normalized = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
      return rgbHexToArgb(normalized);
    }
    // #AARRGGBB
    if (/^#?[0-9a-f]{8}$/i.test(trimmed)) {
      const hex = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
      return `#${hex.toUpperCase()}`;
    }
    return null;
  };

  const createHiddenColorInput = (): HTMLInputElement => {
    const input = document.createElement("input");
    input.type = "color";
    input.tabIndex = -1;
    input.className = "hidden-color-input shell-hidden-input";
    document.body.appendChild(input);
    return input;
  };

  let fontColorPicker: HTMLInputElement | null = null;
  let fillColorPicker: HTMLInputElement | null = null;

  const openColorPicker = (
    input: HTMLInputElement,
    label: string,
    apply: (doc: any, sheetId: string, ranges: CellRange[], argb: string) => void,
  ): void => {
    // Avoid `addEventListener({ once: true })` here.
    //
    // `<input type="color">` does *not* emit a `change` event when the user cancels the native
    // picker. If we used `addEventListener`, we'd accumulate listeners across cancels and the next
    // successful pick would apply formatting multiple times (multiple history entries).
    input.onchange = () => {
      input.onchange = null;
      const argb = rgbHexToArgb(input.value);
      if (!argb) return;
      applyFormattingToSelection(label, (doc, sheetId, ranges) => apply(doc, sheetId, ranges, argb));
    };
    input.click();
  };

  const applyThemePreference = (preference: "system" | "light" | "dark" | "high-contrast"): void => {
    if (!themeController) return;
    try {
      themeController.setThemePreference(preference);
    } catch {
      // Best-effort only.
    }
    refresh();
    focusApp();
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
    "view.theme.light",
    t("command.view.theme.light"),
    () => applyThemePreference("light"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.theme.light"),
      keywords: ["theme", "appearance", "light", "light mode", "color scheme"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.theme.dark",
    t("command.view.theme.dark"),
    () => applyThemePreference("dark"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.theme.dark"),
      keywords: ["theme", "appearance", "dark", "dark mode", "color scheme"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.theme.system",
    t("command.view.theme.system"),
    () => applyThemePreference("system"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.theme.system"),
      keywords: ["theme", "appearance", "system", "dark mode", "light mode", "auto"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.theme.highContrast",
    t("command.view.theme.highContrast"),
    () => applyThemePreference("high-contrast"),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.theme.highContrast"),
      keywords: ["theme", "appearance", "high contrast", "contrast", "accessibility"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.toggleShowFormulas",
    t("command.view.toggleShowFormulas"),
    (next?: boolean) => {
      // When invoked from the Ribbon, the command receives an explicit boolean pressed state.
      // In that case, restore grid focus so users can keep working (Excel-style).
      //
      // When invoked from other surfaces (command palette / keybindings), `next` is typically omitted,
      // so we avoid focusing here (the host owns returning focus after it closes overlays).
      const focusAfter = typeof next === "boolean";

      if (isEditingFn() || getTextEditingTarget()) {
        // Ribbon toggles optimistically update their internal pressed state before the command runs.
        // Refresh so the toggle reflects the actual (unchanged) app state.
        refresh();
        return;
      }

      if (typeof next === "boolean") {
        app.setShowFormulas(next);
        refresh();
        if (focusAfter) focusApp();
        return;
      }

      app.toggleShowFormulas();
      refresh();
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
      refresh();
      if (typeof next === "boolean") focusApp();
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
      if (isEditingFn()) {
        refresh();
        return;
      }
      const currentDirection = layoutController.layout.splitView.direction;
      const shouldEnable = typeof next === "boolean" ? next : currentDirection === "none";

      if (!shouldEnable) {
        layoutController.setSplitDirection("none");
      } else if (currentDirection === "none") {
        // Match ribbon toggle behavior: default to a 50/50 vertical split the first
        // time split view is enabled.
        layoutController.setSplitDirection("vertical", 0.5);
      }

      refresh();
      focusApp();
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
      if (isEditingFn()) return;
      if (getTextEditingTarget()) return;
      app.toggleAuditingPrecedents();
      app.focus();
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
      if (isEditingFn()) return;
      if (getTextEditingTarget()) return;
      app.toggleAuditingDependents();
      app.focus();
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
    "workbench.focusNextRegion",
    t("command.workbench.focusNextRegion"),
    () => {
      const deps = getWorkbenchFocusCycleDeps();
      if (!deps) return;
      cycleWorkbenchFocusRegion(deps, 1);
    },
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbench.focusNextRegion"),
      keywords: ["focus", "region", "next", "f6", "navigation"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbench.focusPrevRegion",
    t("command.workbench.focusPrevRegion"),
    () => {
      const deps = getWorkbenchFocusCycleDeps();
      if (!deps) return;
      cycleWorkbenchFocusRegion(deps, -1);
    },
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbench.focusPrevRegion"),
      keywords: ["focus", "region", "previous", "prev", "shift+f6", "navigation"],
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
    () => {
      if (isEditingFn()) return;
      app.openInlineAiEdit();
    },
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
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // Best-effort (toast root missing in tests/minimal harnesses).
        }
        return;
      }
      // Always call openPanel so we activate docked panels and also trigger a layout re-render
      // even when the panel is already floating (useful for refreshing panel-local state).
      const layout = (() => {
        try {
          return (layoutController as any)?.layout ?? null;
        } catch {
          return null;
        }
      })();
      const placement = (() => {
        try {
          if (layout) return getPanelPlacement(layout, PanelIds.PIVOT_BUILDER);
        } catch {
          // ignore (e.g. minimal unit-test stubs that omit layout state)
        }
        return { kind: "closed" } as const;
      })();
      layoutController.openPanel(PanelIds.PIVOT_BUILDER);

      // Floating panels can be minimized; opening should restore them.
      if (
        placement.kind === "floating" &&
        (layout as any)?.floating?.[PanelIds.PIVOT_BUILDER]?.minimized &&
        typeof (layoutController as any).setFloatingPanelMinimized === "function"
      ) {
        (layoutController as any).setFloatingPanelMinimized(PanelIds.PIVOT_BUILDER, false);
      }
      // If the panel is already open, we still want to refresh its source range from
      // the latest selection.
      try {
        if (typeof window === "undefined") return;
        window.dispatchEvent(new CustomEvent("pivot-builder:use-selection"));
      } catch {
        // ignore (non-DOM contexts/tests)
      }
    },
    {
      category: t("commandCategory.data"),
      icon: null,
      description: t("commandDescription.view.insertPivotTable"),
      keywords: ["pivot", "pivot table", "pivotbuilder"],
    },
  );

  // Alias used by the Ribbon schema (Insert → PivotTable → From Table/Range…).
  //
  // Keep this wired through `CommandRegistry` so generic ribbon enable/disable logic works
  // (and we don't have to special-case it via `createRibbonActionsFromCommands` overrides).
  commandRegistry.registerBuiltinCommand(
    "insert.tables.pivotTable.fromTableRange",
    t("command.insert.tables.pivotTable.fromTableRange"),
    async () => {
      await commandRegistry.executeCommand("view.insertPivotTable");
    },
    {
      category: t("commandCategory.data"),
      icon: null,
      description: t("commandDescription.view.insertPivotTable"),
      keywords: ["pivot", "pivot table", "table", "range"],
      // Hide this ribbon alias from context-aware UI surfaces (command palette, etc) to avoid
      // duplicate "PivotTable…" entries (`view.insertPivotTable` is the canonical command).
      when: "false",
    },
  );

  // Page Layout → Arrange (Selection Pane).
  //
  // The ribbon schema uses this id. Register it as a builtin command so:
  // - the ribbon can rely on CommandRegistry (no `main.ts` switch-case wiring)
  // - the command palette/keybindings can invoke it consistently
  commandRegistry.registerBuiltinCommand(
    "pageLayout.arrange.selectionPane",
    "Selection Pane",
    () => {
      if (isEditingFn()) return;
      openDockPanel(PanelIds.SELECTION_PANE);

      // The panel is a React mount; wait a frame so DOM nodes exist before focusing.
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
      keywords: ["selection pane", "arrange", "objects", "drawings", "charts"],
    },
  );
  // --- What-If Analysis / Solver (ribbon ids) ---------------------------------
  // Register the ribbon ids directly so:
  // - the ribbon doesn't have to rely on `main.ts` fallbacks to stay enabled
  // - commands are available via the command palette/keybindings
  commandRegistry.registerBuiltinCommand(
    "data.forecast.whatIfAnalysis.scenarioManager",
    `${t("whatIf.scenario.title")}…`,
    () => {
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // ignore
        }
        return;
      }
      openDockPanel(PanelIds.SCENARIO_MANAGER);
    },
    {
      category: commandCategoryData,
      icon: null,
      description: t("commandDescription.data.forecast.whatIfAnalysis.scenarioManager"),
      keywords: ["what-if", "what if", "scenario", "manager"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "data.forecast.whatIfAnalysis.monteCarlo",
    `${t("whatIf.monteCarlo.title")}…`,
    () => {
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // ignore
        }
        return;
      }
      openDockPanel(PanelIds.MONTE_CARLO);
    },
    {
      category: commandCategoryData,
      icon: null,
      description: t("commandDescription.data.forecast.whatIfAnalysis.monteCarlo"),
      keywords: ["what-if", "what if", "monte carlo", "simulation"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "data.forecast.whatIfAnalysis.goalSeek",
    `${t("whatIf.goalSeek.title")}…`,
    () => {
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // ignore
        }
        return;
      }
      if (openGoalSeekDialog) {
        openGoalSeekDialog();
        return;
      }
      try {
        showToast("Goal Seek is not available in this build.", "error");
      } catch {
        // In unit-test contexts there may be no #toast-root; fall back to a console warning.
        console.warn("Goal Seek is not available in this build.");
      }
    },
    {
      category: commandCategoryData,
      icon: null,
      description: t("commandDescription.data.forecast.whatIfAnalysis.goalSeek"),
      keywords: ["what-if", "what if", "goal seek"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "formulas.solutions.solver",
    t("panels.solver.title"),
    () => {
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // ignore
        }
        return;
      }
      openDockPanel(PanelIds.SOLVER);
    },
    {
      category: commandCategoryData,
      icon: null,
      description: t("commandDescription.formulas.solutions.solver"),
      keywords: ["solver", "optimization"],
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
    "view.togglePanel.solver",
    t("command.view.togglePanel.solver"),
    () => toggleDockPanel(PanelIds.SOLVER),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.solver"),
      keywords: ["solver", "optimization", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.scenarioManager",
    t("command.view.togglePanel.scenarioManager"),
    () => toggleDockPanel(PanelIds.SCENARIO_MANAGER),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.scenarioManager"),
      keywords: ["what-if", "scenario", "manager", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.monteCarlo",
    t("command.view.togglePanel.monteCarlo"),
    () => toggleDockPanel(PanelIds.MONTE_CARLO),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.monteCarlo"),
      keywords: ["what-if", "monte carlo", "simulation", "panel"],
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
    "view.togglePanel.vbaMigrate",
    t("command.view.togglePanel.vbaMigrate"),
    () => toggleDockPanel(PanelIds.VBA_MIGRATE),
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.togglePanel.vbaMigrate"),
      keywords: ["vba", "migrate", "macros", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.marketplace",
    t("command.view.togglePanel.marketplace"),
    () => toggleDockPanel(PanelIds.MARKETPLACE),
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
      if (isEditingFn()) return;
      // Defense-in-depth: even if some UI surface executes this command while the current
      // role cannot comment (viewer), surface a message and avoid focusing the disabled
      // composer.
      const canComment = (() => {
        // Fail closed for collab sessions: if permissions are unset/unknown or the
        // capability check throws, treat the user as unable to comment so we don't
        // surface comment write actions pre-permissions.
        let session: any = null;
        try {
          session = app.getCollabSession?.() ?? null;
        } catch {
          session = null;
        }
        if (!session) return true;
        try {
          return typeof session.canComment === "function" ? Boolean(session.canComment()) : false;
        } catch {
          return false;
        }
      })();
      if (!canComment) {
        try {
          showToast(t("comments.readOnlyHint"), "warning");
        } catch {
          // Best-effort: do not crash if the toast root is missing (tests/minimal harnesses).
        }
        app.openCommentsPanel();
        return;
      }
      app.openCommentsPanel();
      app.focusNewCommentInput();
    },
    {
      category: t("commandCategory.comments"),
      icon: null,
      description: t("commandDescription.comments.addComment"),
      keywords: ["comment", "add comment", "new comment"],
      // Viewer roles can read comments but cannot create/update them.
      when: "spreadsheet.canComment == true",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezePanes",
    t("command.view.freezePanes"),
    () => {
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      // Hide this dropdown trigger alias from context-aware UI surfaces (command palette, etc)
      // to avoid duplicate "Zoom…" entries (`view.zoom.openPicker` is the canonical command).
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.splitVertical",
    t("command.view.splitVertical"),
    () => {
      if (isEditingFn()) return;
      layoutController.setSplitDirection("vertical", 0.5);
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.splitVertical"),
      keywords: ["split", "view", "vertical"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.splitHorizontal",
    t("command.view.splitHorizontal"),
    () => {
      if (isEditingFn()) return;
      layoutController.setSplitDirection("horizontal", 0.5);
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.splitHorizontal"),
      keywords: ["split", "view", "horizontal"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.splitNone",
    t("command.view.splitNone"),
    () => {
      if (isEditingFn()) return;
      layoutController.setSplitDirection("none", 0.5);
      app.focus();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.view.splitNone"),
      keywords: ["split", "view", "unsplit", "none"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.tracePrecedents",
    t("command.audit.tracePrecedents"),
    () => {
      if (isEditingFn()) return;
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

  // Ribbon command id: keep in sync with `apps/desktop/src/ribbon/ribbonSchema.ts` to avoid schema churn.
  commandRegistry.registerBuiltinCommand(
    "formulas.formulaAuditing.tracePrecedents",
    t("command.audit.tracePrecedents"),
    () => commandRegistry.executeCommand("audit.tracePrecedents"),
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.tracePrecedents"),
      keywords: ["audit", "precedents", "trace", "ribbon"],
      // Hide this ribbon alias from context-aware UI surfaces (command palette, etc)
      // to avoid duplicate "Trace Precedents" entries (`audit.tracePrecedents` is the canonical command).
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceDependents",
    t("command.audit.traceDependents"),
    () => {
      if (isEditingFn()) return;
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

  // Ribbon command id: keep in sync with `apps/desktop/src/ribbon/ribbonSchema.ts` to avoid schema churn.
  commandRegistry.registerBuiltinCommand(
    "formulas.formulaAuditing.traceDependents",
    t("command.audit.traceDependents"),
    () => commandRegistry.executeCommand("audit.traceDependents"),
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.traceDependents"),
      keywords: ["audit", "dependents", "trace", "ribbon"],
      // Hide this ribbon alias from context-aware UI surfaces (command palette, etc)
      // to avoid duplicate "Trace Dependents" entries (`audit.traceDependents` is the canonical command).
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceBoth",
    t("command.audit.traceBoth"),
    () => {
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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

  // Ribbon command id: keep in sync with `apps/desktop/src/ribbon/ribbonSchema.ts` to avoid schema churn.
  commandRegistry.registerBuiltinCommand(
    "formulas.formulaAuditing.removeArrows",
    "Remove Arrows",
    () => {
      if (isEditingFn()) return;
      if (getTextEditingTarget()) return;
      app.clearAuditing();
      app.focus();
    },
    {
      category: t("commandCategory.audit"),
      icon: null,
      description: t("commandDescription.audit.clearAuditing"),
      keywords: ["audit", "clear", "remove", "arrows", "ribbon"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.toggleTransitive",
    t("command.audit.toggleTransitive"),
    () => {
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
    () => commandRegistry.executeCommand("clipboard.paste"),
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.clipboard.pasteSpecial.all"),
      keywords: ["paste", "clipboard", "all"],
      // This command is effectively an alias of `clipboard.paste` (see description). Hide it from
      // context-aware UI surfaces (command palette, etc) to avoid duplicate "Paste" entries.
      when: "false",
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial.values",
    t("clipboard.pasteSpecial.pasteValues"),
    async () => {
      if (getTextEditingTarget()) return;
      if ((app as any).isFormulaBarEditing?.()) return;
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
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
      if (isEditingFn()) return;
      if (isReadOnlyFn()) {
        try {
          showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "warning");
        } catch {
          // Best-effort (toast root missing in tests/headless harnesses).
        }
        focusApp();
        return;
      }
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
    "format.toggleBold",
    t("command.format.toggleBold"),
    (next?: boolean) =>
      applyFormattingToSelection(
        t("command.format.toggleBold"),
        (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges, typeof next === "boolean" ? { next } : {}),
        { forceBatch: true },
      ),
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["bold", "formatting", "font"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleItalic",
    t("command.format.toggleItalic"),
    (next?: boolean) =>
      applyFormattingToSelection(
        t("command.format.toggleItalic"),
        (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges, typeof next === "boolean" ? { next } : {}),
        { forceBatch: true },
      ),
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["italic", "formatting", "font"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleUnderline",
    t("command.format.toggleUnderline"),
    (next?: boolean) =>
      applyFormattingToSelection(
        t("command.format.toggleUnderline"),
        (doc, sheetId, ranges) => toggleUnderline(doc, sheetId, ranges, typeof next === "boolean" ? { next } : {}),
        { forceBatch: true },
      ),
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["underline", "formatting", "font"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleWrapText",
    t("command.format.toggleWrapText"),
    (next?: boolean) =>
      applyFormattingToSelection(
        t("command.format.toggleWrapText"),
        (doc, sheetId, ranges) => toggleWrap(doc, sheetId, ranges, typeof next === "boolean" ? { next } : {}),
        { forceBatch: true },
      ),
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["wrap", "wrap text", "formatting", "alignment"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fontSize.set",
    t("command.format.fontSize.set"),
    async (size?: number) => {
      const resolvedSize = (() => {
        if (typeof size === "number") return size;
        if (typeof size === "string") return Number(size);
        return null;
      })();
      if (resolvedSize == null) {
        if (!shouldOpenFormattingPrompt()) return;
        const picked = await showQuickPick(
          FONT_SIZE_STEPS.map((value) => ({ label: String(value), value })),
          { placeHolder: t("command.format.fontSize.set") },
        );
        if (picked == null) return;
        applyFormattingToSelection(t("command.format.fontSize.set"), (doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, picked));
        return;
      }

      if (!Number.isFinite(resolvedSize) || resolvedSize <= 0) return;
      applyFormattingToSelection(t("command.format.fontSize.set"), (doc, sheetId, ranges) =>
        setFontSize(doc, sheetId, ranges, resolvedSize),
      );
    },
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["font size", "formatting", "size"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fontSize.increase",
    t("command.format.fontSize.increase"),
    () => {
      const current = activeCellFontSizePt();
      const next = stepFontSize(current, "increase");
      if (next === current) return;
      applyFormattingToSelection(t("command.format.fontSize.increase"), (doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
    },
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["font size", "increase", "grow font"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fontSize.decrease",
    t("command.format.fontSize.decrease"),
    () => {
      const current = activeCellFontSizePt();
      const next = stepFontSize(current, "decrease");
      if (next === current) return;
      applyFormattingToSelection(t("command.format.fontSize.decrease"), (doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
    },
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["font size", "decrease", "shrink font"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fontColor",
    t("command.format.fontColor"),
    (color?: string | null) => {
      if (typeof document === "undefined") return;
      if (color === undefined) {
        if (!shouldOpenFormattingPrompt()) return;
        if (!fontColorPicker) fontColorPicker = createHiddenColorInput();
        openColorPicker(fontColorPicker, t("command.format.fontColor"), (doc, sheetId, ranges, argb) =>
          setFontColor(doc, sheetId, ranges, argb),
        );
        return;
      }

      if (color === null) {
        applyFormattingToSelection(t("command.format.fontColor"), (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { font: { color: null } }, { label: "Font color" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }

      const argb = normalizeArgb(color);
      if (!argb) return;
      applyFormattingToSelection(t("command.format.fontColor"), (doc, sheetId, ranges) => setFontColor(doc, sheetId, ranges, argb));
    },
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["font color", "text color", "formatting"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "format.fillColor",
    t("command.format.fillColor"),
    (color?: string | null) => {
      if (typeof document === "undefined") return;
      if (color === undefined) {
        if (!shouldOpenFormattingPrompt()) return;
        if (!fillColorPicker) fillColorPicker = createHiddenColorInput();
        openColorPicker(fillColorPicker, t("command.format.fillColor"), (doc, sheetId, ranges, argb) =>
          setFillColor(doc, sheetId, ranges, argb),
        );
        return;
      }

      if (color === null) {
        applyFormattingToSelection(t("command.format.fillColor"), (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { fill: null }, { label: "Fill color" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }

      const argb = normalizeArgb(color);
      if (!argb) return;
      applyFormattingToSelection(t("command.format.fillColor"), (doc, sheetId, ranges) => setFillColor(doc, sheetId, ranges, argb));
    },
    {
      category: commandCategoryFormat,
      icon: null,
      keywords: ["fill color", "cell color", "formatting"],
    },
  );

  registerNumberFormatCommands({
    commandRegistry,
    applyFormattingToSelection,
    getActiveCellNumberFormat: activeCellNumberFormat,
    t,
    category: commandCategoryFormat,
  });

  // Find/Replace/Go To commands are registered here so they are discoverable early (e.g. command palette),
  // but are expected to be overridden by the UI host (apps/desktop/src/main.ts) once dialogs are mounted.
  commandRegistry.registerBuiltinCommand(
    "edit.find",
    t("command.edit.find"),
    () => {},
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.find"),
      keywords: ["find", "search"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.replace",
    t("command.edit.replace"),
    () => {},
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.replace"),
      keywords: ["replace", "find"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "navigation.goTo",
    t("command.navigation.goTo"),
    () => {},
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.navigation.goTo"),
      keywords: ["go to", "goto", "reference", "name box"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.editCell",
    t("command.edit.editCell"),
    () => {
      if (isEditingFn()) return;
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
    () => {
      if (isEditingFn()) return;
      app.clearSelectionContents();
    },
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
    () => {
      if (isEditingFn()) return;
      app.fillDown();
    },
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
    () => {
      if (isEditingFn()) return;
      app.fillRight();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.fillRight"),
      keywords: ["fill", "fill right", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillUp",
    t("command.edit.fillUp"),
    () => {
      if (isEditingFn()) return;
      app.fillUp();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.fillUp"),
      keywords: ["fill", "fill up", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillLeft",
    t("command.edit.fillLeft"),
    () => {
      if (isEditingFn()) return;
      app.fillLeft();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.fillLeft"),
      keywords: ["fill", "fill left", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.selectCurrentRegion",
    t("command.edit.selectCurrentRegion"),
    () => {
      if (isEditingFn()) return;
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
    () => {
      if (isEditingFn()) return;
      app.autoSum();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      description: t("commandDescription.edit.autoSum"),
      keywords: ["autosum", "sum", "excel"],
    },
  );

  // Ribbon-specific AutoSum dropdown variants (Home → Editing → AutoSum menu items).
  //
  // These are implemented by SpreadsheetApp (and were historically routed via `main.ts`), but are
  // now registered so the ribbon can rely on the CommandRegistry baseline enable/disable behavior.
  commandRegistry.registerBuiltinCommand(
    "home.editing.autoSum.average",
    "AutoSum: Average",
    () => {
      if (isEditingFn()) return;
      app.autoSumAverage();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      keywords: ["autosum", "sum", "average", "excel"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.editing.autoSum.countNumbers",
    "AutoSum: Count Numbers",
    () => {
      if (isEditingFn()) return;
      app.autoSumCountNumbers();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      keywords: ["autosum", "sum", "count", "count numbers", "excel"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.editing.autoSum.max",
    "AutoSum: Max",
    () => {
      if (isEditingFn()) return;
      app.autoSumMax();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      keywords: ["autosum", "sum", "max", "maximum", "excel"],
    },
  );
  commandRegistry.registerBuiltinCommand(
    "home.editing.autoSum.min",
    "AutoSum: Min",
    () => {
      if (isEditingFn()) return;
      app.autoSumMin();
    },
    {
      category: t("commandCategory.editing"),
      icon: null,
      keywords: ["autosum", "sum", "min", "minimum", "excel"],
    },
  );

  if (themeController) {
    const categoryView = t("commandCategory.view");

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.system",
      t("command.view.theme.system"),
      () => commandRegistry.executeCommand("view.theme.system"),
      {
        category: categoryView,
        icon: null,
        description: t("commandDescription.view.theme.system"),
        keywords: ["theme", "appearance", "system", "auto", "os", "dark mode", "light mode"],
        // Ribbon schema uses `view.appearance.theme.*` ids for the View → Appearance menu, but the
        // command palette should surface the canonical `view.theme.*` commands only.
        when: "false",
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.light",
      t("command.view.theme.light"),
      () => commandRegistry.executeCommand("view.theme.light"),
      {
        category: categoryView,
        icon: null,
        description: t("commandDescription.view.theme.light"),
        keywords: ["theme", "appearance", "light", "light mode"],
        // Hide ribbon alias from command palette (canonical: `view.theme.light`).
        when: "false",
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.dark",
      t("command.view.theme.dark"),
      () => commandRegistry.executeCommand("view.theme.dark"),
      {
        category: categoryView,
        icon: null,
        description: t("commandDescription.view.theme.dark"),
        keywords: ["theme", "appearance", "dark", "dark mode"],
        // Hide ribbon alias from command palette (canonical: `view.theme.dark`).
        when: "false",
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme.highContrast",
      t("command.view.theme.highContrast"),
      () => commandRegistry.executeCommand("view.theme.highContrast"),
      {
        category: categoryView,
        icon: null,
        description: t("commandDescription.view.theme.highContrast"),
        keywords: ["theme", "appearance", "high contrast", "contrast", "accessibility", "a11y"],
        // Hide ribbon alias from command palette (canonical: `view.theme.highContrast`).
        when: "false",
      },
    );

    commandRegistry.registerBuiltinCommand(
      "view.appearance.theme",
      t("command.view.appearance.theme"),
      async () => {
        const picked = await showQuickPick(
          [
            { label: t("ribbon.theme.system"), value: "view.appearance.theme.system" },
            { label: t("ribbon.theme.light"), value: "view.appearance.theme.light" },
            { label: t("ribbon.theme.dark"), value: "view.appearance.theme.dark" },
            { label: t("ribbon.theme.highContrast"), value: "view.appearance.theme.highContrast" },
          ],
          { placeHolder: t("quickPick.theme.placeholder") },
        );
        if (!picked) return;
        await commandRegistry.executeCommand(picked);
      },
      {
        category: categoryView,
        icon: null,
        description: t("commandDescription.view.appearance.theme"),
        keywords: ["theme", "appearance", "dark mode", "light mode"],
      },
    );
  }
}
