import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";

export function registerBuiltinCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController;
  ensureExtensionsLoaded?: (() => Promise<void>) | null;
  onExtensionsLoaded?: (() => void) | null;
}): void {
  const { commandRegistry, app, layoutController, ensureExtensionsLoaded = null, onExtensionsLoaded = null } = params;

  const toggleDockPanel = (panelId: string) => {
    const placement = getPanelPlacement(layoutController.layout, panelId);
    if (placement.kind === "closed") layoutController.openPanel(panelId);
    else layoutController.closePanel(panelId);
  };

  const listVisibleSheetIds = (): string[] => {
    const ids = app.getDocument().getSheetIds();
    // DocumentController materializes sheets lazily; mimic the UI fallback behavior so
    // navigation commands are stable even before any edits occur.
    return ids.length > 0 ? ids : ["Sheet1"];
  };

  const activateRelativeSheet = (delta: -1 | 1): void => {
    const sheetIds = listVisibleSheetIds();
    if (sheetIds.length <= 1) return;
    const active = app.getCurrentSheetId();
    const activeIndex = sheetIds.indexOf(active);
    const idx = activeIndex >= 0 ? activeIndex : 0;
    const nextIndex = (idx + delta + sheetIds.length) % sheetIds.length;
    const next = sheetIds[nextIndex];
    if (!next || next === active) return;
    app.activateSheet(next);
    app.focus();
  };

  commandRegistry.registerBuiltinCommand(
    "workbench.showCommandPalette",
    "Show Command Palette",
    () => {
      // Intentionally a no-op: the desktop shell owns opening the palette, but we still
      // register the id so keybinding and menu systems can reference it.
    },
    {
      category: "Navigation",
      icon: null,
      description: "Show the command palette",
      keywords: ["command palette", "commands"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbook.previousSheet",
    "Previous Sheet",
    () => activateRelativeSheet(-1),
    {
      category: "Navigation",
      icon: null,
      description: "Activate the previous visible sheet (wrap around)",
      keywords: ["sheet", "previous", "navigation", "pageup", "pgup"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "workbook.nextSheet",
    "Next Sheet",
    () => activateRelativeSheet(1),
    {
      category: "Navigation",
      icon: null,
      description: "Activate the next visible sheet (wrap around)",
      keywords: ["sheet", "next", "navigation", "pagedown", "pgdn"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.insertPivotTable",
    t("commandPalette.command.insertPivotTable"),
    () => {
      layoutController.openPanel(PanelIds.PIVOT_BUILDER);
      // If the panel is already open, we still want to refresh its source range from
      // the latest selection.
      window.dispatchEvent(new CustomEvent("pivot-builder:use-selection"));
    },
    {
      category: "Data",
      icon: null,
      description: "Open the Pivot Builder panel for the current selection",
      keywords: ["pivot", "pivot table", "pivotbuilder"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.aiChat",
    "Toggle AI Chat",
    () => toggleDockPanel(PanelIds.AI_CHAT),
    {
      category: "AI",
      icon: null,
      description: "Toggle the AI Chat panel",
      keywords: ["ai", "chat", "assistant", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.aiAudit",
    "Toggle AI Audit",
    () => toggleDockPanel(PanelIds.AI_AUDIT),
    {
      category: "AI",
      icon: null,
      description: "Toggle the AI Audit panel",
      keywords: ["ai", "audit", "log", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.extensions",
    "Toggle Extensions",
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
      category: "View",
      icon: null,
      description: "Toggle the Extensions panel",
      keywords: ["extensions", "plugins", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.macros",
    "Toggle Macros",
    () => toggleDockPanel(PanelIds.MACROS),
    {
      category: "View",
      icon: null,
      description: "Toggle the Macros panel",
      keywords: ["macros", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.dataQueries",
    "Toggle Data Queries",
    () => toggleDockPanel(PanelIds.DATA_QUERIES),
    {
      category: "View",
      icon: null,
      description: "Toggle the Data / Queries panel",
      keywords: ["data", "queries", "power query", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.scriptEditor",
    "Toggle Script Editor",
    () => toggleDockPanel(PanelIds.SCRIPT_EDITOR),
    {
      category: "View",
      icon: null,
      description: "Toggle the Script Editor panel",
      keywords: ["script", "editor", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.togglePanel.python",
    "Toggle Python",
    () => toggleDockPanel(PanelIds.PYTHON),
    {
      category: "View",
      icon: null,
      description: "Toggle the Python panel",
      keywords: ["python", "panel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezePanes",
    "Freeze Panes",
    () => {
      app.freezePanes();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Freeze rows and columns based on the current selection",
      keywords: ["freeze", "panes"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezeTopRow",
    "Freeze Top Row",
    () => {
      app.freezeTopRow();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Freeze row 1",
      keywords: ["freeze", "top row"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.freezeFirstColumn",
    "Freeze First Column",
    () => {
      app.freezeFirstColumn();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Freeze column A",
      keywords: ["freeze", "first column"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "view.unfreezePanes",
    "Unfreeze Panes",
    () => {
      app.unfreezePanes();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Unfreeze all panes",
      keywords: ["unfreeze", "panes"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.tracePrecedents",
    "Trace Precedents",
    () => {
      app.clearAuditing();
      app.toggleAuditingPrecedents();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Show precedent arrows for the active cell",
      keywords: ["audit", "precedents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceDependents",
    "Trace Dependents",
    () => {
      app.clearAuditing();
      app.toggleAuditingDependents();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Show dependent arrows for the active cell",
      keywords: ["audit", "dependents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.traceBoth",
    "Trace Precedents + Dependents",
    () => {
      app.clearAuditing();
      app.toggleAuditingPrecedents();
      app.toggleAuditingDependents();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Show both precedent and dependent arrows for the active cell",
      keywords: ["audit", "precedents", "dependents", "trace"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.clearAuditing",
    "Clear Auditing",
    () => {
      app.clearAuditing();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Clear all auditing arrows",
      keywords: ["audit", "clear"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "audit.toggleTransitive",
    "Toggle Transitive Auditing",
    () => {
      app.toggleAuditingTransitive();
      app.focus();
    },
    {
      category: "View",
      icon: null,
      description: "Toggle whether auditing follows references transitively",
      keywords: ["audit", "transitive", "toggle"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.copy",
    t("clipboard.copy"),
    () => app.copyToClipboard(),
    {
      category: "Editing",
      icon: null,
      description: "Copy the current selection",
      keywords: ["copy", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.cut",
    t("clipboard.cut"),
    () => app.cutToClipboard(),
    {
      category: "Editing",
      icon: null,
      description: "Cut the current selection",
      keywords: ["cut", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.paste",
    t("clipboard.paste"),
    () => app.pasteFromClipboard(),
    {
      category: "Editing",
      icon: null,
      description: "Paste from the clipboard into the current selection",
      keywords: ["paste", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "clipboard.pasteSpecial",
    "Paste Special…",
    (mode?: unknown) => app.clipboardPasteSpecial((mode as any) ?? "all"),
    {
      category: "Editing",
      icon: null,
      description: "Paste with a specific mode (values, formulas, formats, etc.)",
      keywords: ["paste", "paste special", "clipboard"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.clearContents",
    "Clear Contents",
    () => app.clearContents(),
    {
      category: "Editing",
      icon: null,
      description: "Clear cell contents in the current selection",
      keywords: ["clear", "contents", "delete"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillDown",
    "Fill Down",
    () => app.fillDown(),
    {
      category: "Editing",
      icon: null,
      description: "Fill the selection down (Excel: Ctrl+D)",
      keywords: ["fill", "fill down", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.fillRight",
    "Fill Right",
    () => app.fillRight(),
    {
      category: "Editing",
      icon: null,
      description: "Fill the selection right (Excel: Ctrl+R)",
      keywords: ["fill", "fill right", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.insertDate",
    "Insert Date",
    () => app.insertDate(),
    {
      category: "Editing",
      icon: null,
      description: "Insert the current date into the selection (Excel: Ctrl+;)",
      keywords: ["date", "insert date", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.insertTime",
    "Insert Time",
    () => app.insertTime(),
    {
      category: "Editing",
      icon: null,
      description: "Insert the current time into the selection (Excel: Ctrl+Shift+;)",
      keywords: ["time", "insert time", "excel"],
    },
  );

  commandRegistry.registerBuiltinCommand(
    "edit.autoSum",
    "AutoSum",
    () => app.autoSum(),
    {
      category: "Editing",
      icon: null,
      description: "Insert a SUM formula based on adjacent numeric cells (Excel: Alt+=)",
      keywords: ["autosum", "sum", "excel"],
    },
  );
}
