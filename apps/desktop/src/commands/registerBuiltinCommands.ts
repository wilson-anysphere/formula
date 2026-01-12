import type { SpreadsheetApp } from "../app/spreadsheetApp";
import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";

export function registerBuiltinCommands(params: {
  commandRegistry: CommandRegistry;
  app: SpreadsheetApp;
  layoutController: LayoutController;
}): void {
  const { commandRegistry, app, layoutController } = params;

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
}
