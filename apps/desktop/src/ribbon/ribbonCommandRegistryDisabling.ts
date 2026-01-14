import type { CommandRegistry } from "../extensions/commandRegistry.js";

import { defaultRibbonSchema, type RibbonSchema } from "./ribbonSchema.js";

/**
 * Ribbon controls intentionally handled outside `CommandRegistry`.
 *
 * These ids will remain enabled even if the command registry doesn't have a matching command.
 *
 * NOTE: Keep this list small and focused — prefer registering real commands in `CommandRegistry`
 * when possible so other UI surfaces (e.g. command palette / keybindings) stay consistent.
 */
const COMMAND_REGISTRY_EXEMPT_IDS = new Set<string>([
  // --- Desktop/file actions ----------------------------------------------------
  "file.save.autoSave",
  "file.new.new",
  "file.new.blankWorkbook",
  "file.open.open",
  "file.save.save",
  "file.save.saveAs",
  "file.save.saveAs.copy",
  "file.save.saveAs.download",
  "file.info.manageWorkbook.versions",
  "file.info.manageWorkbook.branches",
  "file.export.createPdf",
  "file.export.export.pdf",
  "file.export.changeFileType.pdf",
  "file.export.export.csv",
  "file.export.changeFileType.csv",
  "file.export.changeFileType.tsv",
  "file.export.export.xlsx",
  "file.export.changeFileType.xlsx",
  "file.print.print",
  "file.print.printPreview",
  "file.print.pageSetup",
  "file.print.pageSetup.printTitles",
  "file.print.pageSetup.margins",
  "file.options.close",

  // --- Clipboard --------------------------------------------------------------
  "home.clipboard.cut",
  "home.clipboard.copy",
  "home.clipboard.formatPainter",
  "home.clipboard.paste",
  "home.clipboard.paste.default",
  "home.clipboard.paste.values",
  "home.clipboard.paste.formulas",
  "home.clipboard.paste.formats",
  "home.clipboard.paste.transpose",
  "home.clipboard.pasteSpecial",
  "home.clipboard.pasteSpecial.dialog",
  "home.clipboard.pasteSpecial.values",
  "home.clipboard.pasteSpecial.formulas",
  "home.clipboard.pasteSpecial.formats",
  "home.clipboard.pasteSpecial.transpose",

  // --- Formatting (top-level controls) ---------------------------------------
  "home.font.bold",
  "home.font.italic",
  "home.font.underline",
  "home.font.strikethrough",
  "home.font.fontName",
  "home.font.fontSize",
  "home.font.fontColor",
  "home.font.fillColor",
  "home.font.borders",
  "home.font.clearFormatting",
  "home.alignment.wrapText",
  "home.alignment.alignLeft",
  "home.alignment.center",
  "home.alignment.alignRight",
  "home.alignment.topAlign",
  "home.alignment.middleAlign",
  "home.alignment.bottomAlign",
  "home.alignment.increaseIndent",
  "home.alignment.decreaseIndent",
  "home.alignment.orientation.angleCounterclockwise",
  "home.alignment.orientation.angleClockwise",
  "home.alignment.orientation.verticalText",
  "home.alignment.orientation.rotateUp",
  "home.alignment.orientation.rotateDown",
  "home.alignment.orientation.formatCellAlignment",
  // Merge commands are routed via the ribbon fallback handler in main.ts (not CommandRegistry yet).
  "home.alignment.mergeCenter",
  "home.alignment.mergeCenter.mergeCenter",
  "home.alignment.mergeCenter.mergeAcross",
  "home.alignment.mergeCenter.mergeCells",
  "home.alignment.mergeCenter.unmergeCells",
  "home.number.numberFormat",
  "home.number.percent",
  "home.number.accounting",
  "home.number.date",
  "home.number.comma",
  "home.number.increaseDecimal",
  "home.number.decreaseDecimal",
  "home.number.formatCells",
  "home.number.moreFormats.formatCells",
  "home.number.moreFormats.custom",
  "home.cells.format.formatCells",
  "home.cells.format.rowHeight",
  "home.cells.format.columnWidth",

  // --- Home editing -----------------------------------------------------------
  "home.editing.autoSum",
  "home.editing.autoSum.sum",
  "home.editing.fill.down",
  "home.editing.fill.right",
  "home.editing.fill.up",
  "home.editing.fill.left",

  // --- Comments ---------------------------------------------------------------
  "review.comments.newComment",
  "review.comments.showComments",

  // --- Formula auditing -------------------------------------------------------
  "formulas.formulaAuditing.showFormulas",
  "formulas.formulaAuditing.tracePrecedents",
  "formulas.formulaAuditing.traceDependents",
  "formulas.formulaAuditing.removeArrows",

  // --- Pivot table (panel) ----------------------------------------------------
  "insert.tables.pivotTable",

  // --- View -------------------------------------------------------------------
  "view.macros.viewMacros",
  "view.macros.viewMacros.run",
  "view.macros.viewMacros.edit",
  "view.macros.viewMacros.delete",
  "view.macros.recordMacro",
  "view.macros.recordMacro.stop",
  "view.macros.useRelativeReferences",
  "view.show.showFormulas",
  "view.show.performanceStats",
  "view.window.split",
  "view.zoom.zoom",
  "view.zoom.zoom100",
  "view.zoom.zoomToSelection",
  "view.appearance.theme.system",
  "view.appearance.theme.light",
  "view.appearance.theme.dark",
  "view.appearance.theme.highContrast",

  // --- Macro recorder toggles -------------------------------------------------
  "developer.code.useRelativeReferences",
  "developer.code.visualBasic",

  // --- Developer tab macro controls ------------------------------------------
  "developer.code.macros",
  "developer.code.macros.run",
  "developer.code.macros.edit",
  "developer.code.macroSecurity",
  "developer.code.macroSecurity.trustCenter",
  "developer.code.recordMacro",
  "developer.code.recordMacro.stop",
]);

function isExemptViaPattern(commandId: string): boolean {
  // Fill color presets.
  if (commandId.startsWith("home.font.fillColor.")) {
    const preset = commandId.slice("home.font.fillColor.".length);
    return (
      preset === "none" ||
      preset === "noFill" ||
      preset === "lightGray" ||
      preset === "yellow" ||
      preset === "blue" ||
      preset === "green" ||
      preset === "red" ||
      preset === "moreColors"
    );
  }

  // Font color presets.
  if (commandId.startsWith("home.font.fontColor.")) {
    const preset = commandId.slice("home.font.fontColor.".length);
    return preset === "automatic" || preset === "black" || preset === "blue" || preset === "green" || preset === "red" || preset === "moreColors";
  }

  // Clear formatting menu items.
  if (commandId.startsWith("home.font.clearFormatting.")) {
    const kind = commandId.slice("home.font.clearFormatting.".length);
    return kind === "clearFormats" || kind === "clearContents" || kind === "clearAll";
  }

  // Borders menu items.
  if (commandId.startsWith("home.font.borders.")) {
    const kind = commandId.slice("home.font.borders.".length);
    return (
      kind === "none" ||
      kind === "all" ||
      kind === "outside" ||
      kind === "thickBox" ||
      kind === "bottom" ||
      kind === "top" ||
      kind === "left" ||
      kind === "right"
    );
  }

  // Number format dropdown menu items.
  if (commandId.startsWith("home.number.numberFormat.")) {
    const kind = commandId.slice("home.number.numberFormat.".length);
    return (
      kind === "general" ||
      kind === "number" ||
      kind === "currency" ||
      kind === "accounting" ||
      kind === "percentage" ||
      kind === "shortDate" ||
      kind === "longDate" ||
      kind === "time" ||
      kind === "fraction" ||
      kind === "scientific" ||
      kind === "text"
    );
  }

  // Accounting dropdown currency menu items.
  if (commandId.startsWith("home.number.accounting.")) {
    const kind = commandId.slice("home.number.accounting.".length);
    return kind === "usd" || kind === "eur" || kind === "gbp" || kind === "jpy";
  }

  return false;
}

function defaultIsExemptFromCommandRegistry(commandId: string): boolean {
  return COMMAND_REGISTRY_EXEMPT_IDS.has(commandId) || isExemptViaPattern(commandId);
}

function isRegistered(commandRegistry: CommandRegistry, commandId: string): boolean {
  try {
    return commandRegistry.getCommand(commandId) != null;
  } catch {
    return false;
  }
}

export type RibbonCommandRegistryDisablingOptions = {
  schema?: RibbonSchema;
  /**
   * Optional override for which ribbon ids should remain enabled even when not registered.
   */
  isExemptFromCommandRegistry?: (commandId: string) => boolean;
};

/**
 * Compute a baseline `disabledById` override for the ribbon:
 * - If a ribbon control is backed by a registered `CommandRegistry` command, it stays enabled.
 * - If it is not registered, it is disabled by default.
 * - Some ids are intentionally handled outside of `CommandRegistry` and are exempted.
 *
 * Dropdown controls with menus are treated specially:
 * - Their *menu items* are evaluated as commands.
 * - The dropdown trigger itself is disabled only when it is not registered/exempt AND all of its
 *   menu items are disabled (so the menu would be entirely non-functional).
 */
export function computeRibbonDisabledByIdFromCommandRegistry(
  commandRegistry: CommandRegistry,
  options: RibbonCommandRegistryDisablingOptions = {},
): Record<string, boolean> {
  const schema = options.schema ?? defaultRibbonSchema;
  const isExempt = options.isExemptFromCommandRegistry ?? defaultIsExemptFromCommandRegistry;

  const disabledById: Record<string, boolean> = Object.create(null);

  const shouldDisableCommandId = (commandId: string): boolean => !isRegistered(commandRegistry, commandId) && !isExempt(commandId);

  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        const kind = button.kind ?? "button";
        const menuItems = Array.isArray(button.menuItems) ? button.menuItems : null;
        const hasMenu = kind === "dropdown" && Boolean(menuItems?.length);

        if (hasMenu) {
          let allItemsDisabled = true;
          for (const item of menuItems!) {
            if (shouldDisableCommandId(item.id)) {
              disabledById[item.id] = true;
            } else {
              allItemsDisabled = false;
            }
          }

          // Only disable the trigger if the entire menu would be disabled and the trigger itself
          // is not registered/exempt. This keeps "menu-only" controls usable without having to
          // register a separate trigger command id.
          if (allItemsDisabled && shouldDisableCommandId(button.id)) {
            disabledById[button.id] = true;
          }

          continue;
        }

        if (shouldDisableCommandId(button.id)) {
          disabledById[button.id] = true;
        }
      }
    }
  }

  return disabledById;
}
