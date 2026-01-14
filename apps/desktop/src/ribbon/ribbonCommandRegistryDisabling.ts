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
  "home.number.moreFormats.custom",
  "home.cells.format.formatCells",
  "home.cells.format.organizeSheets",
  // Insert/delete cells (not whole sheets). Handled directly by `main.ts` via a dialog-style quick pick.
  "home.cells.insert.insertCells",
  "home.cells.delete.deleteCells",
  // Structural sheet row/col/sheet operations are handled directly by `main.ts` (not CommandRegistry).
  "home.cells.insert.insertSheetRows",
  "home.cells.insert.insertSheetColumns",
  "home.cells.insert.insertSheet",
  "home.cells.delete.deleteSheetRows",
  "home.cells.delete.deleteSheetColumns",
  "home.cells.delete.deleteSheet",

  // --- Home editing -----------------------------------------------------------
  "home.editing.autoSum.average",
  "home.editing.autoSum.countNumbers",
  "home.editing.autoSum.max",
  "home.editing.autoSum.min",
  "home.editing.fill.up",
  "home.editing.fill.left",
  "home.editing.fill.series",
  // Find & Select dropdown menu items are canonical commands; keep these legacy ids enabled in case
  // older ribbon schemas emit them (the current schema uses `edit.find` / `edit.replace` / `navigation.goTo`).
  "home.editing.findSelect.find",
  "home.editing.findSelect.replace",
  "home.editing.findSelect.goTo",

  // Sort/filter (ribbon-only handlers / partially implemented).
  "home.editing.sortFilter.customSort",
  // Data → Sort & Filter (same implementations as Home tab).
  "data.sortFilter.sort.customSort",

  // --- Home → Styles ----------------------------------------------------------
  // Cell Styles are currently handled via `main.ts` (not CommandRegistry yet). Only the
  // Good/Bad/Neutral submenu is implemented today.
  "home.styles.cellStyles.goodBadNeutral",
  "home.styles.formatAsTable",
  "home.styles.formatAsTable.light",
  "home.styles.formatAsTable.medium",
  "home.styles.formatAsTable.dark",
  "home.styles.formatAsTable.newStyle",

  // --- Formula auditing -------------------------------------------------------
  "formulas.formulaAuditing.tracePrecedents",
  "formulas.formulaAuditing.traceDependents",
  "formulas.formulaAuditing.removeArrows",

  // --- Insert pictures --------------------------------------------------------
  // Insert → Pictures menu items are routed via `handleInsertPicturesRibbonCommand` in main.ts.
  "insert.illustrations.pictures",
  "insert.illustrations.pictures.thisDevice",
  "insert.illustrations.pictures.stockImages",
  "insert.illustrations.pictures.onlinePictures",
  "insert.illustrations.onlinePictures",

  // --- View -------------------------------------------------------------------
  "view.appearance.theme.system",
  "view.appearance.theme.light",
  "view.appearance.theme.dark",
  "view.appearance.theme.highContrast",
]);

function defaultIsExemptFromCommandRegistry(commandId: string): boolean {
  return COMMAND_REGISTRY_EXEMPT_IDS.has(commandId);
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
