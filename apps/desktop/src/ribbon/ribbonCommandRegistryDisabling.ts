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
export const COMMAND_REGISTRY_EXEMPT_IDS: ReadonlySet<string> = new Set<string>([
  // --- File tab / backstage actions --------------------------------------------
  //
  // File operations are routed through `RibbonActions.fileActions` and/or special-cased
  // handling in `apps/desktop/src/main.ts` (they are not CommandRegistry ids).
  "file.new.blankWorkbook",
  "file.open.open",
  "file.save.save",
  "file.save.saveAs",
  "file.save.saveAs.copy",
  "file.save.saveAs.download",
  "file.save.autoSave",
  "file.info.manageWorkbook.versions",
  "file.info.manageWorkbook.branches",
  "file.export.createPdf",
  "file.export.export.pdf",
  "file.export.export.csv",
  "file.export.export.xlsx",
  "file.export.changeFileType.pdf",
  "file.export.changeFileType.csv",
  "file.export.changeFileType.tsv",
  "file.export.changeFileType.xlsx",
  "file.print.print",
  "file.print.printPreview",
  "file.print.pageSetup",
  "file.print.pageSetup.printTitles",
  "file.print.pageSetup.margins",
  "file.options.close",

  // --- Ribbon-only handlers (not CommandRegistry yet) --------------------------
  //
  // These ids currently dispatch through `onUnknownCommand` / ribbon overrides in the
  // desktop shell. If/when they become real commands, remove them from this list.

  // Home → Alignment → Merge & Center.
  // Implemented by the desktop ribbon fallback handler (`apps/desktop/src/ribbon/commandHandlers.ts`).
  "home.alignment.mergeCenter.mergeCenter",
  "home.alignment.mergeCenter.mergeAcross",
  "home.alignment.mergeCenter.mergeCells",
  "home.alignment.mergeCenter.unmergeCells",

  // Home → Number → More Formats.
  "home.number.moreFormats.custom",

  // Home → Cells → Format.
  "home.cells.format.organizeSheets",

  // Home → Cells (structural edits).
  "home.cells.insert.insertCells",
  "home.cells.delete.deleteCells",
  "home.cells.insert.insertSheetRows",
  "home.cells.insert.insertSheetColumns",
  "home.cells.insert.insertSheet",
  "home.cells.delete.deleteSheetRows",
  "home.cells.delete.deleteSheetColumns",
  "home.cells.delete.deleteSheet",

  // Home → Editing.
  "home.editing.fill.series",
  "home.editing.sortFilter.customSort",
  "home.editing.sortFilter.filter",
  "home.editing.sortFilter.clear",
  "home.editing.sortFilter.reapply",

  // Data → Sort & Filter.
  "data.sortFilter.sort.customSort",
  "data.sortFilter.filter",
  "data.sortFilter.clear",
  "data.sortFilter.reapply",
  "data.sortFilter.advanced.clearFilter",

  // Home → Styles.
  "home.styles.formatAsTable.light",
  "home.styles.formatAsTable.medium",
  "home.styles.formatAsTable.dark",
  "home.styles.formatAsTable.newStyle",
  "home.styles.cellStyles.goodBadNeutral",
  "home.styles.cellStyles.dataModel",
  "home.styles.cellStyles.titlesHeadings",
  "home.styles.cellStyles.numberFormat",
  "home.styles.cellStyles.newStyle",

  // Insert → Pictures.
  "insert.illustrations.pictures.thisDevice",
  "insert.illustrations.pictures.stockImages",
  "insert.illustrations.pictures.onlinePictures",
  "insert.illustrations.onlinePictures",
]);

function isExemptViaPattern(_commandId: string): boolean {
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
