import { getPrimaryCommandKeybindingDisplay } from "../extensions/keybindings.js";

/**
 * Mapping from ribbon command ids (as defined in `ribbonSchema.ts`) to command ids
 * used by the KeybindingService / command palette.
 *
 * The ribbon uses a mix of Excel-style ids (e.g. `home.clipboard.copy`) and direct
 * built-in command ids (e.g. `format.toggleBold`). This lookup normalizes ribbon
 * ids to the corresponding KeybindingService command id so ribbon tooltips/menus
 * can display keybinding hints.
 */
const KEYBINDING_COMMAND_BY_RIBBON_ID: Record<string, string> = {
  // --- Clipboard --------------------------------------------------------------
  "home.clipboard.cut": "clipboard.cut",
  "home.clipboard.copy": "clipboard.copy",
  "home.clipboard.paste": "clipboard.paste",
  "home.clipboard.paste.default": "clipboard.paste",
  "home.clipboard.pasteSpecial": "clipboard.pasteSpecial",
  "home.clipboard.pasteSpecial.dialog": "clipboard.pasteSpecial",

  // --- Find/Replace/Go To -----------------------------------------------------
  "edit.find": "edit.find",
  "edit.replace": "edit.replace",
  "navigation.goTo": "navigation.goTo",

  // --- Formatting -------------------------------------------------------------
  "format.toggleBold": "format.toggleBold",
  "format.toggleItalic": "format.toggleItalic",
  "format.toggleUnderline": "format.toggleUnderline",

  "home.number.accounting": "format.numberFormat.currency",
  "home.number.numberFormat.currency": "format.numberFormat.currency",
  "home.number.numberFormat.accounting": "format.numberFormat.currency",
  "home.number.percent": "format.numberFormat.percent",
  "home.number.numberFormat.percentage": "format.numberFormat.percent",
  "home.number.date": "format.numberFormat.date",
  "home.number.numberFormat.shortDate": "format.numberFormat.date",
  "home.number.numberFormat.longDate": "format.numberFormat.date",
  "home.number.formatCells": "format.openFormatCells",
  "home.number.moreFormats.formatCells": "format.openFormatCells",

  // --- Comments ---------------------------------------------------------------
  "comments.addComment": "comments.addComment",
  "comments.togglePanel": "comments.togglePanel",

  // --- View ------------------------------------------------------------------
  "view.show.showFormulas": "view.toggleShowFormulas",
  "formulas.formulaAuditing.showFormulas": "view.toggleShowFormulas",
  "open-panel-ai-chat": "view.togglePanel.aiChat",
  "open-inline-ai-edit": "ai.inlineEdit",
};

export function deriveRibbonShortcutById(commandKeybindingDisplayIndex: Map<string, string[]>): Record<string, string> {
  const shortcutById: Record<string, string> = Object.create(null);

  for (const [ribbonId, keybindingCommandId] of Object.entries(KEYBINDING_COMMAND_BY_RIBBON_ID)) {
    const shortcut = getPrimaryCommandKeybindingDisplay(keybindingCommandId, commandKeybindingDisplayIndex);
    if (!shortcut) continue;
    shortcutById[ribbonId] = shortcut;
  }

  return shortcutById;
}
