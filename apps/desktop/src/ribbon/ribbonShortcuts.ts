import { getPrimaryCommandKeybindingDisplay } from "../extensions/keybindings.js";

/**
 * Mapping from ribbon command ids (as defined in `ribbonSchema.ts`) to command ids
 * used by the KeybindingService / command palette.
 *
 * The ribbon uses a mix of canonical CommandRegistry ids (e.g. `clipboard.copy`)
 * and Excel-style ids (e.g. `home.number.accounting`). This lookup normalizes
 * ribbon ids to the corresponding KeybindingService command id so ribbon
 * tooltips/menus can display keybinding hints.
 */
const KEYBINDING_COMMAND_BY_RIBBON_ID: Record<string, string> = {
  // --- Clipboard --------------------------------------------------------------
  "clipboard.cut": "clipboard.cut",
  "clipboard.copy": "clipboard.copy",
  "clipboard.paste": "clipboard.paste",
  "clipboard.pasteSpecial": "clipboard.pasteSpecial",

  // --- Find/Replace/Go To -----------------------------------------------------
  // (These happen to share ids with the built-in commands, but keep the mapping
  // explicit for clarity.)
  "edit.find": "edit.find",
  "edit.replace": "edit.replace",
  "navigation.goTo": "navigation.goTo",

  // --- Formatting -------------------------------------------------------------
  // (Most formatting buttons use built-in ids directly, so they are picked up by
  // the identity mapping in `deriveRibbonShortcutById` below.)

  "home.number.numberFormat.currency": "format.numberFormat.currency",
  "home.number.numberFormat.accounting": "format.numberFormat.currency",
  "home.number.numberFormat.percentage": "format.numberFormat.percent",
  "home.number.numberFormat.shortDate": "format.numberFormat.date",
  "home.number.numberFormat.longDate": "format.numberFormat.date",
  "home.number.formatCells": "format.openFormatCells",
  "home.number.moreFormats.formatCells": "format.openFormatCells",

  // --- Comments ---------------------------------------------------------------
  "comments.addComment": "comments.addComment",
  "comments.togglePanel": "comments.togglePanel",

  // --- View ------------------------------------------------------------------
  "open-panel-ai-chat": "view.togglePanel.aiChat",
  "open-inline-ai-edit": "ai.inlineEdit",
};

export function deriveRibbonShortcutById(commandKeybindingDisplayIndex: Map<string, string[]>): Record<string, string> {
  const shortcutById: Record<string, string> = Object.create(null);

  // First, include the command ids from the KeybindingService index verbatim.
  // Many ribbon buttons use built-in command ids directly.
  for (const [commandId, bindings] of commandKeybindingDisplayIndex.entries()) {
    const shortcut = bindings?.[0];
    if (!shortcut) continue;
    shortcutById[commandId] = shortcut;
  }

  for (const [ribbonId, keybindingCommandId] of Object.entries(KEYBINDING_COMMAND_BY_RIBBON_ID)) {
    const shortcut = getPrimaryCommandKeybindingDisplay(keybindingCommandId, commandKeybindingDisplayIndex);
    if (!shortcut) continue;
    shortcutById[ribbonId] = shortcut;
  }

  return shortcutById;
}
