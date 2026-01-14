import { getPrimaryCommandKeybindingAria, getPrimaryCommandKeybindingDisplay } from "../extensions/keybindings.js";

/**
 * Mapping from ribbon command ids (as defined in `ribbonSchema.ts`) to command ids
 * used by the KeybindingService / command palette.
 *
 * The ribbon uses a mix of canonical CommandRegistry ids (e.g. `clipboard.copy`)
 * and UI-specific ids (or aliases). This lookup normalizes
 * ribbon ids to the corresponding KeybindingService command id so ribbon
 * tooltips/menus can display keybinding hints.
 */
const KEYBINDING_COMMAND_BY_RIBBON_ID: Record<string, string> = {
  // --- File -------------------------------------------------------------------
  // File tab buttons are wired through ribbon-specific ids, but the keyboard shortcuts
  // are registered against the canonical workbench file commands.
  "file.new.new": "workbench.newWorkbook",
  "file.new.blankWorkbook": "workbench.newWorkbook",
  "file.open.open": "workbench.openWorkbook",
  "file.save.save": "workbench.saveWorkbook",
  "file.save.saveAs": "workbench.saveWorkbookAs",
  "file.print.print": "workbench.print",
  "file.options.close": "workbench.closeWorkbook",

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
  // Some ribbon controls use more specific commands (e.g. short/long date) but
  // share the classic Excel preset shortcut with a single canonical command.
  "format.numberFormat.accounting": "format.numberFormat.currency",
  "format.numberFormat.shortDate": "format.numberFormat.date",
  "format.numberFormat.longDate": "format.numberFormat.date",

  // Format Cells dialog entrypoints (Ctrl/Cmd+1).
  // Note: `format.openFormatCells` is used directly in the ribbon schema, so it
  // is picked up by the identity mapping. Keep these mappings for related
  // menu items (and legacy ids) so shortcut hints stay accurate.
  "home.number.formatCells": "format.openFormatCells",
  "home.number.moreFormats.formatCells": "format.openFormatCells",
  "home.alignment.orientation.formatCellAlignment": "format.openFormatCells",
  "format.openAlignmentDialog": "format.openFormatCells",

  // --- Comments ---------------------------------------------------------------
  "comments.addComment": "comments.addComment",
  "comments.togglePanel": "comments.togglePanel",

  // --- View ------------------------------------------------------------------
  "view.togglePanel.aiChat": "view.togglePanel.aiChat",
  "ai.inlineEdit": "ai.inlineEdit",

  // Theme selector ribbon ids (`view.appearance.*`) route to ribbon-specific commands today,
  // but the keyboard shortcuts are registered against the canonical `view.theme.*` commands.
  // Map them so the ribbon tooltip/menu can still display the correct shortcut hint.
  "view.appearance.theme.system": "view.theme.system",
  "view.appearance.theme.light": "view.theme.light",
  "view.appearance.theme.dark": "view.theme.dark",
  "view.appearance.theme.highContrast": "view.theme.highContrast",
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

export function deriveRibbonAriaKeyShortcutsById(commandKeybindingAriaIndex: Map<string, string[]>): Record<string, string> {
  const ariaKeyShortcutsById: Record<string, string> = Object.create(null);

  // First, include the command ids from the KeybindingService index verbatim.
  // Many ribbon buttons use built-in command ids directly.
  for (const [commandId, bindings] of commandKeybindingAriaIndex.entries()) {
    const aria = bindings?.[0];
    if (!aria) continue;
    ariaKeyShortcutsById[commandId] = aria;
  }

  for (const [ribbonId, keybindingCommandId] of Object.entries(KEYBINDING_COMMAND_BY_RIBBON_ID)) {
    const aria = getPrimaryCommandKeybindingAria(keybindingCommandId, commandKeybindingAriaIndex);
    if (!aria) continue;
    ariaKeyShortcutsById[ribbonId] = aria;
  }

  return ariaKeyShortcutsById;
}
