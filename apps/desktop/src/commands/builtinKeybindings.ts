export type BuiltinKeybinding = {
  command: string;
  key: string;
  mac?: string | null;
  when?: string | null;
  /**
   * If true, the command may fire repeatedly while the user holds the key chord down.
   * Defaults to false to avoid accidental repeats for toggle commands.
   */
  allowRepeat?: boolean;
};

// Spreadsheet-affecting shortcuts should fail closed when the focus/edit context keys
// are missing during startup. Prefer explicit `== true/false` checks over `!foo`.
const WHEN_SPREADSHEET_READY = "spreadsheet.isEditing == false && focus.inTextInput == false";
const WHEN_SHEET_NAVIGATION =
  "focus.inSheetTabRename == false && (focus.inTextInput == false || spreadsheet.formulaBarFormulaEditing == true)";
const WHEN_COMMAND_PALETTE_CLOSED = "workbench.commandPaletteOpen == false";
// Dialog-style shortcuts (Find/Replace/Go To, comments panel) should not steal focus while
// the user is typing in a text input (notably the formula bar editor).
const WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT = "workbench.commandPaletteOpen == false && focus.inTextInput == false";
// Find/Replace/Go To should not run while the spreadsheet is editing. This avoids stealing
// focus during formula-bar range selection mode where focus can temporarily return to the grid
// even though an edit session is active.
const WHEN_FIND_REPLACE_GOTO = `${WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT} && spreadsheet.isEditing == false`;
const WHEN_EDIT_CELL = `${WHEN_SPREADSHEET_READY} && focus.inGrid == false`;
const WHEN_OPEN_CONTEXT_MENU = `${WHEN_SPREADSHEET_READY} && ${WHEN_COMMAND_PALETTE_CLOSED} && focus.inSheetTabs == false`;

/**
 * Built-in keybindings that power UI affordances (Command Palette + context menus)
 * without touching SpreadsheetApp's existing keyboard handling yet.
 *
 * These records intentionally mirror VS Code-style extension contributions so the
 * same indexing/lookup logic can be shared with extension-contributed keybindings.
 */
export const builtinKeybindings: BuiltinKeybinding[] = [
  {
    command: "workbench.newWorkbook",
    key: "ctrl+n",
    mac: "cmd+n",
    when: "focus.inTextInput == false",
  },
  // Some environments emit both Ctrl+Meta for a single chord (remote desktop / VM keyboard setups).
  {
    command: "workbench.newWorkbook",
    key: "ctrl+cmd+n",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.openWorkbook",
    key: "ctrl+o",
    mac: "cmd+o",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.openWorkbook",
    key: "ctrl+cmd+o",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.saveWorkbook",
    key: "ctrl+s",
    mac: "cmd+s",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.saveWorkbook",
    key: "ctrl+cmd+s",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.saveWorkbookAs",
    key: "ctrl+shift+s",
    mac: "cmd+shift+s",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.saveWorkbookAs",
    key: "ctrl+cmd+shift+s",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.print",
    key: "ctrl+p",
    mac: "cmd+p",
    when: "focus.inTextInput == false",
  },
  {
    // Some environments emit both Ctrl+Meta for a single chord (remote desktop / VM keyboard setups).
    command: "workbench.print",
    key: "ctrl+cmd+p",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.closeWorkbook",
    key: "ctrl+w",
    mac: "cmd+w",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.closeWorkbook",
    key: "ctrl+cmd+w",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.quit",
    key: "ctrl+q",
    mac: "cmd+q",
    when: "focus.inTextInput == false",
  },
  {
    command: "workbench.quit",
    key: "ctrl+cmd+q",
    when: "focus.inTextInput == false",
  },
  {
    command: "edit.undo",
    key: "ctrl+z",
    mac: "cmd+z",
    // Undo/redo are intentionally left unguarded: the command implementation is
    // responsible for routing to text undo vs workbook history.
    when: null,
  },
  {
    command: "edit.redo",
    key: "ctrl+y",
    mac: "cmd+shift+z",
    when: null,
  },
  {
    command: "edit.redo",
    key: "ctrl+shift+z",
    mac: "cmd+shift+z",
    when: null,
  },
  {
    command: "view.toggleShowFormulas",
    key: "ctrl+`",
    mac: "cmd+`",
    when: WHEN_SPREADSHEET_READY,
  },
  // Theme switching (safe defaults: avoid Excel shortcuts by using Ctrl/Cmd+Alt/Option+Shift+<key>).
  {
    command: "view.theme.light",
    key: "ctrl+alt+shift+l",
    mac: "cmd+option+shift+l",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    command: "view.theme.dark",
    key: "ctrl+alt+shift+d",
    mac: "cmd+option+shift+d",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    command: "view.theme.system",
    key: "ctrl+alt+shift+s",
    mac: "cmd+option+shift+s",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    command: "view.theme.highContrast",
    key: "ctrl+alt+shift+h",
    mac: "cmd+option+shift+h",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    command: "audit.togglePrecedents",
    key: "ctrl+[",
    mac: "cmd+[",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "audit.toggleDependents",
    key: "ctrl+]",
    mac: "cmd+]",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "workbench.showCommandPalette",
    key: "ctrl+shift+p",
    mac: "cmd+shift+p",
    when: WHEN_COMMAND_PALETTE_CLOSED,
  },
  {
    // Some keyboards (and remote desktop setups) can emit both ctrlKey+metaKey for the
    // command palette chord. Add an explicit binding so the palette remains reachable.
    command: "workbench.showCommandPalette",
    key: "ctrl+cmd+shift+p",
    when: WHEN_COMMAND_PALETTE_CLOSED,
  },
  {
    command: "edit.editCell",
    key: "f2",
    mac: "f2",
    // When focus is inside a grid, let the grid's own handler open the correct editor
    // (primary vs split-view secondary). Use the built-in keybinding only as a global
    // fallback when focus is elsewhere (ribbon/menus/etc).
    when: WHEN_EDIT_CELL,
  },
  {
    // Excel-style focus cycling between major UI regions (ribbon/formula bar/grid/etc).
    command: "workbench.focusNextRegion",
    key: "f6",
    mac: "f6",
    // Intentionally permissive: this shortcut should work even when focus is in a text input.
    when: null,
  },
  {
    command: "workbench.focusPrevRegion",
    key: "shift+f6",
    mac: "shift+f6",
    when: null,
  },
  {
    command: "view.togglePanel.aiChat",
    key: "ctrl+shift+a",
    // IMPORTANT: Cmd+I is reserved for toggling the AI sidebar (see instructions/ui.md).
    mac: "cmd+i",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    // Some keyboards (and remote desktop setups) can emit both ctrlKey+metaKey for
    // Cmd-based shortcuts. Add an explicit binding so the AI chat toggle remains reachable.
    command: "view.togglePanel.aiChat",
    key: "ctrl+cmd+i",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  {
    command: "ai.inlineEdit",
    key: "ctrl+k",
    mac: "cmd+k",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.pasteSpecial",
    key: "ctrl+shift+v",
    mac: "cmd+shift+v",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.pasteSpecial",
    key: "ctrl+cmd+shift+v",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.copy",
    key: "ctrl+c",
    mac: "cmd+c",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    // Some keyboards (and remote desktop setups) can emit both ctrlKey+metaKey for
    // common clipboard chords. Add explicit bindings so copy remains reliable.
    command: "clipboard.copy",
    key: "ctrl+cmd+c",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.cut",
    key: "ctrl+x",
    mac: "cmd+x",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.cut",
    key: "ctrl+cmd+x",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.paste",
    key: "ctrl+v",
    mac: "cmd+v",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "clipboard.paste",
    key: "ctrl+cmd+v",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.find",
    key: "ctrl+f",
    mac: "cmd+f",
    when: WHEN_FIND_REPLACE_GOTO,
  },
  {
    command: "edit.replace",
    key: "ctrl+h",
    // Cmd+H is reserved by macOS for "Hide". Use Cmd+Option+F like many native apps.
    mac: "cmd+option+f",
    when: WHEN_FIND_REPLACE_GOTO,
  },
  {
    command: "navigation.goTo",
    key: "ctrl+g",
    mac: "cmd+g",
    when: WHEN_FIND_REPLACE_GOTO,
  },
  {
    command: "edit.clearContents",
    key: "delete",
    // macOS keyboards use Backspace for the "delete backwards" key.
    mac: "backspace",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.fillDown",
    key: "ctrl+d",
    mac: "cmd+d",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.fillRight",
    key: "ctrl+r",
    mac: "cmd+r",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.selectCurrentRegion",
    key: "ctrl+shift+*",
    mac: "cmd+shift+*",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.selectCurrentRegion",
    key: "ctrl+shift+8",
    mac: "cmd+shift+8",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    // Dedicated numpad multiply key. Excel accepts Ctrl/Cmd+* there without Shift.
    command: "edit.selectCurrentRegion",
    key: "ctrl+*",
    mac: "cmd+*",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.insertDate",
    key: "ctrl+;",
    mac: "cmd+;",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.insertTime",
    key: "ctrl+shift+;",
    mac: "cmd+shift+;",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "edit.autoSum",
    key: "alt+=",
    mac: "option+=",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    // Some keyboard layouts require Shift to produce "=". SpreadsheetApp's legacy handler
    // matches via `KeyboardEvent.code === "Equal"` and does not require Shift to be absent.
    //
    // KeybindingService matches modifier sets exactly, so add an explicit Shift variant to
    // keep Excel-compatible AutoSum behavior across layouts.
    command: "edit.autoSum",
    key: "alt+shift+=",
    mac: "option+shift+=",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.toggleBold",
    key: "ctrl+b",
    mac: "cmd+b",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.toggleItalic",
    key: "ctrl+i",
    // IMPORTANT: Cmd+I is reserved for toggling the AI sidebar (see instructions/ui.md).
    // Italic formatting remains Ctrl+I to preserve the AI toggle on macOS.
    mac: "ctrl+i",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.toggleUnderline",
    key: "ctrl+u",
    mac: "cmd+u",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.toggleStrikethrough",
    key: "ctrl+5",
    // Excel uses Ctrl+5 for strikethrough. On macOS, prefer keeping the same chord
    // rather than consuming a Cmd-based shortcut.
    mac: "ctrl+5",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.numberFormat.currency",
    key: "ctrl+shift+$",
    mac: "cmd+shift+$",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.numberFormat.percent",
    key: "ctrl+shift+%",
    mac: "cmd+shift+%",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.numberFormat.date",
    key: "ctrl+shift+#",
    mac: "cmd+shift+#",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "format.openFormatCells",
    key: "ctrl+1",
    mac: "cmd+1",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "workbook.previousSheet",
    key: "ctrl+pageup",
    mac: "cmd+pageup",
    when: WHEN_SHEET_NAVIGATION,
    allowRepeat: true,
  },
  {
    // Some keyboards/remote setups can emit both ctrlKey+metaKey for a single chord.
    command: "workbook.previousSheet",
    key: "ctrl+cmd+pageup",
    when: WHEN_SHEET_NAVIGATION,
    allowRepeat: true,
  },
  {
    command: "workbook.nextSheet",
    key: "ctrl+pagedown",
    mac: "cmd+pagedown",
    when: WHEN_SHEET_NAVIGATION,
    allowRepeat: true,
  },
  {
    // Some keyboards/remote setups can emit both ctrlKey+metaKey for a single chord.
    command: "workbook.nextSheet",
    key: "ctrl+cmd+pagedown",
    when: WHEN_SHEET_NAVIGATION,
    allowRepeat: true,
  },
  {
    command: "comments.addComment",
    key: "shift+f2",
    mac: "shift+f2",
    when: WHEN_SPREADSHEET_READY,
  },
  {
    command: "ui.openContextMenu",
    key: "shift+f10",
    mac: "shift+f10",
    when: WHEN_OPEN_CONTEXT_MENU,
  },
  {
    command: "ui.openContextMenu",
    key: "contextmenu",
    mac: "contextmenu",
    when: WHEN_OPEN_CONTEXT_MENU,
  },
  {
    command: "comments.togglePanel",
    key: "ctrl+shift+m",
    mac: "cmd+shift+m",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
  // Some environments emit both Ctrl+Meta for a single chord (remote desktop / VM keyboard setups).
  {
    command: "comments.togglePanel",
    key: "ctrl+cmd+shift+m",
    when: WHEN_COMMAND_PALETTE_CLOSED_AND_NOT_IN_TEXT_INPUT,
  },
];
