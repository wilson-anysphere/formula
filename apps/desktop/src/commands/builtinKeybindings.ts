export type BuiltinKeybinding = {
  command: string;
  key: string;
  mac?: string | null;
  when?: string | null;
};

/**
 * Built-in keybindings that power UI affordances (Command Palette + context menus)
 * without touching SpreadsheetApp's existing keyboard handling yet.
 *
 * These records intentionally mirror VS Code-style extension contributions so the
 * same indexing/lookup logic can be shared with extension-contributed keybindings.
 */
export const builtinKeybindings: BuiltinKeybinding[] = [
  {
    command: "workbench.showCommandPalette",
    key: "ctrl+shift+p",
    mac: "cmd+shift+p",
    when: null,
  },
  {
    command: "edit.editCell",
    key: "f2",
    mac: "f2",
    when: null,
  },
  {
    command: "view.togglePanel.aiChat",
    key: "ctrl+shift+a",
    mac: "cmd+shift+a",
    when: null,
  },
  {
    command: "clipboard.pasteSpecial",
    key: "ctrl+shift+v",
    mac: "cmd+shift+v",
    when: null,
  },
  {
    command: "clipboard.copy",
    key: "ctrl+c",
    mac: "cmd+c",
    when: null,
  },
  {
    command: "clipboard.cut",
    key: "ctrl+x",
    mac: "cmd+x",
    when: null,
  },
  {
    command: "clipboard.paste",
    key: "ctrl+v",
    mac: "cmd+v",
    when: null,
  },
  {
    command: "edit.find",
    key: "ctrl+f",
    mac: "cmd+f",
    when: null,
  },
  {
    command: "edit.replace",
    key: "ctrl+h",
    mac: "cmd+h",
    when: null,
  },
  {
    command: "navigation.goTo",
    key: "ctrl+g",
    mac: "cmd+g",
    when: null,
  },
  {
    command: "edit.clearContents",
    key: "delete",
    // macOS keyboards use Backspace for the "delete backwards" key.
    mac: "backspace",
    when: null,
  },
  {
    command: "edit.fillDown",
    key: "ctrl+d",
    mac: "cmd+d",
    when: null,
  },
  {
    command: "edit.fillRight",
    key: "ctrl+r",
    mac: "cmd+r",
    when: null,
  },
  {
    command: "edit.insertDate",
    key: "ctrl+;",
    mac: "cmd+;",
    when: null,
  },
  {
    command: "edit.insertTime",
    key: "ctrl+shift+;",
    mac: "cmd+shift+;",
    when: null,
  },
  {
    command: "edit.autoSum",
    key: "alt+=",
    mac: "option+=",
    when: null,
  },
  {
    command: "workbook.previousSheet",
    key: "ctrl+pageup",
    mac: "cmd+pageup",
    when: null,
  },
  {
    command: "workbook.nextSheet",
    key: "ctrl+pagedown",
    mac: "cmd+pagedown",
    when: null,
  },
  {
    command: "comments.addComment",
    key: "shift+f2",
    mac: "shift+f2",
    when: null,
  },
];
