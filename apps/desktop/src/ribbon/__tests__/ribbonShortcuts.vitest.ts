import { describe, expect, it } from "vitest";

import { deriveRibbonAriaKeyShortcutsById, deriveRibbonShortcutById } from "../ribbonShortcuts.js";

describe("ribbonShortcuts", () => {
  it("maps file.* ribbon ids to workbench.* keybindings", () => {
    const displayIndex = new Map<string, string[]>([
      ["workbench.newWorkbook", ["Ctrl+N"]],
      ["workbench.openWorkbook", ["Ctrl+O"]],
      ["workbench.saveWorkbook", ["Ctrl+S"]],
      ["workbench.saveWorkbookAs", ["Ctrl+Shift+S"]],
      ["workbench.print", ["Ctrl+P"]],
      ["workbench.closeWorkbook", ["Ctrl+W"]],
    ]);

    const shortcutById = deriveRibbonShortcutById(displayIndex);

    expect(shortcutById["file.new.new"]).toBe("Ctrl+N");
    expect(shortcutById["file.new.blankWorkbook"]).toBe("Ctrl+N");
    expect(shortcutById["file.open.open"]).toBe("Ctrl+O");
    expect(shortcutById["file.save.save"]).toBe("Ctrl+S");
    expect(shortcutById["file.save.saveAs"]).toBe("Ctrl+Shift+S");
    expect(shortcutById["file.save.saveAs.copy"]).toBe("Ctrl+Shift+S");
    expect(shortcutById["file.save.saveAs.download"]).toBe("Ctrl+Shift+S");
    expect(shortcutById["file.print.print"]).toBe("Ctrl+P");
    expect(shortcutById["file.options.close"]).toBe("Ctrl+W");
  });

  it("maps file.* ribbon ids to workbench.* aria-keyshortcuts", () => {
    const ariaIndex = new Map<string, string[]>([
      ["workbench.newWorkbook", ["Control+N"]],
      ["workbench.openWorkbook", ["Control+O"]],
      ["workbench.saveWorkbook", ["Control+S"]],
      ["workbench.saveWorkbookAs", ["Control+Shift+S"]],
      ["workbench.print", ["Control+P"]],
      ["workbench.closeWorkbook", ["Control+W"]],
    ]);

    const ariaById = deriveRibbonAriaKeyShortcutsById(ariaIndex);

    expect(ariaById["file.new.new"]).toBe("Control+N");
    expect(ariaById["file.new.blankWorkbook"]).toBe("Control+N");
    expect(ariaById["file.open.open"]).toBe("Control+O");
    expect(ariaById["file.save.save"]).toBe("Control+S");
    expect(ariaById["file.save.saveAs"]).toBe("Control+Shift+S");
    expect(ariaById["file.save.saveAs.copy"]).toBe("Control+Shift+S");
    expect(ariaById["file.save.saveAs.download"]).toBe("Control+Shift+S");
    expect(ariaById["file.print.print"]).toBe("Control+P");
    expect(ariaById["file.options.close"]).toBe("Control+W");
  });

  it("maps view.appearance.theme.* ribbon ids to view.theme.* keybindings", () => {
    const displayIndex = new Map<string, string[]>([
      ["view.theme.dark", ["Ctrl+Alt+Shift+D"]],
      ["view.theme.light", ["Ctrl+Alt+Shift+L"]],
      ["view.theme.system", ["Ctrl+Alt+Shift+S"]],
      ["view.theme.highContrast", ["Ctrl+Alt+Shift+H"]],
    ]);

    const shortcutById = deriveRibbonShortcutById(displayIndex);

    expect(shortcutById["view.appearance.theme.dark"]).toBe("Ctrl+Alt+Shift+D");
    expect(shortcutById["view.appearance.theme.light"]).toBe("Ctrl+Alt+Shift+L");
    expect(shortcutById["view.appearance.theme.system"]).toBe("Ctrl+Alt+Shift+S");
    expect(shortcutById["view.appearance.theme.highContrast"]).toBe("Ctrl+Alt+Shift+H");
  });

  it("maps view.appearance.theme.* ribbon ids to view.theme.* aria-keyshortcuts", () => {
    const ariaIndex = new Map<string, string[]>([
      ["view.theme.dark", ["Control+Alt+Shift+D"]],
      ["view.theme.light", ["Control+Alt+Shift+L"]],
      ["view.theme.system", ["Control+Alt+Shift+S"]],
      ["view.theme.highContrast", ["Control+Alt+Shift+H"]],
    ]);

    const ariaById = deriveRibbonAriaKeyShortcutsById(ariaIndex);

    expect(ariaById["view.appearance.theme.dark"]).toBe("Control+Alt+Shift+D");
    expect(ariaById["view.appearance.theme.light"]).toBe("Control+Alt+Shift+L");
    expect(ariaById["view.appearance.theme.system"]).toBe("Control+Alt+Shift+S");
    expect(ariaById["view.appearance.theme.highContrast"]).toBe("Control+Alt+Shift+H");
  });
});
