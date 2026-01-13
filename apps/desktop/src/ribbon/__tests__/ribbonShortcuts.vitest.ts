import { describe, expect, it } from "vitest";

import { deriveRibbonAriaKeyShortcutsById, deriveRibbonShortcutById } from "../ribbonShortcuts.js";

describe("ribbonShortcuts", () => {
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

