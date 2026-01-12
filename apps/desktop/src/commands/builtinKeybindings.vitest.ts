import { describe, expect, it } from "vitest";

import { builtinKeybindings } from "./builtinKeybindings.js";
import { buildCommandKeybindingDisplayIndex, getPrimaryCommandKeybindingDisplay } from "../extensions/keybindings.js";

describe("builtin keybinding catalog", () => {
  it("formats expected display strings per platform", () => {
    const otherIndex = buildCommandKeybindingDisplayIndex({ platform: "other", contributed: [], builtin: builtinKeybindings });
    expect(getPrimaryCommandKeybindingDisplay("workbench.showCommandPalette", otherIndex)).toBe("Ctrl+Shift+P");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", otherIndex)).toBe("Ctrl+Shift+A");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.copy", otherIndex)).toBe("Ctrl+C");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.cut", otherIndex)).toBe("Ctrl+X");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.paste", otherIndex)).toBe("Ctrl+V");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", otherIndex)).toBe("Ctrl+Shift+V");

    const macIndex = buildCommandKeybindingDisplayIndex({ platform: "mac", contributed: [], builtin: builtinKeybindings });
    expect(getPrimaryCommandKeybindingDisplay("workbench.showCommandPalette", macIndex)).toBe("⇧⌘P");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", macIndex)).toBe("⇧⌘A");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.copy", macIndex)).toBe("⌘C");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.cut", macIndex)).toBe("⌘X");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.paste", macIndex)).toBe("⌘V");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", macIndex)).toBe("⇧⌘V");
  });
});
