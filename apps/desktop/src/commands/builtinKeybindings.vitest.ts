import { describe, expect, it } from "vitest";

import { builtinKeybindings } from "./builtinKeybindings.js";
import { buildCommandKeybindingDisplayIndex, getPrimaryCommandKeybindingDisplay } from "../extensions/keybindings.js";

describe("builtin keybinding catalog", () => {
  it("does not contain exact duplicate entries", () => {
    const seen = new Set<string>();
    const dups: string[] = [];
    for (const kb of builtinKeybindings) {
      const signature = `${kb.command}|${kb.key}|${kb.mac ?? ""}|${kb.when ?? ""}`;
      if (seen.has(signature)) dups.push(signature);
      seen.add(signature);
    }
    expect(dups).toEqual([]);
  });

  it("formats expected display strings per platform", () => {
    const otherIndex = buildCommandKeybindingDisplayIndex({ platform: "other", contributed: [], builtin: builtinKeybindings });
    expect(getPrimaryCommandKeybindingDisplay("workbench.showCommandPalette", otherIndex)).toBe("Ctrl+Shift+P");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.copy", otherIndex)).toBe("Ctrl+C");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.cut", otherIndex)).toBe("Ctrl+X");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.paste", otherIndex)).toBe("Ctrl+V");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", otherIndex)).toBe("Ctrl+Shift+V");
    expect(getPrimaryCommandKeybindingDisplay("edit.undo", otherIndex)).toBe("Ctrl+Z");
    expect(getPrimaryCommandKeybindingDisplay("edit.redo", otherIndex)).toBe("Ctrl+Y");
    expect(getPrimaryCommandKeybindingDisplay("edit.replace", otherIndex)).toBe("Ctrl+H");
    expect(getPrimaryCommandKeybindingDisplay("audit.tracePrecedents", otherIndex)).toBe("Ctrl+[");
    expect(getPrimaryCommandKeybindingDisplay("audit.traceDependents", otherIndex)).toBe("Ctrl+]");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", otherIndex)).toBe("Ctrl+Shift+A");
    expect(getPrimaryCommandKeybindingDisplay("ai.inlineEdit", otherIndex)).toBe("Ctrl+K");
    expect(getPrimaryCommandKeybindingDisplay("format.toggleItalic", otherIndex)).toBe("Ctrl+I");

    const macIndex = buildCommandKeybindingDisplayIndex({ platform: "mac", contributed: [], builtin: builtinKeybindings });
    expect(getPrimaryCommandKeybindingDisplay("workbench.showCommandPalette", macIndex)).toBe("⇧⌘P");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.copy", macIndex)).toBe("⌘C");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.cut", macIndex)).toBe("⌘X");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.paste", macIndex)).toBe("⌘V");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", macIndex)).toBe("⇧⌘V");
    expect(getPrimaryCommandKeybindingDisplay("edit.undo", macIndex)).toBe("⌘Z");
    expect(getPrimaryCommandKeybindingDisplay("edit.redo", macIndex)).toBe("⇧⌘Z");
    expect(getPrimaryCommandKeybindingDisplay("edit.replace", macIndex)).toBe("⌘H");
    expect(getPrimaryCommandKeybindingDisplay("audit.tracePrecedents", macIndex)).toBe("⌘[");
    expect(getPrimaryCommandKeybindingDisplay("audit.traceDependents", macIndex)).toBe("⌘]");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", macIndex)).toBe("⇧⌘A");
    expect(getPrimaryCommandKeybindingDisplay("ai.inlineEdit", macIndex)).toBe("⌘K");
    expect(getPrimaryCommandKeybindingDisplay("format.toggleItalic", macIndex)).toBe("⌃I");
  });
});
