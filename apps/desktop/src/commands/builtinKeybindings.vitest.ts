import { describe, expect, it } from "vitest";

import { builtinKeybindings } from "./builtinKeybindings.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { evaluateWhenClause } from "../extensions/whenClause.js";
import { buildCommandKeybindingDisplayIndex, getPrimaryCommandKeybindingDisplay, parseKeybinding } from "../extensions/keybindings.js";

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

  it("does not bind Cmd+H on macOS (reserved for the system Hide shortcut)", () => {
    const offenders = builtinKeybindings
      .filter((kb) => kb.mac)
      .filter((kb) => {
        const parsed = parseKeybinding(kb.command, kb.mac ?? "", kb.when ?? null);
        if (!parsed) return false;
        return parsed.meta && !parsed.ctrl && !parsed.alt && !parsed.shift && parsed.key === "h";
      })
      .map((kb) => kb.command);

    expect(offenders).toEqual([]);
  });

  it("includes explicit AutoSum Shift variants for layouts where '=' requires Shift", () => {
    const autosum = builtinKeybindings.filter((kb) => kb.command === "edit.autoSum");
    expect(autosum.map((kb) => kb.key)).toEqual(expect.arrayContaining(["alt+=", "alt+shift+="]));
    expect(autosum.map((kb) => kb.mac)).toEqual(expect.arrayContaining(["option+=", "option+shift+="]));
  });

  it("gates representative keybindings via focus/edit context keys (capture-phase safe)", () => {
    // Fail-closed behavior is important while migrating to capture-phase routing: if context
    // keys haven't been initialized yet, spreadsheet-affecting shortcuts should not fire.
    const emptyKeys = new ContextKeyService();
    const emptyLookup = emptyKeys.asLookup();

    const copyWhen = builtinKeybindings.find((kb) => kb.command === "clipboard.copy" && kb.key === "ctrl+c")?.when;
    expect(typeof copyWhen).toBe("string");
    expect(evaluateWhenClause(copyWhen, emptyLookup)).toBe(false);

    const paletteWhen = builtinKeybindings.find((kb) => kb.command === "workbench.showCommandPalette" && kb.key === "ctrl+shift+p")
      ?.when;
    expect(typeof paletteWhen).toBe("string");
    expect(evaluateWhenClause(paletteWhen, emptyLookup)).toBe(false);

    const saveWhen = builtinKeybindings.find((kb) => kb.command === "workbench.saveWorkbook" && kb.key === "ctrl+s")?.when;
    expect(typeof saveWhen).toBe("string");
    expect(evaluateWhenClause(saveWhen, emptyLookup)).toBe(false);

    const sheetPrevWhen = builtinKeybindings.find((kb) => kb.command === "workbook.previousSheet" && kb.key === "ctrl+pageup")?.when;
    expect(typeof sheetPrevWhen).toBe("string");
    expect(evaluateWhenClause(sheetPrevWhen, emptyLookup)).toBe(false);

    const contextKeys = new ContextKeyService();
    const lookup = contextKeys.asLookup();

    // Clipboard operations should be disabled while editing or inside text inputs.
    contextKeys.batch({ "spreadsheet.isEditing": false, "focus.inTextInput": false });
    expect(evaluateWhenClause(copyWhen, lookup)).toBe(true);
    expect(evaluateWhenClause(saveWhen, lookup)).toBe(true);

    contextKeys.set("focus.inTextInput", true);
    expect(evaluateWhenClause(copyWhen, lookup)).toBe(false);
    expect(evaluateWhenClause(saveWhen, lookup)).toBe(false);

    contextKeys.batch({ "spreadsheet.isEditing": true, "focus.inTextInput": false });
    expect(evaluateWhenClause(copyWhen, lookup)).toBe(false);

    // Sheet navigation should allow the formula bar "formula editing" exception, but
    // remain blocked while renaming a sheet tab.
    contextKeys.batch({
      "focus.inSheetTabRename": true,
      "focus.inTextInput": false,
      "spreadsheet.formulaBarFormulaEditing": false,
    });
    expect(evaluateWhenClause(sheetPrevWhen, lookup)).toBe(false);

    // Normal grid navigation.
    contextKeys.batch({
      "focus.inSheetTabRename": false,
      "focus.inTextInput": false,
      "spreadsheet.formulaBarFormulaEditing": false,
    });
    expect(evaluateWhenClause(sheetPrevWhen, lookup)).toBe(true);

    // Allow sheet switching while editing a formula in the formula bar.
    contextKeys.batch({
      "focus.inSheetTabRename": false,
      "focus.inTextInput": true,
      "spreadsheet.formulaBarFormulaEditing": true,
    });
    expect(evaluateWhenClause(sheetPrevWhen, lookup)).toBe(true);

    // But not while editing other text inputs.
    contextKeys.batch({
      "focus.inSheetTabRename": false,
      "focus.inTextInput": true,
      "spreadsheet.formulaBarFormulaEditing": false,
    });
    expect(evaluateWhenClause(sheetPrevWhen, lookup)).toBe(false);

    // Command palette should be available even in inputs, but blocked when already open.
    contextKeys.set("workbench.commandPaletteOpen", true);
    expect(evaluateWhenClause(paletteWhen, lookup)).toBe(false);
    contextKeys.set("workbench.commandPaletteOpen", false);
    expect(evaluateWhenClause(paletteWhen, lookup)).toBe(true);
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
    expect(otherIndex.get("edit.redo")).toEqual(expect.arrayContaining(["Ctrl+Y", "Ctrl+Shift+Z"]));
    expect(getPrimaryCommandKeybindingDisplay("view.toggleShowFormulas", otherIndex)).toBe("Ctrl+`");
    expect(getPrimaryCommandKeybindingDisplay("audit.togglePrecedents", otherIndex)).toBe("Ctrl+[");
    expect(getPrimaryCommandKeybindingDisplay("audit.toggleDependents", otherIndex)).toBe("Ctrl+]");
    expect(getPrimaryCommandKeybindingDisplay("edit.replace", otherIndex)).toBe("Ctrl+H");
    expect(getPrimaryCommandKeybindingDisplay("edit.editCell", otherIndex)).toBe("F2");
    expect(getPrimaryCommandKeybindingDisplay("comments.addComment", otherIndex)).toBe("Shift+F2");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", otherIndex)).toBe("Ctrl+Shift+A");
    expect(getPrimaryCommandKeybindingDisplay("ai.inlineEdit", otherIndex)).toBe("Ctrl+K");
    expect(getPrimaryCommandKeybindingDisplay("comments.togglePanel", otherIndex)).toBe("Ctrl+Shift+M");
    expect(otherIndex.get("comments.togglePanel")).toEqual(expect.arrayContaining(["Ctrl+Shift+M", "Ctrl+Shift+Meta+M"]));
    expect(getPrimaryCommandKeybindingDisplay("format.toggleItalic", otherIndex)).toBe("Ctrl+I");
    expect(getPrimaryCommandKeybindingDisplay("edit.fillDown", otherIndex)).toBe("Ctrl+D");
    expect(getPrimaryCommandKeybindingDisplay("edit.fillRight", otherIndex)).toBe("Ctrl+R");
    expect(getPrimaryCommandKeybindingDisplay("edit.selectCurrentRegion", otherIndex)).toBe("Ctrl+Shift+*");
    expect(otherIndex.get("edit.selectCurrentRegion")).toEqual(
      expect.arrayContaining(["Ctrl+Shift+*", "Ctrl+Shift+8", "Ctrl+*"]),
    );
    expect(getPrimaryCommandKeybindingDisplay("edit.insertDate", otherIndex)).toBe("Ctrl+;");
    expect(getPrimaryCommandKeybindingDisplay("edit.insertTime", otherIndex)).toBe("Ctrl+Shift+;");
    expect(getPrimaryCommandKeybindingDisplay("edit.autoSum", otherIndex)).toBe("Alt+=");

    const macIndex = buildCommandKeybindingDisplayIndex({ platform: "mac", contributed: [], builtin: builtinKeybindings });
    expect(getPrimaryCommandKeybindingDisplay("workbench.showCommandPalette", macIndex)).toBe("⇧⌘P");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.copy", macIndex)).toBe("⌘C");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.cut", macIndex)).toBe("⌘X");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.paste", macIndex)).toBe("⌘V");
    expect(getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", macIndex)).toBe("⇧⌘V");
    expect(getPrimaryCommandKeybindingDisplay("edit.undo", macIndex)).toBe("⌘Z");
    expect(getPrimaryCommandKeybindingDisplay("edit.redo", macIndex)).toBe("⇧⌘Z");
    expect(macIndex.get("edit.redo")).toEqual(["⇧⌘Z"]);
    expect(getPrimaryCommandKeybindingDisplay("view.toggleShowFormulas", macIndex)).toBe("⌘`");
    expect(getPrimaryCommandKeybindingDisplay("audit.togglePrecedents", macIndex)).toBe("⌘[");
    expect(getPrimaryCommandKeybindingDisplay("audit.toggleDependents", macIndex)).toBe("⌘]");
    // macOS: avoid Cmd+H which is reserved for the system "Hide" shortcut.
    expect(getPrimaryCommandKeybindingDisplay("edit.replace", macIndex)).toBe("⌥⌘F");
    expect(getPrimaryCommandKeybindingDisplay("edit.editCell", macIndex)).toBe("F2");
    expect(getPrimaryCommandKeybindingDisplay("comments.addComment", macIndex)).toBe("⇧F2");
    expect(getPrimaryCommandKeybindingDisplay("view.togglePanel.aiChat", macIndex)).toBe("⌘I");
    expect(getPrimaryCommandKeybindingDisplay("ai.inlineEdit", macIndex)).toBe("⌘K");
    expect(getPrimaryCommandKeybindingDisplay("comments.togglePanel", macIndex)).toBe("⇧⌘M");
    expect(macIndex.get("comments.togglePanel")).toEqual(expect.arrayContaining(["⇧⌘M", "⌃⇧⌘M"]));
    expect(getPrimaryCommandKeybindingDisplay("format.toggleItalic", macIndex)).toBe("⌃I");
    expect(getPrimaryCommandKeybindingDisplay("edit.fillDown", macIndex)).toBe("⌘D");
    expect(getPrimaryCommandKeybindingDisplay("edit.fillRight", macIndex)).toBe("⌘R");
    expect(getPrimaryCommandKeybindingDisplay("edit.selectCurrentRegion", macIndex)).toBe("⇧⌘*");
    expect(macIndex.get("edit.selectCurrentRegion")).toEqual(expect.arrayContaining(["⇧⌘*", "⇧⌘8", "⌘*"]));
    expect(getPrimaryCommandKeybindingDisplay("edit.insertDate", macIndex)).toBe("⌘;");
    expect(getPrimaryCommandKeybindingDisplay("edit.insertTime", macIndex)).toBe("⇧⌘;");
    expect(getPrimaryCommandKeybindingDisplay("edit.autoSum", macIndex)).toBe("⌥=");
  });
});
