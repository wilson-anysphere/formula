import { describe, expect, test } from "vitest";

import { searchShortcutCommands } from "../shortcutSearch";

describe("command-palette shortcut search", () => {
  test("filters to commands with shortcuts and sorts by category, shortcut, then title", () => {
    const commands = [
      { commandId: "cmd.one", title: "Zebra", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "cmd.two", title: "Alpha", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "cmd.three", title: "Beta", category: "CatB", source: { kind: "builtin" as const } },
      { commandId: "cmd.four", title: "Gamma", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "cmd.noShortcut", title: "No shortcut", category: "CatA", source: { kind: "builtin" as const } },
    ];

    const keybindingIndex = new Map<string, readonly string[]>([
      ["cmd.one", ["ctrl+shift+c"]],
      ["cmd.two", ["ctrl+shift+b"]],
      ["cmd.three", ["ctrl+shift+a"]],
      ["cmd.four", ["ctrl+shift+b"]],
    ]);

    const result = searchShortcutCommands({ commands, keybindingIndex, query: "" });
    expect(result.map((cmd) => cmd.commandId)).toEqual(["cmd.two", "cmd.four", "cmd.one", "cmd.three"]);
  });

  test("matches against title, id, and shortcut display text", () => {
    const commands = [
      { commandId: "sample.hello", title: "Show Greeting", category: "Sample", source: { kind: "builtin" as const } },
      { commandId: "sample.sum", title: "Sum Selection", category: "Sample", source: { kind: "builtin" as const } },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([
      ["sample.hello", ["ctrl+shift+h"]],
      ["sample.sum", ["ctrl+shift+y"]],
    ]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "greeting" }).map((c) => c.commandId)).toEqual(["sample.hello"]);
    expect(searchShortcutCommands({ commands, keybindingIndex, query: "sample.sum" }).map((c) => c.commandId)).toEqual(["sample.sum"]);
    expect(searchShortcutCommands({ commands, keybindingIndex, query: "shift+y" }).map((c) => c.commandId)).toEqual(["sample.sum"]);
  });

  test("matches ascii modifier queries against mac symbol shortcuts (cmd+shift+p vs ⇧⌘P)", () => {
    const commands = [{ commandId: "workbench.showCommandPalette", title: "Show Command Palette", category: "Navigation", source: { kind: "builtin" as const } }];
    const keybindingIndex = new Map<string, readonly string[]>([["workbench.showCommandPalette", ["⇧⌘P"]]]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "cmd+shift+p" }).map((c) => c.commandId)).toEqual([
      "workbench.showCommandPalette",
    ]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "command+shift+p" }).map((c) => c.commandId)).toEqual([
      "workbench.showCommandPalette",
    ]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "meta shift p" }).map((c) => c.commandId)).toEqual([
      "workbench.showCommandPalette",
    ]);
  });

  test("matches pageup/pagedown tokens against mac symbol shortcuts (cmd+pgup vs ⌘⇞)", () => {
    const commands = [
      {
        commandId: "workbook.previousSheet",
        title: "Previous Sheet",
        category: "Navigation",
        source: { kind: "builtin" as const },
      },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([["workbook.previousSheet", ["⌘⇞"]]]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "cmd+pgup" }).map((c) => c.commandId)).toEqual([
      "workbook.previousSheet",
    ]);
    expect(searchShortcutCommands({ commands, keybindingIndex, query: "cmd+pageup" }).map((c) => c.commandId)).toEqual([
      "workbook.previousSheet",
    ]);
  });

  test("accepts ctl synonym in shortcut query (ctl+shift+p)", () => {
    const commands = [
      { commandId: "workbench.showCommandPalette", title: "Show Command Palette", category: "Navigation", source: { kind: "builtin" as const } },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([["workbench.showCommandPalette", ["Ctrl+Shift+P"]]]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "ctl+shift+p" }).map((c) => c.commandId)).toEqual([
      "workbench.showCommandPalette",
    ]);
  });

  test("matches against secondary shortcuts in the keybinding index", () => {
    const commands = [
      { commandId: "edit.redo", title: "Redo", category: "Edit", source: { kind: "builtin" as const } },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([
      ["edit.redo", ["Ctrl+Y", "Ctrl+Shift+Z"]],
    ]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "ctrl+shift+z" }).map((c) => c.commandId)).toEqual([
      "edit.redo",
    ]);
  });

  test("does not drop punctuation tokens (cmd+[ should not match all cmd shortcuts)", () => {
    const commands = [
      { commandId: "audit", title: "Audit", category: "Cat", source: { kind: "builtin" as const } },
      { commandId: "other", title: "Other", category: "Cat", source: { kind: "builtin" as const } },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([
      ["audit", ["⌘["]],
      ["other", ["⌘P"]],
    ]);

    expect(searchShortcutCommands({ commands, keybindingIndex, query: "cmd+[" }).map((c) => c.commandId)).toEqual(["audit"]);
  });

  test("when limits are provided, fills results across categories (doesn't truncate to the first category)", () => {
    const commands = [
      { commandId: "catA.three", title: "Three", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catA.one", title: "One", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catA.two", title: "Two", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catB.four", title: "Four", category: "CatB", source: { kind: "builtin" as const } },
      { commandId: "catB.five", title: "Five", category: "CatB", source: { kind: "builtin" as const } },
    ];

    const keybindingIndex = new Map<string, readonly string[]>([
      ["catA.three", ["ctrl+c"]],
      ["catA.one", ["ctrl+a"]],
      ["catA.two", ["ctrl+b"]],
      ["catB.four", ["ctrl+d"]],
      ["catB.five", ["ctrl+e"]],
    ]);

    const result = searchShortcutCommands({
      commands,
      keybindingIndex,
      query: "ctrl",
      limits: { maxResults: 4, maxResultsPerCategory: 2 },
    });

    // 2 from CatA (ctrl+a, ctrl+b) + 2 from CatB (ctrl+d, ctrl+e).
    expect(result.map((c) => c.commandId)).toEqual(["catA.one", "catA.two", "catB.four", "catB.five"]);
  });

  test("supports result limiting without sorting huge match sets (empty query)", () => {
    const commands = [
      { commandId: "catA.one", title: "Alpha", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catA.two", title: "Beta", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catA.three", title: "Gamma", category: "CatA", source: { kind: "builtin" as const } },
      { commandId: "catB.one", title: "Delta", category: "CatB", source: { kind: "builtin" as const } },
      { commandId: "catB.two", title: "Epsilon", category: "CatB", source: { kind: "builtin" as const } },
    ];
    const keybindingIndex = new Map<string, readonly string[]>([
      ["catA.one", ["ctrl+a"]],
      ["catA.two", ["ctrl+b"]],
      ["catA.three", ["ctrl+c"]],
      ["catB.one", ["ctrl+a"]],
      ["catB.two", ["ctrl+b"]],
    ]);

    const result = searchShortcutCommands({
      commands,
      keybindingIndex,
      query: "",
      limits: { maxResults: 3, maxResultsPerCategory: 2 },
    });

    expect(result.map((cmd) => cmd.commandId)).toEqual(["catA.one", "catA.two", "catB.one"]);
  });
});
