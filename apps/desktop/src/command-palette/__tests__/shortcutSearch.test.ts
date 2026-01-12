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
  });
});
