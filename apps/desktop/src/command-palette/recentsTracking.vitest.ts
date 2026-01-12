import { describe, expect, test } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry";
import { installCommandPaletteRecentsTracking } from "./installCommandPaletteRecentsTracking";
import { readCommandRecents, type StorageLike } from "./recents";

class MemoryStorage implements StorageLike {
  private readonly data = new Map<string, string>();

  getItem(key: string): string | null {
    return this.data.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.data.set(key, value);
  }
}

describe("command-palette/recents tracking", () => {
  test("records executed commands globally and filters noisy ones", async () => {
    const commandRegistry = new CommandRegistry();

    commandRegistry.registerBuiltinCommand("a", "A", () => {});
    commandRegistry.registerBuiltinCommand("b", "B", () => {});
    commandRegistry.registerBuiltinCommand("workbench.showCommandPalette", "Show Command Palette", () => {});
    commandRegistry.registerBuiltinCommand("clipboard.copy", "Copy", () => {});
    commandRegistry.registerBuiltinCommand("clipboard.cut", "Cut", () => {});
    commandRegistry.registerBuiltinCommand("clipboard.paste", "Paste", () => {});

    const storage = new MemoryStorage();
    let nowMs = 1;
    const dispose = installCommandPaletteRecentsTracking(commandRegistry, storage, { now: () => nowMs++ });

    await commandRegistry.executeCommand("a");
    await commandRegistry.executeCommand("clipboard.copy");
    await commandRegistry.executeCommand("b");
    await commandRegistry.executeCommand("a");
    await commandRegistry.executeCommand("workbench.showCommandPalette");
    await commandRegistry.executeCommand("clipboard.paste");

    expect(readCommandRecents(storage).map((entry) => entry.commandId)).toEqual(["a", "b"]);

    dispose();

    await commandRegistry.executeCommand("b");
    expect(readCommandRecents(storage).map((entry) => entry.commandId)).toEqual(["a", "b"]);
  });
});
