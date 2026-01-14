import test from "node:test";
import assert from "node:assert/strict";

import { CommandRegistry } from "../src/extensions/commandRegistry.js";
import { installCommandPaletteRecentsTracking } from "../src/command-palette/installCommandPaletteRecentsTracking.js";
import { readCommandRecents } from "../src/command-palette/recents.js";

class MemoryStorage {
  constructor() {
    /** @type {Map<string, string>} */
    this.data = new Map();
  }
  getItem(key) {
    return this.data.get(key) ?? null;
  }
  setItem(key, value) {
    this.data.set(key, value);
  }
}

test('command palette recents ignore commands hidden via when: "false"', async () => {
  const commandRegistry = new CommandRegistry();
  const storage = new MemoryStorage();

  commandRegistry.registerBuiltinCommand("cmd.visible", "Visible", () => "ok");
  commandRegistry.registerBuiltinCommand("cmd.hidden", "Hidden", () => "ok", { when: "false" });

  const dispose = installCommandPaletteRecentsTracking(commandRegistry, storage, { now: () => 1234 });

  try {
    await commandRegistry.executeCommand("cmd.hidden");
    await commandRegistry.executeCommand("cmd.visible");
  } finally {
    dispose();
  }

  assert.deepEqual(
    readCommandRecents(storage).map((entry) => entry.commandId),
    ["cmd.visible"],
    "Expected cmd.hidden (when:false) to be ignored so it does not crowd out visible command recents",
  );
});

test("command palette recents record canonical command ids when a hidden alias delegates via CommandRegistry", async () => {
  const commandRegistry = new CommandRegistry();
  const storage = new MemoryStorage();

  commandRegistry.registerBuiltinCommand("cmd.canonical", "Canonical", () => "ok");
  commandRegistry.registerBuiltinCommand("cmd.alias", "Alias", () => commandRegistry.executeCommand("cmd.canonical"), { when: "false" });

  const dispose = installCommandPaletteRecentsTracking(commandRegistry, storage, { now: () => 999 });

  try {
    await commandRegistry.executeCommand("cmd.alias");
  } finally {
    dispose();
  }

  assert.deepEqual(
    readCommandRecents(storage).map((entry) => entry.commandId),
    ["cmd.canonical"],
    "Expected cmd.alias (when:false) to be ignored while cmd.canonical is recorded via delegation",
  );
});

