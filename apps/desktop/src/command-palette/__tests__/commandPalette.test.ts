import { describe, expect, test } from "vitest";

import { fuzzyMatchCommand } from "../fuzzy";
import {
  COMMAND_RECENTS_STORAGE_KEY,
  LEGACY_COMMAND_RECENTS_STORAGE_KEY,
  getRecentCommandIdsForDisplay,
  installCommandRecentsTracker,
  readCommandRecents,
  type StorageLike,
} from "../recents";
import { CommandRegistry } from "../../extensions/commandRegistry.js";

class MemoryStorage implements StorageLike {
  private readonly data = new Map<string, string>();

  getItem(key: string): string | null {
    return this.data.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.data.set(key, value);
  }
}

describe("command-palette/fuzzy", () => {
  test("supports abbreviations across words (pvt tbl → Insert Pivot Table)", () => {
    const match = fuzzyMatchCommand("pvt tbl", {
      commandId: "insertPivotTable",
      title: "Insert Pivot Table",
      category: "Insert",
    });

    expect(match).not.toBeNull();
    expect(match!.score).toBeGreaterThan(0);
    // Highlight ranges should exist (some part of the title matched).
    expect(match!.titleRanges.length).toBeGreaterThan(0);
  });

  test("prefers exact title matches (Freeze Panes > Unfreeze Panes)", () => {
    const freeze = fuzzyMatchCommand("Freeze Panes", {
      commandId: "freezePanes",
      title: "Freeze Panes",
      category: "View",
    })!;
    const unfreeze = fuzzyMatchCommand("Freeze Panes", {
      commandId: "unfreezePanes",
      title: "Unfreeze Panes",
      category: "View",
    })!;

    expect(freeze.score).toBeGreaterThan(unfreeze.score);
  });

  test("can match across fields (category + title)", () => {
    const match = fuzzyMatchCommand("view freeze", {
      commandId: "freezePanes",
      title: "Freeze Panes",
      category: "View",
    });
    expect(match).not.toBeNull();
  });

  test("matches on keywords when the title/category/id don't match", () => {
    const match = fuzzyMatchCommand("gizmo", {
      commandId: "ext.openPanel",
      title: "Open Sample Panel",
      category: "Extensions",
      keywords: ["gizmo", "widget"],
    });

    expect(match).not.toBeNull();
    expect(match!.score).toBeGreaterThan(0);
  });
});

describe("command-palette/recents", () => {
  test("executing a command via commandRegistry.executeCommand updates storage", async () => {
    const storage = new MemoryStorage();
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.command", "Test Command", () => "ok");

    installCommandRecentsTracker(commandRegistry, storage, { now: () => 1234 });

    await commandRegistry.executeCommand("test.command");

    expect(readCommandRecents(storage)).toEqual([{ commandId: "test.command", lastUsedMs: 1234, count: 1 }]);
  });

  test("clipboard commands can be ignored", async () => {
    const storage = new MemoryStorage();
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("clipboard.copy", "Copy", () => "ok");
    commandRegistry.registerBuiltinCommand("cmd.normal", "Normal", () => "ok");

    installCommandRecentsTracker(commandRegistry, storage, {
      now: () => 1234,
      ignoreCommandIds: ["clipboard.copy", "clipboard.cut", "clipboard.paste"],
    });

    await commandRegistry.executeCommand("clipboard.copy");
    expect(readCommandRecents(storage)).toEqual([]);

    await commandRegistry.executeCommand("cmd.normal");
    expect(readCommandRecents(storage)).toEqual([{ commandId: "cmd.normal", lastUsedMs: 1234, count: 1 }]);
  });

  test("ignore list matches command ids even when events contain whitespace", async () => {
    const storage = new MemoryStorage();
    const listeners: Array<(evt: any) => void> = [];
    const commandRegistry = {
      onDidExecuteCommand: (listener: (evt: any) => void) => {
        listeners.push(listener);
        return () => {};
      },
    };

    installCommandRecentsTracker(commandRegistry as any, storage, {
      now: () => 1234,
      ignoreCommandIds: ["clipboard.copy"],
    });

    // Simulate a malformed event payload with whitespace.
    listeners[0]!({ commandId: " clipboard.copy ", args: [], result: null });
    expect(readCommandRecents(storage)).toEqual([]);
  });

  test("failed commands are not recorded", async () => {
    const storage = new MemoryStorage();
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("cmd.fail", "Fail", () => {
      throw new Error("boom");
    });

    installCommandRecentsTracker(commandRegistry, storage, { now: () => 1234 });

    await expect(commandRegistry.executeCommand("cmd.fail")).rejects.toThrow("boom");
    expect(readCommandRecents(storage)).toEqual([]);
  });

  test("commands that throw non-Error values (e.g. undefined) are not recorded", async () => {
    const storage = new MemoryStorage();
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("cmd.fail", "Fail", () => {
      // eslint-disable-next-line no-throw-literal
      throw undefined;
    });

    installCommandRecentsTracker(commandRegistry, storage, { now: () => 1234 });

    await expect(commandRegistry.executeCommand("cmd.fail")).rejects.toBeUndefined();
    expect(readCommandRecents(storage)).toEqual([]);
  });

  test("install prunes ignored commands from existing storage", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      COMMAND_RECENTS_STORAGE_KEY,
      JSON.stringify([
        { commandId: "clipboard.copy", lastUsedMs: 2, count: 1 },
        { commandId: "cmd.normal", lastUsedMs: 1, count: 1 },
      ]),
    );
    const commandRegistry = new CommandRegistry();

    installCommandRecentsTracker(commandRegistry, storage, {
      ignoreCommandIds: ["clipboard.copy"],
      now: () => 1234,
    });

    expect(readCommandRecents(storage).map((e) => e.commandId)).toEqual(["cmd.normal"]);
  });

  test("migrates legacy storage key into the new schema (one-time)", () => {
    const storage = new MemoryStorage();
    storage.setItem(LEGACY_COMMAND_RECENTS_STORAGE_KEY, JSON.stringify(["cmd.a", "cmd.b"]));
    const commandRegistry = new CommandRegistry();

    installCommandRecentsTracker(commandRegistry, storage, { now: () => 999 });
    expect(readCommandRecents(storage)).toEqual([
      { commandId: "cmd.a", lastUsedMs: 999, count: 1 },
      { commandId: "cmd.b", lastUsedMs: 999, count: 1 },
    ]);

    // Idempotent: installing again should not overwrite the migrated timestamps.
    installCommandRecentsTracker(commandRegistry, storage, { now: () => 1000 });
    expect(readCommandRecents(storage)).toEqual([
      { commandId: "cmd.a", lastUsedMs: 999, count: 1 },
      { commandId: "cmd.b", lastUsedMs: 999, count: 1 },
    ]);
  });

  test("migration tolerates legacy entries stored as objects", () => {
    const storage = new MemoryStorage();
    storage.setItem(LEGACY_COMMAND_RECENTS_STORAGE_KEY, JSON.stringify([{ commandId: "cmd.a" }, { commandId: "cmd.b" }]));
    const commandRegistry = new CommandRegistry();

    installCommandRecentsTracker(commandRegistry, storage, { now: () => 999 });
    expect(readCommandRecents(storage)).toEqual([
      { commandId: "cmd.a", lastUsedMs: 999, count: 1 },
      { commandId: "cmd.b", lastUsedMs: 999, count: 1 },
    ]);
  });

  test("removed commands are filtered out", () => {
    const storage = new MemoryStorage();
    storage.setItem(
      COMMAND_RECENTS_STORAGE_KEY,
      JSON.stringify([
        { commandId: "cmd.missing", lastUsedMs: 2, count: 1 },
        { commandId: "cmd.exists", lastUsedMs: 1, count: 3 },
      ]),
    );

    expect(getRecentCommandIdsForDisplay(storage, ["cmd.exists"], { limit: 10 })).toEqual(["cmd.exists"]);
  });

  test("cap size is enforced", async () => {
    const storage = new MemoryStorage();
    const commandRegistry = new CommandRegistry();
    for (const id of ["cmd.a", "cmd.b", "cmd.c", "cmd.d"]) {
      commandRegistry.registerBuiltinCommand(id, id, () => undefined);
    }

    let nowMs = 1000;
    installCommandRecentsTracker(commandRegistry, storage, {
      maxEntries: 3,
      now: () => {
        nowMs += 1;
        return nowMs;
      },
    });

    await commandRegistry.executeCommand("cmd.a");
    await commandRegistry.executeCommand("cmd.b");
    await commandRegistry.executeCommand("cmd.c");
    await commandRegistry.executeCommand("cmd.d");

    expect(readCommandRecents(storage).map((entry) => entry.commandId)).toEqual(["cmd.d", "cmd.c", "cmd.b"]);
  });

  test("ignores non-finite counts in stored JSON", () => {
    const storage = new MemoryStorage();
    // `1e309` parses to Infinity in JS. Ensure we don't persist/read non-finite counts.
    storage.setItem(
      COMMAND_RECENTS_STORAGE_KEY,
      '[{"commandId":"cmd.normal","lastUsedMs":1234,"count":1e309}]',
    );
    expect(readCommandRecents(storage)).toEqual([{ commandId: "cmd.normal", lastUsedMs: 1234 }]);
  });

  test("ordering is deterministic when timestamps tie", async () => {
    const storage = new MemoryStorage();
    storage.setItem(
      COMMAND_RECENTS_STORAGE_KEY,
      JSON.stringify([
        { commandId: "cmd.a", lastUsedMs: 1, count: 1 },
        { commandId: "cmd.b", lastUsedMs: 1, count: 1 },
      ]),
    );

    // Parse should preserve original insertion order when timestamps are identical.
    expect(readCommandRecents(storage).map((e) => e.commandId)).toEqual(["cmd.a", "cmd.b"]);

    // Recording with the same timestamp should deterministically keep the new command first.
    const commandRegistry = new CommandRegistry();
    installCommandRecentsTracker(commandRegistry, storage, { now: () => 1 });
    commandRegistry.registerBuiltinCommand("cmd.c", "C", () => "ok");
    await commandRegistry.executeCommand("cmd.c");
    expect(readCommandRecents(storage).map((e) => e.commandId)).toEqual(["cmd.c", "cmd.a", "cmd.b"]);
  });
});
