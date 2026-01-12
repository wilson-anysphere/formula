import { describe, expect, test } from "vitest";

import { fuzzyMatchCommand } from "../fuzzy";
import { readCommandPaletteRecents, recordCommandPaletteRecent, type StorageLike } from "../recents";

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
});

describe("command-palette/recents", () => {
  test("serializes recents as a stable, de-duped MRU list", () => {
    const storage = new MemoryStorage();

    recordCommandPaletteRecent(storage, "a", { limit: 3 });
    recordCommandPaletteRecent(storage, "b", { limit: 3 });
    recordCommandPaletteRecent(storage, "a", { limit: 3 });

    expect(readCommandPaletteRecents(storage)).toEqual(["a", "b"]);
  });

  test("enforces the configured limit", () => {
    const storage = new MemoryStorage();

    recordCommandPaletteRecent(storage, "a", { limit: 2 });
    recordCommandPaletteRecent(storage, "b", { limit: 2 });
    recordCommandPaletteRecent(storage, "c", { limit: 2 });

    expect(readCommandPaletteRecents(storage)).toEqual(["c", "b"]);
  });

  test("treats invalid stored JSON as empty", () => {
    const storage = new MemoryStorage();
    storage.setItem("formula.commandPalette.recents", "{not json");

    expect(readCommandPaletteRecents(storage)).toEqual([]);
  });
});

