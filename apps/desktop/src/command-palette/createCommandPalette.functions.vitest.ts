// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { createCommandPalette } from "./createCommandPalette.js";

function createStorageMock(): Storage {
  const map = new Map<string, string>();
  return {
    get length() {
      return map.size;
    },
    clear() {
      map.clear();
    },
    getItem(key: string) {
      return map.get(key) ?? null;
    },
    key(index: number) {
      return Array.from(map.keys())[index] ?? null;
    },
    removeItem(key: string) {
      map.delete(key);
    },
    setItem(key: string, value: string) {
      map.set(key, String(value));
    },
  };
}

describe("createCommandPalette function results", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    vi.stubGlobal("localStorage", createStorageMock());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("invokes onSelectFunction when selecting a function result", () => {
    vi.useFakeTimers();

    const registry = new CommandRegistry();
    // Add a similarly named command to ensure functions still win ranking.
    registry.registerBuiltinCommand("edit.autoSum", "AutoSum", () => {}, {
      category: "Editing",
      keywords: ["sum"],
    });

    const onSelectFunction = vi.fn();

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      onSelectFunction,
      inputDebounceMs: 1,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    input.value = "sum";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    vi.advanceTimersByTime(1);

    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(onSelectFunction).toHaveBeenCalledTimes(1);
    expect(onSelectFunction).toHaveBeenCalledWith("SUM");

    palette.dispose();
    vi.useRealTimers();
  });
});

