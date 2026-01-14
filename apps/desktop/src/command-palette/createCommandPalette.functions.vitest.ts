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
  const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

  beforeEach(() => {
    document.body.innerHTML = "";
    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless started
    // with `--localstorage-file`. Provide a stable in-memory implementation for tests.
    const storage = createStorageMock();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    localStorage.clear();
  });

  afterEach(() => {
    if (originalGlobalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalGlobalLocalStorage);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).localStorage;
    }

    if (originalWindowLocalStorage) {
      Object.defineProperty(window, "localStorage", originalWindowLocalStorage);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (window as any).localStorage;
    }
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

  it("selects localized function names when the UI locale uses localized formulas (de-DE SUMME)", () => {
    vi.useFakeTimers();

    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
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
      input.value = "summe";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      vi.advanceTimersByTime(1);

      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

      expect(onSelectFunction).toHaveBeenCalledTimes(1);
      expect(onSelectFunction).toHaveBeenCalledWith("SUMME");

      palette.dispose();
    } finally {
      document.documentElement.lang = prevLang;
      vi.useRealTimers();
    }
  });
});
