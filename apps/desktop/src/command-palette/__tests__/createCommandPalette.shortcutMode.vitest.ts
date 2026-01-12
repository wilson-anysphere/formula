/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { ContextKeyService } from "../../extensions/contextKeys.js";
import { createCommandPalette } from "../createCommandPalette.js";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

describe("createCommandPalette shortcut search mode", () => {
  beforeEach(() => {
    document.body.innerHTML = "";

    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless
    // started with `--localstorage-file`. Provide a stable in-memory implementation for tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();

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

  it("enters shortcut mode when query starts with '/' (trimmed) and filters/sorts by shortcut", () => {
    vi.useFakeTimers();
    const registry = new CommandRegistry();

    // Two categories so we can assert category-first ordering.
    registry.registerBuiltinCommand("cmd.b", "Beta", () => {}, { category: "B" });
    registry.registerBuiltinCommand("cmd.a2", "Alpha 2", () => {}, { category: "A" });
    registry.registerBuiltinCommand("cmd.a1", "Alpha 1", () => {}, { category: "A" });
    // No shortcut -> should be excluded in shortcut mode.
    registry.registerBuiltinCommand("cmd.noShortcut", "No Shortcut", () => {}, { category: "A" });

    const keybindingIndex = new Map<string, readonly string[]>([
      ["cmd.b", ["ctrl+shift+c"]],
      // Two keybindings so we can verify shortcut search prefers displaying the matched one.
      ["cmd.a2", ["ctrl+shift+b", "ctrl+shift+z"]],
      ["cmd.a1", ["ctrl+shift+a"]],
    ]);

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex,
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      // Avoid debounce flakiness in this test.
      inputDebounceMs: 0,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;
    const hint = document.querySelector<HTMLElement>(".command-palette__hint")!;

    expect(hint.hidden).toBe(true);

    input.value = " /";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    // Debounced input uses setTimeout(0) when `inputDebounceMs` is 0.
    vi.advanceTimersByTime(0);

    expect(hint.hidden).toBe(false);

    const items = [...list.querySelectorAll(".command-palette__item")].map((el) => el.textContent ?? "");

    // Sorted by category (A then B) and then by shortcut (a before b) and title.
    // Note: each row includes both title and shortcut pill text; we just assert the ids are in order by title presence.
    expect(items.join("\n")).toMatch(/Alpha 1/);
    expect(items.join("\n")).toMatch(/Alpha 2/);
    expect(items.join("\n")).toMatch(/Beta/);
    expect(items.join("\n")).not.toMatch(/No Shortcut/);

    // A category should appear first; within A, ctrl+shift+a before ctrl+shift+b.
    const alpha1Index = items.findIndex((t) => t.includes("Alpha 1"));
    const alpha2Index = items.findIndex((t) => t.includes("Alpha 2"));
    const betaIndex = items.findIndex((t) => t.includes("Beta"));
    expect(alpha1Index).toBeGreaterThanOrEqual(0);
    expect(alpha2Index).toBeGreaterThanOrEqual(0);
    expect(betaIndex).toBeGreaterThanOrEqual(0);
    expect(alpha1Index).toBeLessThan(alpha2Index);
    expect(alpha2Index).toBeLessThan(betaIndex);

    // Shortcut mode should display the shortcut that matched the query when a command has multiple bindings.
    input.value = "/ ctrl+shift+z";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    vi.advanceTimersByTime(0);

    const alpha2Row = [...list.querySelectorAll<HTMLElement>("li.command-palette__item")].find((el) => (el.textContent ?? "").includes("Alpha 2"));
    expect(alpha2Row).toBeTruthy();
    expect(alpha2Row!.querySelector(".command-palette__shortcut")?.textContent).toBe("ctrl+shift+z");

    palette.dispose();
    vi.useRealTimers();
  });
});
