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

describe("command palette performance safeguards", () => {
  const originalGlobalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const originalWindowLocalStorage = Object.getOwnPropertyDescriptor(window, "localStorage");

  beforeEach(() => {
    document.body.innerHTML = "";
    // Node 25 ships an experimental `globalThis.localStorage` accessor that throws unless started
    // with `--localstorage-file`. Provide a stable in-memory implementation for tests.
    const storage = createInMemoryLocalStorage();
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

  it("caps rendered results to the configured max", () => {
    vi.useFakeTimers();

    const registry = new CommandRegistry();
    for (let i = 0; i < 250; i += 1) {
      registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
    }

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      maxResults: 40,
      // Disable per-group limiting so this test exercises the global cap.
      maxResultsPerGroup: 10_000,
      inputDebounceMs: 1,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    input.value = "Command";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    vi.advanceTimersByTime(1);

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;
    const items = list.querySelectorAll(".command-palette__item");
    expect(items.length).toBe(40);

    palette.dispose();
    vi.useRealTimers();
  });

  it("debounces input updates (fake timers)", () => {
    vi.useFakeTimers();

    const registry = new CommandRegistry();
    for (let i = 0; i < 100; i += 1) {
      registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
    }

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      maxResults: 20,
      maxResultsPerGroup: 10_000,
      inputDebounceMs: 70,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;
    expect(list.querySelector(".command-palette__empty")).toBeNull();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    input.value = "zzzz no matches";
    input.dispatchEvent(new Event("input", { bubbles: true }));

    // Debounced: list should not update immediately.
    expect(list.querySelector(".command-palette__empty")).toBeNull();

    vi.advanceTimersByTime(69);
    expect(list.querySelector(".command-palette__empty")).toBeNull();

    vi.advanceTimersByTime(1);
    expect(list.querySelector(".command-palette__empty")).not.toBeNull();

    palette.dispose();
    vi.useRealTimers();
  });

  it("guards scrollIntoView calls when it is missing (jsdom)", async () => {
    const registry = new CommandRegistry();
    for (let i = 0; i < 3; i += 1) {
      registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
    }

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      maxResults: 20,
      maxResultsPerGroup: 20,
      inputDebounceMs: 0,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;

    const selectedBefore = list.querySelector<HTMLElement>('.command-palette__item[aria-selected="true"]');
    expect(selectedBefore?.id).toBe("command-palette-option-0");

    // Ensure the test stays valid even if jsdom implements scrollIntoView in the future:
    // explicitly override the per-element property to a *non-function* value so a direct call
    // (or `scrollIntoView?.(...)`) would throw, and only the `typeof === "function"` guard keeps this safe.
    (selectedBefore as any).scrollIntoView = 1;
    const nextBefore = list.querySelector<HTMLElement>("#command-palette-option-1");
    (nextBefore as any).scrollIntoView = 1;

    // Flush the queued microtask that keeps the selected row in view after rendering.
    await Promise.resolve();

    // Arrow navigation triggers `updateSelection`, which should also guard `scrollIntoView`.
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));

    const selectedAfter = list.querySelector<HTMLElement>('.command-palette__item[aria-selected="true"]');
    expect(selectedAfter?.id).toBe("command-palette-option-1");

    palette.dispose();
  });

  it("scrolls the selected row into view when scrollIntoView is available", async () => {
    const registry = new CommandRegistry();
    for (let i = 0; i < 3; i += 1) {
      registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
    }

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      maxResults: 20,
      maxResultsPerGroup: 20,
      inputDebounceMs: 0,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;

    const selectedBefore = list.querySelector<HTMLElement>('.command-palette__item[aria-selected="true"]');
    expect(selectedBefore?.id).toBe("command-palette-option-0");

    const scrollFirst = vi.fn();
    (selectedBefore as any).scrollIntoView = scrollFirst;

    // Flush the queued microtask that keeps the selected row in view after rendering.
    await Promise.resolve();

    expect(scrollFirst).toHaveBeenCalledTimes(1);
    expect(scrollFirst).toHaveBeenLastCalledWith({ block: "nearest" });

    const second = list.querySelector<HTMLElement>("#command-palette-option-1");
    const scrollSecond = vi.fn();
    (second as any).scrollIntoView = scrollSecond;

    input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));

    expect(scrollSecond).toHaveBeenCalledTimes(1);
    expect(scrollSecond).toHaveBeenLastCalledWith({ block: "nearest" });

    palette.dispose();
  });

  it("shows a searching placeholder during chunked search until the scan completes", async () => {
    vi.useFakeTimers();

    const rafCallbacks: Array<() => void> = [];
    const prevRaf = (globalThis as any).requestAnimationFrame as ((cb: () => void) => any) | undefined;
    (globalThis as any).requestAnimationFrame = (cb: () => void) => {
      rafCallbacks.push(cb);
      return rafCallbacks.length;
    };

    try {
      const registry = new CommandRegistry();
      for (let i = 0; i < 5_001; i += 1) {
        registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
      }

      const palette = createCommandPalette({
        commandRegistry: registry,
        contextKeys: new ContextKeyService(),
        keybindingIndex: new Map(),
        ensureExtensionsLoaded: async () => {},
        onCloseFocus: () => {},
        maxResults: 20,
        maxResultsPerGroup: 20,
        // 0ms debounce so the test doesn't need to advance time.
        inputDebounceMs: 0,
        extensionLoadDelayMs: 60_000,
      });

      palette.open();

      const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
      input.value = "zzzz";
      input.dispatchEvent(new Event("input", { bubbles: true }));

      // Run the debounce timer to kick off chunked search.
      vi.runOnlyPendingTimers();

      const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;
      expect(list.querySelector(".command-palette__empty")?.textContent).toBe("Searching…");

      // Drive the chunked search loop to completion by manually firing rAF callbacks.
      for (let i = 0; i < 50; i += 1) {
        const cb = rafCallbacks.shift();
        if (!cb) break;
        cb();
        // Allow the async loop continuation to run and enqueue the next frame callback.
        await Promise.resolve();
      }

      expect(list.querySelector(".command-palette__empty")?.textContent).toBe("No matching commands");

      palette.dispose();
    } finally {
      (globalThis as any).requestAnimationFrame = prevRaf;
      vi.useRealTimers();
    }
  });
});
