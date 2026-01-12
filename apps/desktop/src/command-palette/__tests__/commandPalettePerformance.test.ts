/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createCommandPalette } from "../createCommandPalette.js";

describe("command palette performance safeguards", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    localStorage.clear();
  });

  it("caps rendered results to the configured max", () => {
    vi.useFakeTimers();

    const registry = new CommandRegistry();
    for (let i = 0; i < 250; i += 1) {
      registry.registerBuiltinCommand(`test.cmd.${i}`, `Command ${i}`, () => {}, { category: "Test" });
    }

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: {} as any,
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
      contextKeys: {} as any,
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
      contextKeys: {} as any,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      maxResults: 20,
      maxResultsPerGroup: 20,
      inputDebounceMs: 0,
      extensionLoadDelayMs: 60_000,
    });

    palette.open();

    // Flush the queued microtask that keeps the selected row in view after rendering.
    await Promise.resolve();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]')!;
    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]')!;

    const selectedBefore = list.querySelector<HTMLElement>('.command-palette__item[aria-selected="true"]');
    expect(selectedBefore?.id).toBe("command-palette-option-0");

    // Arrow navigation triggers `updateSelection`, which should also guard `scrollIntoView`.
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));

    const selectedAfter = list.querySelector<HTMLElement>('.command-palette__item[aria-selected="true"]');
    expect(selectedAfter?.id).toBe("command-palette-option-1");

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
        contextKeys: {} as any,
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
