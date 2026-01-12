/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createCommandPalette } from "../createCommandPalette.js";

describe("createCommandPalette shortcut search mode", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    // Node 22+ exposes an experimental global `localStorage` that can throw unless
    // started with `--localstorage-file`. Use the jsdom window storage explicitly.
    try {
      window.localStorage?.clear();
    } catch {
      // ignore
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
      ["cmd.a2", ["ctrl+shift+b"]],
      ["cmd.a1", ["ctrl+shift+a"]],
    ]);

    const palette = createCommandPalette({
      commandRegistry: registry,
      contextKeys: {} as any,
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

    palette.dispose();
    vi.useRealTimers();
  });
});
