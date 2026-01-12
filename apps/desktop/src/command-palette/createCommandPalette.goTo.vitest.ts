/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import type { GoToWorkbookLookup } from "../../../../packages/search/index.js";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { createCommandPalette } from "./createCommandPalette.js";

describe("createCommandPalette Go to suggestion", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("inserts Go to above command matches when parseGoTo succeeds", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.b3Mode", "B3 Mode", () => {}, { category: "Test" });

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: () => null,
    };
    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      // `createCommandPalette` doesn't currently consult context keys, but keep the type contract.
      contextKeys: {} as any,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.value = "B3";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const items = Array.from(document.querySelectorAll('[data-testid="command-palette-list"] .command-palette__item'));
    expect(items.length).toBeGreaterThanOrEqual(2);
    expect(items[0]?.textContent).toContain("Go to B3");
    expect(items[0]?.textContent).toContain("Sheet1!B3");
    expect(items[1]?.textContent).toContain("B3 Mode");

    palette.dispose();
  });

  it("does not show Go to when parseGoTo fails", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.freezePanes", "Freeze Panes", () => {}, { category: "Test" });

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: () => null,
    };
    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: {} as any,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.value = "Freeze";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const items = Array.from(document.querySelectorAll('[data-testid="command-palette-list"] .command-palette__item'));
    expect(items.length).toBeGreaterThanOrEqual(1);
    expect(items[0]?.textContent).toContain("Freeze Panes");
    expect(items[0]?.textContent).not.toContain("Go to");

    palette.dispose();
  });

  it("supports named ranges (workbook.getName)", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.someCommand", "Some Command", () => {}, { category: "Test" });

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: (name: string) => {
        if (name === "MyRange") {
          return { sheetName: "Sheet1", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } };
        }
        return null;
      },
    };
    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: {} as any,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.value = "MyRange";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const items = Array.from(document.querySelectorAll('[data-testid="command-palette-list"] .command-palette__item'));
    expect(items.length).toBeGreaterThanOrEqual(1);
    expect(items[0]?.textContent).toContain("Go to MyRange");
    expect(items[0]?.textContent).toContain("Sheet1!A1:A2");

    palette.dispose();
  });
});
