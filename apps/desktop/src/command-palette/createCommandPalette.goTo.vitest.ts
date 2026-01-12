/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import type { GoToWorkbookLookup } from "../../../../packages/search/index.js";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";

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
      contextKeys: new ContextKeyService(),
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

  it("executes Go to on Enter and closes the palette", async () => {
    const commandRegistry = new CommandRegistry();

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: () => null,
    };
    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
    });

    palette.open();

    const overlay = document.querySelector<HTMLElement>(".command-palette-overlay");
    expect(overlay).toBeTruthy();
    expect(overlay!.hidden).toBe(false);

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();

    input!.value = "B3";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(onGoTo).toHaveBeenCalledTimes(1);
    expect(onGoTo).toHaveBeenCalledWith({
      type: "range",
      source: "a1",
      sheetName: "Sheet1",
      range: { startRow: 2, endRow: 2, startCol: 1, endCol: 1 },
    });
    expect(overlay!.hidden).toBe(true);

    palette.dispose();
  });

  it("prefers Go to over function results when query matches a named range", async () => {
    const commandRegistry = new CommandRegistry();

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: (name: string) => {
        if (name === "SUM") {
          return { sheetName: "Sheet1", range: { startRow: 2, endRow: 2, startCol: 1, endCol: 1 } };
        }
        return null;
      },
    };
    const onGoTo = vi.fn();
    const onSelectFunction = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
      onSelectFunction,
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();

    input!.value = "SUM";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const items = Array.from(document.querySelectorAll('[data-testid="command-palette-list"] .command-palette__item'));
    expect(items.length).toBeGreaterThanOrEqual(2);
    expect(items[0]?.textContent).toContain("Go to SUM");
    // Function results should still be present, but ranked below the Go to suggestion.
    expect(items[1]?.textContent).toContain("SUM");

    input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    expect(onGoTo).toHaveBeenCalledTimes(1);
    expect(onSelectFunction).not.toHaveBeenCalled();

    palette.dispose();
  });

  it("formats sheet-qualified refs using the explicit sheet name (not the current sheet)", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.someCommand", "Some Command", () => {}, { category: "Test" });

    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: () => null,
    };
    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet2", onGoTo },
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.value = "Sheet1!A1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const items = Array.from(document.querySelectorAll('[data-testid="command-palette-list"] .command-palette__item'));
    expect(items.length).toBeGreaterThanOrEqual(1);
    expect(items[0]?.textContent).toContain("Go to Sheet1!A1");
    expect(items[0]?.textContent).toContain("Sheet1!A1");
    expect(items[0]?.textContent).not.toContain("Sheet2!A1");

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
      contextKeys: new ContextKeyService(),
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
      contextKeys: new ContextKeyService(),
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

  it("hides Go to when the referenced sheet cannot be resolved (workbook.getSheet)", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("test.sheet2a1Mode", "Sheet2!A1 Mode", () => {}, { category: "Test" });

    // `parseGoTo` validates sheet-qualified references when the workbook adapter exposes
    // `getSheet(...)` (throwing for unknown sheets). The palette should therefore only
    // show a Go to suggestion for resolvable sheets.
    const workbook: GoToWorkbookLookup = {
      getTable: () => null,
      getName: () => null,
      getSheet: (name: string) => {
        if (name === "Sheet1") return { name: "Sheet1" };
        throw new Error(`Unknown sheet: ${name}`);
      },
    };

    const onGoTo = vi.fn();

    const palette = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      inputDebounceMs: 0,
      goTo: { workbook, getCurrentSheetName: () => "Sheet1", onGoTo },
    });

    palette.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.value = "Sheet2!A1";
    input!.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 0));

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]');
    expect(list).toBeTruthy();
    expect(list!.textContent).not.toContain("Go to Sheet2!A1");
    expect(list!.textContent).toContain("Sheet2!A1 Mode");

    palette.dispose();
  });
});
