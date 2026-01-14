// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";

import { ContextKeyService } from "../extensions/contextKeys.js";
import { installKeyboardContextKeys, KeyboardContextKeyIds } from "./installKeyboardContextKeys.js";

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => queueMicrotask(resolve));
}

class FakeSpreadsheetApp {
  private baseEditing = false;
  private formulaBarEditing = false;
  private formulaBarFormulaEditing = false;

  private readonly editStateListeners = new Set<(isEditing: boolean) => void>();
  private readonly formulaOverlayListeners = new Set<() => void>();

  isEditing(): boolean {
    return this.baseEditing || this.formulaBarEditing;
  }

  isFormulaBarEditing(): boolean {
    return this.formulaBarEditing;
  }

  isFormulaBarFormulaEditing(): boolean {
    return this.formulaBarFormulaEditing;
  }

  onEditStateChange(listener: (isEditing: boolean) => void): () => void {
    this.editStateListeners.add(listener);
    listener(this.isEditing());
    return () => this.editStateListeners.delete(listener);
  }

  onFormulaBarOverlayChange(listener: () => void): () => void {
    this.formulaOverlayListeners.add(listener);
    listener();
    return () => this.formulaOverlayListeners.delete(listener);
  }

  setBaseEditing(value: boolean): void {
    this.baseEditing = value;
    for (const listener of [...this.editStateListeners]) listener(this.isEditing());
  }

  setFormulaBarEditing(value: boolean): void {
    this.formulaBarEditing = value;
    for (const listener of [...this.editStateListeners]) listener(this.isEditing());
  }

  setFormulaBarFormulaEditing(value: boolean): void {
    this.formulaBarFormulaEditing = value;
    for (const listener of [...this.formulaOverlayListeners]) listener();
  }
}

describe("installKeyboardContextKeys", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("tracks focus and spreadsheet editing state", async () => {
    const gridRoot = document.createElement("div");
    gridRoot.tabIndex = 0;
    document.body.appendChild(gridRoot);

    const formulaBarRoot = document.createElement("div");
    const formulaInput = document.createElement("input");
    formulaBarRoot.appendChild(formulaInput);
    document.body.appendChild(formulaBarRoot);

    const sheetTabsRoot = document.createElement("div");
    const tabStrip = document.createElement("div");
    tabStrip.className = "sheet-tabs";
    const renameInput = document.createElement("input");
    tabStrip.appendChild(renameInput);
    sheetTabsRoot.appendChild(tabStrip);
    const sheetAddButton = document.createElement("button");
    sheetAddButton.textContent = "Add";
    sheetTabsRoot.appendChild(sheetAddButton);
    document.body.appendChild(sheetTabsRoot);

    const outsideButton = document.createElement("button");
    outsideButton.textContent = "Outside";
    document.body.appendChild(outsideButton);
    outsideButton.focus();

    const contextKeys = new ContextKeyService();
    const app = new FakeSpreadsheetApp();

    let commandPaletteOpen = false;
    let splitViewSecondaryEditing = false;

    const dispose = installKeyboardContextKeys({
      contextKeys,
      app,
      formulaBarRoot,
      sheetTabsRoot,
      gridRoot,
      isCommandPaletteOpen: () => commandPaletteOpen,
      isSplitViewSecondaryEditing: () => splitViewSecondaryEditing,
    });

    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInFormulaBar)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabRename)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInGrid)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarEditing)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarFormulaEditing)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.workbenchCommandPaletteOpen)).toBe(false);

    // Focus: grid root.
    gridRoot.focus();
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.focusInGrid)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(false);

    // Focus: formula bar.
    formulaInput.focus();
    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInFormulaBar)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabRename)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInGrid)).toBe(false);

    // Focus: sheet tab rename.
    renameInput.focus();
    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInFormulaBar)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabRename)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInGrid)).toBe(false);

    // Focus: sheet tabs root but outside the tab strip.
    sheetAddButton.focus();
    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInFormulaBar)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabRename)).toBe(false);

    // Focus: non-input element.
    outsideButton.focus();
    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.focusInTextInput)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInFormulaBar)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInSheetTabRename)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.focusInGrid)).toBe(false);

    // Spreadsheet edit state.
    app.setBaseEditing(true);
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(true);

    app.setBaseEditing(false);
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(false);

    // Formula bar edit state.
    app.setFormulaBarEditing(true);
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(true);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarEditing)).toBe(true);

    // Formula bar formula-editing state is driven by overlay changes while typing.
    app.setFormulaBarFormulaEditing(true);
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarFormulaEditing)).toBe(true);

    // Hook-driven state should be recomputed on demand (eg split view, palette).
    app.setFormulaBarEditing(false);
    app.setFormulaBarFormulaEditing(false);
    await flushMicrotasks();

    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarEditing)).toBe(false);
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetFormulaBarFormulaEditing)).toBe(false);

    commandPaletteOpen = true;
    dispose.recompute();
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.workbenchCommandPaletteOpen)).toBe(true);

    splitViewSecondaryEditing = true;
    dispose.recompute();
    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(true);

    dispose();
  });

  it("falls back to the desktop global editing flag when split-view secondary hook is unavailable", async () => {
    const gridRoot = document.createElement("div");
    gridRoot.tabIndex = 0;
    document.body.appendChild(gridRoot);

    const formulaBarRoot = document.createElement("div");
    document.body.appendChild(formulaBarRoot);

    const sheetTabsRoot = document.createElement("div");
    document.body.appendChild(sheetTabsRoot);

    const contextKeys = new ContextKeyService();
    const app = new FakeSpreadsheetApp();

    const dispose = installKeyboardContextKeys({
      contextKeys,
      app,
      formulaBarRoot,
      sheetTabsRoot,
      gridRoot,
      // Intentionally omit `isSplitViewSecondaryEditing`.
    });

    await flushMicrotasks();
    expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(false);

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      dispose.recompute();
      await flushMicrotasks();
      expect(contextKeys.get(KeyboardContextKeyIds.spreadsheetIsEditing)).toBe(true);
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
      dispose();
    }
  });
});
