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

describe("createCommandPalette context keys", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    vi.stubGlobal("localStorage", createStorageMock());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("sets workbench.commandPaletteOpen true on open and false on close/dispose", () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("cmd.test", "Test", () => {});

    const contextKeys = new ContextKeyService();

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      // Keep the extension-load timer from firing in this test (it is cleared on close/dispose anyway).
      extensionLoadDelayMs: 60_000,
    });

    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(false);
    expect(controller.isOpen()).toBe(false);

    controller.open();
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(true);
    expect(controller.isOpen()).toBe(true);

    // Repeated `open()` calls are idempotent and keep the context key authoritative.
    controller.open();
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(true);
    expect(controller.isOpen()).toBe(true);

    // Close via Escape (most common UX path).
    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    expect(input).toBeTruthy();
    input!.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "Escape",
        bubbles: true,
        cancelable: true,
      }),
    );
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(false);
    expect(controller.isOpen()).toBe(false);

    // Close via click outside (overlay background).
    controller.open();
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(true);
    const overlay = document.querySelector<HTMLDivElement>(".command-palette-overlay");
    expect(overlay).toBeTruthy();
    overlay!.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
      }),
    );
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(false);
    expect(controller.isOpen()).toBe(false);

    // Disposal while open should also clear the key.
    controller.open();
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(true);
    controller.dispose();
    expect(contextKeys.get("workbench.commandPaletteOpen")).toBe(false);
  });
});

