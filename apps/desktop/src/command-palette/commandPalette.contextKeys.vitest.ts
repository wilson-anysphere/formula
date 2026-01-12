// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { KeybindingService } from "../extensions/keybindingService.js";
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

  it("marks the command palette overlay as a keybinding barrier (prevents global shortcuts while interacting with the palette)", async () => {
    const commandRegistry = new CommandRegistry();
    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.test", "Test", run);

    const contextKeys = new ContextKeyService();
    const keybindingService = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    keybindingService.setBuiltinKeybindings([{ command: "builtin.test", key: "f2" }]);

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys,
      keybindingIndex: new Map([["builtin.test", ["F2"]]]),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      extensionLoadDelayMs: 60_000,
    });

    controller.open();

    const overlay = document.querySelector<HTMLElement>(".command-palette-overlay");
    expect(overlay).toBeTruthy();
    expect(overlay?.dataset.keybindingBarrier).toBe("true");

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]');
    expect(list).toBeTruthy();

    const makeKeydownEvent = (target: EventTarget | null): KeyboardEvent => {
      const event: any = {
        key: "F2",
        code: "F2",
        ctrlKey: false,
        metaKey: false,
        shiftKey: false,
        altKey: false,
        repeat: false,
        target,
        defaultPrevented: false,
      };
      event.preventDefault = () => {
        event.defaultPrevented = true;
      };
      return event as KeyboardEvent;
    };

    // Inside the palette overlay, global keybindings should be ignored.
    const insideEvent = makeKeydownEvent(list);
    const insideHandled = await keybindingService.dispatchKeydown(insideEvent);
    expect(insideHandled).toBe(false);
    expect(insideEvent.defaultPrevented).toBe(false);
    expect(run).toHaveBeenCalledTimes(0);

    // Outside the palette overlay, the keybinding should fire.
    const outsideEvent = makeKeydownEvent(document.body);
    const outsideHandled = await keybindingService.dispatchKeydown(outsideEvent);
    expect(outsideHandled).toBe(true);
    expect(outsideEvent.defaultPrevented).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);

    controller.dispose();
  });
});
