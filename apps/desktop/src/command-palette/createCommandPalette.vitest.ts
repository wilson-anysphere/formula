// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { createCommandPalette } from "./createCommandPalette.js";
import { COMMAND_RECENTS_STORAGE_KEY, LEGACY_COMMAND_RECENTS_STORAGE_KEY, readCommandRecents } from "./recents.js";

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

function dispatchKey(key: string, opts: { shiftKey?: boolean } = {}): void {
  const target = document.activeElement as HTMLElement | null;
  if (!target) throw new Error("Missing active element for key dispatch");
  target.dispatchEvent(
    new KeyboardEvent("keydown", {
      key,
      bubbles: true,
      cancelable: true,
      shiftKey: opts.shiftKey ?? false,
    }),
  );
}

describe("createCommandPalette focus management", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    vi.stubGlobal("localStorage", createStorageMock());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("traps focus within the palette and restores focus on Escape", () => {
    const outsideButton = document.createElement("button");
    outsideButton.textContent = "Outside";
    document.body.appendChild(outsideButton);
    outsideButton.focus();
    expect(document.activeElement).toBe(outsideButton);

    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("builtin.test", "Test", () => {});

    const onCloseFocus = vi.fn();
    const controller = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus,
      // Keep the extension-load timer from firing in this test (it is cleared on close anyway).
      extensionLoadDelayMs: 60_000,
    });

    controller.open();

    const input = document.querySelector<HTMLInputElement>('[data-testid="command-palette-input"]');
    const list = document.querySelector<HTMLUListElement>('[data-testid="command-palette-list"]');
    expect(input).toBeTruthy();
    expect(list).toBeTruthy();

    expect(document.activeElement).toBe(input);

    // Tab should cycle within the palette (input -> list -> input...).
    dispatchKey("Tab");
    expect(document.activeElement).toBe(list);

    dispatchKey("Tab");
    expect(document.activeElement).toBe(input);

    dispatchKey("Tab", { shiftKey: true });
    expect(document.activeElement).toBe(list);

    // Sanity: repeated tabs never escape to the outside button.
    for (let i = 0; i < 6; i += 1) {
      dispatchKey("Tab");
      expect([input, list]).toContain(document.activeElement);
    }

    // Escape closes the palette and restores focus to the element that was focused before opening.
    dispatchKey("Escape");
    expect(document.activeElement).toBe(outsideButton);
    expect(onCloseFocus).not.toHaveBeenCalled();

    controller.dispose();
  });
});

describe("createCommandPalette recents integration", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    vi.stubGlobal("localStorage", createStorageMock());
    vi.useFakeTimers();
    vi.setSystemTime(1234);
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
  });

  it("does not record ignored noisy commands, but does record normal commands", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("clipboard.copy", "Copy", () => {});
    commandRegistry.registerBuiltinCommand("clipboard.pasteSpecial", "Paste Special", () => {});
    commandRegistry.registerBuiltinCommand("edit.undo", "Undo", () => {});
    commandRegistry.registerBuiltinCommand("cmd.normal", "Normal", () => {});

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      extensionLoadDelayMs: 60_000,
    });

    await commandRegistry.executeCommand("clipboard.copy");
    await commandRegistry.executeCommand("clipboard.pasteSpecial");
    await commandRegistry.executeCommand("edit.undo");
    expect(readCommandRecents(localStorage)).toEqual([]);

    await commandRegistry.executeCommand("cmd.normal");
    expect(readCommandRecents(localStorage)).toEqual([{ commandId: "cmd.normal", lastUsedMs: 1234, count: 1 }]);

    controller.dispose();
  });

  it("does not record failed command executions", async () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("cmd.fail", "Fail", () => {
      throw new Error("boom");
    });

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      extensionLoadDelayMs: 60_000,
    });

    await expect(commandRegistry.executeCommand("cmd.fail")).rejects.toThrow("boom");
    expect(readCommandRecents(localStorage)).toEqual([]);

    controller.dispose();
  });

  it("migrates legacy recents key on install (and filters ignored entries)", () => {
    localStorage.setItem(
      LEGACY_COMMAND_RECENTS_STORAGE_KEY,
      JSON.stringify(["clipboard.copy", "clipboard.pasteSpecial", "edit.undo", "cmd.normal"]),
    );
    expect(localStorage.getItem(COMMAND_RECENTS_STORAGE_KEY)).toBeNull();

    const commandRegistry = new CommandRegistry();
    const controller = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      extensionLoadDelayMs: 60_000,
    });

    expect(readCommandRecents(localStorage)).toEqual([{ commandId: "cmd.normal", lastUsedMs: 1234, count: 1 }]);

    controller.dispose();
  });

  it("prunes ignored commands from existing recents on install", () => {
    localStorage.setItem(
      COMMAND_RECENTS_STORAGE_KEY,
      JSON.stringify([
        { commandId: "clipboard.copy", lastUsedMs: 2, count: 1 },
        { commandId: "edit.undo", lastUsedMs: 1, count: 1 },
        { commandId: "cmd.normal", lastUsedMs: 0, count: 1 },
      ]),
    );

    const commandRegistry = new CommandRegistry();
    const controller = createCommandPalette({
      commandRegistry,
      contextKeys: new ContextKeyService(),
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      extensionLoadDelayMs: 60_000,
    });

    expect(readCommandRecents(localStorage).map((entry) => entry.commandId)).toEqual(["cmd.normal"]);

    controller.dispose();
  });
});

describe("createCommandPalette when clauses", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    vi.stubGlobal("localStorage", createStorageMock());
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("filters commands whose when clause is not satisfied (e.g. permission gated)", () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("comments.addComment", "Add Comment", () => {}, {
      category: "Comments",
      when: "spreadsheet.canComment == true",
    });
    commandRegistry.registerBuiltinCommand("comments.togglePanel", "Toggle Comments", () => {}, { category: "Comments" });

    const contextKeys = new ContextKeyService();
    contextKeys.set("spreadsheet.canComment", false);

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      // Keep the extension-load timer from firing in this test (it is cleared on close/dispose anyway).
      extensionLoadDelayMs: 60_000,
    });

    controller.open();

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]');
    expect(list).toBeTruthy();
    expect(list!.textContent).toContain("Toggle Comments");
    expect(list!.textContent).not.toContain("Add Comment");

    controller.dispose();
  });

  it("includes commands whose when clause is satisfied", () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("comments.addComment", "Add Comment", () => {}, {
      category: "Comments",
      when: "spreadsheet.canComment == true",
    });
    commandRegistry.registerBuiltinCommand("comments.togglePanel", "Toggle Comments", () => {}, { category: "Comments" });

    const contextKeys = new ContextKeyService();
    contextKeys.set("spreadsheet.canComment", true);

    const controller = createCommandPalette({
      commandRegistry,
      contextKeys,
      keybindingIndex: new Map(),
      ensureExtensionsLoaded: async () => {},
      onCloseFocus: () => {},
      // Keep the extension-load timer from firing in this test (it is cleared on close/dispose anyway).
      extensionLoadDelayMs: 60_000,
    });

    controller.open();

    const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]');
    expect(list).toBeTruthy();
    expect(list!.textContent).toContain("Add Comment");
    expect(list!.textContent).toContain("Toggle Comments");

    controller.dispose();
  });

  it("refreshes command visibility when context keys change while open", () => {
    vi.useFakeTimers();
    try {
      const commandRegistry = new CommandRegistry();
      commandRegistry.registerBuiltinCommand("comments.addComment", "Add Comment", () => {}, {
        category: "Comments",
        when: "spreadsheet.canComment == true",
      });
      commandRegistry.registerBuiltinCommand("comments.togglePanel", "Toggle Comments", () => {}, { category: "Comments" });

      const contextKeys = new ContextKeyService();
      contextKeys.set("spreadsheet.canComment", false);

      const controller = createCommandPalette({
        commandRegistry,
        contextKeys,
        keybindingIndex: new Map(),
        ensureExtensionsLoaded: async () => {},
        onCloseFocus: () => {},
        // Keep the extension-load timer from firing in this test.
        extensionLoadDelayMs: 60_000,
      });

      controller.open();

      const list = document.querySelector<HTMLElement>('[data-testid="command-palette-list"]');
      expect(list).toBeTruthy();
      expect(list!.textContent).toContain("Toggle Comments");
      expect(list!.textContent).not.toContain("Add Comment");

      // Upgrade permissions while the palette is open.
      contextKeys.set("spreadsheet.canComment", true);
      // Allow the palette's debounced re-render to run.
      vi.advanceTimersByTime(100);

      expect(list!.textContent).toContain("Add Comment");

      controller.dispose();
    } finally {
      vi.useRealTimers();
    }
  });
});
