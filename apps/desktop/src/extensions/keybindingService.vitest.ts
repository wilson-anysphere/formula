import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "./commandRegistry.js";
import { ContextKeyService } from "./contextKeys.js";
import { KeybindingService } from "./keybindingService.js";

function makeKeydownEvent(opts: {
  key: string;
  code?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
  repeat?: boolean;
  target?: any;
}): KeyboardEvent {
  const event: any = {
    key: opts.key,
    code: opts.code ?? "",
    ctrlKey: Boolean(opts.ctrlKey),
    metaKey: Boolean(opts.metaKey),
    shiftKey: Boolean(opts.shiftKey),
    altKey: Boolean(opts.altKey),
    repeat: Boolean(opts.repeat),
    target: opts.target ?? null,
    defaultPrevented: false,
  };
  event.preventDefault = () => {
    event.defaultPrevented = true;
  };
  return event as KeyboardEvent;
}

describe("KeybindingService", () => {
  it("prefers built-in bindings over extension bindings for the same key", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.doThing", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.doThing", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.doThing", key: "ctrl+j" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.doThing", key: "ctrl+j", mac: null, when: null }]);

    await service.dispatchKeydown(makeKeydownEvent({ key: "j", ctrlKey: true }));

    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(extRun).not.toHaveBeenCalled();
  });

  it("filters by when-clause and falls back to lower-priority bindings", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.conditional", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.fallback", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.conditional", key: "ctrl+j", when: "hasSelection" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.fallback", key: "ctrl+j", mac: null, when: null }]);

    // when-clause false -> should skip builtin and run extension.
    contextKeys.set("hasSelection", false);
    await service.dispatchKeydown(makeKeydownEvent({ key: "j", ctrlKey: true }));
    expect(builtinRun).not.toHaveBeenCalled();
    expect(extRun).toHaveBeenCalledTimes(1);

    // when-clause true -> builtin should win.
    extRun.mockClear();
    contextKeys.set("hasSelection", true);
    await service.dispatchKeydown(makeKeydownEvent({ key: "j", ctrlKey: true }));
    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(extRun).not.toHaveBeenCalled();
  });

  it("never dispatches reserved shortcuts to extensions", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const extRun = vi.fn();
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.stealCopy", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setExtensionKeybindings([
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+c", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+cmd+c", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+shift+v", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+cmd+shift+v", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "ctrl+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "cmd+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "ctrl+cmd+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "ctrl+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "cmd+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "ctrl+cmd+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+shift+a", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "cmd+h", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+cmd+h", mac: null, when: null },
    ]);

    const event = makeKeydownEvent({ key: "c", ctrlKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event2 = makeKeydownEvent({ key: "v", ctrlKey: true, shiftKey: true });
    const handled2 = await service.dispatchKeydown(event2);
    expect(handled2).toBe(false);
    expect(event2.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event3 = makeKeydownEvent({ key: "c", ctrlKey: true, metaKey: true });
    const handled3 = await service.dispatchKeydown(event3);
    expect(handled3).toBe(false);
    expect(event3.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event4 = makeKeydownEvent({ key: "v", ctrlKey: true, metaKey: true, shiftKey: true });
    const handled4 = await service.dispatchKeydown(event4);
    expect(handled4).toBe(false);
    expect(event4.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // macOS system shortcut: Hide (Cmd+H) should never be claimable by extensions.
    const event5 = makeKeydownEvent({ key: "h", metaKey: true });
    const handled5 = await service.dispatchKeydown(event5);
    expect(handled5).toBe(false);
    expect(event5.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Some environments emit both Ctrl+Meta for a single chord.
    const event6 = makeKeydownEvent({ key: "h", ctrlKey: true, metaKey: true });
    const handled6 = await service.dispatchKeydown(event6);
    expect(handled6).toBe(false);
    expect(event6.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Quick Open (Tauri global): Ctrl/Cmd+Shift+O.
    const event7 = makeKeydownEvent({ key: "O", ctrlKey: true, shiftKey: true });
    const handled7 = await service.dispatchKeydown(event7);
    expect(handled7).toBe(false);
    expect(event7.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event8 = makeKeydownEvent({ key: "O", metaKey: true, shiftKey: true });
    const handled8 = await service.dispatchKeydown(event8);
    expect(handled8).toBe(false);
    expect(event8.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Inline AI edit: Ctrl/Cmd+K.
    const event9 = makeKeydownEvent({ key: "k", ctrlKey: true });
    const handled9 = await service.dispatchKeydown(event9);
    expect(handled9).toBe(false);
    expect(event9.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event10 = makeKeydownEvent({ key: "k", metaKey: true });
    const handled10 = await service.dispatchKeydown(event10);
    expect(handled10).toBe(false);
    expect(event10.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // AI Chat toggle: Ctrl+Shift+A / Cmd+I.
    const event11 = makeKeydownEvent({ key: "A", ctrlKey: true, shiftKey: true });
    const handled11 = await service.dispatchKeydown(event11);
    expect(handled11).toBe(false);
    expect(event11.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event12 = makeKeydownEvent({ key: "i", metaKey: true });
    const handled12 = await service.dispatchKeydown(event12);
    expect(handled12).toBe(false);
    expect(event12.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Some environments emit both Ctrl+Meta for a single chord.
    const event13 = makeKeydownEvent({ key: "i", ctrlKey: true, metaKey: true });
    const handled13 = await service.dispatchKeydown(event13);
    expect(handled13).toBe(false);
    expect(event13.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();
  });

  it("does not advertise reserved shortcuts in the command keybinding display index", () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();
    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });

    service.setExtensionKeybindings([
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+c", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+cmd+shift+v", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "ctrl+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "ctrl+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.allowed", key: "ctrl+j", mac: null, when: null },
    ]);

    const index = service.getCommandKeybindingDisplayIndex();
    expect(index.get("ext.allowed")).toEqual(["Ctrl+J"]);

    // Reserved bindings should not be surfaced as hints since they will never fire.
    expect(index.get("ext.stealCopy")).toBeUndefined();
    expect(index.get("ext.stealPasteSpecial")).toBeUndefined();
    expect(index.get("ext.stealQuickOpen")).toBeUndefined();
    expect(index.get("ext.stealInlineAI")).toBeUndefined();
  });

  it("matches shifted punctuation keybindings via KeyboardEvent.code fallback", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const extRun = vi.fn();
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.punct", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.punct", key: "ctrl+shift+;", mac: null, when: null }]);

    const event = makeKeydownEvent({ key: ":", code: "Semicolon", ctrlKey: true, shiftKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(extRun).toHaveBeenCalledTimes(1);
  });

  it("supports mac-specific builtins while still allowing the base keybinding as a fallback", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("cmd.replace", "Replace", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "mac" });
    service.setBuiltinKeybindings([{ command: "cmd.replace", key: "ctrl+h", mac: "cmd+option+f" }]);

    // mac-specific binding works.
    const handledMac = await service.dispatchKeydown(makeKeydownEvent({ key: "f", metaKey: true, altKey: true }));
    expect(handledMac).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);

    // Base binding still works on macOS as a fallback (Windows/Linux-style shortcut).
    const handledFallback = await service.dispatchKeydown(makeKeydownEvent({ key: "h", ctrlKey: true }));
    expect(handledFallback).toBe(true);
    expect(run).toHaveBeenCalledTimes(2);

    // System-reserved Cmd+H should not match either binding.
    const handledCmdH = await service.dispatchKeydown(makeKeydownEvent({ key: "h", metaKey: true }));
    expect(handledCmdH).toBe(false);
    expect(run).toHaveBeenCalledTimes(2);
  });

  it("does not dispatch on repeated keydown events", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.repeatTest", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.repeatTest", key: "ctrl+k" }]);

    const event = makeKeydownEvent({ key: "k", ctrlKey: true, repeat: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(builtinRun).not.toHaveBeenCalled();
  });

  it("ignores keybindings when the target is an input/textarea/contenteditable", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.inputGuard", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.inputGuard", key: "ctrl+k" }]);

    const event = makeKeydownEvent({
      key: "k",
      ctrlKey: true,
      target: { tagName: "INPUT", isContentEditable: false },
    });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(builtinRun).not.toHaveBeenCalled();
  });

  it("respects weight when multiple built-in bindings match the same chord", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const lowRun = vi.fn();
    const highRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.low", "Low", lowRun);
    commandRegistry.registerBuiltinCommand("builtin.high", "High", highRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([
      { command: "builtin.low", key: "ctrl+k", weight: 0 },
      { command: "builtin.high", key: "ctrl+k", weight: 10 },
    ]);

    const event = makeKeydownEvent({ key: "k", ctrlKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(highRun).toHaveBeenCalledTimes(1);
    expect(lowRun).not.toHaveBeenCalled();
  });

  it("accepts mac-specific builtin keybindings as alternates on other platforms (playwright compat)", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.replace", "Replace", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.replace", key: "ctrl+h", mac: "cmd+option+f" }]);

    const event = makeKeydownEvent({ key: "f", metaKey: true, altKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);
  });

  it("accepts non-mac builtin keybindings as fallbacks on macOS", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.replace", "Replace", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "mac" });
    service.setBuiltinKeybindings([{ command: "builtin.replace", key: "ctrl+h", mac: "cmd+option+f" }]);

    // Prefer Cmd+Option+F on macOS.
    const event = makeKeydownEvent({ key: "f", metaKey: true, altKey: true });
    const handled = await service.dispatchKeydown(event);
    expect(handled).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);

    // But still accept Ctrl+H as a fallback.
    const fallbackEvent = makeKeydownEvent({ key: "h", ctrlKey: true });
    const handledFallback = await service.dispatchKeydown(fallbackEvent);
    expect(handledFallback).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(2);
  });
});
