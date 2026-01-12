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
    service.setBuiltinKeybindings([{ command: "builtin.doThing", key: "ctrl+k" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.doThing", key: "ctrl+k", mac: null, when: null }]);

    await service.dispatchKeydown(makeKeydownEvent({ key: "k", ctrlKey: true }));

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
    service.setBuiltinKeybindings([{ command: "builtin.conditional", key: "ctrl+k", when: "hasSelection" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.fallback", key: "ctrl+k", mac: null, when: null }]);

    // when-clause false -> should skip builtin and run extension.
    contextKeys.set("hasSelection", false);
    await service.dispatchKeydown(makeKeydownEvent({ key: "k", ctrlKey: true }));
    expect(builtinRun).not.toHaveBeenCalled();
    expect(extRun).toHaveBeenCalledTimes(1);

    // when-clause true -> builtin should win.
    extRun.mockClear();
    contextKeys.set("hasSelection", true);
    await service.dispatchKeydown(makeKeydownEvent({ key: "k", ctrlKey: true }));
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
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+shift+v", mac: null, when: null },
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
