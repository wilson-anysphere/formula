import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "./commandRegistry.js";
import { ContextKeyService } from "./contextKeys.js";
import { KeybindingService } from "./keybindingService.js";
import { parseKeybinding } from "./keybindings.js";
import { builtinKeybindings } from "../commands/builtinKeybindings.js";

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
      { extensionId: "ext", command: "ext.stealEscape", key: "escape", mac: null, when: null },
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
      { extensionId: "ext", command: "ext.stealEditCell", key: "f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAddComment", key: "shift+f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+shift+a", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "cmd+h", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+cmd+h", mac: null, when: null },
      // File shortcuts (core UX): Ctrl/Cmd + N/O/S/W/Q.
      { extensionId: "ext", command: "ext.stealNew", key: "ctrl+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "cmd+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "cmd+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "cmd+q", mac: null, when: null },
    ]);

    // Safety net: even if an extension keybinding for a reserved shortcut slips through filtering,
    // it should never dispatch at runtime.
    (service as any).extensions.push({
      source: { kind: "extension", extensionId: "ext" },
      binding: parseKeybinding("ext.stealCopy", "f2", null)!,
      weight: 0,
      order: 9999,
    });
    (service as any).extensions.push({
      source: { kind: "extension", extensionId: "ext" },
      binding: parseKeybinding("ext.stealCopy", "shift+f2", null)!,
      weight: 0,
      order: 10000,
    });
    const escapeEvent = makeKeydownEvent({ key: "Escape" });
    const escapeHandled = await service.dispatchKeydown(escapeEvent);
    expect(escapeHandled).toBe(false);
    expect(escapeEvent.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

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

    // Edit Cell (Excel-style): F2.
    const event14 = makeKeydownEvent({ key: "F2" });
    const handled14 = await service.dispatchKeydown(event14);
    expect(handled14).toBe(false);
    expect(event14.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Add Comment (Excel-style): Shift+F2.
    const event15 = makeKeydownEvent({ key: "F2", shiftKey: true });
    const handled15 = await service.dispatchKeydown(event15);
    expect(handled15).toBe(false);
    expect(event15.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Toggle Comments Panel (core UX): Ctrl/Cmd+Shift+M.
    const event16 = makeKeydownEvent({ key: "M", ctrlKey: true, shiftKey: true });
    const handled16 = await service.dispatchKeydown(event16);
    expect(handled16).toBe(false);
    expect(event16.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event17 = makeKeydownEvent({ key: "M", metaKey: true, shiftKey: true });
    const handled17 = await service.dispatchKeydown(event17);
    expect(handled17).toBe(false);
    expect(event17.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Some environments emit both Ctrl+Meta for a single chord.
    const event18 = makeKeydownEvent({ key: "M", ctrlKey: true, metaKey: true, shiftKey: true });
    const handled18 = await service.dispatchKeydown(event18);
    expect(handled18).toBe(false);
    expect(event18.defaultPrevented).toBe(false);
    // File shortcuts: Ctrl/Cmd + N/O/S/W/Q.
    const fileEvent1 = makeKeydownEvent({ key: "n", ctrlKey: true });
    const fileHandled1 = await service.dispatchKeydown(fileEvent1);
    expect(fileHandled1).toBe(false);
    expect(fileEvent1.defaultPrevented).toBe(false);

    const fileEvent2 = makeKeydownEvent({ key: "n", metaKey: true });
    const fileHandled2 = await service.dispatchKeydown(fileEvent2);
    expect(fileHandled2).toBe(false);
    expect(fileEvent2.defaultPrevented).toBe(false);

    const fileEvent3 = makeKeydownEvent({ key: "o", ctrlKey: true });
    const fileHandled3 = await service.dispatchKeydown(fileEvent3);
    expect(fileHandled3).toBe(false);
    expect(fileEvent3.defaultPrevented).toBe(false);

    const fileEvent4 = makeKeydownEvent({ key: "o", metaKey: true });
    const fileHandled4 = await service.dispatchKeydown(fileEvent4);
    expect(fileHandled4).toBe(false);
    expect(fileEvent4.defaultPrevented).toBe(false);

    const fileEvent5 = makeKeydownEvent({ key: "s", ctrlKey: true });
    const fileHandled5 = await service.dispatchKeydown(fileEvent5);
    expect(fileHandled5).toBe(false);
    expect(fileEvent5.defaultPrevented).toBe(false);

    const fileEvent6 = makeKeydownEvent({ key: "s", metaKey: true });
    const fileHandled6 = await service.dispatchKeydown(fileEvent6);
    expect(fileHandled6).toBe(false);
    expect(fileEvent6.defaultPrevented).toBe(false);

    const fileEvent7 = makeKeydownEvent({ key: "S", ctrlKey: true, shiftKey: true });
    const fileHandled7 = await service.dispatchKeydown(fileEvent7);
    expect(fileHandled7).toBe(false);
    expect(fileEvent7.defaultPrevented).toBe(false);

    const fileEvent8 = makeKeydownEvent({ key: "S", metaKey: true, shiftKey: true });
    const fileHandled8 = await service.dispatchKeydown(fileEvent8);
    expect(fileHandled8).toBe(false);
    expect(fileEvent8.defaultPrevented).toBe(false);

    const fileEvent9 = makeKeydownEvent({ key: "w", ctrlKey: true });
    const fileHandled9 = await service.dispatchKeydown(fileEvent9);
    expect(fileHandled9).toBe(false);
    expect(fileEvent9.defaultPrevented).toBe(false);

    const fileEvent10 = makeKeydownEvent({ key: "w", metaKey: true });
    const fileHandled10 = await service.dispatchKeydown(fileEvent10);
    expect(fileHandled10).toBe(false);
    expect(fileEvent10.defaultPrevented).toBe(false);

    const fileEvent11 = makeKeydownEvent({ key: "q", ctrlKey: true });
    const fileHandled11 = await service.dispatchKeydown(fileEvent11);
    expect(fileHandled11).toBe(false);
    expect(fileEvent11.defaultPrevented).toBe(false);

    const fileEvent12 = makeKeydownEvent({ key: "q", metaKey: true });
    const fileHandled12 = await service.dispatchKeydown(fileEvent12);
    expect(fileHandled12).toBe(false);
    expect(fileEvent12.defaultPrevented).toBe(false);

    expect(extRun).not.toHaveBeenCalled();
  });

  it("does not advertise reserved shortcuts in the command keybinding display index", () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();
    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });

    service.setExtensionKeybindings([
      { extensionId: "ext", command: "ext.stealEscape", key: "escape", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+c", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+cmd+shift+v", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "ctrl+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "ctrl+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealEditCell", key: "f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAddComment", key: "shift+f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+shift+a", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "ctrl+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "cmd+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "cmd+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "cmd+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.allowed", key: "ctrl+j", mac: null, when: null },
    ]);

    const index = service.getCommandKeybindingDisplayIndex();
    expect(index.get("ext.allowed")).toEqual(["Ctrl+J"]);

    // Reserved bindings should not be surfaced as hints since they will never fire.
    expect(index.get("ext.stealCopy")).toBeUndefined();
    expect(index.get("ext.stealPasteSpecial")).toBeUndefined();
    expect(index.get("ext.stealQuickOpen")).toBeUndefined();
    expect(index.get("ext.stealInlineAI")).toBeUndefined();
    expect(index.get("ext.stealEditCell")).toBeUndefined();
    expect(index.get("ext.stealAddComment")).toBeUndefined();
    expect(index.get("ext.stealAIChat")).toBeUndefined();
    expect(index.get("ext.stealCommentsPanel")).toBeUndefined();
    expect(index.get("ext.stealEscape")).toBeUndefined();
    expect(index.get("ext.stealNew")).toBeUndefined();
    expect(index.get("ext.stealOpen")).toBeUndefined();
    expect(index.get("ext.stealSave")).toBeUndefined();
    expect(index.get("ext.stealSaveAs")).toBeUndefined();
    expect(index.get("ext.stealClose")).toBeUndefined();
    expect(index.get("ext.stealQuit")).toBeUndefined();
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

  it("matches Excel-style number-format shortcuts via KeyboardEvent.code fallback (Ctrl+Shift+$)", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.currency", "Currency", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.currency", key: "ctrl+shift+$" }]);

    // Keep the display string Excel-like (show "$", not "4").
    expect(service.getCommandKeybindingDisplayIndex().get("builtin.currency")).toEqual(["Ctrl+Shift+$"]);

    // Simulate a non-US layout where the physical Digit4 key does not produce "$" for the chord.
    const event = makeKeydownEvent({ key: "4", code: "Digit4", ctrlKey: true, shiftKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("matches AutoSum on layouts where '=' requires Shift (Alt+Shift+Equal)", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("edit.autoSum", "AutoSum", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings(builtinKeybindings.filter((kb) => kb.command === "edit.autoSum"));

    // Simulate a layout where pressing the physical Equal key requires Shift to reach "=".
    // KeybindingService requires Shift to match exactly, so we should still resolve to AutoSum.
    const event = makeKeydownEvent({ key: "+", code: "Equal", altKey: true, shiftKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);
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

  it("dispatches on repeated keydown events when a builtin binding opts into repeats", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.repeatTest", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.repeatTest", key: "ctrl+k", allowRepeat: true }]);

    const event = makeKeydownEvent({ key: "k", ctrlKey: true, repeat: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);
  });

  it("does not dispatch on repeated keydown events by default", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.repeatDefault", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.repeatDefault", key: "ctrl+k" }]);

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

    // Simulate focus being inside a text input via context keys (preferred over inspecting `event.target`).
    contextKeys.set("focus.inTextInput", true);
    const event = makeKeydownEvent({ key: "k", ctrlKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(builtinRun).not.toHaveBeenCalled();
  });

  it('dispatches builtins but not extensions in inputs when ignoreInputTargets is "extensions"', async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.inputAllowed", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.inputBlocked", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other", ignoreInputTargets: "extensions" });
    service.setBuiltinKeybindings([{ command: "builtin.inputAllowed", key: "ctrl+j", when: "builtinEnabled" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.inputBlocked", key: "ctrl+j", mac: null, when: null }]);

    // Simulate focus being in a text input via context keys.
    contextKeys.set("focus.inTextInput", true);

    // Builtins still dispatch from inputs when allowed by their when-clause.
    contextKeys.set("builtinEnabled", true);
    const event1 = makeKeydownEvent({ key: "j", ctrlKey: true });
    const handled1 = await service.dispatchKeydown(event1);
    expect(handled1).toBe(true);
    expect(event1.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(extRun).not.toHaveBeenCalled();

    // Extensions should never dispatch from inputs in this mode, even if the builtin does not match.
    contextKeys.set("builtinEnabled", false);
    const event2 = makeKeydownEvent({ key: "j", ctrlKey: true });
    const handled2 = await service.dispatchKeydown(event2);
    expect(handled2).toBe(false);
    expect(event2.defaultPrevented).toBe(false);
    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(extRun).not.toHaveBeenCalled();
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

  it("does not allow extensions to steal grid typing when builtins dispatch in capture phase", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const extRun = vi.fn();
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.stealTyping", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.stealTyping", key: "c", mac: null, when: null }]);

    const event = makeKeydownEvent({ key: "c" });

    // Capture-phase dispatch: builtins only. Extensions must *not* run here.
    const captureHandled = await service.dispatchKeydown(event, { allowBuiltins: true, allowExtensions: false });
    expect(captureHandled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Bubble-phase "grid" handler consumes the key (e.g. start typing to edit).
    event.preventDefault();

    // Bubble-phase dispatch: extensions only. Must respect SpreadsheetApp preventDefault.
    const bubbleHandled = await service.dispatchKeydown(event, { allowBuiltins: false, allowExtensions: true });
    expect(bubbleHandled).toBe(false);
    expect(extRun).not.toHaveBeenCalled();
  });

  it("dispatches builtins in capture phase without dispatching extensions", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.editCell", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.stealTyping", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.editCell", key: "f2" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.stealTyping", key: "c", mac: null, when: null }]);

    const event = makeKeydownEvent({ key: "F2" });

    const captureHandled = await service.dispatchKeydown(event, { allowBuiltins: true, allowExtensions: false });
    expect(captureHandled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(extRun).not.toHaveBeenCalled();

    // Bubble-phase extension dispatch should not fire because the built-in handler already
    // consumed the key with preventDefault().
    const bubbleHandled = await service.dispatchKeydown(event, { allowBuiltins: false, allowExtensions: true });
    expect(bubbleHandled).toBe(false);
    expect(extRun).not.toHaveBeenCalled();
  });
});
