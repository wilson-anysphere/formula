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

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
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
      [
        { extensionId: "ext", command: "ext.stealCopy", title: "Ext" },
        { extensionId: "ext", command: "ext.stealPreviousSheet", title: "Ext" },
        { extensionId: "ext", command: "ext.stealNextSheet", title: "Ext" },
        { extensionId: "ext", command: "ext.stealContextMenu", title: "Ext" },
      ],
      async (commandId) => extRun(commandId),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setExtensionKeybindings([
      { extensionId: "ext", command: "ext.stealEscape", key: "escape", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealEnter", key: "enter", mac: null, when: null },
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
      { extensionId: "ext", command: "ext.stealCopy", key: "f6", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "shift+f6", mac: null, when: null },
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
      { extensionId: "ext", command: "ext.stealNew", key: "ctrl+cmd+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "cmd+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+cmd+shift+s", mac: null, when: null },
      // Print (core UX): Ctrl/Cmd+P.
      { extensionId: "ext", command: "ext.stealPrint", key: "ctrl+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPrint", key: "cmd+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPrint", key: "ctrl+cmd+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "cmd+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+cmd+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealContextMenu", key: "shift+f10", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealContextMenu", key: "contextmenu", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "ctrl+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "cmd+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "ctrl+cmd+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "ctrl+pagedown", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "cmd+pagedown", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "ctrl+cmd+pagedown", mac: null, when: null },
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
    (service as any).extensions.push({
      source: { kind: "extension", extensionId: "ext" },
      binding: parseKeybinding("ext.stealCopy", "enter", null)!,
      weight: 0,
      order: 10001,
    });
    (service as any).extensions.push({
      source: { kind: "extension", extensionId: "ext" },
      binding: parseKeybinding("ext.stealCopy", "f6", null)!,
      weight: 0,
      order: 10002,
    });
    (service as any).extensions.push({
      source: { kind: "extension", extensionId: "ext" },
      binding: parseKeybinding("ext.stealCopy", "shift+f6", null)!,
      weight: 0,
      order: 10003,
    });
    const escapeEvent = makeKeydownEvent({ key: "Escape" });
    const escapeHandled = await service.dispatchKeydown(escapeEvent);
    expect(escapeHandled).toBe(false);
    expect(escapeEvent.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const enterEvent = makeKeydownEvent({ key: "Enter" });
    const enterHandled = await service.dispatchKeydown(enterEvent);
    expect(enterHandled).toBe(false);
    expect(enterEvent.defaultPrevented).toBe(false);
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

    // Focus cycling between major regions (Excel-style): F6 / Shift+F6.
    const eventF6 = makeKeydownEvent({ key: "F6" });
    const handledF6 = await service.dispatchKeydown(eventF6);
    expect(handledF6).toBe(false);
    expect(eventF6.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const eventShiftF6 = makeKeydownEvent({ key: "F6", shiftKey: true });
    const handledShiftF6 = await service.dispatchKeydown(eventShiftF6);
    expect(handledShiftF6).toBe(false);
    expect(eventShiftF6.defaultPrevented).toBe(false);
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
    expect(extRun).not.toHaveBeenCalled();

    // Workbook sheet navigation (Excel-style): Ctrl/Cmd+PgUp/PgDn.
    const event19 = makeKeydownEvent({ key: "PageUp", ctrlKey: true });
    const handled19 = await service.dispatchKeydown(event19);
    expect(handled19).toBe(false);
    expect(event19.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event20 = makeKeydownEvent({ key: "PageUp", metaKey: true });
    const handled20 = await service.dispatchKeydown(event20);
    expect(handled20).toBe(false);
    expect(event20.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Some environments emit both Ctrl+Meta for a single chord.
    const event21 = makeKeydownEvent({ key: "PageUp", ctrlKey: true, metaKey: true });
    const handled21 = await service.dispatchKeydown(event21);
    expect(handled21).toBe(false);
    expect(event21.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event22 = makeKeydownEvent({ key: "PageDown", ctrlKey: true });
    const handled22 = await service.dispatchKeydown(event22);
    expect(handled22).toBe(false);
    expect(event22.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const event23 = makeKeydownEvent({ key: "PageDown", metaKey: true });
    const handled23 = await service.dispatchKeydown(event23);
    expect(handled23).toBe(false);
    expect(event23.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    // Some environments emit both Ctrl+Meta for a single chord.
    const event24 = makeKeydownEvent({ key: "PageDown", ctrlKey: true, metaKey: true });
    const handled24 = await service.dispatchKeydown(event24);
    expect(handled24).toBe(false);
    expect(event24.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();


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

    const fileEvent8b = makeKeydownEvent({ key: "p", ctrlKey: true });
    const fileHandled8b = await service.dispatchKeydown(fileEvent8b);
    expect(fileHandled8b).toBe(false);
    expect(fileEvent8b.defaultPrevented).toBe(false);

    const fileEvent8c = makeKeydownEvent({ key: "p", metaKey: true });
    const fileHandled8c = await service.dispatchKeydown(fileEvent8c);
    expect(fileHandled8c).toBe(false);
    expect(fileEvent8c.defaultPrevented).toBe(false);

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

    // Some environments emit both Ctrl+Meta for a single chord.
    const fileEvent13 = makeKeydownEvent({ key: "n", ctrlKey: true, metaKey: true });
    const fileHandled13 = await service.dispatchKeydown(fileEvent13);
    expect(fileHandled13).toBe(false);
    expect(fileEvent13.defaultPrevented).toBe(false);

    const fileEvent14 = makeKeydownEvent({ key: "o", ctrlKey: true, metaKey: true });
    const fileHandled14 = await service.dispatchKeydown(fileEvent14);
    expect(fileHandled14).toBe(false);
    expect(fileEvent14.defaultPrevented).toBe(false);

    const fileEvent15 = makeKeydownEvent({ key: "s", ctrlKey: true, metaKey: true });
    const fileHandled15 = await service.dispatchKeydown(fileEvent15);
    expect(fileHandled15).toBe(false);
    expect(fileEvent15.defaultPrevented).toBe(false);

    const fileEvent16 = makeKeydownEvent({ key: "S", ctrlKey: true, metaKey: true, shiftKey: true });
    const fileHandled16 = await service.dispatchKeydown(fileEvent16);
    expect(fileHandled16).toBe(false);
    expect(fileEvent16.defaultPrevented).toBe(false);

    const fileEvent16b = makeKeydownEvent({ key: "p", ctrlKey: true, metaKey: true });
    const fileHandled16b = await service.dispatchKeydown(fileEvent16b);
    expect(fileHandled16b).toBe(false);
    expect(fileEvent16b.defaultPrevented).toBe(false);

    const fileEvent17 = makeKeydownEvent({ key: "w", ctrlKey: true, metaKey: true });
    const fileHandled17 = await service.dispatchKeydown(fileEvent17);
    expect(fileHandled17).toBe(false);
    expect(fileEvent17.defaultPrevented).toBe(false);

    const fileEvent18 = makeKeydownEvent({ key: "q", ctrlKey: true, metaKey: true });
    const fileHandled18 = await service.dispatchKeydown(fileEvent18);
    expect(fileHandled18).toBe(false);
    expect(fileEvent18.defaultPrevented).toBe(false);

    expect(extRun).not.toHaveBeenCalled();

    // Open context menu: Shift+F10 / ContextMenu key.
    const contextMenuEvent1 = makeKeydownEvent({ key: "F10", shiftKey: true });
    const contextMenuHandled1 = await service.dispatchKeydown(contextMenuEvent1);
    expect(contextMenuHandled1).toBe(false);
    expect(contextMenuEvent1.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();

    const contextMenuEvent2 = makeKeydownEvent({ key: "ContextMenu", code: "ContextMenu" });
    const contextMenuHandled2 = await service.dispatchKeydown(contextMenuEvent2);
    expect(contextMenuHandled2).toBe(false);
    expect(contextMenuEvent2.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();
  });

  it("does not advertise reserved shortcuts in the command keybinding display index", () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();
    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });

    service.setExtensionKeybindings([
      { extensionId: "ext", command: "ext.stealEscape", key: "escape", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealEnter", key: "enter", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCopy", key: "ctrl+c", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPasteSpecial", key: "ctrl+cmd+shift+v", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuickOpen", key: "ctrl+shift+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealInlineAI", key: "ctrl+k", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealEditCell", key: "f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAddComment", key: "shift+f2", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "ctrl+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "cmd+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPreviousSheet", key: "ctrl+cmd+pageup", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "ctrl+pagedown", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "cmd+pagedown", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNextSheet", key: "ctrl+cmd+pagedown", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+shift+a", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealAIChat", key: "ctrl+cmd+i", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealCommentsPanel", key: "ctrl+cmd+shift+m", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "ctrl+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "cmd+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealNew", key: "ctrl+cmd+n", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealOpen", key: "ctrl+cmd+o", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSave", key: "ctrl+cmd+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "cmd+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealSaveAs", key: "ctrl+cmd+shift+s", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPrint", key: "ctrl+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPrint", key: "cmd+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealPrint", key: "ctrl+cmd+p", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealClose", key: "ctrl+cmd+w", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "cmd+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealQuit", key: "ctrl+cmd+q", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealContextMenu", key: "shift+f10", mac: null, when: null },
      { extensionId: "ext", command: "ext.stealContextMenu", key: "contextmenu", mac: null, when: null },
      { extensionId: "ext", command: "ext.allowed", key: "ctrl+j", mac: null, when: null },
    ]);

    const index = service.getCommandKeybindingDisplayIndex();
    const ariaIndex = service.getCommandKeybindingAriaIndex();
    expect(index.get("ext.allowed")).toEqual(["Ctrl+J"]);
    expect(ariaIndex.get("ext.allowed")).toEqual(["Control+J"]);

    // Reserved bindings should not be surfaced as hints since they will never fire.
    expect(index.get("ext.stealCopy")).toBeUndefined();
    expect(ariaIndex.get("ext.stealCopy")).toBeUndefined();
    expect(index.get("ext.stealPasteSpecial")).toBeUndefined();
    expect(ariaIndex.get("ext.stealPasteSpecial")).toBeUndefined();
    expect(index.get("ext.stealQuickOpen")).toBeUndefined();
    expect(ariaIndex.get("ext.stealQuickOpen")).toBeUndefined();
    expect(index.get("ext.stealInlineAI")).toBeUndefined();
    expect(ariaIndex.get("ext.stealInlineAI")).toBeUndefined();
    expect(index.get("ext.stealEditCell")).toBeUndefined();
    expect(ariaIndex.get("ext.stealEditCell")).toBeUndefined();
    expect(index.get("ext.stealAddComment")).toBeUndefined();
    expect(ariaIndex.get("ext.stealAddComment")).toBeUndefined();
    expect(index.get("ext.stealPreviousSheet")).toBeUndefined();
    expect(ariaIndex.get("ext.stealPreviousSheet")).toBeUndefined();
    expect(index.get("ext.stealNextSheet")).toBeUndefined();
    expect(ariaIndex.get("ext.stealNextSheet")).toBeUndefined();
    expect(index.get("ext.stealAIChat")).toBeUndefined();
    expect(ariaIndex.get("ext.stealAIChat")).toBeUndefined();
    expect(index.get("ext.stealCommentsPanel")).toBeUndefined();
    expect(ariaIndex.get("ext.stealCommentsPanel")).toBeUndefined();
    expect(index.get("ext.stealEscape")).toBeUndefined();
    expect(ariaIndex.get("ext.stealEscape")).toBeUndefined();
    expect(index.get("ext.stealEnter")).toBeUndefined();
    expect(ariaIndex.get("ext.stealEnter")).toBeUndefined();
    expect(index.get("ext.stealNew")).toBeUndefined();
    expect(ariaIndex.get("ext.stealNew")).toBeUndefined();
    expect(index.get("ext.stealOpen")).toBeUndefined();
    expect(ariaIndex.get("ext.stealOpen")).toBeUndefined();
    expect(index.get("ext.stealSave")).toBeUndefined();
    expect(ariaIndex.get("ext.stealSave")).toBeUndefined();
    expect(index.get("ext.stealSaveAs")).toBeUndefined();
    expect(ariaIndex.get("ext.stealSaveAs")).toBeUndefined();
    expect(index.get("ext.stealPrint")).toBeUndefined();
    expect(ariaIndex.get("ext.stealPrint")).toBeUndefined();
    expect(index.get("ext.stealClose")).toBeUndefined();
    expect(ariaIndex.get("ext.stealClose")).toBeUndefined();
    expect(index.get("ext.stealQuit")).toBeUndefined();
    expect(ariaIndex.get("ext.stealQuit")).toBeUndefined();
    expect(index.get("ext.stealContextMenu")).toBeUndefined();
    expect(ariaIndex.get("ext.stealContextMenu")).toBeUndefined();
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

  it.each([
    { symbol: "$", code: "Digit4", eventKey: "4", command: "builtin.currency", title: "Currency" },
    { symbol: "%", code: "Digit5", eventKey: "5", command: "builtin.percent", title: "Percent" },
    { symbol: "#", code: "Digit3", eventKey: "3", command: "builtin.date", title: "Date" },
  ])(
    "matches Excel-style number-format shortcuts via KeyboardEvent.code fallback (Ctrl+Shift+$symbol)",
    async ({ symbol, code, eventKey, command, title }) => {
      const contextKeys = new ContextKeyService();
      const commandRegistry = new CommandRegistry();

      const run = vi.fn();
      commandRegistry.registerBuiltinCommand(command, title, run);

      const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
      service.setBuiltinKeybindings([{ command, key: `ctrl+shift+${symbol}` }]);

      // Keep the display string Excel-like (show "$/%/#", not "4/5/3").
      expect(service.getCommandKeybindingDisplayIndex().get(command)).toEqual([`Ctrl+Shift+${symbol}`]);

      // Simulate a non-US layout where the physical Digit key does not produce the literal symbol for the chord.
      const event = makeKeydownEvent({ key: eventKey, code, ctrlKey: true, shiftKey: true });
      const handled = await service.dispatchKeydown(event);

      expect(handled).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(run).toHaveBeenCalledTimes(1);
    },
  );

  it.each([
    { symbol: "$", code: "Digit4", eventKey: "4", command: "builtin.currency", title: "Currency" },
    { symbol: "%", code: "Digit5", eventKey: "5", command: "builtin.percent", title: "Percent" },
    { symbol: "#", code: "Digit3", eventKey: "3", command: "builtin.date", title: "Date" },
  ])(
    "matches Excel-style number-format shortcuts via KeyboardEvent.code fallback on macOS (Cmd+Shift+$symbol)",
    async ({ symbol, code, eventKey, command, title }) => {
      const contextKeys = new ContextKeyService();
      const commandRegistry = new CommandRegistry();

      const run = vi.fn();
      commandRegistry.registerBuiltinCommand(command, title, run);

      const service = new KeybindingService({ commandRegistry, contextKeys, platform: "mac" });
      service.setBuiltinKeybindings([{ command, key: `ctrl+shift+${symbol}`, mac: `cmd+shift+${symbol}` }]);

      // Ensure the platform display string stays Excel-like (show "$/%/#", not "4/5/3").
      expect(service.getCommandKeybindingDisplayIndex().get(command)).toEqual([`⇧⌘${symbol}`]);

      // Simulate a non-US layout where the physical Digit key does not produce the literal symbol for the chord.
      const event = makeKeydownEvent({ key: eventKey, code, metaKey: true, shiftKey: true });
      const handled = await service.dispatchKeydown(event);

      expect(handled).toBe(true);
      expect(event.defaultPrevented).toBe(true);
      expect(run).toHaveBeenCalledTimes(1);
    },
  );

  it("matches AutoSum on layouts where '=' requires Shift (Alt+Shift+Equal)", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("edit.autoSum", "AutoSum", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings(builtinKeybindings.filter((kb) => kb.command === "edit.autoSum"));

    // Builtin spreadsheet shortcuts are gated by context keys and should fail closed.
    contextKeys.batch({ "spreadsheet.isEditing": false, "focus.inTextInput": false });

    // Simulate a layout where pressing the physical Equal key requires Shift to reach "=".
    // KeybindingService requires Shift to match exactly, so we should still resolve to AutoSum.
    const event = makeKeydownEvent({ key: "+", code: "Equal", altKey: true, shiftKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("matches AutoSum on macOS layouts where '=' requires Shift (Option+Shift+Equal)", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("edit.autoSum", "AutoSum", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "mac" });
    service.setBuiltinKeybindings(builtinKeybindings.filter((kb) => kb.command === "edit.autoSum"));

    // Builtin spreadsheet shortcuts are gated by context keys and should fail closed.
    contextKeys.batch({ "spreadsheet.isEditing": false, "focus.inTextInput": false });

    // On macOS, Option is reported as `altKey`. Simulate a layout where the physical Equal key
    // needs Shift to reach "=".
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

  it('blocks built-in keybindings in inputs when ignoreInputTargets="all" (default)', async () => {
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

  it('allows built-ins and extensions in inputs when ignoreInputTargets="none"', async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.doThing", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.doThing", title: "Ext" }],
      async () => extRun(),
    );

    const inputTarget = { tagName: "INPUT", isContentEditable: false };
    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other", ignoreInputTargets: "none" });
    service.setBuiltinKeybindings([{ command: "builtin.doThing", key: "ctrl+j" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.doThing", key: "ctrl+l", mac: null, when: null }]);

    const builtinEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: inputTarget });
    const builtinHandled = await service.dispatchKeydown(builtinEvent);
    expect(builtinHandled).toBe(true);
    expect(builtinEvent.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);

    const extEvent = makeKeydownEvent({ key: "l", ctrlKey: true, target: inputTarget });
    const extHandled = await service.dispatchKeydown(extEvent);
    expect(extHandled).toBe(true);
    expect(extEvent.defaultPrevented).toBe(true);
    expect(extRun).toHaveBeenCalledTimes(1);
  });

  it('ignores keybindings when the target is inside a `data-keybinding-barrier="true"` subtree', async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.barrier", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other", ignoreInputTargets: "none" });
    service.setBuiltinKeybindings([{ command: "builtin.barrier", key: "ctrl+k" }]);

    const barrier = { dataset: { keybindingBarrier: "true" } };
    const target = { tagName: "DIV", isContentEditable: false, parentElement: barrier };

    const event = makeKeydownEvent({ key: "k", ctrlKey: true, target });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(builtinRun).not.toHaveBeenCalled();
  });

  it("does not run extension keybindings from capture-phase listeners", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    const extRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.doThing", "Builtin", builtinRun);
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext", command: "ext.doThing", title: "Ext" }],
      async () => extRun(),
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other", ignoreInputTargets: "none" });
    service.setBuiltinKeybindings([{ command: "builtin.doThing", key: "ctrl+j" }]);
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.doThing", key: "ctrl+l", mac: null, when: null }]);

    const listeners: Array<{ capture: boolean; handler: (evt: KeyboardEvent) => void }> = [];
    const fakeWindow = {
      addEventListener: (_type: string, handler: any, options?: any) => {
        listeners.push({ capture: Boolean(options?.capture), handler });
      },
      removeEventListener: () => {},
    };
    service.installWindowListener(fakeWindow as any, { capture: true });

    const captureHandler = listeners.find((l) => l.capture)?.handler;
    const bubbleHandler = listeners.find((l) => !l.capture)?.handler;
    expect(captureHandler).toBeDefined();
    expect(bubbleHandler).toBeDefined();

    // Extension-only chord should not fire during capture, but should fire during bubble.
    const event = makeKeydownEvent({ key: "l", ctrlKey: true });
    captureHandler!(event);
    await flushMicrotasks();
    expect(extRun).not.toHaveBeenCalled();
    expect(event.defaultPrevented).toBe(false);

    bubbleHandler!(event);
    await flushMicrotasks();
    expect(extRun).toHaveBeenCalledTimes(1);
    expect(event.defaultPrevented).toBe(true);

    // Builtins should still win overall.
    const event2 = makeKeydownEvent({ key: "j", ctrlKey: true });
    captureHandler!(event2);
    await flushMicrotasks();
    expect(builtinRun).toHaveBeenCalledTimes(1);
    expect(event2.defaultPrevented).toBe(true);

    bubbleHandler!(event2);
    await flushMicrotasks();
    expect(extRun).toHaveBeenCalledTimes(1);
  });

  it("dispatches repeated keydown events when the binding opts into repeat", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const builtinRun = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.repeatAllowed", "Builtin", builtinRun);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.repeatAllowed", key: "ctrl+k", allowRepeat: true }]);

    const event = makeKeydownEvent({ key: "k", ctrlKey: true, repeat: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(true);
    expect(event.defaultPrevented).toBe(true);
    expect(builtinRun).toHaveBeenCalledTimes(1);
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
