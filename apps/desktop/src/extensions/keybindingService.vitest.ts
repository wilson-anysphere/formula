import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "./commandRegistry.js";
import { ContextKeyService } from "./contextKeys.js";
import { KeybindingService } from "./keybindingService.js";

function makeKeydownEvent(opts: {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
  target?: any;
}): KeyboardEvent {
  const event: any = {
    key: opts.key,
    ctrlKey: Boolean(opts.ctrlKey),
    metaKey: Boolean(opts.metaKey),
    shiftKey: Boolean(opts.shiftKey),
    altKey: Boolean(opts.altKey),
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
    service.setExtensionKeybindings([{ extensionId: "ext", command: "ext.stealCopy", key: "ctrl+c", mac: null, when: null }]);

    const event = makeKeydownEvent({ key: "c", ctrlKey: true });
    const handled = await service.dispatchKeydown(event);

    expect(handled).toBe(false);
    expect(event.defaultPrevented).toBe(false);
    expect(extRun).not.toHaveBeenCalled();
  });
});

