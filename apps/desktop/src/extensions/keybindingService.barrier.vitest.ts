// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "./commandRegistry.js";
import { ContextKeyService } from "./contextKeys.js";
import { KeybindingService } from "./keybindingService.js";

function makeKeydownEvent(opts: { key: string; ctrlKey?: boolean; target?: EventTarget | null }): KeyboardEvent {
  const event: any = {
    key: opts.key,
    code: "",
    ctrlKey: Boolean(opts.ctrlKey),
    metaKey: false,
    shiftKey: false,
    altKey: false,
    repeat: false,
    target: opts.target ?? null,
    defaultPrevented: false,
  };
  event.preventDefault = () => {
    event.defaultPrevented = true;
  };
  return event as KeyboardEvent;
}

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i += 1) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("KeybindingService keybinding barrier", () => {
  it("suppresses dispatch for events originating inside a keybinding barrier subtree", async () => {
    const contextKeys = new ContextKeyService();
    const commandRegistry = new CommandRegistry();

    const runBuiltin = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.test", "Test", runBuiltin);

    const runExtension = vi.fn();
    commandRegistry.setExtensionCommands(
      [{ extensionId: "ext.test", command: "extension.test", title: "Extension Test" }],
      async () => {
        runExtension();
      },
    );

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.test", key: "ctrl+j" }]);
    service.setExtensionKeybindings([{ extensionId: "ext.test", command: "extension.test", key: "arrowdown", mac: null, when: null }]);

    // Sanity check: outside a barrier, the binding should fire.
    const outsideEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: document.body });
    const outsideHandled = await service.dispatchKeydown(outsideEvent);
    expect(outsideHandled).toBe(true);
    expect(outsideEvent.defaultPrevented).toBe(true);
    expect(runBuiltin).toHaveBeenCalledTimes(1);

    // Sanity check: extension keybindings also fire outside a barrier.
    const outsideExtEvent = makeKeydownEvent({ key: "ArrowDown", target: document.body });
    const outsideExtHandled = await service.dispatchKeydown(outsideExtEvent);
    expect(outsideExtHandled).toBe(true);
    expect(outsideExtEvent.defaultPrevented).toBe(true);
    expect(runExtension).toHaveBeenCalledTimes(1);

    const barrierRoot = document.createElement("div");
    barrierRoot.setAttribute("data-keybinding-barrier", "true");
    const inner = document.createElement("button");
    barrierRoot.appendChild(inner);
    document.body.appendChild(barrierRoot);

    const insideEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: inner });
    const insideHandled = await service.dispatchKeydown(insideEvent);

    expect(insideHandled).toBe(false);
    expect(insideEvent.defaultPrevented).toBe(false);
    expect(runBuiltin).toHaveBeenCalledTimes(1);

    const insideExtEvent = makeKeydownEvent({ key: "ArrowDown", target: inner });
    const insideExtHandled = await service.dispatchKeydown(insideExtEvent);
    expect(insideExtHandled).toBe(false);
    expect(insideExtEvent.defaultPrevented).toBe(false);
    expect(runExtension).toHaveBeenCalledTimes(1);

    // Also ensure the synchronous helper respects the barrier and does not schedule execution.
    const insideSyncEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: inner });
    const syncHandled = service.handleKeydown(insideSyncEvent);
    expect(syncHandled).toBe(false);
    expect(insideSyncEvent.defaultPrevented).toBe(false);
    await flushMicrotasks();
    expect(runBuiltin).toHaveBeenCalledTimes(1);

    const insideSyncExtEvent = makeKeydownEvent({ key: "ArrowDown", target: inner });
    const syncExtHandled = service.handleKeydown(insideSyncExtEvent);
    expect(syncExtHandled).toBe(false);
    expect(insideSyncExtEvent.defaultPrevented).toBe(false);
    await flushMicrotasks();
    expect(runExtension).toHaveBeenCalledTimes(1);

    barrierRoot.remove();
  });
});
