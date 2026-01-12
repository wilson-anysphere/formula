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

    const run = vi.fn();
    commandRegistry.registerBuiltinCommand("builtin.test", "Test", run);

    const service = new KeybindingService({ commandRegistry, contextKeys, platform: "other" });
    service.setBuiltinKeybindings([{ command: "builtin.test", key: "ctrl+j" }]);

    // Sanity check: outside a barrier, the binding should fire.
    const outsideEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: document.body });
    const outsideHandled = await service.dispatchKeydown(outsideEvent);
    expect(outsideHandled).toBe(true);
    expect(outsideEvent.defaultPrevented).toBe(true);
    expect(run).toHaveBeenCalledTimes(1);

    const barrierRoot = document.createElement("div");
    barrierRoot.setAttribute("data-keybinding-barrier", "true");
    const inner = document.createElement("button");
    barrierRoot.appendChild(inner);
    document.body.appendChild(barrierRoot);

    const insideEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: inner });
    const insideHandled = await service.dispatchKeydown(insideEvent);

    expect(insideHandled).toBe(false);
    expect(insideEvent.defaultPrevented).toBe(false);
    expect(run).toHaveBeenCalledTimes(1);

    // Also ensure the synchronous helper respects the barrier and does not schedule execution.
    const insideSyncEvent = makeKeydownEvent({ key: "j", ctrlKey: true, target: inner });
    const syncHandled = service.handleKeydown(insideSyncEvent);
    expect(syncHandled).toBe(false);
    expect(insideSyncEvent.defaultPrevented).toBe(false);
    await flushMicrotasks();
    expect(run).toHaveBeenCalledTimes(1);

    barrierRoot.remove();
  });
});

