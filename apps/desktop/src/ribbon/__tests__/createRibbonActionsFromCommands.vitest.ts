// @vitest-environment jsdom
import { describe, expect, it, vi, afterEach } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createRibbonActionsFromCommands } from "../createRibbonActionsFromCommands";

async function flushMicrotasks(): Promise<void> {
  // Allow any `queueMicrotask` / async IIFE work to run.
  await Promise.resolve();
  await Promise.resolve();
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("createRibbonActionsFromCommands", () => {
  it("dispatches registered ribbon commands through CommandRegistry", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("ribbon.test", "Test", run);

    const actions = createRibbonActionsFromCommands({ commandRegistry: registry });
    actions.onCommand?.("ribbon.test");
    await flushMicrotasks();

    expect(run).toHaveBeenCalledTimes(1);
  });

  it("dispatches registered toggle commands with the pressed state and suppresses the follow-up onCommand", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("ribbon.toggle", "Toggle", run);

    const actions = createRibbonActionsFromCommands({ commandRegistry: registry });
    actions.onToggle?.("ribbon.toggle", true);
    // Ribbon toggles invoke both callbacks; simulate that contract.
    actions.onCommand?.("ribbon.toggle");
    await flushMicrotasks();

    expect(run).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledWith(true);
  });

  it("uses overrides when provided (even if the command is not registered)", async () => {
    const registry = new CommandRegistry();
    const override = vi.fn();

    const actions = createRibbonActionsFromCommands({
      commandRegistry: registry,
      commandOverrides: { "ribbon.special": override },
    });

    actions.onCommand?.("ribbon.special");
    await flushMicrotasks();

    expect(override).toHaveBeenCalledTimes(1);
  });

  it("routes command errors to onCommandError", async () => {
    const registry = new CommandRegistry();
    registry.registerBuiltinCommand("ribbon.fail", "Fail", () => {
      throw new Error("boom");
    });

    const onCommandError = vi.fn();
    const actions = createRibbonActionsFromCommands({ commandRegistry: registry, onCommandError });

    actions.onCommand?.("ribbon.fail");
    await flushMicrotasks();

    expect(onCommandError).toHaveBeenCalledTimes(1);
    expect(onCommandError.mock.calls[0]?.[0]).toBe("ribbon.fail");
    expect(onCommandError.mock.calls[0]?.[1]).toBeInstanceOf(Error);
  });

  it("shows a toast for unknown commands by default", async () => {
    const registry = new CommandRegistry();
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const actions = createRibbonActionsFromCommands({ commandRegistry: registry });
    actions.onCommand?.("ribbon.unknown");
    await flushMicrotasks();

    const toast = toastRoot.querySelector<HTMLElement>("[data-testid=\"toast\"]");
    expect(toast?.textContent).toBe("Ribbon: ribbon.unknown");
  });
});

