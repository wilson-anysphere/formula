// @vitest-environment jsdom
import { describe, expect, it, vi, afterEach } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createRibbonActionsFromCommands, createRibbonFileActionsFromCommands } from "../createRibbonActionsFromCommands";

async function flushMicrotasks(): Promise<void> {
  // Allow any `queueMicrotask` / async IIFE work to run.
  // Nested async boundaries (CommandRegistry -> createRibbonActionsFromCommands) can
  // require multiple microtask turns in jsdom/vitest, so flush a small batch.
  for (let i = 0; i < 8; i += 1) {
    await Promise.resolve();
  }
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

  it("runs onBeforeExecuteCommand before executing registered commands", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("ribbon.before", "Before", run);

    const onBeforeExecuteCommand = vi.fn();
    const actions = createRibbonActionsFromCommands({ commandRegistry: registry, onBeforeExecuteCommand });

    actions.onCommand?.("ribbon.before");
    await flushMicrotasks();

    expect(onBeforeExecuteCommand).toHaveBeenCalledTimes(1);
    expect(onBeforeExecuteCommand.mock.calls[0]?.[0]).toBe("ribbon.before");
    expect(onBeforeExecuteCommand.mock.calls[0]?.[1]).toEqual({ kind: "builtin" });
    expect(run).toHaveBeenCalledTimes(1);
  });

  it("routes onBeforeExecuteCommand errors to onCommandError and does not execute the command", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("ribbon.beforeFail", "BeforeFail", run);

    const onCommandError = vi.fn();
    const actions = createRibbonActionsFromCommands({
      commandRegistry: registry,
      onCommandError,
      onBeforeExecuteCommand: () => {
        throw new Error("before boom");
      },
    });

    actions.onCommand?.("ribbon.beforeFail");
    await flushMicrotasks();

    expect(run).not.toHaveBeenCalled();
    expect(onCommandError).toHaveBeenCalledTimes(1);
    expect(onCommandError.mock.calls[0]?.[0]).toBe("ribbon.beforeFail");
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

  it("allows unknown toggles to fall through to onUnknownCommand when onUnknownToggle returns false", async () => {
    const registry = new CommandRegistry();
    const onUnknownCommand = vi.fn();

    const actions = createRibbonActionsFromCommands({
      commandRegistry: registry,
      onUnknownCommand,
      onUnknownToggle: () => false,
    });

    actions.onToggle?.("ribbon.toggle.unknown", true);
    actions.onCommand?.("ribbon.toggle.unknown");
    await flushMicrotasks();

    expect(onUnknownCommand).toHaveBeenCalledTimes(1);
    expect(onUnknownCommand).toHaveBeenCalledWith("ribbon.toggle.unknown");
  });

  it("suppresses the follow-up onCommand for unknown toggles when onUnknownToggle returns true", async () => {
    const registry = new CommandRegistry();
    const onUnknownCommand = vi.fn();

    const actions = createRibbonActionsFromCommands({
      commandRegistry: registry,
      onUnknownCommand,
      onUnknownToggle: () => true,
    });

    actions.onToggle?.("ribbon.toggle.unknown", true);
    actions.onCommand?.("ribbon.toggle.unknown");
    await flushMicrotasks();

    expect(onUnknownCommand).not.toHaveBeenCalled();
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

describe("createRibbonFileActionsFromCommands", () => {
  it("dispatches file actions through CommandRegistry", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("workbench.newWorkbook", "New", run);

    const fileActions = createRibbonFileActionsFromCommands({
      commandRegistry: registry,
      commandIds: { newWorkbook: "workbench.newWorkbook" },
    });

    fileActions.newWorkbook?.();
    await flushMicrotasks();

    expect(run).toHaveBeenCalledTimes(1);
  });

  it("dispatches toggleAutoSave with the enabled state", async () => {
    const registry = new CommandRegistry();
    const run = vi.fn();
    registry.registerBuiltinCommand("workbench.setAutoSaveEnabled", "AutoSave", run);

    const fileActions = createRibbonFileActionsFromCommands({
      commandRegistry: registry,
      commandIds: { toggleAutoSave: "workbench.setAutoSaveEnabled" },
    });

    fileActions.toggleAutoSave?.(true);
    await flushMicrotasks();

    expect(run).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledWith(true);
  });

  it("routes file action errors to onCommandError", async () => {
    const registry = new CommandRegistry();
    registry.registerBuiltinCommand("workbench.fail", "Fail", () => {
      throw new Error("boom");
    });

    const onCommandError = vi.fn();
    const fileActions = createRibbonFileActionsFromCommands({
      commandRegistry: registry,
      onCommandError,
      commandIds: { newWorkbook: "workbench.fail" },
    });

    fileActions.newWorkbook?.();
    await flushMicrotasks();

    expect(onCommandError).toHaveBeenCalledTimes(1);
    expect(onCommandError.mock.calls[0]?.[0]).toBe("workbench.fail");
    expect(onCommandError.mock.calls[0]?.[1]).toBeInstanceOf(Error);
  });
});
