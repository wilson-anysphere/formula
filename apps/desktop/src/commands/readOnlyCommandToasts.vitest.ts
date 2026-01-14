/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { registerDataQueriesCommands, DATA_QUERIES_RIBBON_COMMANDS } from "./registerDataQueriesCommands.js";
import { registerRibbonMacroCommands } from "./registerRibbonMacroCommands.js";
import { registerFormatPainterCommand, FORMAT_PAINTER_COMMAND_ID } from "./formatPainterCommand.js";

describe("read-only command toasts", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(1_000);
    document.body.innerHTML = `<div id="toast-root"></div>`;
  });

  afterEach(() => {
    vi.clearAllTimers();
    vi.useRealTimers();
  });

  it("shows a read-only toast when Data Queries refresh is invoked in read-only mode", async () => {
    const commandRegistry = new CommandRegistry();
    const focusAfterExecute = vi.fn();
    const refreshAll = vi.fn(() => ({ promise: Promise.resolve() }));

    registerDataQueriesCommands({
      commandRegistry,
      layoutController: null,
      getPowerQueryService: () => ({ ready: Promise.resolve(), getQueries: () => [{ id: "q1" }], refreshAll } as any),
      showToast: () => {},
      notify: () => {},
      focusAfterExecute,
      isReadOnly: () => true,
    });

    await commandRegistry.executeCommand(DATA_QUERIES_RIBBON_COMMANDS.refreshAll);
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("refresh queries");
    expect(refreshAll).not.toHaveBeenCalled();
    expect(focusAfterExecute).toHaveBeenCalledTimes(1);
  });

  it("shows a read-only toast when a macro command is invoked in read-only mode", async () => {
    const commandRegistry = new CommandRegistry();
    const openPanel = vi.fn();

    registerRibbonMacroCommands({
      commandRegistry,
      isReadOnly: () => true,
      handlers: {
        openPanel,
        focusScriptEditorPanel: vi.fn(),
        focusVbaMigratePanel: vi.fn(),
        setPendingMacrosPanelFocus: vi.fn(),
        startMacroRecorder: vi.fn(),
        stopMacroRecorder: vi.fn(),
        isTauri: () => false,
      },
    });

    await commandRegistry.executeCommand("view.macros.viewMacros.run");
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("run macros");
    expect(openPanel).not.toHaveBeenCalled();
  });

  it("shows a read-only toast when Format Painter is invoked in read-only mode", async () => {
    const commandRegistry = new CommandRegistry();
    const arm = vi.fn();
    const disarm = vi.fn();

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => false,
      arm,
      disarm,
      isReadOnly: () => true,
    });

    await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Format Painter");
    expect(arm).not.toHaveBeenCalled();
    expect(disarm).not.toHaveBeenCalled();
  });
});
