import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { FORMAT_PAINTER_COMMAND_ID, registerFormatPainterCommand } from "./formatPainterCommand.js";

describe("format painter command", () => {
  it("toggles format painter armed state via arm/disarm callbacks", async () => {
    const commandRegistry = new CommandRegistry();

    let armed = false;
    const arm = vi.fn(() => {
      armed = true;
    });
    const disarm = vi.fn(() => {
      armed = false;
    });

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => armed,
      arm,
      disarm,
    });

    await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    expect(arm).toHaveBeenCalledTimes(1);
    expect(disarm).toHaveBeenCalledTimes(0);
    expect(armed).toBe(true);

    await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    expect(arm).toHaveBeenCalledTimes(1);
    expect(disarm).toHaveBeenCalledTimes(1);
    expect(armed).toBe(false);
  });

  it("does not arm while editing", async () => {
    const commandRegistry = new CommandRegistry();

    const arm = vi.fn();
    const disarm = vi.fn();

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => false,
      arm,
      disarm,
      isEditing: () => true,
    });

    await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    expect(arm).not.toHaveBeenCalled();
    expect(disarm).not.toHaveBeenCalled();
  });

  it("does not arm in read-only mode", async () => {
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
    expect(arm).not.toHaveBeenCalled();
    expect(disarm).not.toHaveBeenCalled();
  });

  it("still disarms while blocked when already armed", async () => {
    const commandRegistry = new CommandRegistry();

    let armed = true;
    const arm = vi.fn(() => {
      armed = true;
    });
    const disarm = vi.fn(() => {
      armed = false;
    });

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => armed,
      arm,
      disarm,
      isEditing: () => true,
      isReadOnly: () => true,
    });

    await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    expect(arm).not.toHaveBeenCalled();
    expect(disarm).toHaveBeenCalledTimes(1);
    expect(armed).toBe(false);
  });

  it("does not arm while the spreadsheet is editing (split-view secondary editor via global flag)", async () => {
    const commandRegistry = new CommandRegistry();

    const arm = vi.fn();
    const disarm = vi.fn();

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => false,
      arm,
      disarm,
    });

    (globalThis as any).__formulaSpreadsheetIsEditing = true;
    try {
      await commandRegistry.executeCommand(FORMAT_PAINTER_COMMAND_ID);
    } finally {
      delete (globalThis as any).__formulaSpreadsheetIsEditing;
    }

    expect(arm).not.toHaveBeenCalled();
    expect(disarm).not.toHaveBeenCalled();
  });
});
