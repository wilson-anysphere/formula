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
});

