import { describe, expect, it, vi } from "vitest";

import { createRibbonActions } from "../ribbonCommandRouter.js";

describe("createRibbonActions (core command routing)", () => {
  it("routes Home tab commands through CommandRegistry.executeCommand", () => {
    const executeCommand = vi.fn(async () => {});
    const onCommandFallback = vi.fn();
    const onToggleFallback = vi.fn();

    const actions = createRibbonActions({
      commandRegistry: { executeCommand } as any,
      onCommandFallback,
      onToggleFallback,
      onError: () => {},
    });

    actions.onCommand?.("clipboard.copy");
    expect(executeCommand).toHaveBeenCalledWith("clipboard.copy");

    actions.onCommand?.("clipboard.pasteSpecial.values");
    expect(executeCommand).toHaveBeenCalledWith("clipboard.pasteSpecial.values");

    actions.onCommand?.("format.numberFormat.percent");
    expect(executeCommand).toHaveBeenCalledWith("format.numberFormat.percent");

    actions.onCommand?.("edit.find");
    expect(executeCommand).toHaveBeenCalledWith("edit.find");

    actions.onCommand?.("format.toggleBold");
    expect(executeCommand).toHaveBeenCalledWith("format.toggleBold");

    actions.onToggle?.("format.toggleWrapText", true);
    expect(executeCommand).toHaveBeenCalledWith("format.toggleWrapText", true);

    actions.onCommand?.("some.unknown.command");
    expect(onCommandFallback).toHaveBeenCalledWith("some.unknown.command");
    expect(onToggleFallback).not.toHaveBeenCalled();
  });
});
