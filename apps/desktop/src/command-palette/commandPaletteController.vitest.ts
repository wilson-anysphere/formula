import { beforeEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { CommandPaletteController } from "./commandPaletteController.js";

describe("CommandPaletteController", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("calls ensureExtensionsLoaded exactly once across multiple opens", () => {
    const commandRegistry = new CommandRegistry();
    commandRegistry.registerBuiltinCommand("builtin.test", "Test", () => {});

    const ensureExtensionsLoaded = vi.fn(async () => {});

    const extensionHostManager = {
      ready: false,
      error: null as unknown,
      subscribe: () => () => {},
    };

    const controller = new CommandPaletteController({
      commandRegistry,
      ensureExtensionsLoaded,
      extensionHostManager,
    });

    controller.open();
    controller.open();
    controller.close();
    controller.open();

    expect(ensureExtensionsLoaded).toHaveBeenCalledTimes(1);
  });
});

