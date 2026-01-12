import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";

import { WORKBENCH_FILE_COMMANDS, registerWorkbenchFileCommands } from "./registerWorkbenchFileCommands.js";

describe("registerWorkbenchFileCommands", () => {
  it("registers expected commands and routes execution to handlers", async () => {
    const commandRegistry = new CommandRegistry();

    const handlers = {
      newWorkbook: vi.fn(),
      openWorkbook: vi.fn(),
      saveWorkbook: vi.fn(),
      saveWorkbookAs: vi.fn(),
      print: vi.fn(),
      printPreview: vi.fn(),
      closeWorkbook: vi.fn(),
      quit: vi.fn(),
    };

    registerWorkbenchFileCommands({ commandRegistry, handlers });

    for (const commandId of Object.values(WORKBENCH_FILE_COMMANDS)) {
      expect(commandRegistry.getCommand(commandId)).toBeTruthy();
    }

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.newWorkbook);
    expect(handlers.newWorkbook).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.openWorkbook);
    expect(handlers.openWorkbook).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.saveWorkbook);
    expect(handlers.saveWorkbook).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.saveWorkbookAs);
    expect(handlers.saveWorkbookAs).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.print);
    expect(handlers.print).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.printPreview);
    expect(handlers.printPreview).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.closeWorkbook);
    expect(handlers.closeWorkbook).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.quit);
    expect(handlers.quit).toHaveBeenCalledTimes(1);
  });
});
