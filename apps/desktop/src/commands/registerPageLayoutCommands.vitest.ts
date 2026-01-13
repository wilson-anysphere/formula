import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import type { PageSetup } from "../print/index.js";

import { PAGE_LAYOUT_COMMANDS, registerPageLayoutCommands } from "./registerPageLayoutCommands.js";

describe("registerPageLayoutCommands", () => {
  it("registers expected commands and routes execution to handlers", async () => {
    const commandRegistry = new CommandRegistry();

    const handlers = {
      openPageSetupDialog: vi.fn(async () => {}),
      updatePageSetup: vi.fn(async (_patch: (current: PageSetup) => PageSetup) => {}),
      setPrintArea: vi.fn(async () => {}),
      clearPrintArea: vi.fn(async () => {}),
      addToPrintArea: vi.fn(async () => {}),
      exportPdf: vi.fn(async () => {}),
    };

    registerPageLayoutCommands({ commandRegistry, handlers });

    const allCommandIds: string[] = [
      PAGE_LAYOUT_COMMANDS.pageSetupDialog,
      ...Object.values(PAGE_LAYOUT_COMMANDS.margins),
      ...Object.values(PAGE_LAYOUT_COMMANDS.orientation),
      ...Object.values(PAGE_LAYOUT_COMMANDS.size),
      ...Object.values(PAGE_LAYOUT_COMMANDS.printArea),
      PAGE_LAYOUT_COMMANDS.exportPdf,
    ];

    for (const id of allCommandIds) {
      expect(commandRegistry.getCommand(id), `Expected command to be registered: ${id}`).toBeTruthy();
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.pageSetupDialog);
    expect(handlers.openPageSetupDialog).toHaveBeenCalledTimes(1);

    const sample: PageSetup = {
      orientation: "portrait",
      paperSize: 123,
      margins: { left: 0, right: 0, top: 0, bottom: 0, header: 0.5, footer: 0.5 },
      scaling: { kind: "percent", percent: 100 },
    };

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.margins.normal);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(1);
    const normalPatch = handlers.updatePageSetup.mock.calls[0]?.[0];
    expect(typeof normalPatch).toBe("function");
    if (typeof normalPatch === "function") {
      const next = normalPatch(sample);
      expect(next.margins).toEqual({ ...sample.margins, left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 });
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.margins.wide);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(2);
    const widePatch = handlers.updatePageSetup.mock.calls[1]?.[0];
    expect(typeof widePatch).toBe("function");
    if (typeof widePatch === "function") {
      const next = widePatch(sample);
      expect(next.margins).toEqual({ ...sample.margins, left: 1, right: 1, top: 1, bottom: 1 });
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.margins.narrow);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(3);
    const narrowPatch = handlers.updatePageSetup.mock.calls[2]?.[0];
    expect(typeof narrowPatch).toBe("function");
    if (typeof narrowPatch === "function") {
      const next = narrowPatch(sample);
      expect(next.margins).toEqual({ ...sample.margins, left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 });
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.margins.custom);
    expect(handlers.openPageSetupDialog).toHaveBeenCalledTimes(2);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.orientation.portrait);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(4);
    const portraitPatch = handlers.updatePageSetup.mock.calls[3]?.[0];
    expect(typeof portraitPatch).toBe("function");
    if (typeof portraitPatch === "function") {
      expect(portraitPatch({ ...sample, orientation: "landscape" }).orientation).toBe("portrait");
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.orientation.landscape);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(5);
    const landscapePatch = handlers.updatePageSetup.mock.calls[4]?.[0];
    expect(typeof landscapePatch).toBe("function");
    if (typeof landscapePatch === "function") {
      expect(landscapePatch(sample).orientation).toBe("landscape");
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.size.letter);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(6);
    const letterPatch = handlers.updatePageSetup.mock.calls[5]?.[0];
    expect(typeof letterPatch).toBe("function");
    if (typeof letterPatch === "function") {
      expect(letterPatch(sample).paperSize).toBe(1);
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.size.a4);
    expect(handlers.updatePageSetup).toHaveBeenCalledTimes(7);
    const a4Patch = handlers.updatePageSetup.mock.calls[6]?.[0];
    expect(typeof a4Patch).toBe("function");
    if (typeof a4Patch === "function") {
      expect(a4Patch(sample).paperSize).toBe(9);
    }

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.size.more);
    expect(handlers.openPageSetupDialog).toHaveBeenCalledTimes(3);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.printArea.setPrintArea);
    expect(handlers.setPrintArea).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.printArea.set);
    expect(handlers.setPrintArea).toHaveBeenCalledTimes(2);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.printArea.clearPrintArea);
    expect(handlers.clearPrintArea).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.printArea.clear);
    expect(handlers.clearPrintArea).toHaveBeenCalledTimes(2);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.printArea.addTo);
    expect(handlers.addToPrintArea).toHaveBeenCalledTimes(1);

    await commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.exportPdf);
    expect(handlers.exportPdf).toHaveBeenCalledTimes(1);
  });
});

