import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { COMMAND_REGISTRY_EXEMPT_IDS, computeRibbonDisabledByIdFromCommandRegistry } from "../../ribbon/ribbonCommandRegistryDisabling.js";
import { registerDesktopCommands } from "../registerDesktopCommands.js";
import { HOME_STYLES_COMMAND_IDS } from "../registerHomeStylesCommands.js";

describe("Home → Styles CommandRegistry commands", () => {
  it("registers the implemented Home → Styles commands and keeps them enabled via CommandRegistry baseline", () => {
    const commandRegistry = new CommandRegistry();

    registerDesktopCommands({
      commandRegistry,
      app: {} as any,
      layoutController: null,
      applyFormattingToSelection: () => {},
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
      openFormatCells: () => {},
      showQuickPick: async () => null,
      findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
      sheetStructureHandlers: {
        openOrganizeSheets: () => {},
        insertSheet: () => {},
        deleteSheet: () => {},
      },
      workbenchFileHandlers: {
        newWorkbook: () => {},
        openWorkbook: () => {},
        saveWorkbook: () => {},
        saveWorkbookAs: () => {},
        setAutoSaveEnabled: () => {},
        print: () => {},
        printPreview: () => {},
        closeWorkbook: () => {},
        quit: () => {},
      },
    });

    const ids = [
      HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral,
      HOME_STYLES_COMMAND_IDS.formatAsTable.light,
      HOME_STYLES_COMMAND_IDS.formatAsTable.medium,
      HOME_STYLES_COMMAND_IDS.formatAsTable.dark,
    ] as const;

    for (const id of ids) {
      expect(commandRegistry.getCommand(id), `Expected '${id}' to be registered`).toBeDefined();
      expect(COMMAND_REGISTRY_EXEMPT_IDS.has(id), `Expected '${id}' to not be exempt`).toBe(false);
    }

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    for (const id of ids) {
      expect(disabledById[id], `Expected '${id}' to not be disabled by baseline`).toBeUndefined();
    }
  });
});
