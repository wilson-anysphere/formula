import { describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbon/ribbonCommandRegistryDisabling.js";
import type { RibbonSchema } from "../ribbon/ribbonSchema.js";

import { registerDesktopCommands } from "./registerDesktopCommands.js";
import { HOME_STYLES_COMMAND_IDS } from "./registerHomeStylesCommands.js";

describe("registerHomeStylesCommands (via registerDesktopCommands)", () => {
  it("registers Home → Styles command ids and keeps them enabled in CommandRegistry-backed ribbon disabling", async () => {
    const commandRegistry = new CommandRegistry();

    const focus = vi.fn();
    const app = {
      isEditing: () => false,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 100, maxCols: 100 }),
      focus,
    } as any;

    const applyFormattingToSelection = vi.fn();
    const showQuickPick = vi.fn(async () => "good");

    registerDesktopCommands({
      commandRegistry,
      app,
      layoutController: null,
      isEditing: () => false,
      applyFormattingToSelection,
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
      openFormatCells: () => {},
      showQuickPick: showQuickPick as any,
      findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
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

    for (const id of [
      HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral,
      HOME_STYLES_COMMAND_IDS.formatAsTable.light,
      HOME_STYLES_COMMAND_IDS.formatAsTable.medium,
      HOME_STYLES_COMMAND_IDS.formatAsTable.dark,
    ]) {
      const cmd = commandRegistry.getCommand(id);
      expect(cmd, id).toBeTruthy();
      expect(cmd?.source.kind, id).toBe("builtin");
    }

    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "styles",
              label: "Styles",
              buttons: [
                {
                  id: "home.styles.cellStyles",
                  label: "Cell Styles",
                  ariaLabel: "Cell Styles",
                  kind: "dropdown",
                  menuItems: [
                    {
                      id: HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral,
                      label: "Good, Bad, and Neutral",
                      ariaLabel: "Good, Bad, and Neutral",
                    },
                  ],
                },
                {
                  id: "home.styles.formatAsTable",
                  label: "Format as Table",
                  ariaLabel: "Format as Table",
                  kind: "dropdown",
                  menuItems: [
                    { id: HOME_STYLES_COMMAND_IDS.formatAsTable.light, label: "Light", ariaLabel: "Light" },
                    { id: HOME_STYLES_COMMAND_IDS.formatAsTable.medium, label: "Medium", ariaLabel: "Medium" },
                    { id: HOME_STYLES_COMMAND_IDS.formatAsTable.dark, label: "Dark", ariaLabel: "Dark" },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const disabledWithoutCommands = computeRibbonDisabledByIdFromCommandRegistry(new CommandRegistry(), {
      schema,
      // Override the global exemption list so this test verifies the *registration* behavior:
      // these ids should be disabled when missing and enabled when present.
      isExemptFromCommandRegistry: () => false,
    });
    expect(disabledWithoutCommands[HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral]).toBe(true);
    expect(disabledWithoutCommands[HOME_STYLES_COMMAND_IDS.formatAsTable.light]).toBe(true);

    const disabledWithCommands = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, {
      schema,
      isExemptFromCommandRegistry: () => false,
    });
    expect(disabledWithCommands[HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral]).toBeUndefined();
    expect(disabledWithCommands[HOME_STYLES_COMMAND_IDS.formatAsTable.light]).toBeUndefined();
    expect(disabledWithCommands[HOME_STYLES_COMMAND_IDS.formatAsTable.medium]).toBeUndefined();
    expect(disabledWithCommands[HOME_STYLES_COMMAND_IDS.formatAsTable.dark]).toBeUndefined();

    await commandRegistry.executeCommand(HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral);
    expect(showQuickPick).toHaveBeenCalledTimes(1);
    const items = showQuickPick.mock.calls[0]?.[0] as any[];
    const placeholder = showQuickPick.mock.calls[0]?.[1]?.placeHolder;
    expect(items.map((i) => i.label)).toEqual(expect.arrayContaining(["Good", "Bad", "Neutral"]));
    expect(placeholder).toBe("Good, Bad, and Neutral");
    expect(applyFormattingToSelection).toHaveBeenCalledTimes(1);
    const [label, applyFn] = applyFormattingToSelection.mock.calls[0] ?? [];
    expect(label).toBe("Cell style: Good");
    expect(typeof applyFn).toBe("function");
    const doc = { setRangeFormat: vi.fn(() => true) } as any;
    applyFn(doc, "sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    expect(doc.setRangeFormat).toHaveBeenCalled();

    applyFormattingToSelection.mockClear();

    await commandRegistry.executeCommand(HOME_STYLES_COMMAND_IDS.formatAsTable.light);
    expect(applyFormattingToSelection).toHaveBeenCalledTimes(1);
    const [tableLabel, tableFn, tableOptions] = applyFormattingToSelection.mock.calls[0] ?? [];
    expect(tableLabel).toBe("Format as Table");
    expect(tableOptions).toEqual({ forceBatch: true });
    const tableDoc = { setRangeFormat: vi.fn(() => true) } as any;
    tableFn(tableDoc, "sheet1", [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }]);
    expect(tableDoc.setRangeFormat).toHaveBeenCalled();
  });
});
