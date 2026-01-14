import { describe, expect, it } from "vitest";

import { registerDesktopCommands } from "../../commands/registerDesktopCommands.js";
import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState.js";
import { panelRegistry } from "../../panels/panelRegistry.js";

import { COMMAND_REGISTRY_EXEMPT_IDS } from "../ribbonCommandRegistryDisabling.js";
import { defaultRibbonSchema } from "../ribbonSchema.js";

describe("Ribbon CommandRegistry exemptions drift-guards", () => {
  function collectDefaultRibbonIds(): Set<string> {
    const ids = new Set<string>();
    for (const tab of defaultRibbonSchema.tabs) {
      for (const group of tab.groups) {
        for (const button of group.buttons) {
          ids.add(button.id);
          for (const menuItem of button.menuItems ?? []) {
            ids.add(menuItem.id);
          }
        }
      }
    }
    return ids;
  }

  function createDesktopCommandRegistry(): CommandRegistry {
    const commandRegistry = new CommandRegistry();

    // Mirror `apps/desktop/src/main.ts` so the test stays representative as the command catalog evolves.
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
      // Some registered commands reference split APIs; keep them present so registrations
      // don't drift when invoked from unit tests.
      setSplitDirection(direction: string, ratio?: number) {
        this.layout = {
          ...this.layout,
          splitView: {
            ...(this.layout as any).splitView,
            direction,
            ratio: typeof ratio === "number" ? ratio : (this.layout as any)?.splitView?.ratio ?? 0.5,
          },
        };
      },
    } as any;

    registerDesktopCommands({
      commandRegistry,
      app: {} as any,
      layoutController,
      themeController: { setThemePreference: () => {} } as any,
      refreshRibbonUiState: () => {},
      applyFormattingToSelection: () => {},
      getActiveCellNumberFormat: () => null,
      getActiveCellIndentLevel: () => 0,
      formatPainter: { isArmed: () => false, arm: () => {}, disarm: () => {}, onCancel: null },
      openFormatCells: () => {},
      showQuickPick: async () => null,
      findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
      ribbonMacroHandlers: {
        openPanel: () => {},
        focusScriptEditorPanel: () => {},
        focusVbaMigratePanel: () => {},
        setPendingMacrosPanelFocus: () => {},
        startMacroRecorder: () => {},
        stopMacroRecorder: () => {},
        isTauri: () => false,
      },
      dataQueriesHandlers: {
        getPowerQueryService: () => null,
        showToast: () => {},
        notify: () => {},
        focusAfterExecute: () => {},
      },
      pageLayoutHandlers: {
        openPageSetupDialog: () => {},
        updatePageSetup: () => {},
        setPrintArea: () => {},
        clearPrintArea: () => {},
        addToPrintArea: () => {},
        exportPdf: () => {},
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
      openCommandPalette: () => {},
    });

    return commandRegistry;
  }

  it("ensures COMMAND_REGISTRY_EXEMPT_IDS stays in sync with defaultRibbonSchema", () => {
    const ribbonIds = collectDefaultRibbonIds();
    const staleExemptions = [...COMMAND_REGISTRY_EXEMPT_IDS].filter((id) => !ribbonIds.has(id)).sort();
    expect(
      staleExemptions,
      `Exemptions contain ids that are no longer present in defaultRibbonSchema:\n${staleExemptions.map((id) => `- ${id}`).join("\n")}`,
    ).toEqual([]);
  });

  it("ensures COMMAND_REGISTRY_EXEMPT_IDS does not contain registered CommandRegistry ids", () => {
    const commandRegistry = createDesktopCommandRegistry();
    const implementedExemptions = [...COMMAND_REGISTRY_EXEMPT_IDS].filter((id) => commandRegistry.getCommand(id) != null).sort();
    expect(
      implementedExemptions,
      [
        "Exemptions contain ids that are now registered commands (please remove them from COMMAND_REGISTRY_EXEMPT_IDS):",
        ...implementedExemptions.map((id) => `- ${id}`),
      ].join("\n"),
    ).toEqual([]);
  });
});

