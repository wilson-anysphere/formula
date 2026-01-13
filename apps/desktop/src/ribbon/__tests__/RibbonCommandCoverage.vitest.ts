import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState.js";
import { panelRegistry } from "../../panels/panelRegistry.js";
import { registerDesktopCommands } from "../../commands/registerDesktopCommands.js";
import { registerPageLayoutCommands } from "../../commands/registerPageLayoutCommands.js";

import { defaultRibbonSchema } from "../ribbonSchema";

/**
 * Ribbon ids are user-facing integration points (e2e selectors, docs, etc) and the CommandRegistry
 * is the central dispatch mechanism for keyboard shortcuts + command palette.
 *
 * This test ensures that when a ribbon control opts in to a *canonical command namespace*
 * (e.g. `clipboard.*`, `edit.*`, `view.*`), that id is actually registered in CommandRegistry.
 *
 * Adding exemptions:
 * - If a ribbon id is intentionally present but not implemented (yet), add it to
 *   `INTENTIONALLY_UNIMPLEMENTED_RIBBON_COMMAND_IDS` below.
 * - Keep the exemption list small and remove entries as soon as commands are implemented.
 */

const CANONICAL_RIBBON_COMMAND_RE = /^(clipboard|edit|format|view|comments|workbench|ai)\./;

// IDs in canonical namespaces that exist in the ribbon schema but are intentionally not
// registered as commands yet (typically UI placeholders).
const INTENTIONALLY_UNIMPLEMENTED_RIBBON_COMMAND_IDS = new Set<string>([
  // View → Show (placeholders; not wired to SpreadsheetApp/CommandRegistry yet)
  "view.show.ruler",
  "view.show.gridlines",
  "view.show.formulaBar",
  "view.show.headings",

  // View → Workbook Views (placeholders)
  "view.workbookViews.normal",
  "view.workbookViews.pageBreakPreview",
  "view.workbookViews.pageLayout",
  "view.workbookViews.customViews",
  "view.workbookViews.customViews.manage",

  // View → Window (placeholders; most are not implemented in the desktop shell yet)
  "view.window.newWindow",
  "view.window.newWindow.newWindowForActiveSheet",
  "view.window.arrangeAll",
  "view.window.arrangeAll.tiled",
  "view.window.arrangeAll.horizontal",
  "view.window.arrangeAll.vertical",
  "view.window.arrangeAll.cascade",
  "view.window.hide",
  "view.window.unhide",
  "view.window.viewSideBySide",
  "view.window.synchronousScrolling",
  "view.window.resetWindowPosition",
  "view.window.switchWindows",
  "view.window.switchWindows.window1",
  "view.window.switchWindows.window2",

  // View → Macros (implemented via ribbon-only handlers today; not exposed as CommandRegistry commands yet)
  "view.macros.viewMacros",
  "view.macros.viewMacros.run",
  "view.macros.viewMacros.edit",
  "view.macros.viewMacros.delete",
  "view.macros.recordMacro",
  "view.macros.recordMacro.stop",
  "view.macros.useRelativeReferences",
]);

const REQUIRED_PAGE_LAYOUT_COMMAND_IDS = [
  "pageLayout.pageSetup.pageSetupDialog",
  "pageLayout.pageSetup.margins.normal",
  "pageLayout.pageSetup.margins.wide",
  "pageLayout.pageSetup.margins.narrow",
  "pageLayout.pageSetup.margins.custom",
  "pageLayout.pageSetup.orientation.portrait",
  "pageLayout.pageSetup.orientation.landscape",
  "pageLayout.pageSetup.size.letter",
  "pageLayout.pageSetup.size.a4",
  "pageLayout.pageSetup.size.more",
  "pageLayout.printArea.setPrintArea",
  "pageLayout.printArea.clearPrintArea",
  "pageLayout.pageSetup.printArea.set",
  "pageLayout.pageSetup.printArea.clear",
  "pageLayout.pageSetup.printArea.addTo",
  "pageLayout.export.exportPdf",
];

function collectRibbonCommandIds(): string[] {
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
  return [...ids];
}

describe("Ribbon ↔ CommandRegistry coverage", () => {
  it("registers canonical ribbon command ids in CommandRegistry", () => {
    const ribbonIds = collectRibbonCommandIds();
    const idsToCheck = ribbonIds
      .filter((id) => CANONICAL_RIBBON_COMMAND_RE.test(id))
      .filter((id) => !INTENTIONALLY_UNIMPLEMENTED_RIBBON_COMMAND_IDS.has(id))
      // Page Layout is not a fully canonical command namespace yet (many schema ids are still placeholders),
      // but these specific Page Setup/Print Area/PDF Export controls should be real commands so they can
      // be invoked from the command palette/extensions and covered by generic command-disable logic.
      .concat(REQUIRED_PAGE_LAYOUT_COMMAND_IDS.filter((id) => ribbonIds.includes(id)))
      .sort((a, b) => a.localeCompare(b));

    const commandRegistry = new CommandRegistry();

    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
      // Some ribbon-exposed commands (e.g. split view) reference split APIs; keep them present so the
      // registered commands are executable in unit tests if needed.
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
      openFormatCells: () => {},
      showQuickPick: async () => null,
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
      openCommandPalette: () => {},
    });

    registerPageLayoutCommands({
      commandRegistry,
      handlers: {
        openPageSetupDialog: () => {},
        updatePageSetup: () => {},
        setPrintArea: () => {},
        clearPrintArea: () => {},
        addToPrintArea: () => {},
        exportPdf: () => {},
      },
    });
    const missing = idsToCheck.filter((id) => commandRegistry.getCommand(id) == null);

    expect(missing, `Missing CommandRegistry registrations for:\n${missing.map((id) => `- ${id}`).join("\n")}`).toEqual([]);
  });
});
