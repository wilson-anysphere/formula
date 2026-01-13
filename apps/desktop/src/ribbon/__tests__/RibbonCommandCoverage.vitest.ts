import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState.js";
import { panelRegistry } from "../../panels/panelRegistry.js";
import { registerBuiltinCommands } from "../../commands/registerBuiltinCommands.js";
import { registerNumberFormatCommands } from "../../commands/registerNumberFormatCommands.js";
import { registerWorkbenchFileCommands } from "../../commands/registerWorkbenchFileCommands.js";

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

    // Register the same built-in commands that the desktop shell wires up.
    registerBuiltinCommands({
      commandRegistry,
      app: {} as any,
      layoutController,
      themeController: { setThemePreference: () => {} } as any,
      refreshRibbonUiState: () => {},
    });
    registerWorkbenchFileCommands({
      commandRegistry,
      handlers: {
        newWorkbook: () => {},
        openWorkbook: () => {},
        saveWorkbook: () => {},
        saveWorkbookAs: () => {},
        print: () => {},
        printPreview: () => {},
        closeWorkbook: () => {},
        quit: () => {},
      },
    });

    // Number format commands are registered in the desktop shell via `registerNumberFormatCommands(...)`
    // (today invoked from `apps/desktop/src/main.ts`). Register them here so ribbon ids like
    // `format.numberFormat.accounting.*` stay covered without needing a long stub list.
    registerNumberFormatCommands({
      commandRegistry,
      applyFormattingToSelection: () => {},
      getActiveCellNumberFormat: () => null,
      t: (key) => key,
      category: null,
    });

    // Some canonical commands are still registered inline in `apps/desktop/src/main.ts`
    // (because they depend on UI dialogs / selection helpers). Mirror those ids here so
    // this test remains a Ribbon↔CommandRegistry drift guard even before command
    // registration is fully centralized.
    for (const id of [
      "edit.find",
      "edit.replace",
      "format.toggleBold",
      "format.toggleItalic",
      "format.toggleUnderline",
      "format.toggleStrikethrough",
      "format.toggleWrapText",
      "format.openFormatCells",
    ]) {
      if (commandRegistry.getCommand(id)) continue;
      commandRegistry.registerBuiltinCommand(id, id, () => {});
    }

    const missing = idsToCheck.filter((id) => commandRegistry.getCommand(id) == null);

    expect(missing, `Missing CommandRegistry registrations for:\n${missing.map((id) => `- ${id}`).join("\n")}`).toEqual([]);
  });
});
