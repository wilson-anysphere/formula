// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState";
import { panelRegistry } from "../../panels/panelRegistry";
import { registerDesktopCommands } from "../../commands/registerDesktopCommands";
import { registerDataQueriesCommands } from "../../commands/registerDataQueriesCommands";
import { registerFormatPainterCommand } from "../../commands/formatPainterCommand";
import { registerRibbonMacroCommands } from "../../commands/registerRibbonMacroCommands";
import { Ribbon } from "../Ribbon";
import type { RibbonSchema } from "../ribbonSchema";
import { defaultRibbonSchema } from "../ribbonSchema";
import { COMMAND_REGISTRY_EXEMPT_IDS } from "../ribbonCommandRegistryDisabling";
import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbonCommandRegistryDisabling";
import { setRibbonUiState } from "../ribbonUiState";

afterEach(() => {
  act(() => {
    setRibbonUiState({
      pressedById: Object.create(null),
      labelById: Object.create(null),
      disabledById: Object.create(null),
      shortcutById: Object.create(null),
      ariaKeyShortcutsById: Object.create(null),
    });
  });
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("CommandRegistry-backed ribbon disabling", () => {
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
      openFormatCells: () => {},
      showQuickPick: async () => null,
      pageLayoutHandlers: {
        openPageSetupDialog: () => {},
        updatePageSetup: () => {},
        setPrintArea: () => {},
        clearPrintArea: () => {},
        addToPrintArea: () => {},
        exportPdf: () => {},
      },
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

    registerRibbonMacroCommands({
      commandRegistry,
      handlers: {
        openPanel: () => {},
        focusScriptEditorPanel: () => {},
        focusVbaMigratePanel: () => {},
        setPendingMacrosPanelFocus: () => {},
        startMacroRecorder: () => {},
        stopMacroRecorder: () => {},
        isTauri: () => false,
      },
    });

    registerFormatPainterCommand({
      commandRegistry,
      isArmed: () => false,
      arm: () => {},
      disarm: () => {},
    });

    registerDataQueriesCommands({
      commandRegistry,
      layoutController,
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: () => {},
      refreshRibbonUiState: () => {},
      focusAfterExecute: () => {},
    });

    return commandRegistry;
  }

  it("keeps the CommandRegistry exemption list in sync with the ribbon schema", () => {
    const ribbonIds = collectDefaultRibbonIds();
    const staleExemptions = [...COMMAND_REGISTRY_EXEMPT_IDS].filter((id) => !ribbonIds.has(id)).sort();
    expect(
      staleExemptions,
      `Exemptions contain ids that are no longer present in defaultRibbonSchema:\n${staleExemptions.map((id) => `- ${id}`).join("\n")}`,
    ).toEqual([]);
  });

  it("ensures CommandRegistry exemptions are truly non-registry ids (catches drift)", () => {
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

  it("renders an unknown command id as disabled when the baseline override is applied", () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "group",
              label: "Group",
              buttons: [{ id: "unknown.command", label: "Unknown", ariaLabel: "Unknown" }],
            },
          ],
        },
      ],
    };

    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema });

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: baselineDisabledById,
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    act(() => {
      root.render(React.createElement(Ribbon, { actions: {}, schema }));
    });

    const button = container.querySelector<HTMLButtonElement>('[data-command-id="unknown.command"]');
    expect(button).toBeInstanceOf(HTMLButtonElement);
    expect(button?.disabled).toBe(true);

    act(() => root.unmount());
  });

  it("keeps implemented ribbon-only commands enabled even though they are not registered", () => {
    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    // These are currently handled directly by the desktop ribbon command handler (not via CommandRegistry),
    // so they must be exempt from the registry-backed disabling allowlist.
    expect(baselineDisabledById["home.editing.fill.up"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.fill.left"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.fill.series"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.sortFilter.customSort"]).toBeUndefined();
    expect(baselineDisabledById["data.sortFilter.sort.customSort"]).toBeUndefined();

    // Home → Cells structural edit commands are also handled directly in `main.ts`.
    expect(baselineDisabledById["home.cells.insert.insertCells"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.insert.insertSheetRows"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.insert.insertSheetColumns"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.insert.insertSheet"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteCells"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteSheetRows"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteSheetColumns"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteSheet"]).toBeUndefined();
  });

  it("keeps AutoSum dropdown variants enabled even though they are not registered", () => {
    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    // AutoSum dropdown variants are wired via `apps/desktop/src/main.ts` (not CommandRegistry), so they must
    // be exempt from the registry-backed disabling allowlist to stay clickable in the ribbon.
    expect(baselineDisabledById["home.editing.autoSum.average"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.autoSum.countNumbers"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.autoSum.max"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.autoSum.min"]).toBeUndefined();
  });

  it("keeps exempt menu items enabled even when the CommandRegistry does not register them", () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "cells",
              label: "Cells",
              buttons: [
                {
                  id: "home.cells.format",
                  label: "Format",
                  ariaLabel: "Format Cells",
                  kind: "dropdown",
                  menuItems: [
                    { id: "home.cells.format.organizeSheets", label: "Organize Sheets", ariaLabel: "Organize Sheets" },
                  ],
                },
                { id: "home.editing.autoSum.average", label: "Average", ariaLabel: "Average" },
                // Non-exempt id to prove the baseline is still working.
                { id: "totally.unknown", label: "Unknown", ariaLabel: "Unknown" },
              ],
            },
          ],
        },
      ],
    };

    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema });

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: baselineDisabledById,
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    act(() => {
      // Avoid JSX in a `.ts` test file (esbuild treats `.ts` as non-JSX).
      root.render(React.createElement(Ribbon, { actions: {}, schema }));
    });

    const trigger = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.format"]');
    expect(trigger).toBeInstanceOf(HTMLButtonElement);
    expect(trigger?.disabled).toBe(false);

    const average = container.querySelector<HTMLButtonElement>('[data-command-id="home.editing.autoSum.average"]');
    expect(average).toBeInstanceOf(HTMLButtonElement);
    expect(average?.disabled).toBe(false);

    const unknown = container.querySelector<HTMLButtonElement>('[data-command-id="totally.unknown"]');
    expect(unknown).toBeInstanceOf(HTMLButtonElement);
    expect(unknown?.disabled).toBe(true);

    act(() => {
      trigger?.click();
    });

    const organize = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.format.organizeSheets"]');
    expect(organize).toBeInstanceOf(HTMLButtonElement);
    expect(organize?.disabled).toBe(false);

    act(() => root.unmount());
  });

  it("disables unimplemented Clear dropdown menu items while keeping registered ones enabled", () => {
    const commandRegistry = createDesktopCommandRegistry();
    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "home.editing",
              label: "Editing",
              buttons: [
                {
                  id: "home.editing.clear",
                  label: "Clear",
                  ariaLabel: "Clear",
                  kind: "dropdown",
                  menuItems: [
                    { id: "format.clearAll", label: "Clear All", ariaLabel: "Clear All" },
                    { id: "format.clearFormats", label: "Clear Formats", ariaLabel: "Clear Formats" },
                    { id: "edit.clearContents", label: "Clear Contents", ariaLabel: "Clear Contents" },
                    { id: "home.editing.clear.clearComments", label: "Clear Comments", ariaLabel: "Clear Comments" },
                    { id: "home.editing.clear.clearHyperlinks", label: "Clear Hyperlinks", ariaLabel: "Clear Hyperlinks" },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema });

    // These are canonical ids registered in the desktop CommandRegistry.
    expect(disabledById["format.clearAll"]).not.toBe(true);
    expect(disabledById["format.clearFormats"]).not.toBe(true);
    expect(disabledById["edit.clearContents"]).not.toBe(true);

    // These are not yet implemented, so they should remain disabled.
    expect(disabledById["home.editing.clear.clearComments"]).toBe(true);
    expect(disabledById["home.editing.clear.clearHyperlinks"]).toBe(true);

    // The trigger should not be disabled because at least one menu item is enabled.
    expect(disabledById["home.editing.clear"]).not.toBe(true);
  });
});
