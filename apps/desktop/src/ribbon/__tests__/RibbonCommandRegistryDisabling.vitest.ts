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
import { COMMAND_REGISTRY_EXEMPT_IDS, computeRibbonDisabledByIdFromCommandRegistry } from "../ribbonCommandRegistryDisabling";
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

  it("keeps ribbon-only commands enabled via the exemption list (not CommandRegistry)", () => {
    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    // These are currently handled directly by the desktop ribbon command handler (not via CommandRegistry),
    // so they must be exempt from the registry-backed disabling allowlist.
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

  it("registers Fill Up/Left/Series ribbon ids as CommandRegistry commands (no exemptions needed)", () => {
    const commandRegistry = createDesktopCommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    const ids = ["home.editing.fill.up", "home.editing.fill.left", "home.editing.fill.series"] as const;
    for (const id of ids) {
      expect(commandRegistry.getCommand(id), `Expected '${id}' to be registered`).toBeDefined();
      expect(COMMAND_REGISTRY_EXEMPT_IDS.has(id), `Expected '${id}' to not be exempt`).toBe(false);
      expect(baselineDisabledById[id], `Expected '${id}' to not be disabled by baseline`).toBeUndefined();
    }
  });

  it("registers AutoSum dropdown variants as CommandRegistry commands (no exemptions needed)", () => {
    const commandRegistry = createDesktopCommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    const ids = [
      "home.editing.autoSum.average",
      "home.editing.autoSum.countNumbers",
      "home.editing.autoSum.max",
      "home.editing.autoSum.min",
    ] as const;
    for (const id of ids) {
      expect(commandRegistry.getCommand(id), `Expected '${id}' to be registered`).toBeDefined();
      expect(COMMAND_REGISTRY_EXEMPT_IDS.has(id), `Expected '${id}' to not be exempt`).toBe(false);
      expect(baselineDisabledById[id], `Expected '${id}' to not be disabled by baseline`).toBeUndefined();
    }
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
                // Exempt command id to prove the exemption list keeps implemented ribbon-only
                // actions enabled even when the CommandRegistry doesn't register them.
                { id: "home.cells.insert.insertCells", label: "Insert Cells", ariaLabel: "Insert Cells" },
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

    const insertCells = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.insert.insertCells"]');
    expect(insertCells).toBeInstanceOf(HTMLButtonElement);
    expect(insertCells?.disabled).toBe(false);

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
