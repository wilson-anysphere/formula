// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { Ribbon } from "../Ribbon";
import type { RibbonSchema } from "../ribbonSchema";
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

  it("keeps implemented ribbon-only Fill Up/Left commands enabled even though they are not registered", () => {
    const commandRegistry = new CommandRegistry();
    const baselineDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);

    // These are currently handled directly by the desktop ribbon command handler (not via CommandRegistry),
    // so they must be exempt from the registry-backed disabling allowlist.
    expect(baselineDisabledById["home.editing.fill.up"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.fill.left"]).toBeUndefined();
    expect(baselineDisabledById["home.editing.fill.series"]).toBeUndefined();

    // Home → Cells structural edit commands are also handled directly in `main.ts`.
    expect(baselineDisabledById["home.cells.insert.insertSheetRows"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.insert.insertSheetColumns"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteSheetRows"]).toBeUndefined();
    expect(baselineDisabledById["home.cells.delete.deleteSheetColumns"]).toBeUndefined();
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
                    { id: "home.cells.format.formatCells", label: "Format Cells…", ariaLabel: "Format Cells" },
                    { id: "home.cells.format.organizeSheets", label: "Organize Sheets", ariaLabel: "Organize Sheets" },
                  ],
                },
                { id: "data.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z" },
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

    const sort = container.querySelector<HTMLButtonElement>('[data-command-id="data.sortFilter.sortAtoZ"]');
    expect(sort).toBeInstanceOf(HTMLButtonElement);
    expect(sort?.disabled).toBe(false);

    const unknown = container.querySelector<HTMLButtonElement>('[data-command-id="totally.unknown"]');
    expect(unknown).toBeInstanceOf(HTMLButtonElement);
    expect(unknown?.disabled).toBe(true);

    act(() => {
      trigger?.click();
    });

    const formatCells = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.format.formatCells"]');
    const organizeSheets = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.format.organizeSheets"]');
    expect(formatCells).toBeInstanceOf(HTMLButtonElement);
    expect(organizeSheets).toBeInstanceOf(HTMLButtonElement);
    expect(formatCells?.disabled).toBe(false);
    expect(organizeSheets?.disabled).toBe(false);

    act(() => root.unmount());
  });

  it("keeps Home → Editing → Clear menu items enabled via the exemption list", () => {
    const commandRegistry = new CommandRegistry();
    const schema: RibbonSchema = {
      tabs: [
        {
          id: "home",
          label: "Home",
          groups: [
            {
              id: "editing",
              label: "Editing",
              buttons: [
                {
                  id: "home.editing.clear",
                  label: "Clear",
                  ariaLabel: "Clear",
                  kind: "dropdown",
                  menuItems: [
                    { id: "home.editing.clear.clearAll", label: "Clear All", ariaLabel: "Clear All" },
                    { id: "home.editing.clear.clearFormats", label: "Clear Formats", ariaLabel: "Clear Formats" },
                    { id: "home.editing.clear.clearContents", label: "Clear Contents", ariaLabel: "Clear Contents" },
                    { id: "home.editing.clear.clearComments", label: "Clear Comments", ariaLabel: "Clear Comments" },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema });

    // These menu items are ribbon-specific ids but are routed to real `format.clear*` commands
    // in `main.ts` (via `resolveHomeEditingClearCommandTarget`). They should not be auto-disabled
    // just because the command registry doesn't have matching ids.
    expect(disabledById["home.editing.clear.clearAll"]).not.toBe(true);
    expect(disabledById["home.editing.clear.clearFormats"]).not.toBe(true);
    expect(disabledById["home.editing.clear.clearContents"]).not.toBe(true);

    // Unimplemented variants should remain disabled by default.
    expect(disabledById["home.editing.clear.clearComments"]).toBe(true);

    // The trigger should not be disabled because at least one menu item is enabled.
    expect(disabledById["home.editing.clear"]).not.toBe(true);
  });
});
