// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { closePanel, createDefaultLayout, openPanel } from "../layout/layoutState.js";
import { panelRegistry } from "../panels/panelRegistry.js";
import { registerBuiltinCommands } from "../commands/registerBuiltinCommands.js";

describe("workbench focus region commands", () => {
  it("cycles focus between ribbon, formula bar, grid, sheet tabs, and status bar", async () => {
    document.body.innerHTML = `
      <div id="ribbon">
        <button class="ribbon__tab" role="tab" aria-selected="true">Home</button>
      </div>
      <div id="formula-bar">
        <input data-testid="formula-address" />
      </div>
      <div id="grid">
        <button id="grid-focus">Grid</button>
      </div>
      <div id="sheet-tabs">
        <button role="tab" aria-selected="true">Sheet1</button>
      </div>
      <div class="statusbar">
        <select data-testid="zoom-control">
          <option value="100">100%</option>
        </select>
      </div>
    `;

    const ribbonTab = document.querySelector<HTMLElement>("#ribbon .ribbon__tab")!;
    const formulaAddress = document.querySelector<HTMLInputElement>('#formula-bar [data-testid="formula-address"]')!;
    const gridFocus = document.querySelector<HTMLButtonElement>("#grid-focus")!;
    const sheetTab = document.querySelector<HTMLElement>('#sheet-tabs button[role="tab"]')!;
    const zoomControl = document.querySelector<HTMLElement>('.statusbar [data-testid="zoom-control"]')!;

    const commandRegistry = new CommandRegistry();
    const layoutController = {
      layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
      openPanel(panelId: string) {
        this.layout = openPanel(this.layout, panelId, { panelRegistry });
      },
      closePanel(panelId: string) {
        this.layout = closePanel(this.layout, panelId);
      },
    } as any;

    const app = {
      focus: () => {
        gridFocus.focus();
      },
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    ribbonTab.focus();
    expect(document.activeElement).toBe(ribbonTab);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(formulaAddress);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(gridFocus);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(sheetTab);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(zoomControl);

    // Wrap.
    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(ribbonTab);

    // Reverse direction.
    await commandRegistry.executeCommand("workbench.focusPrevRegion");
    expect(document.activeElement).toBe(zoomControl);
  });
});

