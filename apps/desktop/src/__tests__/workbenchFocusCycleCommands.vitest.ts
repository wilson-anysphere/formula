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

  it("falls back to the first focusable status bar control when zoom is disabled", async () => {
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
        <select data-testid="zoom-control" disabled>
          <option value="100">100%</option>
        </select>
        <button data-testid="open-version-history-panel">Version history</button>
      </div>
    `;

    const sheetTab = document.querySelector<HTMLElement>('#sheet-tabs button[role="tab"]')!;
    const gridFocus = document.querySelector<HTMLButtonElement>("#grid-focus")!;
    const versionHistory = document.querySelector<HTMLElement>('[data-testid="open-version-history-panel"]')!;

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

    // From sheet tabs, next region is status bar (since order is ribbon -> formula -> grid -> sheet tabs -> status bar).
    sheetTab.focus();
    expect(document.activeElement).toBe(sheetTab);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(versionHistory);
  });

  it("starts cycling from the ribbon when focus is outside any known region (and reverses to the status bar)", async () => {
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
      <button id="outside">Outside</button>
    `;

    const ribbonTab = document.querySelector<HTMLElement>("#ribbon .ribbon__tab")!;
    const zoomControl = document.querySelector<HTMLElement>('.statusbar [data-testid="zoom-control"]')!;
    const outside = document.querySelector<HTMLButtonElement>("#outside")!;
    const gridFocus = document.querySelector<HTMLButtonElement>("#grid-focus")!;

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

    outside.focus();
    expect(document.activeElement).toBe(outside);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(ribbonTab);

    outside.focus();
    expect(document.activeElement).toBe(outside);

    await commandRegistry.executeCommand("workbench.focusPrevRegion");
    expect(document.activeElement).toBe(zoomControl);
  });

  it("treats the secondary grid root as part of the grid region when cycling focus", async () => {
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
      <div id="grid-secondary">
        <button id="grid-secondary-focus">Secondary</button>
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

    const secondaryFocus = document.querySelector<HTMLButtonElement>("#grid-secondary-focus")!;
    const sheetTab = document.querySelector<HTMLElement>('#sheet-tabs button[role="tab"]')!;
    const gridFocus = document.querySelector<HTMLButtonElement>("#grid-focus")!;

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

    secondaryFocus.focus();
    expect(document.activeElement).toBe(secondaryFocus);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(sheetTab);
  });

  it("does not steal focus when a modal dialog is open", async () => {
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

    const gridFocus = document.querySelector<HTMLButtonElement>("#grid-focus")!;
    const app = {
      focus: () => {
        gridFocus.focus();
      },
    } as any;

    registerBuiltinCommands({ commandRegistry, app, layoutController });

    const dialog = document.createElement("dialog");
    dialog.setAttribute("open", "true");
    const input = document.createElement("input");
    dialog.appendChild(input);
    document.body.appendChild(dialog);
    input.focus();
    expect(document.activeElement).toBe(input);

    await commandRegistry.executeCommand("workbench.focusNextRegion");
    expect(document.activeElement).toBe(input);

    dialog.remove();
  });
});
