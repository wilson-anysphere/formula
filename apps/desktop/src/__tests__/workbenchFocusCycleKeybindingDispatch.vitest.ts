// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { KeybindingService } from "../extensions/keybindingService.js";
import { builtinKeybindings } from "../commands/builtinKeybindings.js";
import { registerBuiltinCommands } from "../commands/registerBuiltinCommands.js";
import { closePanel, createDefaultLayout, openPanel } from "../layout/layoutState.js";
import { panelRegistry } from "../panels/panelRegistry.js";

function makeKeydownEvent(opts: {
  key: string;
  shiftKey?: boolean;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
  target: EventTarget | null;
}): KeyboardEvent {
  const event: any = {
    key: opts.key,
    code: "",
    ctrlKey: Boolean(opts.ctrlKey),
    metaKey: Boolean(opts.metaKey),
    shiftKey: Boolean(opts.shiftKey),
    altKey: Boolean(opts.altKey),
    repeat: false,
    target: opts.target,
    defaultPrevented: false,
  };
  event.preventDefault = () => {
    event.defaultPrevented = true;
  };
  return event as KeyboardEvent;
}

afterEach(() => {
  document.body.innerHTML = "";
});

function createHarness(opts: { zoomDisabled?: boolean; withSecondaryGrid?: boolean } = {}) {
  const gridSecondary = opts.withSecondaryGrid ? `<div id="grid-secondary" tabindex="0"></div>` : "";
  const statusBarExtra = opts.zoomDisabled
    ? `<button type="button" data-testid="open-version-history-panel">Version history</button>`
    : "";
  const zoomDisabledAttr = opts.zoomDisabled ? "disabled" : "";

  document.body.innerHTML = `
      <div id="ribbon">
        <button class="ribbon__tab" role="tab" aria-selected="true">Home</button>
      </div>
      <div id="formula-bar">
        <input data-testid="formula-address" />
      </div>
      <div id="grid" tabindex="0"></div>
      ${gridSecondary}
      <div id="sheet-tabs">
        <button role="tab" aria-selected="true">Sheet1</button>
      </div>
      <div class="statusbar">
        <select data-testid="zoom-control" ${zoomDisabledAttr}>
          <option value="100">100%</option>
        </select>
        ${statusBarExtra}
      </div>
      <button type="button" id="outside-focus">Outside</button>
    `;

  const ribbonTab = document.querySelector<HTMLButtonElement>("#ribbon .ribbon__tab")!;
  const formulaAddress = document.querySelector<HTMLInputElement>('#formula-bar [data-testid="formula-address"]')!;
  const grid = document.querySelector<HTMLElement>("#grid")!;
  const gridSecondaryEl = document.querySelector<HTMLElement>("#grid-secondary");
  const sheetTab = document.querySelector<HTMLButtonElement>('#sheet-tabs button[role="tab"]')!;
  const zoomControl = document.querySelector<HTMLElement>('.statusbar [data-testid="zoom-control"]')!;
  const versionHistoryButton = document.querySelector<HTMLButtonElement>('[data-testid="open-version-history-panel"]');
  const outsideButton = document.querySelector<HTMLButtonElement>("#outside-focus")!;

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
      grid.focus();
    },
  } as any;

  registerBuiltinCommands({ commandRegistry, app, layoutController });

  const contextKeys = new ContextKeyService();
  const service = new KeybindingService({
    commandRegistry,
    contextKeys,
    platform: "other",
    // Match main.ts behavior: allow builtins even when focus is in a text input.
    ignoreInputTargets: "extensions",
  });
  service.setBuiltinKeybindings(builtinKeybindings);

  return {
    service,
    elements: {
      ribbonTab,
      formulaAddress,
      grid,
      gridSecondary: gridSecondaryEl,
      sheetTab,
      zoomControl,
      versionHistoryButton,
      outsideButton,
    },
  };
}

async function dispatchF6(
  service: KeybindingService,
  target: EventTarget | null,
  opts: { shiftKey?: boolean; ctrlKey?: boolean; altKey?: boolean; metaKey?: boolean } = {},
): Promise<{ handled: boolean; event: KeyboardEvent }> {
  const event = makeKeydownEvent({ key: "F6", target, ...opts });
  const handled = await service.dispatchKeydown(event, { allowBuiltins: true, allowExtensions: false });
  return { handled, event };
}

describe("F6 focus cycling keybinding dispatch", () => {
  it("cycles focus across ribbon, formula bar, grid, sheet tabs, and status bar (wraps)", async () => {
    const { service, elements } = createHarness();

    // Start in the grid.
    elements.grid.focus();
    expect(document.activeElement).toBe(elements.grid);

    // Forward cycle (region order: ribbon -> formula bar -> grid -> sheet tabs -> status bar):
    // From grid: sheet tabs -> status bar -> ribbon -> formula bar -> grid.
    let res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.sheetTab);

    res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.zoomControl);

    res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.ribbonTab);

    res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.formulaAddress);

    res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.grid);

    // Wrap.
    res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.sheetTab);

    // Reverse cycle from grid: formula bar -> ribbon -> status bar -> sheet tabs -> grid.
    elements.grid.focus();
    expect(document.activeElement).toBe(elements.grid);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.formulaAddress);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.ribbonTab);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.zoomControl);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.sheetTab);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.grid);
  });

  it("wraps from unknown focus targets (outside known regions)", async () => {
    const { service, elements } = createHarness();

    elements.outsideButton.focus();
    expect(document.activeElement).toBe(elements.outsideButton);

    let res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.ribbonTab);

    elements.outsideButton.focus();
    expect(document.activeElement).toBe(elements.outsideButton);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.zoomControl);
  });

  it("treats the secondary grid root as part of the grid region", async () => {
    const { service, elements } = createHarness({ withSecondaryGrid: true });
    const secondary = elements.gridSecondary;
    expect(secondary).not.toBeNull();
    secondary!.focus();
    expect(document.activeElement).toBe(secondary);

    let res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.sheetTab);

    res = await dispatchF6(service, document.activeElement, { shiftKey: true });
    expect(res.handled).toBe(true);
    // Focus restoration for the grid uses the primary `app.focus()` hook.
    expect(document.activeElement).toBe(elements.grid);
  });

  it("falls back to the first enabled status bar control when the zoom dropdown is disabled", async () => {
    const { service, elements } = createHarness({ zoomDisabled: true });
    expect(elements.versionHistoryButton).not.toBeNull();

    elements.sheetTab.focus();
    expect(document.activeElement).toBe(elements.sheetTab);

    const res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(elements.versionHistoryButton);
  });

  it("ignores modifier chords (Ctrl/Alt/Meta+F6)", async () => {
    const { service, elements } = createHarness();

    elements.grid.focus();
    expect(document.activeElement).toBe(elements.grid);

    const ctrl = await dispatchF6(service, document.activeElement, { ctrlKey: true });
    expect(ctrl.handled).toBe(false);
    expect(ctrl.event.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(elements.grid);

    const alt = await dispatchF6(service, document.activeElement, { altKey: true });
    expect(alt.handled).toBe(false);
    expect(alt.event.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(elements.grid);

    const meta = await dispatchF6(service, document.activeElement, { metaKey: true });
    expect(meta.handled).toBe(false);
    expect(meta.event.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(elements.grid);
  });

  it("does not move focus when executed from within an open <dialog> (focus trap friendly)", async () => {
    const { service } = createHarness();

    const dialog = document.createElement("dialog");
    dialog.setAttribute("open", "true");
    document.body.appendChild(dialog);

    const input = document.createElement("input");
    input.type = "text";
    dialog.appendChild(input);

    input.focus();
    expect(document.activeElement).toBe(input);

    const res = await dispatchF6(service, document.activeElement);
    expect(res.handled).toBe(true);
    expect(document.activeElement).toBe(input);
  });

  it("dispatches through KeybindingService (including from inputs) and respects keybinding barriers", async () => {
    const { service, elements } = createHarness();
    const { formulaAddress, grid } = elements;
    // Use a nested focusable inside the grid to ensure `isInputTarget` stays false.
    const gridFocus = document.createElement("button");
    gridFocus.id = "grid-focus";
    grid.appendChild(gridFocus);

    // Focus in the formula bar input and press F6 -> should move to grid.
    formulaAddress.focus();
    expect(document.activeElement).toBe(formulaAddress);
    const handled = await service.dispatchKeydown(makeKeydownEvent({ key: "F6", target: formulaAddress }), {
      allowBuiltins: true,
      allowExtensions: false,
    });
    expect(handled).toBe(true);
    // The command uses `app.focus()` which defaults to focusing the grid root.
    expect(document.activeElement).toBe(grid);

    // Shift+F6 from grid -> should move back to formula bar.
    const handledBack = await service.dispatchKeydown(makeKeydownEvent({ key: "F6", shiftKey: true, target: grid }), {
      allowBuiltins: true,
      allowExtensions: false,
    });
    expect(handledBack).toBe(true);
    expect(document.activeElement).toBe(formulaAddress);

    // Keybinding barrier suppresses dispatch.
    const barrierRoot = document.createElement("div");
    barrierRoot.setAttribute("data-keybinding-barrier", "true");
    const barrierButton = document.createElement("button");
    barrierButton.textContent = "Barrier";
    barrierRoot.appendChild(barrierButton);
    document.body.appendChild(barrierRoot);
    barrierButton.focus();
    expect(document.activeElement).toBe(barrierButton);

    const blocked = await service.dispatchKeydown(makeKeydownEvent({ key: "F6", target: barrierButton }), {
      allowBuiltins: true,
      allowExtensions: false,
    });
    expect(blocked).toBe(false);
    expect(document.activeElement).toBe(barrierButton);

    barrierRoot.remove();
  });
});
