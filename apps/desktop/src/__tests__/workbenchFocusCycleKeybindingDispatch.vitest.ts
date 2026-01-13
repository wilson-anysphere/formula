// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ContextKeyService } from "../extensions/contextKeys.js";
import { KeybindingService } from "../extensions/keybindingService.js";
import { builtinKeybindings } from "../commands/builtinKeybindings.js";
import { registerBuiltinCommands } from "../commands/registerBuiltinCommands.js";
import { closePanel, createDefaultLayout, openPanel } from "../layout/layoutState.js";
import { panelRegistry } from "../panels/panelRegistry.js";

function makeKeydownEvent(opts: { key: string; shiftKey?: boolean; target: EventTarget | null }): KeyboardEvent {
  const event: any = {
    key: opts.key,
    code: "",
    ctrlKey: false,
    metaKey: false,
    shiftKey: Boolean(opts.shiftKey),
    altKey: false,
    repeat: false,
    target: opts.target,
    defaultPrevented: false,
  };
  event.preventDefault = () => {
    event.defaultPrevented = true;
  };
  return event as KeyboardEvent;
}

describe("F6 focus cycling keybinding dispatch", () => {
  it("dispatches through KeybindingService (including from inputs) and respects keybinding barriers", async () => {
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

    const formulaAddress = document.querySelector<HTMLInputElement>('#formula-bar [data-testid="formula-address"]')!;
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

    const contextKeys = new ContextKeyService();
    const service = new KeybindingService({
      commandRegistry,
      contextKeys,
      platform: "other",
      // Match main.ts behavior: allow builtins even when focus is in a text input.
      ignoreInputTargets: "extensions",
    });
    service.setBuiltinKeybindings(builtinKeybindings);

    // Focus in the formula bar input and press F6 -> should move to grid.
    formulaAddress.focus();
    expect(document.activeElement).toBe(formulaAddress);
    const handled = await service.dispatchKeydown(makeKeydownEvent({ key: "F6", target: formulaAddress }), {
      allowBuiltins: true,
      allowExtensions: false,
    });
    expect(handled).toBe(true);
    expect(document.activeElement).toBe(gridFocus);

    // Shift+F6 from grid -> should move back to formula bar.
    const handledBack = await service.dispatchKeydown(makeKeydownEvent({ key: "F6", shiftKey: true, target: gridFocus }), {
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

