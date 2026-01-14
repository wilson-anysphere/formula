// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
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

function renderRibbon() {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions: {} }));
  });
  return { container, root };
}

describe("Ribbon UI state overrides (dropdown menu items)", () => {
  it("applies disabled + label overrides to dropdown menu items", async () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const getValuesItem = () =>
      container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.pasteSpecial.values"]');

    const getValuesLabel = () => getValuesItem()?.querySelector(".ribbon-dropdown__label")?.textContent?.trim() ?? "";

    expect(getValuesItem()).toBeInstanceOf(HTMLButtonElement);
    expect(getValuesItem()?.disabled).toBe(false);
    expect(getValuesLabel()).toBe("Paste Values");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: { "clipboard.pasteSpecial.values": "Paste Values (Web)" },
        disabledById: { "clipboard.pasteSpecial.values": true },
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(getValuesItem()?.disabled).toBe(true);
    expect(getValuesLabel()).toBe("Paste Values (Web)");

    act(() => root.unmount());
  });

  it("allows disabledById overrides to re-enable schema-disabled menu items", async () => {
    const { container, root } = renderRibbon();

    const insertTab = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-tab-insert"]');
    expect(insertTab).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      insertTab?.click();
    });

    const pivotTable = container.querySelector<HTMLButtonElement>('[data-command-id="view.insertPivotTable"]');
    expect(pivotTable).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      pivotTable?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
    expect(menu).toBeInstanceOf(HTMLElement);
    if (!menu) throw new Error("Missing dropdown menu");

    const schemaDisabledId = "insert.tables.pivotTable.fromExternal";
    const schemaDisabledItem = menu.querySelector<HTMLButtonElement>(`[data-command-id="${schemaDisabledId}"]`);
    expect(schemaDisabledItem).toBeInstanceOf(HTMLButtonElement);
    expect(schemaDisabledItem?.disabled).toBe(true);

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: { [schemaDisabledId]: false },
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const updated = menu.querySelector<HTMLButtonElement>(`[data-command-id="${schemaDisabledId}"]`);
    expect(updated).toBeInstanceOf(HTMLButtonElement);
    expect(updated?.disabled).toBe(false);
    expect(updated?.hasAttribute("disabled")).toBe(false);
    expect(menu.querySelector(`.ribbon-dropdown__menuitem:not(:disabled)[data-command-id="${schemaDisabledId}"]`)).toBeInstanceOf(
      HTMLButtonElement,
    );

    act(() => root.unmount());
  });
});
