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
});
