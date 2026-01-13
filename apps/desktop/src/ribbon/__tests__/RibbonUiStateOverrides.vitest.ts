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

describe("Ribbon UI state overrides", () => {
  it("updates toggle aria-pressed when pressed overrides change", () => {
    const { container, root } = renderRibbon();
    const bold = container.querySelector<HTMLButtonElement>('[data-command-id="format.toggleBold"]');
    expect(bold).toBeInstanceOf(HTMLButtonElement);
    expect(bold?.getAttribute("aria-pressed")).toBe("false");

    act(() => {
      setRibbonUiState({
        pressedById: { "format.toggleBold": true },
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(bold?.getAttribute("aria-pressed")).toBe("true");
    act(() => root.unmount());
  });

  it("updates number-format dropdown label via label overrides", () => {
    const { container, root } = renderRibbon();
    const numberFormat = container.querySelector<HTMLButtonElement>('[data-command-id="home.number.numberFormat"]');
    expect(numberFormat).toBeInstanceOf(HTMLButtonElement);

    const labelSpan = () => numberFormat?.querySelector(".ribbon-button__label")?.textContent?.trim() ?? "";
    expect(labelSpan()).toBe("General");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: { "home.number.numberFormat": "Percent" },
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(labelSpan()).toBe("Percent");
    act(() => root.unmount());
  });

  it("includes shortcut hints in the button title when provided", () => {
    const { container, root } = renderRibbon();
    const copy = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.copy"]');
    expect(copy).toBeInstanceOf(HTMLButtonElement);
    expect(copy?.getAttribute("aria-label")).toBe("Copy");
    expect(copy?.getAttribute("title")).toBe("Copy");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: { "clipboard.copy": "Ctrl+C" },
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(copy?.getAttribute("aria-label")).toBe("Copy");
    expect(copy?.getAttribute("title")).toBe("Copy (Ctrl+C)");
    act(() => root.unmount());
  });

  it("sets aria-keyshortcuts when ariaKeyShortcuts overrides change", () => {
    const { container, root } = renderRibbon();
    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);
    expect(paste?.getAttribute("aria-keyshortcuts")).toBeNull();

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: {
          "clipboard.paste": "Control+V",
          "clipboard.pasteSpecial.values": "Alt+V",
        },
      });
    });

    expect(paste?.getAttribute("aria-keyshortcuts")).toBe("Control+V");

    act(() => {
      paste?.click();
    });

    const menuItem = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.pasteSpecial.values"]');
    expect(menuItem).toBeInstanceOf(HTMLButtonElement);
    expect(menuItem?.getAttribute("aria-keyshortcuts")).toBe("Alt+V");

    act(() => root.unmount());
  });
});
