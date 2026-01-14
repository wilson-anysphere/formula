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
  // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
  // https://react.dev/reference/react/act#configuring-your-test-environment
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions: {} }));
  });
  return { container, root };
}

describe("Ribbon aria-label overrides", () => {
  it("applies labelById .ariaLabel overrides to dropdown triggers and menus", async () => {
    const { container, root } = renderRibbon();

    const trigger = container.querySelector<HTMLButtonElement>('[data-command-id="home.number.numberFormat"]');
    expect(trigger).toBeInstanceOf(HTMLButtonElement);
    // Baseline comes from the schema (English-only today).
    expect(trigger?.getAttribute("aria-label")).toBe("Number Format");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: { "home.number.numberFormat.ariaLabel": "Zahlenformat" },
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(trigger?.getAttribute("aria-label")).toBe("Zahlenformat");
    expect(trigger?.getAttribute("title")).toBe("Zahlenformat");

    await act(async () => {
      trigger?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
    expect(menu).toBeInstanceOf(HTMLElement);
    expect(menu?.getAttribute("aria-label")).toBe("Zahlenformat");

    act(() => root.unmount());
  });
});

