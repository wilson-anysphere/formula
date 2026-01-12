// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";

afterEach(() => {
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function renderRibbon(actions: any, initialTabId: string) {
  // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
  // https://react.dev/reference/react/act#configuring-your-test-environment
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);

  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions, initialTabId }));
  });

  return { container, root };
}

describe("Ribbon theme selector", () => {
  it("renders in the View tab and emits theme command ids", async () => {
    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onCommand }, "view");

    // Let effects run (e.g. density + event listeners).
    await act(async () => {
      await Promise.resolve();
    });

    const trigger = container.querySelector<HTMLButtonElement>('[data-testid="theme-selector"]');
    expect(trigger).toBeTruthy();

    act(() => {
      trigger!.click();
    });

    const dark = container.querySelector<HTMLButtonElement>('[data-testid="theme-option-dark"]');
    expect(dark).toBeTruthy();

    act(() => {
      dark!.click();
    });

    expect(onCommand).toHaveBeenCalledWith("view.appearance.theme.dark");

    act(() => root.unmount());
  });
});

