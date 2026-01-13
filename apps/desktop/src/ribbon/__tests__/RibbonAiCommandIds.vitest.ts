// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import type { RibbonActions } from "../ribbonSchema";

afterEach(() => {
  document.body.innerHTML = "";
  try {
    globalThis.localStorage?.removeItem?.("formula.ui.ribbonCollapsed");
  } catch {
    // Ignore storage failures.
  }
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function renderRibbon(actions: RibbonActions = {}) {
  // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
  // https://react.dev/reference/react/act#configuring-your-test-environment
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions }));
  });
  return { container, root };
}

describe("Ribbon AI buttons", () => {
  it("emits canonical command ids when AI ribbon buttons are clicked (matches keybindings/command palette)", () => {
    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onCommand });

    const aiChatButton = container.querySelector<HTMLButtonElement>('[data-testid="open-panel-ai-chat"]');
    expect(aiChatButton).toBeInstanceOf(HTMLButtonElement);
    expect(aiChatButton?.getAttribute("data-command-id")).toBe("view.togglePanel.aiChat");

    act(() => {
      aiChatButton?.click();
    });
    expect(onCommand).toHaveBeenCalledWith("view.togglePanel.aiChat");

    const inlineEditButton = container.querySelector<HTMLButtonElement>('[data-testid="open-inline-ai-edit"]');
    expect(inlineEditButton).toBeInstanceOf(HTMLButtonElement);
    expect(inlineEditButton?.getAttribute("data-command-id")).toBe("ai.inlineEdit");

    act(() => {
      inlineEditButton?.click();
    });
    expect(onCommand).toHaveBeenCalledWith("ai.inlineEdit");

    act(() => root.unmount());
  });
});

