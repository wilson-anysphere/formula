// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import type { RibbonActions } from "../ribbonSchema";

afterEach(() => {
  document.body.innerHTML = "";
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

function getTabs(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(container.querySelectorAll('[role="tab"]')) as HTMLButtonElement[];
}

function getSelectedTab(container: HTMLElement): HTMLButtonElement {
  const selected = getTabs(container).find((tab) => tab.getAttribute("aria-selected") === "true");
  if (!selected) throw new Error("No selected tab");
  return selected;
}

function getActivePanel(container: HTMLElement): HTMLElement {
  const panels = Array.from(container.querySelectorAll('[role="tabpanel"]')) as HTMLElement[];
  const active = panels.find((p) => !p.hidden);
  if (!active) throw new Error("No active tabpanel");
  return active;
}

describe("Ribbon a11y + keyboard navigation", () => {
  it("uses roving tabindex for tabs", () => {
    const { container, root } = renderRibbon();

    const tabs = getTabs(container);
    expect(tabs.length).toBeGreaterThan(0);

    const selected = getSelectedTab(container);
    expect(selected.textContent?.trim()).toBe("Home");

    for (const tab of tabs) {
      if (tab === selected) expect(tab.tabIndex).toBe(0);
      else expect(tab.tabIndex).toBe(-1);
    }

    act(() => root.unmount());
  });

  it("supports Left/Right/Home/End keyboard navigation (wraps around)", () => {
    const { container, root } = renderRibbon();

    // Start on "Home".
    const home = getSelectedTab(container);
    home.focus();
    expect(document.activeElement).toBe(home);

    // Right -> Insert
    act(() => {
      home.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    });
    const insert = getSelectedTab(container);
    expect(insert.textContent?.trim()).toBe("Insert");
    expect(document.activeElement).toBe(insert);

    // End -> Help
    act(() => {
      insert.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
    });
    const help = getSelectedTab(container);
    expect(help.textContent?.trim()).toBe("Help");
    expect(document.activeElement).toBe(help);

    // Right on last tab wraps -> File (index 0 in schema)
    act(() => {
      help.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    });
    const file = getSelectedTab(container);
    expect(file.textContent?.trim()).toBe("File");
    expect(document.activeElement).toBe(file);

    // Home -> File (stays)
    act(() => {
      file.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true }));
    });
    expect(getSelectedTab(container).textContent?.trim()).toBe("File");

    // Left on first tab wraps -> Help
    const fileTab = getSelectedTab(container);
    act(() => {
      fileTab.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true }));
    });
    expect(getSelectedTab(container).textContent?.trim()).toBe("Help");

    act(() => root.unmount());
  });

  it("renders focusable ribbon buttons with aria-labels in the active tabpanel", () => {
    const { container, root } = renderRibbon();

    const panel = getActivePanel(container);
    const buttons = Array.from(panel.querySelectorAll("button"));

    expect(buttons.length).toBeGreaterThan(0);
    for (const btn of buttons) {
      const label = btn.getAttribute("aria-label") ?? "";
      expect(label.trim().length).toBeGreaterThan(0);
    }

    act(() => root.unmount());
  });
});

