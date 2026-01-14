// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import type { RibbonActions } from "../ribbonSchema";
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
  it("only sets aria-haspopup=\"menu\" when a dropdown button actually has menu items", () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);
    expect(paste?.getAttribute("aria-haspopup")).toBe("menu");

    const insertTab = getTabs(container).find((tab) => tab.textContent?.trim() === "Insert");
    if (!insertTab) throw new Error("Missing Insert tab");

    act(() => {
      insertTab.click();
    });

    // This command is marked as kind="dropdown" in the schema but does not have an attached menu.
    const pivotChart = container.querySelector<HTMLButtonElement>('[data-command-id="insert.pivotcharts.pivotChart"]');
    expect(pivotChart).toBeInstanceOf(HTMLButtonElement);
    expect(pivotChart?.getAttribute("aria-haspopup")).toBeNull();

    act(() => root.unmount());
  });

  it("closes an open ribbon dropdown menu when pressing Tab from a menuitem", async () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
    expect(menu).toBeInstanceOf(HTMLElement);
    if (!menu) throw new Error("Missing dropdown menu");

    const firstItem = menu.querySelector<HTMLButtonElement>(".ribbon-dropdown__menuitem:not(:disabled)");
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);
    firstItem?.focus();

    await act(async () => {
      firstItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    expect(container.querySelector(".ribbon-dropdown__menu")).toBeNull();

    act(() => root.unmount());
  });

  it("skips disabled dropdown menu items during ArrowDown navigation", async () => {
    const { container, root } = renderRibbon();

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: { "edit.clearContents": true },
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const clearFormatting = container.querySelector<HTMLButtonElement>('[data-command-id="home.font.clearFormatting"]');
    expect(clearFormatting).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      clearFormatting?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
    expect(menu).toBeInstanceOf(HTMLElement);
    if (!menu) throw new Error("Missing dropdown menu");

    const clearFormats = menu.querySelector<HTMLButtonElement>('[data-command-id="format.clearFormats"]');
    const clearContents = menu.querySelector<HTMLButtonElement>('[data-command-id="edit.clearContents"]');
    const clearAll = menu.querySelector<HTMLButtonElement>('[data-command-id="format.clearAll"]');
    expect(clearFormats).toBeInstanceOf(HTMLButtonElement);
    expect(clearContents).toBeInstanceOf(HTMLButtonElement);
    expect(clearAll).toBeInstanceOf(HTMLButtonElement);
    expect(clearContents?.disabled).toBe(true);

    // Focus should land on the first enabled item automatically, but explicitly set it
    // so the test doesn't depend on timing quirks.
    clearFormats?.focus();
    expect(document.activeElement).toBe(clearFormats);

    act(() => {
      clearFormats?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });

    // ArrowDown should skip the disabled middle item and land on the next enabled one.
    expect(document.activeElement).toBe(clearAll);

    act(() => root.unmount());
  });

  it("wires tabs to tabpanels via aria-controls / aria-labelledby", () => {
    const { container, root } = renderRibbon();

    const tablist = container.querySelector('[role="tablist"]');
    expect(tablist).toBeInstanceOf(HTMLElement);

    const tabs = getTabs(container);
    expect(tabs.length).toBeGreaterThan(0);

    for (const tab of tabs) {
      const controls = tab.getAttribute("aria-controls");
      expect(controls).toBeTruthy();

      const panel = container.querySelector(`#${controls}`) as HTMLElement | null;
      expect(panel).toBeInstanceOf(HTMLElement);
      if (!panel) continue;

      expect(panel.getAttribute("role")).toBe("tabpanel");
      expect(panel.getAttribute("aria-labelledby")).toBe(tab.id);
      expect(panel.tabIndex).toBe(-1);
    }

    act(() => root.unmount());
  });

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

  it("supports Left/Right/Home/End keyboard navigation (wraps around)", async () => {
    const { container, root } = renderRibbon();

    // Start on "Home".
    const home = getSelectedTab(container);
    home.focus();
    expect(document.activeElement).toBe(home);

    // Right -> Insert
    await act(async () => {
      home.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    const insert = getSelectedTab(container);
    expect(insert.textContent?.trim()).toBe("Insert");
    expect(document.activeElement).toBe(insert);

    // End -> Help
    await act(async () => {
      insert.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    const help = getSelectedTab(container);
    expect(help.textContent?.trim()).toBe("Help");
    expect(document.activeElement).toBe(help);

    // Right on last tab wraps -> File (index 0 in schema)
    await act(async () => {
      help.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    const file = getSelectedTab(container);
    expect(file.textContent?.trim()).toBe("File");
    expect(document.activeElement).toBe(file);

    // Home -> File (stays)
    await act(async () => {
      file.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    expect(getSelectedTab(container).textContent?.trim()).toBe("File");

    // Left on first tab wraps -> Help
    const fileTab = getSelectedTab(container);
    await act(async () => {
      fileTab.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true }));
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    expect(getSelectedTab(container).textContent?.trim()).toBe("Help");

    act(() => root.unmount());
  });

  it("moves focus into the active tabpanel when pressing Tab on the active tab", () => {
    const { container, root } = renderRibbon();

    const home = getSelectedTab(container);
    home.focus();
    expect(document.activeElement).toBe(home);

    const panel = getActivePanel(container);
    const firstButton = panel.querySelector<HTMLButtonElement>("button:not(:disabled)");
    expect(firstButton).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      home.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true }));
    });

    expect(document.activeElement).toBe(firstButton);

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
