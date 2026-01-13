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
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions }));
  });
  return { container, root };
}

describe("Ribbon icon rendering", () => {
  it("renders SVG icons for schema buttons with iconId (and no icon otherwise)", () => {
    const { container, root } = renderRibbon();

    const bold = container.querySelector<HTMLButtonElement>('[data-command-id="format.toggleBold"]');
    expect(bold).toBeInstanceOf(HTMLButtonElement);
    expect(bold?.querySelector(".ribbon-button__icon svg")).toBeInstanceOf(SVGSVGElement);

    const fontName = container.querySelector<HTMLButtonElement>('[data-command-id="home.font.fontName"]');
    expect(fontName).toBeInstanceOf(HTMLButtonElement);
    expect(fontName?.querySelector(".ribbon-button__icon")).toBeNull();
    expect(fontName?.querySelector(".ribbon-button__label")?.textContent?.trim()).toBe("Font");

    act(() => root.unmount());
  });

  it("does not render icons for dropdown menu items when they have no iconId", async () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await Promise.resolve();
    });

    const pasteValues = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.pasteSpecial.values"]');
    expect(pasteValues).toBeInstanceOf(HTMLButtonElement);

    expect(pasteValues?.querySelector(".ribbon-dropdown__icon")).toBeNull();
    expect(pasteValues?.querySelector(".ribbon-dropdown__label")?.textContent?.trim()).toBe("Paste Values");

    act(() => root.unmount());
  });

  it("renders SVG icons for dropdown menu items when they have an iconId", async () => {
    const { container, root } = renderRibbon();

    const viewTab = Array.from(container.querySelectorAll<HTMLButtonElement>('[role="tab"]')).find(
      (tab) => tab.textContent?.trim() === "View",
    );
    if (!viewTab) throw new Error("Missing View tab");

    act(() => {
      viewTab.click();
    });

    const freezePanes = container.querySelector<HTMLButtonElement>('[data-command-id="view.window.freezePanes"]');
    expect(freezePanes).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      freezePanes?.click();
      await Promise.resolve();
    });

    // This command id also exists as a standalone Home tab button; scope to the open dropdown menu
    // so we assert against the menu item UI (which should include a `.ribbon-dropdown__icon`).
    const openMenus = Array.from(container.querySelectorAll<HTMLElement>(".ribbon-dropdown__menu"));
    const activeMenu = openMenus.at(-1);
    const freezeTopRow = activeMenu?.querySelector<HTMLButtonElement>('[data-command-id="view.freezeTopRow"]') ?? null;
    expect(freezeTopRow).toBeInstanceOf(HTMLButtonElement);
    expect(freezeTopRow?.querySelector(".ribbon-dropdown__icon svg")).toBeInstanceOf(SVGSVGElement);

    act(() => root.unmount());
  });
});
