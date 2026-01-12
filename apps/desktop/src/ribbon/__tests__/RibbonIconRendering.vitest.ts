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
  it("renders SVG icons for mapped command ids (and falls back to glyph icons otherwise)", () => {
    const { container, root } = renderRibbon();

    const bold = container.querySelector<HTMLButtonElement>('[data-command-id="home.font.bold"]');
    expect(bold).toBeInstanceOf(HTMLButtonElement);
    expect(bold?.querySelector(".ribbon-button__icon svg")).toBeInstanceOf(SVGSVGElement);

    const fontName = container.querySelector<HTMLButtonElement>('[data-command-id="home.font.fontName"]');
    expect(fontName).toBeInstanceOf(HTMLButtonElement);
    expect(fontName?.querySelector(".ribbon-button__icon svg")).toBeNull();
    expect(fontName?.querySelector(".ribbon-button__icon")?.textContent?.trim()).toBe("A");

    act(() => root.unmount());
  });

  it("renders fallback glyph icons for dropdown menu items when no SVG mapping exists", async () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="home.clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await Promise.resolve();
    });

    const pasteValues = container.querySelector<HTMLButtonElement>('[data-command-id="home.clipboard.paste.values"]');
    expect(pasteValues).toBeInstanceOf(HTMLButtonElement);

    expect(pasteValues?.querySelector(".ribbon-dropdown__icon svg")).toBeNull();
    expect(pasteValues?.querySelector(".ribbon-dropdown__icon")?.textContent?.trim()).toBe("123");

    act(() => root.unmount());
  });
});

