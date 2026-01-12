// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";

afterEach(() => {
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

describe("Ribbon icons", () => {
  it("renders an SVG icon when the schema provides an iconId", () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>("[data-testid=\"ribbon-paste\"]");
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    const svg = paste?.querySelector(".ribbon-button__icon svg");
    expect(svg).toBeInstanceOf(SVGElement);

    act(() => root.unmount());
  });
});
