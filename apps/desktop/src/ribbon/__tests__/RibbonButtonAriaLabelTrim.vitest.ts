// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";

import { RibbonButton } from "../RibbonButton";
import type { RibbonButtonDefinition } from "../ribbonSchema";

let root: Root | null = null;

afterEach(() => {
  if (root) {
    act(() => root?.unmount());
    root = null;
  }
  document.body.innerHTML = "";
});

describe("RibbonButton aria-label trimming", () => {
  it("trims labelOverride when used as the fallback aria-label for icon-only buttons", () => {
    // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
    // https://react.dev/reference/react/act#configuring-your-test-environment
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const container = document.createElement("div");
    document.body.appendChild(container);
    root = createRoot(container);

    const button: RibbonButtonDefinition = {
      id: "test.copy",
      label: "Copy",
      ariaLabel: "Copy (default)",
      size: "icon",
    };

    act(() => {
      root?.render(React.createElement(RibbonButton, { button, labelOverride: "  Copy  " }));
    });

    const el = container.querySelector("button");
    expect(el).toBeInstanceOf(HTMLButtonElement);
    expect(el?.getAttribute("aria-label")).toBe("Copy");
    expect(el?.getAttribute("title")).toBe("Copy");
  });
});
