// @vitest-environment jsdom

import { act } from "react";
import { describe, expect, it } from "vitest";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

describe("mountTitlebar", () => {
  it("mounts a Titlebar into a container and unmounts cleanly", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const { mountTitlebar } = await import("./mountTitlebar.js");

    let dispose: (() => void) | null = null;
    await act(async () => {
      dispose = mountTitlebar(container, {
        appName: "Formula",
        documentName: "Untitled.xlsx",
        actions: [
          { label: "Share", ariaLabel: "Share document" },
          { label: "Comments", ariaLabel: "Open comments" },
        ],
      });
    });

    expect(container.querySelector(".formula-titlebar")).toBeInstanceOf(HTMLDivElement);
    expect(container.querySelector(".formula-titlebar--component")).toBeInstanceOf(HTMLDivElement);

    const dragRegion = container.querySelector<HTMLElement>(".formula-titlebar__drag-region");
    expect(dragRegion).toBeTruthy();
    expect(dragRegion?.getAttribute("data-tauri-drag-region")).not.toBeNull();

    expect(container.querySelector(".formula-titlebar__app-name")?.textContent).toBe("Formula");
    expect(container.querySelector(".formula-titlebar__document-name")?.textContent).toBe("Untitled.xlsx");

    // Window controls exist with accessible labels.
    expect(container.querySelector('[aria-label="Close window"]')).toBeInstanceOf(HTMLButtonElement);
    expect(container.querySelector('[aria-label="Minimize window"]')).toBeInstanceOf(HTMLButtonElement);
    expect(container.querySelector('[aria-label="Maximize window"]')).toBeInstanceOf(HTMLButtonElement);

    // Actions exist with aria labels.
    expect(container.querySelector('[aria-label="Share document"]')).toBeInstanceOf(HTMLButtonElement);
    expect(container.querySelector('[aria-label="Open comments"]')).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      dispose?.();
    });

    expect(container.childElementCount).toBe(0);
  });
});

