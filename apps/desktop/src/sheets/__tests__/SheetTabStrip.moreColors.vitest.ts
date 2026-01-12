// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { SheetTabStrip } from "../SheetTabStrip";
import { WorkbookSheetStore } from "../workbookSheetStore";

afterEach(() => {
  document.body.innerHTML = "";
  // React 18 act env flag is set per-test in `renderSheetTabStrip`.
  delete (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
  vi.restoreAllMocks();
});

function renderSheetTabStrip(store: WorkbookSheetStore) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);

  act(() => {
    root.render(
      React.createElement(SheetTabStrip, {
        store,
        activeSheetId: store.listVisible()[0]?.id ?? "",
        onActivateSheet: () => {},
        onAddSheet: () => {},
      }),
    );
  });

  return { container, root };
}

describe("SheetTabStrip tab color picker", () => {
  it("seeds the More Colors picker from a theme-based TabColor", async () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible", tabColor: { theme: 4 } },
    ]);

    const { container, root } = renderSheetTabStrip(store);

    // Flush effects (the hidden <input type="color"> is created in a useEffect).
    await act(async () => {
      await Promise.resolve();
    });

    const tab = container.querySelector<HTMLButtonElement>('[data-testid="sheet-tab-s1"]');
    expect(tab).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tab!.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: 10,
          clientY: 10,
        }),
      );
    });

    const overlay = document.querySelector<HTMLDivElement>('[data-testid="sheet-tab-context-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    const tabColorButton = Array.from(overlay!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "Tab Color",
    );
    expect(tabColorButton).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tabColorButton!.click();
    });

    const submenu = overlay!.querySelector<HTMLDivElement>(".context-menu__submenu");
    expect(submenu).toBeInstanceOf(HTMLDivElement);

    const moreColorsButton = Array.from(submenu!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "More Colors…",
    );
    expect(moreColorsButton).toBeInstanceOf(HTMLButtonElement);

    const colorInput = document.querySelector<HTMLInputElement>('input[type="color"]');
    expect(colorInput).toBeInstanceOf(HTMLInputElement);
    const clickSpy = vi.spyOn(colorInput as HTMLInputElement, "click").mockImplementation(() => {});

    act(() => {
      moreColorsButton!.click();
    });

    // theme:4 (accent1) maps to #5b9bd5 in normalizeExcelColorToCss.
    expect((colorInput as HTMLInputElement).value.toLowerCase()).toBe("#5b9bd5");
    expect(clickSpy).toHaveBeenCalledTimes(1);

    act(() => root.unmount());
  });

  it("stores palette colors as Excel/OOXML ARGB even when CSS tokens are unavailable", async () => {
    // In jsdom there are no real CSS variables; SheetTabStrip should fall back to
    // hardcoded hex values and still be able to convert to ARGB.
    const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    const setTabColorSpy = vi.spyOn(store, "setTabColor");

    const { container, root } = renderSheetTabStrip(store);

    await act(async () => {
      await Promise.resolve();
    });

    const tab = container.querySelector<HTMLButtonElement>('[data-testid="sheet-tab-s1"]');
    expect(tab).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tab!.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: 10,
          clientY: 10,
        }),
      );
    });

    const overlay = document.querySelector<HTMLDivElement>('[data-testid="sheet-tab-context-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    const tabColorButton = Array.from(overlay!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "Tab Color",
    );
    expect(tabColorButton).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tabColorButton!.click();
    });

    const submenu = overlay!.querySelector<HTMLDivElement>(".context-menu__submenu");
    expect(submenu).toBeInstanceOf(HTMLDivElement);

    const redButton = Array.from(submenu!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "Red",
    );
    expect(redButton).toBeInstanceOf(HTMLButtonElement);
    const redFill = redButton!.querySelector("rect")?.getAttribute("fill")?.toLowerCase() ?? null;
    expect(redFill).toBe("#ff0000");

    act(() => {
      redButton!.click();
    });

    // Clicking a submenu action should close the menu synchronously.
    expect(overlay?.hidden).toBe(true);

    await act(async () => {
      // ContextMenu invokes item actions asynchronously. Palette selection awaits the
      // persistence hook (even when undefined), so flush at least two microtasks so
      // the `store.setTabColor(...)` step has run before we assert.
      await Promise.resolve();
      await Promise.resolve();
    });

    expect(setTabColorSpy).toHaveBeenCalled();
    expect(store.getById("s1")?.tabColor?.rgb).toBe("FFFF0000");

    act(() => root.unmount());
  });
});
