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

function createStorageMock(initial: Record<string, string> = {}): Storage {
  const map = new Map<string, string>(Object.entries(initial));
  return {
    get length() {
      return map.size;
    },
    clear() {
      map.clear();
    },
    getItem(key: string) {
      return map.get(key) ?? null;
    },
    key(index: number) {
      return Array.from(map.keys())[index] ?? null;
    },
    removeItem(key: string) {
      map.delete(key);
    },
    setItem(key: string, value: string) {
      map.set(key, String(value));
    },
  };
}

function renderRibbon(actions: React.ComponentProps<typeof Ribbon>["actions"]) {
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

describe("Ribbon Pictures test ids", () => {
  it("renders stable test ids for Insert → Illustrations → Pictures dropdown + menu items", async () => {
    vi.stubGlobal("localStorage", createStorageMock());

    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onCommand });

    const insertTab = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-tab-insert"]');
    expect(insertTab).toBeInstanceOf(HTMLButtonElement);
    if (!insertTab) throw new Error("Missing Insert tab");

    await act(async () => {
      insertTab.click();
      await Promise.resolve();
    });

    const picturesDropdown = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-insert-pictures"]');
    expect(picturesDropdown).toBeInstanceOf(HTMLButtonElement);
    if (!picturesDropdown) throw new Error("Missing Insert → Pictures dropdown");

    const panel = picturesDropdown.closest<HTMLElement>('[role="tabpanel"]');
    expect(panel).toBeInstanceOf(HTMLElement);
    expect(panel?.hidden).toBe(false);

    await act(async () => {
      picturesDropdown.click();
      await Promise.resolve();
    });

    const thisDevice = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-insert-pictures-this-device"]');
    const stockImages = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-insert-pictures-stock-images"]');
    const onlinePictures = container.querySelector<HTMLButtonElement>(
      '[data-testid="ribbon-insert-pictures-online-pictures"]',
    );

    expect(thisDevice).toBeInstanceOf(HTMLButtonElement);
    expect(stockImages).toBeInstanceOf(HTMLButtonElement);
    expect(onlinePictures).toBeInstanceOf(HTMLButtonElement);

    if (!thisDevice) throw new Error("Missing Insert → Pictures → This Device menu item");

    await act(async () => {
      thisDevice.click();
      await Promise.resolve();
    });

    expect(onCommand).toHaveBeenCalledWith("insert.illustrations.pictures.thisDevice");

    act(() => root.unmount());
  });
});

