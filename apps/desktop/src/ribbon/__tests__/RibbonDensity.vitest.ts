// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";

const STORAGE_KEY = "formula.ui.ribbonCollapsed";

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

function renderRibbon() {
  // React 18+ requires this flag for `act` to behave correctly in non-Jest runners.
  // https://react.dev/reference/react/act#configuring-your-test-environment
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions: {} }));
  });
  const ribbonRoot = container.querySelector<HTMLElement>('[data-testid="ribbon-root"]');
  if (!ribbonRoot) throw new Error("Missing ribbon root");
  return { container, root, ribbonRoot };
}

function setRibbonWidth(ribbonRoot: HTMLElement, width: number): void {
  ribbonRoot.getBoundingClientRect = (() => ({ width })) as any;
}

describe("Ribbon density + collapse toggle", () => {
  it("sets data-density based on ribbon width breakpoints", async () => {
    vi.stubGlobal("localStorage", createStorageMock());
    const { root, ribbonRoot } = renderRibbon();

    // Ensure effects have installed the resize listener.
    await act(async () => {
      await Promise.resolve();
    });

    act(() => {
      setRibbonWidth(ribbonRoot, 1300);
      window.dispatchEvent(new Event("resize"));
    });
    expect(ribbonRoot.dataset.density).toBe("full");

    act(() => {
      setRibbonWidth(ribbonRoot, 1000);
      window.dispatchEvent(new Event("resize"));
    });
    expect(ribbonRoot.dataset.density).toBe("compact");

    act(() => {
      setRibbonWidth(ribbonRoot, 700);
      window.dispatchEvent(new Event("resize"));
    });
    expect(ribbonRoot.dataset.density).toBe("hidden");

    act(() => root.unmount());
  });

  it("persists collapse state to localStorage and overrides full/compact widths", async () => {
    vi.stubGlobal("localStorage", createStorageMock({ [STORAGE_KEY]: "true" }));
    const { root, ribbonRoot, container } = renderRibbon();

    await act(async () => {
      await Promise.resolve();
    });

    act(() => {
      setRibbonWidth(ribbonRoot, 1300);
      window.dispatchEvent(new Event("resize"));
    });
    expect(ribbonRoot.dataset.density).toBe("hidden");

    const toggle = container.querySelector<HTMLButtonElement>(".ribbon__collapse-toggle");
    if (!toggle) throw new Error("Missing collapse toggle");

    act(() => {
      toggle.click();
    });

    expect(globalThis.localStorage.getItem(STORAGE_KEY)).toBe("false");
    expect(ribbonRoot.dataset.density).toBe("full");

    act(() => root.unmount());
  });

  it("supports double-clicking the active tab to toggle collapse", async () => {
    vi.stubGlobal("localStorage", createStorageMock());
    const { root, ribbonRoot, container } = renderRibbon();

    await act(async () => {
      await Promise.resolve();
    });

    act(() => {
      setRibbonWidth(ribbonRoot, 1300);
      window.dispatchEvent(new Event("resize"));
    });
    expect(ribbonRoot.dataset.density).toBe("full");

    const activeTab = container.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]');
    if (!activeTab) throw new Error("Missing active tab");

    act(() => {
      activeTab.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    });
    expect(globalThis.localStorage.getItem(STORAGE_KEY)).toBe("true");
    expect(ribbonRoot.dataset.density).toBe("hidden");

    act(() => {
      activeTab.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    });
    expect(globalThis.localStorage.getItem(STORAGE_KEY)).toBe("false");
    expect(ribbonRoot.dataset.density).toBe("full");

    act(() => root.unmount());
  });
});
