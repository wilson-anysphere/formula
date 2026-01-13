// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import type { RibbonActions } from "../ribbonSchema";

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

describe("Ribbon toggle activation", () => {
  it("invokes onToggle (not onCommand) for toggle buttons; buttons/menu items still call onCommand once", async () => {
    vi.stubGlobal("localStorage", createStorageMock());

    const onToggle = vi.fn();
    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onToggle, onCommand });

    // Let effects run (e.g. density + event listeners).
    await act(async () => {
      await Promise.resolve();
    });

    const bold = container.querySelector<HTMLButtonElement>('[data-command-id="format.toggleBold"]');
    expect(bold).toBeInstanceOf(HTMLButtonElement);

    // First click: toggles on.
    await act(async () => {
      bold?.click();
      await Promise.resolve();
    });

    expect(onToggle).toHaveBeenCalledTimes(1);
    expect(onToggle).toHaveBeenCalledWith("format.toggleBold", true);
    expect(onCommand).not.toHaveBeenCalled();
    expect(bold?.getAttribute("aria-pressed")).toBe("true");

    // Second click: toggles off.
    await act(async () => {
      bold?.click();
      await Promise.resolve();
    });

    expect(onToggle).toHaveBeenCalledTimes(2);
    expect(onToggle).toHaveBeenNthCalledWith(2, "format.toggleBold", false);
    expect(onCommand).not.toHaveBeenCalled();
    expect(bold?.getAttribute("aria-pressed")).toBe("false");

    const cut = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.cut"]');
    expect(cut).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      cut?.click();
      await Promise.resolve();
    });

    expect(onCommand).toHaveBeenCalledTimes(1);
    expect(onCommand).toHaveBeenCalledWith("clipboard.cut");
    expect(onToggle).toHaveBeenCalledTimes(2);

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await Promise.resolve();
    });

    // Opening the dropdown should not dispatch a command.
    expect(onCommand).toHaveBeenCalledTimes(1);

    const pasteValues = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.pasteSpecial.values"]');
    expect(pasteValues).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      pasteValues?.click();
      await Promise.resolve();
    });

    expect(onCommand).toHaveBeenCalledTimes(2);
    expect(onCommand).toHaveBeenNthCalledWith(2, "clipboard.pasteSpecial.values");
    expect(onToggle).toHaveBeenCalledTimes(2);

    act(() => root.unmount());
  });
});
