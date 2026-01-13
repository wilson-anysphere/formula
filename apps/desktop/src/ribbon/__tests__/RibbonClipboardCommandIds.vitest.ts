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

function getMenuItems(container: HTMLElement): HTMLButtonElement[] {
  const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
  if (!menu) return [];
  return Array.from(menu.querySelectorAll<HTMLButtonElement>(".ribbon-dropdown__menuitem"));
}

describe("Ribbon clipboard command ids", () => {
  it("uses canonical CommandRegistry clipboard.* ids in Home → Clipboard", async () => {
    const onCommand = vi.fn();
    const { container, root } = renderRibbon({ onCommand });

    const cut = container.querySelector<HTMLButtonElement>('button[aria-label="Cut"]');
    expect(cut).toBeInstanceOf(HTMLButtonElement);
    expect(cut?.dataset.commandId).toBe("clipboard.cut");
    act(() => cut?.click());
    expect(onCommand).toHaveBeenLastCalledWith("clipboard.cut");

    const copy = container.querySelector<HTMLButtonElement>('button[aria-label="Copy"]');
    expect(copy).toBeInstanceOf(HTMLButtonElement);
    expect(copy?.dataset.commandId).toBe("clipboard.copy");
    act(() => copy?.click());
    expect(onCommand).toHaveBeenLastCalledWith("clipboard.copy");

    const paste = container.querySelector<HTMLButtonElement>('[data-testid="ribbon-paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);
    expect(paste?.dataset.commandId).toBe("clipboard.paste");

    // Open Paste dropdown and assert menu item ids.
    await act(async () => {
      paste?.click();
      await Promise.resolve();
    });
    expect(getMenuItems(container).length).toBeGreaterThan(0);

    const pasteItem = getMenuItems(container).find((item) => item.textContent?.trim() === "Paste");
    expect(pasteItem).toBeInstanceOf(HTMLButtonElement);
    expect(pasteItem?.dataset.commandId).toBe("clipboard.paste");
    onCommand.mockClear();
    act(() => pasteItem?.click());
    expect(onCommand).toHaveBeenCalledWith("clipboard.paste");

    // Open Paste dropdown again and validate Paste Values id.
    await act(async () => {
      paste?.click();
      await Promise.resolve();
    });
    const pasteValues = getMenuItems(container).find((item) => item.textContent?.trim() === "Paste Values");
    expect(pasteValues).toBeInstanceOf(HTMLButtonElement);
    expect(pasteValues?.dataset.commandId).toBe("clipboard.pasteSpecial.values");

    // Paste Special dropdown includes a "Paste Special…" menu item wired to clipboard.pasteSpecial.
    const pasteSpecial = Array.from(container.querySelectorAll<HTMLButtonElement>("button.ribbon-button")).find(
      (btn) => btn.getAttribute("aria-label") === "Paste Special",
    );
    expect(pasteSpecial).toBeInstanceOf(HTMLButtonElement);
    expect(pasteSpecial?.dataset.commandId).toBe("clipboard.pasteSpecial");

    await act(async () => {
      pasteSpecial?.click();
      await Promise.resolve();
    });
    const pasteSpecialDialog = getMenuItems(container).find((item) => item.textContent?.trim() === "Paste Special…");
    expect(pasteSpecialDialog).toBeInstanceOf(HTMLButtonElement);
    expect(pasteSpecialDialog?.dataset.commandId).toBe("clipboard.pasteSpecial");

    const transpose = getMenuItems(container).find((item) => item.textContent?.trim() === "Transpose");
    expect(transpose).toBeInstanceOf(HTMLButtonElement);
    expect(transpose?.dataset.commandId).toBe("clipboard.pasteSpecial.transpose");

    act(() => root.unmount());
  });
});

