// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import type { RibbonActions } from "../ribbonSchema";
import { setRibbonUiState } from "../ribbonUiState";

afterEach(() => {
  act(() => {
    setRibbonUiState({
      pressedById: Object.create(null),
      labelById: Object.create(null),
      disabledById: Object.create(null),
      shortcutById: Object.create(null),
      ariaKeyShortcutsById: Object.create(null),
    });
  });
  document.body.innerHTML = "";
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function renderRibbon(actions: RibbonActions) {
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

function enableFileActions(): RibbonActions {
  return {
    fileActions: {
      newWorkbook: vi.fn(),
      openWorkbook: vi.fn(),
      saveWorkbook: vi.fn(),
      saveWorkbookAs: vi.fn(),
      toggleAutoSave: vi.fn(),
      versionHistory: vi.fn(),
      branchManager: vi.fn(),
      print: vi.fn(),
      pageSetup: vi.fn(),
      closeWindow: vi.fn(),
      quit: vi.fn(),
    },
  };
}

function openFileBackstage(container: HTMLElement) {
  const fileTab = Array.from(container.querySelectorAll<HTMLButtonElement>('[role="tab"]')).find(
    (tab) => tab.textContent?.trim() === "File",
  );
  if (!fileTab) throw new Error("Missing File tab");

  act(() => {
    fileTab.click();
  });

  const overlay = container.querySelector<HTMLElement>(".ribbon-backstage-overlay");
  if (!overlay) throw new Error("Expected File backstage overlay to render");
  return { overlay, fileTab };
}

describe("FileBackstage", () => {
  it("opens from the File tab and focuses the first enabled action", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    const { container, root } = renderRibbon(enableFileActions());

    const { overlay } = openFileBackstage(container);
    expect(overlay.getAttribute("role")).toBe("dialog");

    const firstItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-new"]');
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);
    expect(document.activeElement).toBe(firstItem);

    act(() => root.unmount());
  });

  it("uses RibbonUiState shortcut overrides for workbench file commands", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: { "workbench.newWorkbook": "Ctrl+Alt+N" },
        ariaKeyShortcutsById: { "workbench.newWorkbook": "Control+Alt+N" },
      });
    });

    const { container, root } = renderRibbon(enableFileActions());
    const { overlay } = openFileBackstage(container);

    const firstItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-new"]');
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);
    expect(firstItem?.getAttribute("aria-keyshortcuts")).toBe("Control+Alt+N");
    expect(firstItem?.querySelector(".ribbon-backstage__hint")?.textContent?.trim()).toBe("Ctrl+Alt+N");

    act(() => root.unmount());
  });

  it("supports ArrowUp/ArrowDown focus movement within the backstage menu", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    const { container, root } = renderRibbon(enableFileActions());
    const { overlay } = openFileBackstage(container);

    const newItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-new"]');
    const openItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-open"]');
    const saveItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-save"]');
    const saveAsItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-save-as"]');
    const autoSaveItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-auto-save"]');
    const versionHistoryItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-version-history"]');
    const branchManagerItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-branch-manager"]');
    const quitItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-quit"]');
    expect(newItem).toBeInstanceOf(HTMLButtonElement);
    expect(openItem).toBeInstanceOf(HTMLButtonElement);
    expect(saveItem).toBeInstanceOf(HTMLButtonElement);
    expect(saveAsItem).toBeInstanceOf(HTMLButtonElement);
    expect(autoSaveItem).toBeInstanceOf(HTMLButtonElement);
    expect(versionHistoryItem).toBeInstanceOf(HTMLButtonElement);
    expect(branchManagerItem).toBeInstanceOf(HTMLButtonElement);
    expect(quitItem).toBeInstanceOf(HTMLButtonElement);

    expect(document.activeElement).toBe(newItem);

    act(() => {
      newItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(openItem);

    act(() => {
      openItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    });
    expect(document.activeElement).toBe(newItem);

    act(() => {
      newItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(openItem);

    act(() => {
      openItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(saveItem);

    act(() => {
      saveItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(saveAsItem);

    act(() => {
      saveAsItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(autoSaveItem);

    act(() => {
      autoSaveItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(versionHistoryItem);

    act(() => {
      versionHistoryItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });
    expect(document.activeElement).toBe(branchManagerItem);

    act(() => {
      branchManagerItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    });
    expect(document.activeElement).toBe(versionHistoryItem);

    // Wrap around on ArrowUp from the first item.
    act(() => {
      newItem?.focus();
      newItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    });
    expect(document.activeElement).toBe(quitItem);

    act(() => root.unmount());
  });

  it("invokes Version History + Branch Manager backstage actions", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    const actions = enableFileActions();
    const { container, root } = renderRibbon(actions);

    const { overlay } = openFileBackstage(container);
    const versionHistoryItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-version-history"]');
    expect(versionHistoryItem).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      versionHistoryItem?.click();
    });

    expect(actions.fileActions?.versionHistory).toHaveBeenCalledTimes(1);
    expect(container.querySelector(".ribbon-backstage-overlay")).toBeNull();

    const { overlay: overlay2 } = openFileBackstage(container);
    const branchManagerItem = overlay2.querySelector<HTMLButtonElement>('[data-testid="file-branch-manager"]');
    expect(branchManagerItem).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      branchManagerItem?.click();
    });

    expect(actions.fileActions?.branchManager).toHaveBeenCalledTimes(1);
    expect(container.querySelector(".ribbon-backstage-overlay")).toBeNull();

    act(() => root.unmount());
  });

  it("traps Tab navigation inside the backstage menu", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    const { container, root } = renderRibbon(enableFileActions());
    const { overlay } = openFileBackstage(container);

    const firstItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-new"]');
    const quitItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-quit"]');
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);
    expect(quitItem).toBeInstanceOf(HTMLButtonElement);

    quitItem?.focus();
    expect(document.activeElement).toBe(quitItem);

    act(() => {
      quitItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true }));
    });
    expect(document.activeElement).toBe(firstItem);

    act(() => {
      firstItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, bubbles: true }));
    });
    expect(document.activeElement).toBe(quitItem);

    act(() => root.unmount());
  });

  it("closes on Escape and restores focus to the last non-File tab", () => {
    vi.stubGlobal("requestAnimationFrame", ((cb: FrameRequestCallback) => {
      cb(0);
      return 0 as any;
    }) as any);

    const { container, root } = renderRibbon(enableFileActions());
    const { overlay } = openFileBackstage(container);

    const firstItem = overlay.querySelector<HTMLButtonElement>('[data-testid="file-new"]');
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      firstItem?.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });

    expect(container.querySelector(".ribbon-backstage-overlay")).toBeNull();
    const selected = container.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]');
    expect(selected?.textContent?.trim()).toBe("Home");
    expect(document.activeElement).toBe(selected);

    act(() => root.unmount());
  });
});
