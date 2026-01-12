// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";

import { mountTitlebar } from "./mountTitlebar.js";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

const TEST_TIMEOUT_MS = 30_000;

afterEach(() => {
  document.body.replaceChildren();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("mountTitlebar", () => {
  it("mounts a Titlebar into a container and unmounts cleanly", async () => {
    const container = document.createElement("div");
    container.classList.add("formula-titlebar");
    document.body.appendChild(container);

    let titlebar: ReturnType<typeof mountTitlebar> | null = null;
    await act(async () => {
      titlebar = mountTitlebar(container, {
        appName: "Formula",
        documentName: "— Untitled.xlsx",
        actions: [
          { label: "Share", ariaLabel: "Share document", disabled: true },
          { label: "Comments", ariaLabel: "Open comments" },
        ],
      });
    });

    // mountTitlebar should remove any existing `.formula-titlebar` class from the container
    // so the rendered titlebar isn't nested inside another styled titlebar element.
    expect(container.classList.contains("formula-titlebar")).toBe(false);

    const titlebarRoot = container.querySelector<HTMLElement>('[data-testid="titlebar-component"]');
    expect(titlebarRoot).toBeInstanceOf(HTMLDivElement);
    expect(titlebarRoot?.getAttribute("role")).toBe("banner");
    expect(titlebarRoot?.getAttribute("aria-label")).toBe("Titlebar");

    expect(container.querySelector(".formula-titlebar")).toBeInstanceOf(HTMLDivElement);

    const dragRegion = container.querySelector<HTMLElement>('[data-testid="titlebar-drag-region"]');
    expect(dragRegion).toBeTruthy();
    expect(dragRegion?.getAttribute("data-tauri-drag-region")).not.toBeNull();

    expect(container.querySelector('[data-testid="titlebar-app-name"]')?.textContent).toBe("Formula");
    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.textContent).toBe("Untitled.xlsx");
    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.getAttribute("title")).toBe("Untitled.xlsx");

    // Window controls exist with accessible labels.
    const windowControls = container.querySelector<HTMLElement>('[data-testid="titlebar-window-controls"]');
    expect(windowControls?.getAttribute("role")).toBe("group");
    expect(windowControls?.getAttribute("aria-label")).toBe("Window controls");

    const closeButton = container.querySelector('[data-testid="titlebar-window-close"]');
    const minimizeButton = container.querySelector('[data-testid="titlebar-window-minimize"]');
    const maximizeButton = container.querySelector('[data-testid="titlebar-window-maximize"]');
    expect(closeButton).toBeInstanceOf(HTMLButtonElement);
    expect(minimizeButton).toBeInstanceOf(HTMLButtonElement);
    expect(maximizeButton).toBeInstanceOf(HTMLButtonElement);

    // Without callbacks, window controls should be disabled.
    expect((closeButton as HTMLButtonElement).disabled).toBe(true);
    expect((minimizeButton as HTMLButtonElement).disabled).toBe(true);
    expect((maximizeButton as HTMLButtonElement).disabled).toBe(true);

    // Actions exist with aria labels.
    const actionsToolbar = container.querySelector<HTMLElement>('[data-testid="titlebar-actions"]');
    expect(actionsToolbar).toBeTruthy();
    expect(actionsToolbar?.getAttribute("role")).toBe("toolbar");
    expect(actionsToolbar?.getAttribute("aria-label")).toBe("Titlebar actions");
    const shareButton = container.querySelector<HTMLButtonElement>('[aria-label="Share document"]');
    expect(shareButton).toBeInstanceOf(HTMLButtonElement);
    expect(shareButton?.disabled).toBe(true);
    expect(container.querySelector('[aria-label="Open comments"]')).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      titlebar?.();
    });

    expect(container.childElementCount).toBe(0);

    // Disposing should restore the container class.
    expect(container.classList.contains("formula-titlebar")).toBe(true);

    // It should also be safe to call .dispose() explicitly.
    act(() => {
      titlebar?.dispose();
    });
  }, TEST_TIMEOUT_MS);

  it("wires window control callbacks when provided", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const calls = { close: 0, minimize: 0, maximize: 0 };
    let handle: ReturnType<typeof mountTitlebar> | null = null;
    await act(async () => {
      handle = mountTitlebar(container, {
        actions: [],
        windowControls: {
          onClose: () => {
            calls.close += 1;
          },
          onMinimize: () => {
            calls.minimize += 1;
          },
          onToggleMaximize: () => {
            calls.maximize += 1;
          },
        },
        undoRedo: {
          canUndo: true,
          canRedo: true,
          undoLabel: null,
          redoLabel: null,
        },
      });
    });

    expect(container.querySelector('[data-testid="titlebar-actions"]')).toBeNull();
    expect(container.querySelector('[data-testid="titlebar-quick-access"]')).toBeTruthy();

    // Without callbacks, undo/redo buttons should be disabled even when canUndo/canRedo is true.
    expect(container.querySelector<HTMLButtonElement>('[data-testid="undo"]')?.disabled).toBe(true);
    expect(container.querySelector<HTMLButtonElement>('[data-testid="redo"]')?.disabled).toBe(true);

    const closeButton = container.querySelector<HTMLButtonElement>('[data-testid="titlebar-window-close"]');
    const minimizeButton = container.querySelector<HTMLButtonElement>('[data-testid="titlebar-window-minimize"]');
    const maximizeButton = container.querySelector<HTMLButtonElement>('[data-testid="titlebar-window-maximize"]');

    expect(closeButton?.disabled).toBe(false);
    expect(minimizeButton?.disabled).toBe(false);
    expect(maximizeButton?.disabled).toBe(false);

    closeButton?.click();
    minimizeButton?.click();
    maximizeButton?.click();

    // Double-clicking the drag region should also toggle maximize.
    container.querySelector<HTMLElement>('[data-testid="titlebar-drag-region"]')?.dispatchEvent(
      new MouseEvent("dblclick", { bubbles: true }),
    );

    expect(calls).toEqual({ close: 1, minimize: 1, maximize: 2 });

    act(() => {
      handle?.dispose();
    });
  }, TEST_TIMEOUT_MS);

  it("omits document name span when documentName is empty/separator-only", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    let handle: ReturnType<typeof mountTitlebar> | null = null;
    await act(async () => {
      handle = mountTitlebar(container, { documentName: "—", actions: [] });
    });

    expect(container.querySelector('[data-testid="titlebar-document-name"]')).toBeNull();

    act(() => {
      handle?.dispose();
    });
  }, TEST_TIMEOUT_MS);

  it("does not strip a leading hyphen from a real document name", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    let handle: ReturnType<typeof mountTitlebar> | null = null;
    await act(async () => {
      handle = mountTitlebar(container, { documentName: "-report.xlsx", actions: [] });
    });

    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.textContent).toBe("-report.xlsx");
    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.getAttribute("title")).toBe("-report.xlsx");

    act(() => {
      handle?.dispose();
    });
  }, TEST_TIMEOUT_MS);

  it("supports updating props via handle.update()", async () => {
    const container = document.createElement("div");
    container.classList.add("formula-titlebar");
    document.body.appendChild(container);

    let handle: ReturnType<typeof mountTitlebar> | null = null;
    await act(async () => {
      handle = mountTitlebar(container, { documentName: "Doc 1", actions: [] });
    });
    if (!handle) throw new Error("Expected mountTitlebar to return a handle");
    expect(container.classList.contains("formula-titlebar")).toBe(false);
    expect(container.querySelector('[data-testid="titlebar-actions"]')).toBeNull();
    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.textContent).toBe("Doc 1");

    act(() => {
      handle?.update({
        documentName: "Doc 2",
        actions: [{ id: "share", label: "Share", ariaLabel: "Share document", disabled: true }],
      });
    });

    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.textContent).toBe("Doc 2");
    expect(container.querySelector('[data-testid="titlebar-document-name"]')?.getAttribute("title")).toBe("Doc 2");
    expect(container.querySelector('[data-testid="titlebar-actions"]')).toBeTruthy();
    expect(container.querySelector('[data-testid="titlebar-action-share"]')).toBeTruthy();
    expect(container.querySelector<HTMLButtonElement>('[aria-label="Share document"]')?.disabled).toBe(true);

    act(() => {
      handle?.dispose();
    });
    expect(container.classList.contains("formula-titlebar")).toBe(true);
  }, TEST_TIMEOUT_MS);
});
