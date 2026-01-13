/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it } from "vitest";

import { installExcelStyleFocusCycle } from "./installExcelStyleFocusCycle.js";

function dispatchF6(
  opts: { shiftKey?: boolean; altKey?: boolean; ctrlKey?: boolean; metaKey?: boolean } = {},
): KeyboardEvent {
  const target = (document.activeElement as HTMLElement | null) ?? window;
  const event = new KeyboardEvent("keydown", {
    key: "F6",
    code: "F6",
    shiftKey: Boolean(opts.shiftKey),
    altKey: Boolean(opts.altKey),
    ctrlKey: Boolean(opts.ctrlKey),
    metaKey: Boolean(opts.metaKey),
    bubbles: true,
    cancelable: true,
  });
  target.dispatchEvent(event);
  return event;
}

type ShellFixture = {
  dispose: () => void;
  ribbonTab: HTMLButtonElement;
  formulaAddress: HTMLInputElement;
  grid: HTMLDivElement;
  gridSecondary: HTMLDivElement;
  sheetTab: HTMLButtonElement;
  zoomControl: HTMLSelectElement;
  versionHistoryButton: HTMLButtonElement;
  outsideButton: HTMLButtonElement;
};

function setupShell(opts: { zoomDisabled?: boolean; withSecondaryGrid?: boolean } = {}): ShellFixture {
  document.body.innerHTML = "";

  const ribbonRoot = document.createElement("div");
  ribbonRoot.id = "ribbon";
  document.body.appendChild(ribbonRoot);

  const ribbonTab = document.createElement("button");
  ribbonTab.type = "button";
  ribbonTab.className = "ribbon__tab";
  ribbonTab.setAttribute("role", "tab");
  ribbonTab.setAttribute("aria-selected", "true");
  ribbonTab.tabIndex = 0;
  ribbonTab.dataset.testid = "ribbon-tab-home";
  ribbonTab.textContent = "Home";
  ribbonRoot.appendChild(ribbonTab);

  const formulaBarRoot = document.createElement("div");
  formulaBarRoot.id = "formula-bar";
  document.body.appendChild(formulaBarRoot);

  const formulaAddress = document.createElement("input");
  formulaAddress.type = "text";
  formulaAddress.dataset.testid = "formula-address";
  formulaAddress.value = "A1";
  formulaBarRoot.appendChild(formulaAddress);

  const grid = document.createElement("div");
  grid.id = "grid";
  grid.tabIndex = 0;
  document.body.appendChild(grid);

  const gridSecondary = document.createElement("div");
  gridSecondary.id = "grid-secondary";
  gridSecondary.tabIndex = 0;
  if (opts.withSecondaryGrid) {
    document.body.appendChild(gridSecondary);
  }

  const sheetTabsRoot = document.createElement("div");
  sheetTabsRoot.id = "sheet-tabs";
  document.body.appendChild(sheetTabsRoot);

  const sheetTab = document.createElement("button");
  sheetTab.type = "button";
  sheetTab.setAttribute("role", "tab");
  sheetTab.setAttribute("aria-selected", "true");
  sheetTab.tabIndex = 0;
  sheetTab.dataset.testid = "sheet-tab-Sheet1";
  sheetTab.textContent = "Sheet1";
  sheetTabsRoot.appendChild(sheetTab);

  const statusBarRoot = document.createElement("div");
  statusBarRoot.className = "statusbar";
  document.body.appendChild(statusBarRoot);

  const zoomControl = document.createElement("select");
  zoomControl.dataset.testid = "zoom-control";
  zoomControl.disabled = Boolean(opts.zoomDisabled);
  statusBarRoot.appendChild(zoomControl);

  const versionHistoryButton = document.createElement("button");
  versionHistoryButton.type = "button";
  versionHistoryButton.dataset.testid = "open-version-history-panel";
  versionHistoryButton.textContent = "Version history";
  statusBarRoot.appendChild(versionHistoryButton);

  const outsideButton = document.createElement("button");
  outsideButton.type = "button";
  outsideButton.textContent = "Outside";
  document.body.appendChild(outsideButton);

  const dispose = installExcelStyleFocusCycle({
    ribbonRoot,
    formulaBarRoot,
    gridRoot: grid,
    sheetTabsRoot,
    statusBarRoot,
    focusGrid: () => grid.focus(),
    gridSecondaryRoot: opts.withSecondaryGrid ? gridSecondary : null,
  });

  return {
    dispose,
    ribbonTab,
    formulaAddress,
    grid,
    gridSecondary,
    sheetTab,
    zoomControl,
    versionHistoryButton,
    outsideButton,
  };
}

afterEach(() => {
  document.body.innerHTML = "";
});

describe("Excel-style focus cycling (F6 / Shift+F6)", () => {
  it("cycles focus across ribbon -> formula bar -> grid -> sheet tabs -> status bar (shared grid / zoom enabled)", () => {
    const fixture = setupShell({ zoomDisabled: false });
    try {
      fixture.grid.focus();
      expect(document.activeElement).toBe(fixture.grid);

      // Forward cycle (starting from grid): grid -> sheet tabs -> status bar -> ribbon -> formula bar -> grid.
      dispatchF6();
      expect(document.activeElement).toBe(fixture.sheetTab);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.zoomControl);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.ribbonTab);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.formulaAddress);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.grid);

      // Wrap back to sheet tabs.
      dispatchF6();
      expect(document.activeElement).toBe(fixture.sheetTab);

      // Reverse cycle (starting from grid): grid -> formula bar -> ribbon -> status bar -> sheet tabs -> grid.
      fixture.grid.focus();
      expect(document.activeElement).toBe(fixture.grid);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.formulaAddress);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.ribbonTab);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.zoomControl);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.sheetTab);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.grid);
    } finally {
      fixture.dispose();
    }
  });

  it("falls back to the first enabled status bar control when zoom is disabled (legacy grid)", () => {
    const fixture = setupShell({ zoomDisabled: true });
    try {
      fixture.grid.focus();
      expect(document.activeElement).toBe(fixture.grid);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.sheetTab);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.versionHistoryButton);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.ribbonTab);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.formulaAddress);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.grid);

      fixture.grid.focus();
      expect(document.activeElement).toBe(fixture.grid);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.formulaAddress);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.ribbonTab);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.versionHistoryButton);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.sheetTab);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.grid);
    } finally {
      fixture.dispose();
    }
  });

  it("ignores modified F6 chords (Ctrl/Alt/Meta)", () => {
    const fixture = setupShell({ zoomDisabled: false });
    try {
      fixture.grid.focus();
      expect(document.activeElement).toBe(fixture.grid);

      const ctrlEvent = dispatchF6({ ctrlKey: true });
      expect(ctrlEvent.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(fixture.grid);

      const altEvent = dispatchF6({ altKey: true });
      expect(altEvent.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(fixture.grid);

      const metaEvent = dispatchF6({ metaKey: true });
      expect(metaEvent.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(fixture.grid);
    } finally {
      fixture.dispose();
    }
  });

  it("does not cycle focus when a modal dialog is open (focus trap friendly)", () => {
    const fixture = setupShell({ zoomDisabled: false });
    try {
      const dialog = document.createElement("dialog");
      dialog.setAttribute("open", "true");
      document.body.appendChild(dialog);

      const input = document.createElement("input");
      input.type = "text";
      dialog.appendChild(input);

      input.focus();
      expect(document.activeElement).toBe(input);

      const event = dispatchF6();
      expect(event.defaultPrevented).toBe(false);
      expect(document.activeElement).toBe(input);
    } finally {
      fixture.dispose();
    }
  });

  it("starts cycling from the ribbon when focus is outside any known region", () => {
    const fixture = setupShell({ zoomDisabled: false });
    try {
      fixture.outsideButton.focus();
      expect(document.activeElement).toBe(fixture.outsideButton);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.ribbonTab);

      // Reverse from an unknown element should wrap to the last region (status bar).
      fixture.outsideButton.focus();
      expect(document.activeElement).toBe(fixture.outsideButton);

      dispatchF6({ shiftKey: true });
      expect(document.activeElement).toBe(fixture.zoomControl);
    } finally {
      fixture.dispose();
    }
  });

  it("treats the secondary grid root as part of the grid region when cycling focus", () => {
    const fixture = setupShell({ zoomDisabled: false, withSecondaryGrid: true });
    try {
      fixture.gridSecondary.focus();
      expect(document.activeElement).toBe(fixture.gridSecondary);

      dispatchF6();
      expect(document.activeElement).toBe(fixture.sheetTab);

      dispatchF6({ shiftKey: true });
      // The focus cycle uses the primary grid focus handler.
      expect(document.activeElement).toBe(fixture.grid);
    } finally {
      fixture.dispose();
    }
  });
});
