export type ExcelStyleFocusCycleRegionId = "ribbon" | "formulaBar" | "grid" | "sheetTabs" | "statusBar";

type FocusRegion = {
  id: ExcelStyleFocusCycleRegionId;
  contains: (active: Element | null) => boolean;
  focus: () => void;
};

export type ExcelStyleFocusCycleParams = {
  ribbonRoot: HTMLElement;
  formulaBarRoot: HTMLElement;
  gridRoot: HTMLElement;
  sheetTabsRoot: HTMLElement;
  statusBarRoot: HTMLElement;
  /**
   * Focus handler for the grid region.
   *
   * This is typically `SpreadsheetApp.focus()` so the grid can restore its internal
   * editing affordances (canvas, cell editor overlays, etc).
   */
  focusGrid: () => void;
  /**
   * Optional secondary grid root (split-view).
   */
  gridSecondaryRoot?: HTMLElement | null;
};

export type ExcelStyleFocusCycleDisposer = () => void;

function focusWithoutScroll(el: HTMLElement): void {
  try {
    el.focus({ preventScroll: true });
  } catch {
    el.focus();
  }
}

function findFirstFocusable(root: HTMLElement): HTMLElement | null {
  return root.querySelector<HTMLElement>(
    'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
  );
}

/**
 * Install Excel-style `F6` / `Shift+F6` focus cycling.
 *
 * Excel uses `F6` to cycle focus through major UI regions. In Formula, we mirror the
 * UX so Tab/Shift+Tab can remain dedicated to in-grid navigation.
 *
 * Focus order (forward):
 * ribbon -> formula bar -> grid -> sheet tabs -> status bar -> (wrap)
 */
export function installExcelStyleFocusCycle(params: ExcelStyleFocusCycleParams): ExcelStyleFocusCycleDisposer {
  if (typeof window === "undefined" || typeof document === "undefined") return () => {};

  const { ribbonRoot, formulaBarRoot, gridRoot, sheetTabsRoot, statusBarRoot, focusGrid, gridSecondaryRoot } = params;

  function focusRibbonRegion(): void {
    // Prefer the active ribbon tab (ARIA tablist semantics).
    const activeTab =
      ribbonRoot.querySelector<HTMLElement>('.ribbon__tab[role="tab"][aria-selected="true"]') ??
      ribbonRoot.querySelector<HTMLElement>('.ribbon__tab[role="tab"][tabindex="0"]');
    if (activeTab) {
      focusWithoutScroll(activeTab);
      return;
    }

    const firstFocusable = findFirstFocusable(ribbonRoot);
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  }

  function focusFormulaBarRegion(): void {
    // Focus the Name Box (address input) so we don't accidentally start formula editing
    // just by cycling focus.
    const address = formulaBarRoot.querySelector<HTMLElement>('[data-testid="formula-address"]');
    if (address) {
      focusWithoutScroll(address);
      return;
    }

    const firstFocusable = findFirstFocusable(formulaBarRoot);
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  }

  function focusGridRegion(): void {
    focusGrid();
  }

  function focusSheetTabsRegion(): void {
    const activeTab =
      sheetTabsRoot.querySelector<HTMLElement>('button[role="tab"][aria-selected="true"]') ??
      sheetTabsRoot.querySelector<HTMLElement>('button[role="tab"]');
    if (activeTab) {
      focusWithoutScroll(activeTab);
      return;
    }

    const fallback = findFirstFocusable(sheetTabsRoot);
    if (fallback) focusWithoutScroll(fallback);
  }

  function focusStatusBarRegion(): void {
    const zoom = statusBarRoot.querySelector<HTMLElement>('[data-testid="zoom-control"]:not([disabled])');
    if (zoom) {
      focusWithoutScroll(zoom);
      return;
    }

    const firstFocusable = findFirstFocusable(statusBarRoot);
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  }

  const focusRegions: FocusRegion[] = [
    {
      id: "ribbon",
      contains: (active) => Boolean(active && ribbonRoot.contains(active)),
      focus: focusRibbonRegion,
    },
    {
      id: "formulaBar",
      contains: (active) => Boolean(active && formulaBarRoot.contains(active)),
      focus: focusFormulaBarRegion,
    },
    {
      id: "grid",
      contains: (active) => Boolean(active && (gridRoot.contains(active) || gridSecondaryRoot?.contains(active))),
      focus: focusGridRegion,
    },
    {
      id: "sheetTabs",
      contains: (active) => Boolean(active && sheetTabsRoot.contains(active)),
      focus: focusSheetTabsRegion,
    },
    {
      id: "statusBar",
      contains: (active) => Boolean(active && statusBarRoot.contains(active)),
      focus: focusStatusBarRegion,
    },
  ];

  function cycleFocus(dir: 1 | -1): void {
    const active = document.activeElement as Element | null;
    const currentIndex = focusRegions.findIndex((region) => region.contains(active));
    const nextIndex =
      currentIndex === -1
        ? dir === 1
          ? 0
          : focusRegions.length - 1
        : (currentIndex + dir + focusRegions.length) % focusRegions.length;
    focusRegions[nextIndex]?.focus();
  }

  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.defaultPrevented) return;
    if (event.key !== "F6") return;
    // Avoid collisions with OS/browser-specific modified-F6 shortcuts.
    if (event.ctrlKey || event.metaKey || event.altKey) return;

    // Don't break modal focus traps (find/replace, go to, etc).
    const active = document.activeElement as Element | null;
    if (active?.closest("dialog[open]")) return;

    event.preventDefault();
    cycleFocus(event.shiftKey ? -1 : 1);
  };

  window.addEventListener("keydown", onKeyDown, { capture: true });

  return () => {
    window.removeEventListener("keydown", onKeyDown, { capture: true });
  };
}

