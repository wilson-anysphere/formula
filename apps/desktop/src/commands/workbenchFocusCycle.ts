export type WorkbenchFocusCycleDeps = {
  ribbonRootEl: HTMLElement;
  formulaBarRootEl: HTMLElement;
  gridRootEl: HTMLElement;
  statusBarRootEl: HTMLElement;
  focusGrid: () => void;
  getSecondaryGridRoot?: (() => HTMLElement | null) | null;
  getSheetTabsRoot?: (() => HTMLElement | null) | null;
};

function focusWithoutScroll(el: HTMLElement): void {
  try {
    el.focus({ preventScroll: true });
  } catch {
    el.focus();
  }
}

type FocusRegion = {
  id: "ribbon" | "formulaBar" | "grid" | "sheetTabs" | "statusBar";
  contains: (active: Element | null) => boolean;
  focus: () => void;
};

/**
 * Cycles focus between major spreadsheet workbench regions (Excel-style F6).
 *
 * Region order:
 * ribbon → formula bar → grid → sheet tabs → status bar
 */
export function cycleWorkbenchFocusRegion(deps: WorkbenchFocusCycleDeps, dir: 1 | -1): void {
  if (typeof document === "undefined") return;

  const active = document.activeElement as Element | null;
  // Avoid breaking modal focus traps. KeybindingService normally blocks keybindings inside
  // `data-keybinding-barrier` roots, but this is an additional safeguard for `<dialog>`.
  try {
    if (active?.closest?.("dialog[open]")) return;
  } catch {
    // ignore (non-Element activeElement/test doubles)
  }

  const getSecondaryGridRoot =
    deps.getSecondaryGridRoot ??
    (() => (typeof document === "undefined" ? null : (document.getElementById("grid-secondary") as HTMLElement | null)));
  const getSheetTabsRoot =
    deps.getSheetTabsRoot ??
    (() => (typeof document === "undefined" ? null : (document.getElementById("sheet-tabs") as HTMLElement | null)));

  const focusRibbonRegion = (): void => {
    // Prefer the active ribbon tab (ARIA tablist semantics).
    const activeTab =
      deps.ribbonRootEl.querySelector<HTMLElement>('.ribbon__tab[role="tab"][aria-selected="true"]') ??
      deps.ribbonRootEl.querySelector<HTMLElement>('.ribbon__tab[role="tab"][tabindex="0"]');
    if (activeTab) {
      focusWithoutScroll(activeTab);
      return;
    }
    const firstFocusable = deps.ribbonRootEl.querySelector<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    );
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  };

  const focusFormulaBarRegion = (): void => {
    // Focus the Name Box (address input) so we don't accidentally start formula editing
    // just by cycling focus.
    const address = deps.formulaBarRootEl.querySelector<HTMLElement>('[data-testid="formula-address"]');
    if (address) {
      focusWithoutScroll(address);
      return;
    }
    const firstFocusable = deps.formulaBarRootEl.querySelector<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    );
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  };

  const focusGridRegion = (): void => {
    deps.focusGrid();
  };

  const focusSheetTabsRegion = (): void => {
    const root = getSheetTabsRoot();
    if (!root) return;
    const activeTab =
      root.querySelector<HTMLElement>('button[role="tab"][aria-selected="true"]') ??
      root.querySelector<HTMLElement>('button[role="tab"]');
    if (activeTab) {
      focusWithoutScroll(activeTab);
      return;
    }
    const fallback = root.querySelector<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    );
    if (fallback) focusWithoutScroll(fallback);
  };

  const focusStatusBarRegion = (): void => {
    const zoom = deps.statusBarRootEl.querySelector<HTMLElement>('[data-testid="zoom-control"]:not([disabled])');
    if (zoom) {
      focusWithoutScroll(zoom);
      return;
    }
    const firstFocusable = deps.statusBarRootEl.querySelector<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    );
    if (firstFocusable) focusWithoutScroll(firstFocusable);
  };

  const focusRegions: FocusRegion[] = [
    {
      id: "ribbon",
      contains: (active) => Boolean(active && deps.ribbonRootEl.contains(active)),
      focus: focusRibbonRegion,
    },
    {
      id: "formulaBar",
      contains: (active) => Boolean(active && deps.formulaBarRootEl.contains(active)),
      focus: focusFormulaBarRegion,
    },
    {
      id: "grid",
      contains: (active) =>
        Boolean(active && (deps.gridRootEl.contains(active) || getSecondaryGridRoot()?.contains(active))),
      focus: focusGridRegion,
    },
    {
      id: "sheetTabs",
      contains: (active) => Boolean(active && getSheetTabsRoot()?.contains(active)),
      focus: focusSheetTabsRegion,
    },
    {
      id: "statusBar",
      contains: (active) => Boolean(active && deps.statusBarRootEl.contains(active)),
      focus: focusStatusBarRegion,
    },
  ];

  const currentIndex = focusRegions.findIndex((region) => region.contains(active));
  const nextIndex =
    currentIndex === -1
      ? dir === 1
        ? 0
        : focusRegions.length - 1
      : (currentIndex + dir + focusRegions.length) % focusRegions.length;
  focusRegions[nextIndex]?.focus();
}

