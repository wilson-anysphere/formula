import React from "react";

import type { RibbonActions, RibbonButtonDefinition, RibbonSchema } from "./ribbonSchema.js";
import { defaultRibbonSchema } from "./ribbonSchema.js";
import { RibbonGroup } from "./RibbonGroup.js";
import { getRibbonUiStateSnapshot, subscribeRibbonUiState } from "./ribbonUiState.js";

import "../styles/ribbon.css";

type RibbonDensity = "full" | "compact" | "hidden";

const RIBBON_COLLAPSED_STORAGE_KEY = "formula.ui.ribbonCollapsed";

function readRibbonCollapsedFromStorage(): boolean {
  try {
    const value = localStorage.getItem(RIBBON_COLLAPSED_STORAGE_KEY);
    return value === "true" || value === "1";
  } catch {
    return false;
  }
}

function writeRibbonCollapsedToStorage(collapsed: boolean): void {
  try {
    localStorage.setItem(RIBBON_COLLAPSED_STORAGE_KEY, collapsed ? "true" : "false");
  } catch {
    // Ignore storage errors (e.g. disabled storage).
  }
}

function densityFromWidth(width: number): RibbonDensity {
  // In non-layout environments (tests/SSR) `getBoundingClientRect()` is often 0.
  // Default to "full" rather than collapsing the ribbon in those cases.
  if (!Number.isFinite(width) || width <= 0) return "full";
  if (width < 800) return "hidden";
  if (width < 1200) return "compact";
  return "full";
}

function computeInitialPressed(schema: RibbonSchema): Record<string, boolean> {
  const pressed: Record<string, boolean> = {};
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        if (button.kind === "toggle") {
          pressed[button.id] = Boolean(button.defaultPressed);
        }
      }
    }
  }
  return pressed;
}

function computeToggleIds(schema: RibbonSchema): Set<string> {
  const ids = new Set<string>();
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        if (button.kind === "toggle") ids.add(button.id);
      }
    }
  }
  return ids;
}

export interface RibbonProps {
  actions: RibbonActions;
  schema?: RibbonSchema;
  initialTabId?: string;
}

export function Ribbon({ actions, schema = defaultRibbonSchema, initialTabId }: RibbonProps) {
  const rootRef = React.useRef<HTMLDivElement | null>(null);

  const tabs = schema.tabs;
  const defaultTabId = React.useMemo(() => {
    if (initialTabId && tabs.some((tab) => tab.id === initialTabId)) {
      return initialTabId;
    }

    return tabs.find((tab) => tab.id === "home")?.id ?? tabs[0]?.id ?? "home";
  }, [initialTabId, tabs]);

  const [activeTabId, setActiveTabId] = React.useState<string>(defaultTabId);
  const [pressedById, setPressedById] = React.useState<Record<string, boolean>>(() => computeInitialPressed(schema));
  const pressedByIdRef = React.useRef<Record<string, boolean>>(pressedById);

  const uiState = React.useSyncExternalStore(
    subscribeRibbonUiState,
    getRibbonUiStateSnapshot,
    getRibbonUiStateSnapshot,
  );
  const [ribbonWidth, setRibbonWidth] = React.useState<number>(0);
  const [userCollapsed, setUserCollapsed] = React.useState<boolean>(() => readRibbonCollapsedFromStorage());

  const toggleUserCollapsed = React.useCallback(() => {
    setUserCollapsed((prev) => !prev);
  }, []);

  React.useEffect(() => {
    // Persist the user-controlled "Collapse Ribbon" toggle.
    writeRibbonCollapsedToStorage(userCollapsed);
  }, [userCollapsed]);

  const tabButtonRefs = React.useRef<Record<string, HTMLButtonElement | null>>({});

  React.useEffect(() => {
    pressedByIdRef.current = pressedById;
  }, [pressedById]);

  const mergedPressedById = React.useMemo(() => {
    // Spread here is fine: there are only ~tens of ribbon controls.
    return { ...pressedById, ...uiState.pressedById };
  }, [pressedById, uiState.pressedById]);

  React.useEffect(() => {
    const root = rootRef.current;
    if (!root) return;

    const update = () => {
      const width = root.getBoundingClientRect().width;
      setRibbonWidth((prev) => (prev === width ? prev : width));
    };

    update();

    if (typeof ResizeObserver !== "undefined") {
      const observer = new ResizeObserver((entries) => {
        const entry = entries[entries.length - 1];
        const width = entry?.contentRect?.width ?? root.getBoundingClientRect().width;
        setRibbonWidth((prev) => (prev === width ? prev : width));
      });
      observer.observe(root);
      return () => observer.disconnect();
    }

    window.addEventListener("resize", update);
    return () => window.removeEventListener("resize", update);
  }, []);

  React.useEffect(() => {
    // Keep internal toggle state in sync with schema changes (e.g. when tabs/groups
    // are swapped out by the host app).
    const toggleIds = computeToggleIds(schema);
    const defaults = computeInitialPressed(schema);
    setPressedById((prev) => {
      const next: Record<string, boolean> = {};
      for (const id of toggleIds) {
        next[id] = Object.prototype.hasOwnProperty.call(prev, id) ? prev[id]! : defaults[id]!;
      }
      return next;
    });
  }, [schema]);

  React.useEffect(() => {
    if (!tabs.some((tab) => tab.id === activeTabId)) {
      setActiveTabId(defaultTabId);
    }
  }, [activeTabId, defaultTabId, tabs]);

  const activateButton = React.useCallback(
    (button: RibbonButtonDefinition) => {
      const kind = button.kind ?? "button";

      if (kind === "toggle") {
        const currentPressed = Object.prototype.hasOwnProperty.call(uiState.pressedById, button.id)
          ? uiState.pressedById[button.id]
          : pressedByIdRef.current[button.id];
        const nextPressed = !currentPressed;
        setPressedById((prev) => ({ ...prev, [button.id]: nextPressed }));
        actions.onToggle?.(button.id, nextPressed);
        actions.onCommand?.(button.id);
        return;
      }

      actions.onCommand?.(button.id);
    },
    [actions, uiState.pressedById],
  );

  const selectTabByIndex = React.useCallback(
    (nextIndex: number) => {
      const tab = tabs[nextIndex];
      if (!tab) return;
      setActiveTabId(tab.id);
      actions.onTabChange?.(tab.id);
      const button = tabButtonRefs.current[tab.id];
      // Prefer keeping the viewport stable while still ensuring the newly-focused
      // tab is visible within the horizontally-scrollable tab strip.
      button?.focus({ preventScroll: true });
      button?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    },
    [actions, tabs],
  );

  const focusFirstControl = React.useCallback((tabId: string) => {
    const panel = document.getElementById(`ribbon-panel-${tabId}`);
    if (!panel) return;
    const first = panel.querySelector<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    );
    if (first) {
      first.focus();
      return;
    }
    // Ensure Tab from the tab strip doesn't become a focus trap when a panel has
    // no tabbable controls (e.g. all commands are disabled).
    (panel as HTMLElement).focus?.();
  }, []);

  const responsiveDensity = React.useMemo(() => densityFromWidth(ribbonWidth), [ribbonWidth]);
  const density: RibbonDensity = responsiveDensity === "hidden" ? "hidden" : userCollapsed ? "hidden" : responsiveDensity;
  const contentVisible = density !== "hidden";
  const collapseLabel = userCollapsed ? "Expand ribbon" : "Collapse ribbon";

  return (
    <div
      className="ribbon"
      data-testid="ribbon-root"
      data-density={density}
      ref={rootRef}
      onKeyDownCapture={(event) => {
        if (event.key !== "Escape" && event.key !== "ArrowUp") return;
        const target = event.target as HTMLElement | null;
        if (!target) return;
        if (target.closest(".ribbon__tabs")) return;
        if (!target.closest(".ribbon__content")) return;
        const dropdownRoot = target.closest(".ribbon-dropdown");
        // If a dropdown menu is open (or we're on a dropdown trigger), let the
        // dropdown own Escape/ArrowUp handling (menu closes on Escape; ArrowUp may
        // open/focus within the menu).
        if (dropdownRoot) {
          const menuOpen = dropdownRoot.querySelector(".ribbon-dropdown__menu");
          if (menuOpen) return;
          if (event.key === "ArrowUp") return;
        }
        event.preventDefault();
        tabButtonRefs.current[activeTabId]?.focus();
      }}
    >
      <div className="ribbon__tabstrip">
        <div className="ribbon__tabs" role="tablist" aria-label="Ribbon tabs" aria-orientation="horizontal">
          {tabs.map((tab, index) => {
            const isActive = tab.id === activeTabId;
            const isFile = Boolean(tab.isFile);
            return (
              <button
                key={tab.id}
                type="button"
                className={[
                  "ribbon__tab",
                  isActive ? "is-active" : null,
                  isFile ? "ribbon__tab--file" : null,
                ]
                  .filter(Boolean)
                  .join(" ")}
                role="tab"
                id={`ribbon-tab-${tab.id}`}
                aria-selected={isActive}
                aria-controls={`ribbon-panel-${tab.id}`}
                tabIndex={isActive ? 0 : -1}
                ref={(el) => {
                  tabButtonRefs.current[tab.id] = el;
                }}
                onClick={() => {
                  setActiveTabId(tab.id);
                  actions.onTabChange?.(tab.id);
                }}
                onKeyDown={(event) => {
                  if (event.key === "ArrowRight") {
                    event.preventDefault();
                    selectTabByIndex((index + 1) % tabs.length);
                    return;
                  }
                  if (event.key === "ArrowLeft") {
                    event.preventDefault();
                    selectTabByIndex((index - 1 + tabs.length) % tabs.length);
                    return;
                  }
                  if (event.key === "Home") {
                    event.preventDefault();
                    selectTabByIndex(0);
                    return;
                  }
                  if (event.key === "End") {
                    event.preventDefault();
                    selectTabByIndex(tabs.length - 1);
                    return;
                  }
                  if (event.key === "ArrowDown") {
                    if (!contentVisible) return;
                    event.preventDefault();
                    if (!isActive) {
                      setActiveTabId(tab.id);
                      actions.onTabChange?.(tab.id);
                      requestAnimationFrame(() => focusFirstControl(tab.id));
                      return;
                    }
                    focusFirstControl(tab.id);
                    return;
                  }
                  if (event.key === "Tab" && !event.shiftKey) {
                    if (!contentVisible) return;
                    // Excel-style: Tab from the active tab moves focus into the tab panel.
                    // Shift+Tab should keep browser default behavior (move focus backwards).
                    event.preventDefault();
                    if (!isActive) {
                      setActiveTabId(tab.id);
                      actions.onTabChange?.(tab.id);
                      requestAnimationFrame(() => focusFirstControl(tab.id));
                      return;
                    }
                    focusFirstControl(tab.id);
                    return;
                  }
                }}
                onDoubleClick={() => {
                  if (isActive) toggleUserCollapsed();
                }}
              >
                {tab.label}
              </button>
            );
          })}
        </div>
        <div className="ribbon__tabstrip-right">
          <button
            type="button"
            className="ribbon__collapse-toggle"
            aria-label={collapseLabel}
            title={collapseLabel}
            aria-pressed={userCollapsed}
            onClick={toggleUserCollapsed}
          >
            {userCollapsed ? "▾" : "▴"}
          </button>
        </div>
      </div>

      <div className="ribbon__content" hidden={!contentVisible}>
        {tabs.map((tab) => {
          const isActive = tab.id === activeTabId;
          return (
            <div
              key={tab.id}
              id={`ribbon-panel-${tab.id}`}
              role="tabpanel"
              aria-labelledby={`ribbon-tab-${tab.id}`}
              aria-label={tab.label}
              tabIndex={-1}
              hidden={!isActive}
              className="ribbon__tabpanel"
            >
              {tab.groups.map((group) => (
                <RibbonGroup
                  key={group.id}
                  group={group}
                  pressedById={mergedPressedById}
                  labelById={uiState.labelById}
                  disabledById={uiState.disabledById}
                  onActivateButton={activateButton}
                />
              ))}
            </div>
          );
        })}
      </div>
    </div>
  );
}
