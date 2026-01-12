import React from "react";

import type { RibbonActions, RibbonButtonDefinition, RibbonSchema } from "./ribbonSchema.js";
import { defaultRibbonSchema } from "./ribbonSchema.js";
import { RibbonGroup } from "./RibbonGroup.js";
import { getRibbonUiStateSnapshot, subscribeRibbonUiState } from "./ribbonUiState.js";
import { FileBackstage } from "./FileBackstage.js";
import { RibbonIcon } from "./icons/RibbonIcon.js";

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
  const reactInstanceId = React.useId();
  const domInstanceId = React.useMemo(() => reactInstanceId.replace(/[^a-zA-Z0-9_-]/g, "-"), [reactInstanceId]);

  const tabDomId = React.useCallback((tabId: string) => `ribbon-tab-${domInstanceId}-${tabId}`, [domInstanceId]);
  const panelDomId = React.useCallback((tabId: string) => `ribbon-panel-${domInstanceId}-${tabId}`, [domInstanceId]);

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
  const [backstageOpen, setBackstageOpen] = React.useState(false);

  const uiState = React.useSyncExternalStore(
    subscribeRibbonUiState,
    getRibbonUiStateSnapshot,
    getRibbonUiStateSnapshot,
  );
  const [ribbonWidth, setRibbonWidth] = React.useState<number>(0);
  const [userCollapsed, setUserCollapsed] = React.useState<boolean>(() => readRibbonCollapsedFromStorage());
  const [flyoutOpen, setFlyoutOpen] = React.useState(false);
  const [tabMenuOpen, setTabMenuOpen] = React.useState(false);

  const toggleUserCollapsed = React.useCallback(() => {
    setUserCollapsed((prev) => !prev);
  }, []);

  const tabMenuButtonRef = React.useRef<HTMLButtonElement | null>(null);
  const tabMenuRef = React.useRef<HTMLDivElement | null>(null);
  const tabMenuId = React.useMemo(() => `ribbon-tab-menu-${domInstanceId}`, [domInstanceId]);

  const closeTabMenu = React.useCallback(() => {
    setTabMenuOpen(false);
  }, []);

  React.useEffect(() => {
    // Persist the user-controlled "Collapse Ribbon" toggle.
    writeRibbonCollapsedToStorage(userCollapsed);
  }, [userCollapsed]);

  const tabButtonRefs = React.useRef<Record<string, HTMLButtonElement | null>>({});
  const lastNonFileTabId = React.useRef<string>(defaultTabId);

  React.useEffect(() => {
    pressedByIdRef.current = pressedById;
  }, [pressedById]);

  const mergedPressedById = React.useMemo(() => {
    // Spread here is fine: there are only ~tens of ribbon controls.
    return { ...pressedById, ...uiState.pressedById };
  }, [pressedById, uiState.pressedById]);

  // Use a layout effect so narrow windows don't briefly flash the full ribbon on mount
  // before we measure width.
  React.useLayoutEffect(() => {
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

  React.useEffect(() => {
    const active = tabs.find((tab) => tab.id === activeTabId);
    if (active && !active.isFile) {
      lastNonFileTabId.current = activeTabId;
    }
  }, [activeTabId, tabs]);

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
        setFlyoutOpen(false);
        return;
      }

      actions.onCommand?.(button.id);
      setFlyoutOpen(false);
    },
    [actions, uiState.pressedById],
  );

  const selectTabByIndex = React.useCallback(
    (nextIndex: number) => {
      const tab = tabs[nextIndex];
      if (!tab) return;
      setActiveTabId(tab.id);
      setBackstageOpen(Boolean(tab.isFile));
      if (tab.isFile) {
        setFlyoutOpen(false);
      }
      actions.onTabChange?.(tab.id);
      // Focus needs to occur after React has committed updates so that the newly
      // selected tab isn't still hidden by responsive styles (e.g. narrow tab strip).
      requestAnimationFrame(() => {
        const button = tabButtonRefs.current[tab.id];
        // Prefer keeping the viewport stable while still ensuring the newly-focused
        // tab is visible within the horizontally-scrollable tab strip.
        button?.focus({ preventScroll: true });
        button?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
      });
    },
    [actions, tabs],
  );

  const focusFirstControl = React.useCallback((tabId: string) => {
    const panel = document.getElementById(panelDomId(tabId));
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
  }, [panelDomId]);

  const responsiveDensity = React.useMemo(() => densityFromWidth(ribbonWidth), [ribbonWidth]);
  const density: RibbonDensity = responsiveDensity === "hidden" ? "hidden" : userCollapsed ? "hidden" : responsiveDensity;
  const contentVisible = density !== "hidden";
  const showTabMenuToggle = responsiveDensity === "hidden";
  const collapseLabel = userCollapsed ? "Expand ribbon" : "Collapse ribbon";
  const showContent = contentVisible || flyoutOpen;

  React.useEffect(() => {
    if (showTabMenuToggle) return;
    setTabMenuOpen(false);
  }, [showTabMenuToggle]);

  React.useEffect(() => {
    if (contentVisible) setFlyoutOpen(false);
  }, [contentVisible]);

  const selectTabFromMenu = React.useCallback(
    (tabId: string) => {
      const tab = tabs.find((candidate) => candidate.id === tabId);
      if (!tab) return;
      closeTabMenu();
      setActiveTabId(tab.id);
      if (tab.isFile) {
        setBackstageOpen(true);
        setFlyoutOpen(false);
        actions.onTabChange?.(tab.id);
        return;
      }
      setBackstageOpen(false);
      actions.onTabChange?.(tab.id);
      if (!contentVisible) {
        setFlyoutOpen(true);
        requestAnimationFrame(() => focusFirstControl(tab.id));
      }
    },
    [actions, closeTabMenu, contentVisible, focusFirstControl, tabs],
  );

  const focusActiveTabMenuItem = React.useCallback(() => {
    const menu = tabMenuRef.current;
    if (!menu) return;
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>(".ribbon__tab-menuitem:not(:disabled)"));
    if (items.length === 0) return;
    const activeItem = items.find((item) => item.dataset.tabId === activeTabId) ?? items[0];
    activeItem?.focus();
  }, [activeTabId]);

  React.useEffect(() => {
    if (!tabMenuOpen) return;
    requestAnimationFrame(() => focusActiveTabMenuItem());
  }, [focusActiveTabMenuItem, tabMenuOpen]);

  React.useEffect(() => {
    if (!tabMenuOpen) return;

    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      const menu = tabMenuRef.current;
      const button = tabMenuButtonRef.current;
      if (menu?.contains(target)) return;
      if (button?.contains(target)) return;
      closeTabMenu();
    };

    const onFocusIn = (event: FocusEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      const menu = tabMenuRef.current;
      const button = tabMenuButtonRef.current;
      if (menu?.contains(target)) return;
      if (button?.contains(target)) return;
      closeTabMenu();
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      closeTabMenu();
      tabMenuButtonRef.current?.focus();
    };

    document.addEventListener("pointerdown", onPointerDown);
    document.addEventListener("focusin", onFocusIn);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("focusin", onFocusIn);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [closeTabMenu, tabMenuOpen]);

  React.useEffect(() => {
    if (!flyoutOpen) return;
    if (contentVisible) return;
    const root = rootRef.current;
    if (!root) return;

    const close = () => setFlyoutOpen(false);

    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (root.contains(target)) return;
      close();
    };

    const onFocusIn = (event: FocusEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (root.contains(target)) return;
      close();
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      close();
      tabButtonRefs.current[activeTabId]?.focus();
    };

    document.addEventListener("pointerdown", onPointerDown);
    document.addEventListener("focusin", onFocusIn);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("focusin", onFocusIn);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [activeTabId, contentVisible, flyoutOpen]);

  const closeBackstage = React.useCallback(() => {
    setBackstageOpen(false);
    setFlyoutOpen(false);
    const fallback =
      tabs.find((tab) => tab.id === lastNonFileTabId.current && !tab.isFile)?.id ??
      tabs.find((tab) => !tab.isFile)?.id ??
      defaultTabId;
    setActiveTabId(fallback);
    requestAnimationFrame(() => {
      const button = tabButtonRefs.current[fallback];
      button?.focus({ preventScroll: true });
      button?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    });
  }, [defaultTabId, tabs]);

  return (
    <div
      className="ribbon"
      data-testid="ribbon-root"
      data-responsive-density={responsiveDensity}
      data-density={density}
      data-flyout-open={flyoutOpen && !contentVisible ? "true" : undefined}
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
        if (event.key === "Escape" && flyoutOpen && !contentVisible) {
          setFlyoutOpen(false);
        }
        tabButtonRefs.current[activeTabId]?.focus();
      }}
    >
      <div className="ribbon__tabstrip">
        {showTabMenuToggle ? (
          <div className="ribbon__tabstrip-left">
            <button
              type="button"
              className="ribbon__tab-menu-toggle"
              aria-label="Open ribbon menu"
              title="Open ribbon menu"
              aria-haspopup="menu"
              aria-expanded={tabMenuOpen}
              aria-controls={tabMenuOpen ? tabMenuId : undefined}
              data-testid="ribbon-tab-menu-toggle"
              ref={tabMenuButtonRef}
              onClick={() => {
                const next = !tabMenuOpen;
                if (next) setFlyoutOpen(false);
                setTabMenuOpen(next);
              }}
              onKeyDown={(event) => {
                if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
                  event.preventDefault();
                  setFlyoutOpen(false);
                  setTabMenuOpen(true);
                }
              }}
            >
              <RibbonIcon id="menu" width={16} height={16} />
            </button>
            {tabMenuOpen ? (
              <div
                id={tabMenuId}
                className="ribbon__tab-menu"
                role="menu"
                aria-label="Ribbon tabs"
                data-keybinding-barrier="true"
                data-testid="ribbon-tab-menu"
                ref={tabMenuRef}
                onKeyDown={(event) => {
                  const menu = tabMenuRef.current;
                  if (!menu) return;
                  const items = Array.from(
                    menu.querySelectorAll<HTMLButtonElement>(".ribbon__tab-menuitem:not(:disabled)"),
                  );
                  if (items.length === 0) return;
                  const currentIndex = items.findIndex((el) => el === document.activeElement);

                  if (event.key === "Tab") {
                    // Let the browser move focus (Tab / Shift+Tab), then close the menu on the
                    // next frame so we don't unmount the focused element mid-navigation.
                    requestAnimationFrame(() => closeTabMenu());
                    return;
                  }

                  if (event.key === "ArrowDown") {
                    event.preventDefault();
                    const next = currentIndex >= 0 ? (currentIndex + 1) % items.length : 0;
                    items[next]?.focus();
                    return;
                  }

                  if (event.key === "ArrowUp") {
                    event.preventDefault();
                    const next =
                      currentIndex >= 0 ? (currentIndex - 1 + items.length) % items.length : items.length - 1;
                    items[next]?.focus();
                    return;
                  }

                  if (event.key === "Home") {
                    event.preventDefault();
                    items[0]?.focus();
                    return;
                  }

                  if (event.key === "End") {
                    event.preventDefault();
                    items.at(-1)?.focus();
                  }
                }}
              >
                {tabs.map((tab) => {
                  const isActive = tab.id === activeTabId;
                  return (
                    <button
                      key={tab.id}
                      type="button"
                      role="menuitemradio"
                      aria-checked={isActive}
                      className={["ribbon__tab-menuitem", isActive ? "is-active" : null].filter(Boolean).join(" ")}
                      title={tab.label}
                      data-tab-id={tab.id}
                      tabIndex={-1}
                      onClick={() => selectTabFromMenu(tab.id)}
                    >
                      {tab.label}
                    </button>
                  );
                })}
              </div>
            ) : null}
          </div>
        ) : null}
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
                data-testid={`ribbon-tab-${tab.id}`}
                id={tabDomId(tab.id)}
                aria-selected={isActive}
                aria-controls={panelDomId(tab.id)}
                tabIndex={isActive ? 0 : -1}
                ref={(el) => {
                  tabButtonRefs.current[tab.id] = el;
                }}
                onClick={() => {
                  setActiveTabId(tab.id);
                  if (isFile) {
                    setBackstageOpen(true);
                    setFlyoutOpen(false);
                    actions.onTabChange?.(tab.id);
                    return;
                  }
                  setBackstageOpen(false);
                  actions.onTabChange?.(tab.id);
                  if (!contentVisible) setFlyoutOpen((prev) => (isActive ? !prev : true));
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
                    if (isFile) {
                      event.preventDefault();
                      if (!isActive) {
                        setActiveTabId(tab.id);
                        setBackstageOpen(true);
                        setFlyoutOpen(false);
                        actions.onTabChange?.(tab.id);
                        return;
                      }
                      setBackstageOpen(true);
                      setFlyoutOpen(false);
                      return;
                    }
                    event.preventDefault();
                    if (!contentVisible) {
                      if (!isActive) {
                        setActiveTabId(tab.id);
                        setBackstageOpen(false);
                        actions.onTabChange?.(tab.id);
                      }
                      setFlyoutOpen(true);
                      requestAnimationFrame(() => focusFirstControl(tab.id));
                      return;
                    }
                    if (!isActive) {
                      setActiveTabId(tab.id);
                      setBackstageOpen(false);
                      actions.onTabChange?.(tab.id);
                      requestAnimationFrame(() => focusFirstControl(tab.id));
                      return;
                    }
                    focusFirstControl(tab.id);
                    return;
                  }
                  if (event.key === "Tab" && !event.shiftKey) {
                    // Excel-style: Tab from the active tab moves focus into the tab panel.
                    // Shift+Tab should keep browser default behavior (move focus backwards).
                    if (isFile) {
                      event.preventDefault();
                      if (!isActive) {
                        setActiveTabId(tab.id);
                        setBackstageOpen(true);
                        setFlyoutOpen(false);
                        actions.onTabChange?.(tab.id);
                        return;
                      }
                      setBackstageOpen(true);
                      setFlyoutOpen(false);
                      return;
                    }
                    event.preventDefault();
                    if (!contentVisible) {
                      if (!isActive) {
                        setActiveTabId(tab.id);
                        setBackstageOpen(false);
                        actions.onTabChange?.(tab.id);
                      }
                      setFlyoutOpen(true);
                      requestAnimationFrame(() => focusFirstControl(tab.id));
                      return;
                    }
                    if (!isActive) {
                      setActiveTabId(tab.id);
                      setBackstageOpen(false);
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
            <RibbonIcon id={userCollapsed ? "arrowDown" : "arrowUp"} width={14} height={14} />
          </button>
        </div>
      </div>

      <div className="ribbon__content" hidden={!showContent || backstageOpen}>
        {tabs.map((tab) => {
          const isActive = tab.id === activeTabId;
          return (
            <div
              key={tab.id}
              id={panelDomId(tab.id)}
              role="tabpanel"
              aria-labelledby={tabDomId(tab.id)}
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

      <FileBackstage
        open={backstageOpen && Boolean(tabs.find((tab) => tab.id === activeTabId)?.isFile)}
        actions={actions.fileActions}
        onClose={closeBackstage}
      />
    </div>
  );
}
