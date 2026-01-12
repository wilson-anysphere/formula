import React from "react";

import type { RibbonActions, RibbonButtonDefinition, RibbonSchema } from "./ribbonSchema.js";
import { defaultRibbonSchema } from "./ribbonSchema.js";
import { RibbonGroup } from "./RibbonGroup.js";

import "../styles/ribbon.css";

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
  const tabs = schema.tabs;
  const defaultTabId = React.useMemo(() => {
    if (initialTabId && tabs.some((tab) => tab.id === initialTabId)) {
      return initialTabId;
    }

    return tabs.find((tab) => tab.id === "home")?.id ?? tabs[0]?.id ?? "home";
  }, [initialTabId, tabs]);

  const [activeTabId, setActiveTabId] = React.useState<string>(defaultTabId);
  const [pressedById, setPressedById] = React.useState<Record<string, boolean>>(() => computeInitialPressed(schema));

  const tabButtonRefs = React.useRef<Record<string, HTMLButtonElement | null>>({});

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
        const nextPressed = !pressedById[button.id];
        setPressedById((prev) => ({ ...prev, [button.id]: !prev[button.id] }));
        actions.onToggle?.(button.id, nextPressed);
        actions.onCommand?.(button.id);
        return;
      }

      actions.onCommand?.(button.id);
    },
    [actions, pressedById],
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
    const first = panel.querySelector<HTMLButtonElement>("button:not(:disabled)");
    first?.focus();
  }, []);

  return (
    <div
      className="ribbon"
      data-testid="ribbon-root"
      onKeyDownCapture={(event) => {
        if (event.key !== "Escape" && event.key !== "ArrowUp") return;
        const target = event.target as HTMLElement | null;
        if (!target) return;
        if (target.closest(".ribbon__tabs")) return;
        if (!target.closest(".ribbon__content")) return;
        event.preventDefault();
        tabButtonRefs.current[activeTabId]?.focus();
      }}
    >
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
            >
              {tab.label}
            </button>
          );
        })}
      </div>

      <div className="ribbon__content">
        {tabs.map((tab) => {
          const isActive = tab.id === activeTabId;
          return (
            <div
              key={tab.id}
              id={`ribbon-panel-${tab.id}`}
              role="tabpanel"
              aria-labelledby={`ribbon-tab-${tab.id}`}
              aria-label={tab.label}
              hidden={!isActive}
              className="ribbon__tabpanel"
            >
              {tab.groups.map((group) => (
                <RibbonGroup key={group.id} group={group} pressedById={pressedById} onActivateButton={activateButton} />
              ))}
            </div>
          );
        })}
      </div>
    </div>
  );
}
