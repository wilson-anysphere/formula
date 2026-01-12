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

export interface RibbonProps {
  actions: RibbonActions;
  schema?: RibbonSchema;
  initialTabId?: string;
}

export function Ribbon({ actions, schema = defaultRibbonSchema, initialTabId }: RibbonProps) {
  const tabs = schema.tabs;
  const firstTabId = tabs[0]?.id ?? "home";

  const [activeTabId, setActiveTabId] = React.useState<string>(initialTabId ?? firstTabId);
  const [pressedById, setPressedById] = React.useState<Record<string, boolean>>(() => computeInitialPressed(schema));

  const tabButtonRefs = React.useRef<Record<string, HTMLButtonElement | null>>({});

  const activateButton = React.useCallback(
    (button: RibbonButtonDefinition) => {
      const kind = button.kind ?? "button";

      if (kind === "toggle") {
        setPressedById((prev) => {
          const nextPressed = !prev[button.id];
          const next = { ...prev, [button.id]: nextPressed };
          actions.onToggle?.(button.id, nextPressed);
          actions.onCommand?.(button.id);
          return next;
        });
        return;
      }

      actions.onCommand?.(button.id);
    },
    [actions],
  );

  const selectTabByIndex = React.useCallback(
    (nextIndex: number) => {
      const tab = tabs[nextIndex];
      if (!tab) return;
      setActiveTabId(tab.id);
      actions.onTabChange?.(tab.id);
      tabButtonRefs.current[tab.id]?.focus();
    },
    [actions, tabs],
  );

  return (
    <div className="ribbon" data-testid="ribbon-root">
      <div className="ribbon__tabs" role="tablist" aria-label="Ribbon tabs">
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
