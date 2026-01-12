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

  const activeTab = tabs.find((t) => t.id === activeTabId) ?? tabs[0];

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

  return (
    <div className="ribbon" data-testid="ribbon-root">
      <div className="ribbon__tabs" role="tablist" aria-label="Ribbon tabs">
        {tabs.map((tab) => {
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
              aria-selected={isActive}
              onClick={() => {
                setActiveTabId(tab.id);
                actions.onTabChange?.(tab.id);
              }}
            >
              {tab.label}
            </button>
          );
        })}
      </div>

      <div className="ribbon__content" role="tabpanel" aria-label={activeTab?.label ?? "Ribbon"}>
        {activeTab?.groups.map((group) => (
          <RibbonGroup key={group.id} group={group} pressedById={pressedById} onActivateButton={activateButton} />
        ))}
      </div>
    </div>
  );
}

