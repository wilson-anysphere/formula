import React from "react";

import type { RibbonButtonDefinition, RibbonGroupDefinition } from "./ribbonSchema.js";
import { RibbonButton } from "./RibbonButton.js";

export interface RibbonGroupProps {
  group: RibbonGroupDefinition;
  pressedById: Record<string, boolean>;
  onActivateButton?: (button: RibbonButtonDefinition) => void;
}

export function RibbonGroup({ group, pressedById, onActivateButton }: RibbonGroupProps) {
  return (
    <section className="ribbon-group" aria-label={group.label}>
      <div className="ribbon-group__content">
        {group.buttons.map((button) => (
          <RibbonButton key={button.id} button={button} pressed={pressedById[button.id]} onActivate={onActivateButton} />
        ))}
      </div>
      <div className="ribbon-group__label">{group.label}</div>
    </section>
  );
}

