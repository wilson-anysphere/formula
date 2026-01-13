import React from "react";

import type { RibbonButtonDefinition, RibbonGroupDefinition } from "./ribbonSchema.js";
import { RibbonButton } from "./RibbonButton.js";

export interface RibbonGroupProps {
  group: RibbonGroupDefinition;
  pressedById: Record<string, boolean>;
  labelById?: Record<string, string>;
  disabledById?: Record<string, boolean>;
  shortcutById?: Record<string, string>;
  onActivateButton?: (button: RibbonButtonDefinition) => void;
}

export function RibbonGroup({ group, pressedById, labelById, disabledById, shortcutById, onActivateButton }: RibbonGroupProps) {
  return (
    <section className="ribbon-group" role="group" aria-label={group.label}>
      <div className="ribbon-group__content">
        {group.buttons.map((button) => (
          <RibbonButton
            key={button.id}
            button={button}
            pressed={pressedById[button.id]}
            labelOverride={labelById?.[button.id]}
            disabledOverride={disabledById?.[button.id]}
            shortcutOverride={shortcutById?.[button.id]}
            shortcutById={shortcutById}
            onActivate={onActivateButton}
          />
        ))}
      </div>
      <div className="ribbon-group__label">{group.label}</div>
    </section>
  );
}
