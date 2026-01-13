import React from "react";

import type { RibbonButtonDefinition, RibbonGroupDefinition } from "./ribbonSchema.js";
import { RibbonButton } from "./RibbonButton.js";

export interface RibbonGroupProps {
  group: RibbonGroupDefinition;
  pressedById: Record<string, boolean>;
  labelById?: Record<string, string>;
  disabledById?: Record<string, boolean>;
  shortcutById?: Record<string, string>;
  ariaKeyShortcutsById?: Record<string, string>;
  onActivateButton?: (button: RibbonButtonDefinition) => void;
}

export function RibbonGroup({
  group,
  pressedById,
  labelById,
  disabledById,
  shortcutById,
  ariaKeyShortcutsById,
  onActivateButton,
}: RibbonGroupProps) {
  return (
    <section className="ribbon-group" role="group" aria-label={group.label}>
      <div className="ribbon-group__content">
        {group.buttons.map((button) => (
          <RibbonButton
            // Multiple controls can legitimately map to the same command id (e.g. legacy
            // variants that keep stable `testId`s for e2e). Use `testId` as the stable
            // React key when available to avoid collisions.
            key={button.testId ?? button.id}
            button={button}
            pressed={pressedById[button.id]}
            labelOverride={labelById?.[button.id]}
            disabledOverride={disabledById?.[button.id]}
            shortcutOverride={shortcutById?.[button.id]}
            shortcutById={shortcutById}
            ariaKeyShortcutsOverride={ariaKeyShortcutsById?.[button.id]}
            ariaKeyShortcutsById={ariaKeyShortcutsById}
            onActivate={onActivateButton}
          />
        ))}
      </div>
      <div className="ribbon-group__label">{group.label}</div>
    </section>
  );
}
