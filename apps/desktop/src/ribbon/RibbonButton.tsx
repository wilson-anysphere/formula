import React from "react";

import type { RibbonButtonDefinition, RibbonButtonKind, RibbonButtonSize } from "./ribbonSchema.js";

export interface RibbonButtonProps {
  button: RibbonButtonDefinition;
  pressed?: boolean;
  onActivate?: (button: RibbonButtonDefinition) => void;
}

function classForKind(kind: RibbonButtonKind): string {
  switch (kind) {
    case "toggle":
      return "ribbon-button--toggle";
    case "dropdown":
      return "ribbon-button--dropdown";
    case "button":
    default:
      return "ribbon-button--button";
  }
}

function classForSize(size: RibbonButtonSize): string {
  switch (size) {
    case "large":
      return "ribbon-button--large";
    case "icon":
      return "ribbon-button--icon";
    case "small":
    default:
      return "ribbon-button--small";
  }
}

export function RibbonButton({ button, pressed, onActivate }: RibbonButtonProps) {
  const kind = button.kind ?? "button";
  const size = button.size ?? "small";
  const ariaPressed = kind === "toggle" ? Boolean(pressed) : undefined;
  const ariaHaspopup = kind === "dropdown" ? ("menu" as const) : undefined;

  return (
    <button
      type="button"
      className={[
        "ribbon-button",
        classForKind(kind),
        classForSize(size),
        ariaPressed ? "is-pressed" : null,
      ]
        .filter(Boolean)
        .join(" ")}
      aria-label={button.ariaLabel}
      aria-pressed={ariaPressed}
      aria-haspopup={ariaHaspopup}
      disabled={button.disabled}
      data-testid={button.testId}
      onClick={() => onActivate?.(button)}
      title={button.ariaLabel}
    >
      {button.icon ? (
        <span className="ribbon-button__icon" aria-hidden="true">
          {button.icon}
        </span>
      ) : null}
      <span className="ribbon-button__label">{button.label}</span>
      {kind === "dropdown" ? (
        <span className="ribbon-button__caret" aria-hidden="true">
          ▾
        </span>
      ) : null}
    </button>
  );
}
