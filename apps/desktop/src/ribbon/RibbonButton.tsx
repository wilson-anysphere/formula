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
  const hasMenu = kind === "dropdown" && Boolean(button.menuItems?.length);

  const [menuOpen, setMenuOpen] = React.useState(false);
  const dropdownRef = React.useRef<HTMLDivElement | null>(null);
  const buttonRef = React.useRef<HTMLButtonElement | null>(null);

  const closeMenu = React.useCallback(() => {
    setMenuOpen(false);
  }, []);

  React.useEffect(() => {
    if (!menuOpen) return;

    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      const root = dropdownRef.current;
      if (!root) return;
      if (root.contains(target)) return;
      closeMenu();
    };

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      closeMenu();
      buttonRef.current?.focus();
    };

    document.addEventListener("pointerdown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [closeMenu, menuOpen]);

  const buttonEl = (
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
      aria-expanded={hasMenu ? menuOpen : undefined}
      disabled={button.disabled}
      data-testid={button.testId}
      ref={buttonRef}
      onClick={() => {
        if (hasMenu) {
          setMenuOpen((prev) => !prev);
          return;
        }
        onActivate?.(button);
      }}
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

  if (!hasMenu) {
    return buttonEl;
  }

  return (
    <div className="ribbon-dropdown" ref={dropdownRef}>
      {buttonEl}
      {menuOpen ? (
        <div className="ribbon-dropdown__menu" role="menu" aria-label={button.ariaLabel}>
          {button.menuItems?.map((item) => (
            <button
              key={item.id}
              type="button"
              role="menuitem"
              className="ribbon-dropdown__menuitem"
              disabled={item.disabled}
              data-testid={item.testId}
              onClick={() => {
                closeMenu();
                onActivate?.({
                  id: item.id,
                  label: item.label,
                  ariaLabel: item.ariaLabel,
                  icon: item.icon,
                  kind: "button",
                  size: "small",
                  testId: item.testId,
                  disabled: item.disabled,
                });
              }}
            >
              {item.icon ? (
                <span className="ribbon-dropdown__icon" aria-hidden="true">
                  {item.icon}
                </span>
              ) : null}
              <span className="ribbon-dropdown__label">{item.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}
