import React from "react";

import type { RibbonButtonDefinition, RibbonButtonKind, RibbonButtonSize } from "./ribbonSchema.js";
import { RibbonIcon } from "./icons/RibbonIcon.js";

export interface RibbonButtonProps {
  button: RibbonButtonDefinition;
  pressed?: boolean;
  labelOverride?: string;
  disabledOverride?: boolean;
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

function getRibbonIconNode(iconId?: RibbonButtonDefinition["iconId"]): React.ReactNode {
  if (!iconId) return null;
  return <RibbonIcon id={iconId} />;
}

export const RibbonButton = React.memo(function RibbonButton({
  button,
  pressed,
  labelOverride,
  disabledOverride,
  onActivate,
}: RibbonButtonProps) {
  const kind = button.kind ?? "button";
  const size = button.size ?? "small";
  const isPressed = Boolean(pressed);
  const iconNode = getRibbonIconNode(button.iconId);
  const hasIcon = Boolean(iconNode);
  const ariaPressed = kind === "toggle" ? isPressed : undefined;
  const hasMenu = kind === "dropdown" && Boolean(button.menuItems?.length);
  const ariaHaspopup = hasMenu ? ("menu" as const) : undefined;
  const reactInstanceId = React.useId();
  const domInstanceId = React.useMemo(() => reactInstanceId.replace(/[^a-zA-Z0-9_-]/g, "-"), [reactInstanceId]);
  const menuId = React.useMemo(
    () => `ribbon-menu-${domInstanceId}-${button.id.replace(/[^a-zA-Z0-9_-]/g, "-")}`,
    [button.id, domInstanceId],
  );
  const label = labelOverride ?? button.label;
  const disabled = typeof disabledOverride === "boolean" ? disabledOverride : Boolean(button.disabled);

  const [menuOpen, setMenuOpen] = React.useState(false);
  const dropdownRef = React.useRef<HTMLDivElement | null>(null);
  const buttonRef = React.useRef<HTMLButtonElement | null>(null);

  const closeMenu = React.useCallback(() => {
    setMenuOpen(false);
  }, []);

  const focusFirstMenuItem = React.useCallback(() => {
    const root = dropdownRef.current;
    if (!root) return;
    const first = root.querySelector<HTMLButtonElement>(".ribbon-dropdown__menuitem:not(:disabled)");
    first?.focus();
  }, []);

  const focusLastMenuItem = React.useCallback(() => {
    const root = dropdownRef.current;
    if (!root) return;
    const items = Array.from(root.querySelectorAll<HTMLButtonElement>(".ribbon-dropdown__menuitem:not(:disabled)"));
    items.at(-1)?.focus();
  }, []);

  React.useEffect(() => {
    if (!menuOpen) return;
    // Defer so the menu is mounted before trying to move focus.
    requestAnimationFrame(() => focusFirstMenuItem());
  }, [focusFirstMenuItem, menuOpen]);

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

    const onFocusIn = (event: FocusEvent) => {
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
    document.addEventListener("focusin", onFocusIn);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("focusin", onFocusIn);
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
        hasIcon ? "ribbon-button--has-icon" : null,
        isPressed ? "is-pressed" : null,
        hasMenu && menuOpen ? "is-open" : null,
      ]
        .filter(Boolean)
        .join(" ")}
      aria-label={button.ariaLabel}
      aria-pressed={ariaPressed}
      aria-haspopup={ariaHaspopup}
      aria-expanded={hasMenu ? menuOpen : undefined}
      aria-controls={hasMenu && menuOpen ? menuId : undefined}
      disabled={disabled}
      data-testid={button.testId}
      data-command-id={button.id}
      ref={buttonRef}
      onClick={() => {
        if (hasMenu) {
          setMenuOpen((prev) => !prev);
          return;
        }
        onActivate?.(button);
      }}
      onKeyDown={(event) => {
        if (!hasMenu) return;
        if (event.key === "ArrowDown") {
          event.preventDefault();
          setMenuOpen(true);
          requestAnimationFrame(() => focusFirstMenuItem());
          return;
        }
        if (event.key === "ArrowUp") {
          event.preventDefault();
          setMenuOpen(true);
          requestAnimationFrame(() => focusLastMenuItem());
        }
      }}
      title={button.ariaLabel}
    >
      {iconNode ? (
        <span className="ribbon-button__icon" aria-hidden="true">
          {iconNode}
        </span>
      ) : null}
      <span className="ribbon-button__label">{label}</span>
      {kind === "dropdown" ? (
        <span className="ribbon-button__caret" aria-hidden="true">
          <RibbonIcon id="arrowDown" />
        </span>
      ) : null}
    </button>
  );

  if (!hasMenu) {
    return buttonEl;
  }

  return (
    <div
      className={["ribbon-dropdown", `ribbon-dropdown--${size}`].join(" ")}
      ref={dropdownRef}
      data-keybinding-barrier={menuOpen ? "true" : undefined}
    >
      {buttonEl}
      {menuOpen ? (
        <div
          id={menuId}
          className="ribbon-dropdown__menu"
          data-keybinding-barrier="true"
          role="menu"
          aria-label={button.ariaLabel}
          onKeyDown={(event) => {
            const root = dropdownRef.current;
            if (!root) return;
            const items = Array.from(root.querySelectorAll<HTMLButtonElement>(".ribbon-dropdown__menuitem:not(:disabled)"));
            if (items.length === 0) return;
            const currentIndex = items.findIndex((el) => el === document.activeElement);

            if (event.key === "Tab") {
              // Allow the browser to perform normal sequential focus navigation (Tab / Shift+Tab),
              // then close the menu on the next frame. This avoids unmounting the focused menuitem
              // before the browser has a chance to move focus.
              requestAnimationFrame(() => closeMenu());
              return;
            }

            if (event.key === "ArrowDown") {
              event.preventDefault();
              const nextIndex = currentIndex >= 0 ? (currentIndex + 1) % items.length : 0;
              items[nextIndex]?.focus();
              return;
            }

            if (event.key === "ArrowUp") {
              event.preventDefault();
              const nextIndex = currentIndex >= 0 ? (currentIndex - 1 + items.length) % items.length : items.length - 1;
              items[nextIndex]?.focus();
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
          {button.menuItems?.map((item) => (
            <button
              key={item.id}
              type="button"
              role="menuitem"
              className="ribbon-dropdown__menuitem"
              aria-label={item.ariaLabel}
              tabIndex={-1}
              disabled={item.disabled}
              data-testid={item.testId}
              data-command-id={item.id}
              onClick={() => {
                closeMenu();
                onActivate?.({
                  id: item.id,
                  label: item.label,
                  ariaLabel: item.ariaLabel,
                  iconId: item.iconId,
                  kind: "button",
                  size: "small",
                  testId: item.testId,
                  disabled: item.disabled,
                });
                buttonRef.current?.focus();
              }}
            >
              {(() => {
                const menuIconNode = getRibbonIconNode(item.iconId);
                return menuIconNode ? (
                  <span className="ribbon-dropdown__icon" aria-hidden="true">
                    {menuIconNode}
                  </span>
                ) : null;
              })()}
              <span className="ribbon-dropdown__label">{item.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
});

RibbonButton.displayName = "RibbonButton";
