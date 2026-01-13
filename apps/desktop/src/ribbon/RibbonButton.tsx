import React from "react";

import type { RibbonButtonDefinition, RibbonButtonKind, RibbonButtonSize } from "./ribbonSchema.js";
import { RibbonIcon } from "./icons/RibbonIcon.js";

export interface RibbonButtonProps {
  button: RibbonButtonDefinition;
  pressed?: boolean;
  labelOverride?: string;
  disabledOverride?: boolean;
  /**
   * Optional shortcut display override for the top-level ribbon control.
   *
   * This is plumbed through `RibbonGroup` from `RibbonUiState.shortcutById`.
   */
  shortcutOverride?: string;
  /**
   * Optional lookup table for dropdown menu items (keyed by menu item `id`).
   */
  shortcutById?: Record<string, string>;
  /**
   * Optional `aria-keyshortcuts` override for the top-level ribbon control.
   *
   * This is plumbed through `RibbonGroup` from `RibbonUiState.ariaKeyShortcutsById`.
   */
  ariaKeyShortcutsOverride?: string;
  /**
   * Optional lookup table for dropdown menu items (keyed by menu item `id`).
   */
  ariaKeyShortcutsById?: Record<string, string>;
  /**
   * Per-menu-item disabled overrides, aligned with `button.menuItems`.
   *
   * This is passed as an array (instead of the full `disabledById` record) to avoid
   * invalidating `React.memo(...)` on every ribbon UI state update.
   */
  menuItemDisabledOverrides?: ReadonlyArray<boolean | undefined>;
  /**
   * Per-menu-item label overrides, aligned with `button.menuItems`.
   *
   * This is passed as an array (instead of the full `labelById` record) to avoid
   * invalidating `React.memo(...)` on every ribbon UI state update.
   */
  menuItemLabelOverrides?: ReadonlyArray<string | undefined>;
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

function formatTooltipTitle(base: string, shortcut: string | null | undefined): string {
  const baseLabel = String(base ?? "");
  const hint = typeof shortcut === "string" && shortcut.trim() ? shortcut.trim() : "";
  if (!hint) return baseLabel;

  // If the base tooltip already contains the shortcut (e.g. some ariaLabels include a hint),
  // avoid duplicating it.
  if (baseLabel.includes(hint)) return baseLabel;

  // Special-case the "Ctrl/Cmd" placeholder pattern used in a small number of ribbon ariaLabels.
  // Replace the placeholder with the platform-specific binding from KeybindingService so the
  // tooltip stays accurate without rendering a duplicate "(...)" suffix.
  const replaced = baseLabel.replace(/\(Ctrl\/Cmd\+[^)]*\)/, `(${hint})`);
  if (replaced !== baseLabel) return replaced;

  return `${baseLabel} (${hint})`;
}

export const RibbonButton = React.memo(function RibbonButton({
  button,
  pressed,
  labelOverride,
  disabledOverride,
  shortcutOverride,
  shortcutById,
  ariaKeyShortcutsOverride,
  ariaKeyShortcutsById,
  menuItemDisabledOverrides,
  menuItemLabelOverrides,
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
  const shortcut = shortcutOverride ?? shortcutById?.[button.id];
  const title = formatTooltipTitle(button.ariaLabel, shortcut);
  const ariaKeyShortcuts = ariaKeyShortcutsOverride ?? ariaKeyShortcutsById?.[button.id];

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
      aria-keyshortcuts={ariaKeyShortcuts || undefined}
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
      title={title}
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
          {button.menuItems?.map((item, idx) => {
            const menuItemLabelOverride = menuItemLabelOverrides?.[idx];
            const menuItemLabel = menuItemLabelOverride ?? item.label;
            const menuItemDisabledOverride = menuItemDisabledOverrides?.[idx];
            const menuItemDisabled =
              typeof menuItemDisabledOverride === "boolean" ? menuItemDisabledOverride : Boolean(item.disabled);
            const itemShortcut = shortcutById?.[item.id];
            const itemTitle = formatTooltipTitle(item.ariaLabel, itemShortcut);
            const itemAriaKeyShortcuts = ariaKeyShortcutsById?.[item.id];

            return (
              <button
                key={item.id}
                type="button"
                role="menuitem"
                className="ribbon-dropdown__menuitem"
                aria-label={item.ariaLabel}
                aria-keyshortcuts={itemAriaKeyShortcuts || undefined}
                title={itemTitle}
                data-shortcut={itemShortcut || undefined}
                tabIndex={-1}
                disabled={menuItemDisabled}
                data-testid={item.testId}
                data-command-id={item.id}
                onClick={() => {
                  closeMenu();
                  onActivate?.({
                    id: item.id,
                    label: menuItemLabel,
                    ariaLabel: item.ariaLabel,
                    iconId: item.iconId,
                    kind: "button",
                    size: "small",
                    testId: item.testId,
                    disabled: menuItemDisabled,
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
                <span className="ribbon-dropdown__label">{menuItemLabel}</span>
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
},
function areRibbonButtonPropsEqual(prev, next) {
  if (prev.button !== next.button) return false;
  if (prev.pressed !== next.pressed) return false;
  if (prev.labelOverride !== next.labelOverride) return false;
  if (prev.disabledOverride !== next.disabledOverride) return false;
  if (prev.shortcutOverride !== next.shortcutOverride) return false;
  if (prev.shortcutById !== next.shortcutById) return false;
  if (prev.ariaKeyShortcutsOverride !== next.ariaKeyShortcutsOverride) return false;
  if (prev.ariaKeyShortcutsById !== next.ariaKeyShortcutsById) return false;
  if (prev.onActivate !== next.onActivate) return false;

  const arraysShallowEqual = <T,>(a?: ReadonlyArray<T>, b?: ReadonlyArray<T>): boolean => {
    if (a === b) return true;
    if (!a || !b) return false;
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  };

  if (!arraysShallowEqual(prev.menuItemDisabledOverrides, next.menuItemDisabledOverrides)) return false;
  if (!arraysShallowEqual(prev.menuItemLabelOverrides, next.menuItemLabelOverrides)) return false;

  return true;
});

RibbonButton.displayName = "RibbonButton";
