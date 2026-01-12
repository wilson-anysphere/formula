export type ContextMenuSeparator = { type: "separator" };

export type ContextMenuLeading = {
  type: "swatch";
  /**
   * Swatch fill color.
   *
   * Prefer passing a CSS variable token (e.g. `--sheet-tab-red`) so the swatch
   * stays token/theme driven. Literal colors (e.g. `#RRGGBB`) are also supported.
   */
  color: string;
};

export type ContextMenuActionItem = {
  type: "item";
  label: string;
  enabled?: boolean;
  shortcut?: string;
  leading?: ContextMenuLeading;
  /**
   * Invoked when the item is selected.
   *
   * Note: the menu closes before `onSelect` is called.
   */
  onSelect: () => void | Promise<void>;
};

export type ContextMenuSubmenuItem = {
  type: "submenu";
  label: string;
  enabled?: boolean;
  shortcut?: string;
  leading?: ContextMenuLeading;
  items: ContextMenuItem[];
};

export type ContextMenuItem = ContextMenuActionItem | ContextMenuSubmenuItem | ContextMenuSeparator;

export type ContextMenuOpenOptions = {
  x: number;
  y: number;
  items: ContextMenuItem[];
};

type ContextMenuButtonItem = ContextMenuActionItem | ContextMenuSubmenuItem;

type Point = { x: number; y: number };

export class ContextMenu {
  private readonly overlay: HTMLDivElement;
  private readonly menu: HTMLDivElement;

  private submenu: HTMLDivElement | null = null;
  private submenuParent: HTMLButtonElement | null = null;
  private isShown = false;
  private lastAnchor: Point | null = null;
  /**
   * Ignore external scroll events for a brief grace period after opening.
   *
   * Some browsers/WebView environments can emit a `scroll` event shortly after
   * focusing an element (even when the scroll offsets don't meaningfully change).
   * If a context menu is opened immediately after that focus, the delayed scroll
   * event would otherwise instantly dismiss the menu, which feels broken and can
   * flake e2e tests.
   */
  private ignoreExternalScrollUntil: number | null = null;

  private readonly onClose: (() => void) | null;
  private keydownListener: ((e: KeyboardEvent) => void) | null = null;
  private pointerDownListener: ((e: PointerEvent) => void) | null = null;
  private scrollListener: ((e: Event) => void) | null = null;
  private wheelListener: ((e: WheelEvent) => void) | null = null;
  private resizeListener: (() => void) | null = null;
  private blurListener: (() => void) | null = null;

  private readonly buttonItems = new WeakMap<HTMLButtonElement, ContextMenuButtonItem>();

  constructor(options: { onClose?: () => void; testId?: string } = {}) {
    this.onClose = options.onClose ?? null;

    const overlay = document.createElement("div");
    overlay.dataset.testid = options.testId ?? "context-menu";
    overlay.className = "context-menu-overlay";
    overlay.hidden = true;

    const menu = document.createElement("div");
    menu.className = "context-menu";
    // We intentionally keep menu items as native <button> elements without overriding
    // their implicit ARIA role so Playwright can locate them via role="button".
    // (Some e2e tests depend on this.)
    menu.setAttribute("role", "menu");
    menu.setAttribute("aria-orientation", "vertical");
    // Let us focus the menu container as a fallback when there are no enabled items.
    menu.tabIndex = -1;

    overlay.appendChild(menu);
    document.body.appendChild(overlay);

    // Prevent showing the native browser context menu when right-clicking inside ours.
    menu.addEventListener("contextmenu", (e) => e.preventDefault());

    this.overlay = overlay;
    this.menu = menu;
  }

  isOpen(): boolean {
    return this.isShown;
  }

  open({ x, y, items }: ContextMenuOpenOptions): void {
    this.close();
    this.isShown = true;
    this.lastAnchor = { x, y };
    // Brief grace period to avoid immediately closing due to delayed scroll events
    // (e.g. focus-induced scrollIntoView).
    this.ignoreExternalScrollUntil = performance.now() + 100;

    this.menu.replaceChildren();
    this.closeSubmenu();

    this.buildMenuContents(this.menu, items, { level: "menu" });

    this.overlay.hidden = false;
    this.positionMenu(x, y);

    // Close on outside clicks without swallowing the click (Excel-like behavior).
    this.pointerDownListener = (e: PointerEvent) => {
      if (!this.isShown) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (this.menu.contains(target)) return;
      if (this.submenu && this.submenu.contains(target)) return;
      this.close();
    };
    window.addEventListener("pointerdown", this.pointerDownListener, true);

    // Close on scroll events that originate outside the menu. This covers cases like
    // dragging scrollbars (which may not emit wheel events).
    this.scrollListener = (e: Event) => {
      if (!this.isShown) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (this.submenu && this.submenu.contains(target)) return;
      if (this.menu.contains(target)) {
        // If the main menu scrolls while a submenu is open, the submenu would no longer
        // be anchored to the correct parent item. Close it to avoid mis-positioning.
        if (this.submenu) {
          const parent = this.submenuParent;
          const active = document.activeElement as HTMLElement | null;
          const focusInSubmenu = Boolean(active && this.submenu.contains(active));
          this.closeSubmenu();
          if (focusInSubmenu) parent?.focus({ preventScroll: true });
        }
        return;
      }
      if (this.ignoreExternalScrollUntil != null && performance.now() < this.ignoreExternalScrollUntil) {
        return;
      }
      this.close();
    };
    window.addEventListener("scroll", this.scrollListener, true);

    // Close on wheel scrolling the underlying surface (keep it open when scrolling
    // within the menu/submenu itself).
    this.wheelListener = (e: WheelEvent) => {
      if (!this.isShown) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (this.submenu && this.submenu.contains(target)) return;
      if (this.menu.contains(target)) {
        if (this.submenu) {
          const parent = this.submenuParent;
          const active = document.activeElement as HTMLElement | null;
          const focusInSubmenu = Boolean(active && this.submenu.contains(active));
          this.closeSubmenu();
          if (focusInSubmenu) parent?.focus({ preventScroll: true });
        }
        return;
      }
      this.close();
    };
    window.addEventListener("wheel", this.wheelListener, { capture: true, passive: true });

    // Close on window focus changes / resizes (prevents awkward positioning after
    // viewport changes).
    this.resizeListener = () => {
      if (!this.isShown) return;
      this.close();
    };
    window.addEventListener("resize", this.resizeListener);

    this.blurListener = () => {
      if (!this.isShown) return;
      this.close();
    };
    window.addEventListener("blur", this.blurListener);

    this.keydownListener = (e) => this.onKeyDown(e);
    window.addEventListener("keydown", this.keydownListener, true);
    this.focusFirst();
  }

  focusFirst(): void {
    if (!this.isShown) return;
    this.focusFirstItem(this.menu);
  }

  /**
   * Re-render the menu in-place (used when the item model changes while open).
   */
  update(items: ContextMenuItem[]): void {
    if (!this.isShown) return;
    if (!this.lastAnchor) return;

    const active = document.activeElement;
    const activeIsMenuButton = active instanceof HTMLButtonElement && this.menu.contains(active);
    const activeLabelText = activeIsMenuButton
      ? (active.querySelector<HTMLElement>(".context-menu__label")?.textContent ?? active.textContent ?? "")
      : "";
    const focusInSubmenu = Boolean(this.submenu && active instanceof Node && this.submenu.contains(active));
    const submenuParentLabelText = focusInSubmenu
      ? (this.submenuParent?.querySelector<HTMLElement>(".context-menu__label")?.textContent ??
          this.submenuParent?.textContent ??
          "")
      : "";

    this.menu.replaceChildren();
    this.closeSubmenu();
    this.buildMenuContents(this.menu, items, { level: "menu" });
    this.positionMenu(this.lastAnchor.x, this.lastAnchor.y);

    if (activeIsMenuButton || focusInSubmenu) {
      const enabled = this.getEnabledButtons(this.menu);
      const match =
        (activeIsMenuButton ? activeLabelText : submenuParentLabelText).trim() === ""
          ? null
          : enabled.find((btn) => {
              const label = btn.querySelector<HTMLElement>(".context-menu__label")?.textContent ?? btn.textContent ?? "";
              return label === (activeIsMenuButton ? activeLabelText : submenuParentLabelText);
            }) ?? null;
      if (match) {
        this.setRovingTabIndex(enabled, match);
        match.focus({ preventScroll: true });
        return;
      }
      this.focusFirstItem(this.menu);
    }
  }

  close(): void {
    if (!this.isShown) return;
    this.isShown = false;
    this.lastAnchor = null;
    this.ignoreExternalScrollUntil = null;

    this.overlay.hidden = true;
    this.menu.replaceChildren();
    this.closeSubmenu();

    if (this.keydownListener) {
      window.removeEventListener("keydown", this.keydownListener, true);
      this.keydownListener = null;
    }

    if (this.pointerDownListener) {
      window.removeEventListener("pointerdown", this.pointerDownListener, true);
      this.pointerDownListener = null;
    }

    if (this.wheelListener) {
      window.removeEventListener("wheel", this.wheelListener, true);
      this.wheelListener = null;
    }

    if (this.scrollListener) {
      window.removeEventListener("scroll", this.scrollListener, true);
      this.scrollListener = null;
    }

    if (this.resizeListener) {
      window.removeEventListener("resize", this.resizeListener);
      this.resizeListener = null;
    }

    if (this.blurListener) {
      window.removeEventListener("blur", this.blurListener);
      this.blurListener = null;
    }

    try {
      this.onClose?.();
    } catch {
      // ignore
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (!this.isShown) return;

    const activeEl = document.activeElement as HTMLElement | null;
    const focusInSubmenu = Boolean(this.submenu && activeEl && this.submenu.contains(activeEl));
    const focusContainer = focusInSubmenu && this.submenu ? this.submenu : this.menu;

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      if (this.submenu) {
        const parent = this.submenuParent;
        this.closeSubmenu();
        if (focusInSubmenu) parent?.focus({ preventScroll: true });
        return;
      }
      this.close();
      return;
    }

    if (e.key === "Tab") {
      e.preventDefault();
      e.stopPropagation();
      this.close();
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      this.moveFocus(1);
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      e.stopPropagation();
      this.moveFocus(-1);
      return;
    }

    if (e.key === "Home" || e.key === "End") {
      e.preventDefault();
      e.stopPropagation();
      if (!focusInSubmenu) this.closeSubmenu();

      const enabled = this.getEnabledButtons(focusContainer);
      if (enabled.length === 0) return;
      const next = e.key === "Home" ? enabled[0]! : enabled[enabled.length - 1]!;
      this.setRovingTabIndex(enabled, next);
      next.focus({ preventScroll: true });
      return;
    }

    if (e.key === "Enter" || e.key === " " || e.key === "Spacebar") {
      const isInMenu =
        activeEl != null && (this.menu.contains(activeEl) || (this.submenu != null && this.submenu.contains(activeEl)));
      if (!isInMenu) return;

      e.preventDefault();
      e.stopPropagation();

      const activeBtn = activeEl instanceof HTMLButtonElement ? activeEl : null;
      if (!activeBtn || activeBtn.disabled) return;
      activeBtn.click();
      return;
    }

    if (e.key === "ArrowRight") {
      e.preventDefault();
      e.stopPropagation();

      if (focusInSubmenu) return;
      const activeBtn = activeEl instanceof HTMLButtonElement ? activeEl : null;
      if (!activeBtn || activeBtn.disabled) return;

      const item = this.buttonItems.get(activeBtn);
      if (item?.type !== "submenu") return;
      this.openSubmenu(activeBtn, item.items, { focus: true });
      return;
    }

    if (e.key === "ArrowLeft") {
      e.preventDefault();
      e.stopPropagation();

      if (!this.submenu) return;

      if (!focusInSubmenu) {
        this.closeSubmenu();
        return;
      }

      const parent = this.submenuParent;
      this.closeSubmenu();
      parent?.focus({ preventScroll: true });
    }
  }

  private moveFocus(delta: 1 | -1): void {
    const active = document.activeElement;
    const inSubmenu = this.submenu != null && active instanceof Node && this.submenu.contains(active);
    const container = inSubmenu && this.submenu ? this.submenu : this.menu;

    const enabled = this.getEnabledButtons(container);
    if (enabled.length === 0) return;

    // Navigating the main menu should close any open submenu.
    if (!inSubmenu) this.closeSubmenu();

    const activeBtn = active instanceof HTMLButtonElement ? active : null;
    const idx = activeBtn ? enabled.indexOf(activeBtn) : -1;
    const nextIdx = idx === -1 ? (delta > 0 ? 0 : enabled.length - 1) : (idx + delta + enabled.length) % enabled.length;
    const next = enabled[nextIdx]!;

    this.setRovingTabIndex(enabled, next);
    next.focus({ preventScroll: true });
  }

  private focusFirstItem(container: HTMLDivElement): void {
    const enabled = this.getEnabledButtons(container);
    if (enabled.length === 0) {
      container.focus();
      return;
    }

    this.setRovingTabIndex(enabled, enabled[0]!);
    enabled[0]!.focus({ preventScroll: true });
  }

  private getEnabledButtons(container: HTMLDivElement): HTMLButtonElement[] {
    return Array.from(container.querySelectorAll<HTMLButtonElement>(".context-menu__item:not(:disabled)"));
  }

  private setRovingTabIndex(buttons: HTMLButtonElement[], active: HTMLButtonElement): void {
    for (const btn of buttons) {
      btn.tabIndex = btn === active ? 0 : -1;
    }
  }

  private buildMenuContents(container: HTMLDivElement, items: ContextMenuItem[], opts: { level: "menu" | "submenu" }): void {
    for (const item of items) {
      if (item.type === "separator") {
        const sep = document.createElement("div");
        sep.className = "context-menu__separator";
        sep.setAttribute("role", "separator");
        container.appendChild(sep);
        continue;
      }

      const enabled = item.enabled ?? true;
      const isSubmenu = opts.level === "menu" && item.type === "submenu";
      // Only one-level submenus are supported; if a submenu item itself contains a
      // submenu, render it as disabled text.
      const shouldDisable = !enabled || (opts.level === "submenu" && item.type === "submenu");

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "context-menu__item";
      // Intentionally keep the native <button> role so Playwright e2e tests can
      // locate items via `getByRole("button")`. Do NOT override to role=menuitem.
      btn.tabIndex = -1;

      this.buttonItems.set(btn, item);

      btn.disabled = shouldDisable;
      if (shouldDisable) {
        btn.classList.add("context-menu__item--disabled");
        btn.setAttribute("aria-disabled", "true");
      }

      if (isSubmenu) {
        btn.dataset.contextMenuSubmenu = "true";
        btn.setAttribute("aria-haspopup", "menu");
        btn.setAttribute("aria-expanded", "false");
      }

      if (item.leading?.type === "swatch") {
        const swatch = document.createElement("span");
        swatch.className = "context-menu__leading context-menu__leading--swatch";
        swatch.setAttribute("aria-hidden", "true");
        const svgNs = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(svgNs, "svg");
        svg.setAttribute("viewBox", "0 0 14 14");
        svg.setAttribute("width", "14");
        svg.setAttribute("height", "14");
        svg.setAttribute("focusable", "false");
        svg.setAttribute("aria-hidden", "true");

        const rect = document.createElementNS(svgNs, "rect");
        rect.setAttribute("x", "0");
        rect.setAttribute("y", "0");
        rect.setAttribute("width", "14");
        rect.setAttribute("height", "14");
        rect.setAttribute("rx", "3");
        rect.setAttribute("ry", "3");

        const rawColor = String(item.leading.color ?? "").trim();
        const fill = rawColor.startsWith("--") ? `var(${rawColor}, none)` : rawColor;
        rect.setAttribute("fill", fill || "none");

        svg.appendChild(rect);
        swatch.appendChild(svg);
        btn.appendChild(swatch);
      }

      const label = document.createElement("span");
      label.className = "context-menu__label";
      label.textContent = item.label;
      btn.appendChild(label);

      if (item.shortcut) {
        const shortcut = document.createElement("span");
        shortcut.className = "context-menu__shortcut";
        shortcut.setAttribute("aria-hidden", "true");
        shortcut.textContent = item.shortcut;
        btn.appendChild(shortcut);
      }

      if (isSubmenu) {
        const arrow = document.createElement("span");
        arrow.className = "context-menu__submenu-arrow";
        arrow.setAttribute("aria-hidden", "true");
        arrow.textContent = "›";
        btn.appendChild(arrow);

        btn.addEventListener("mouseenter", () => {
          if (btn.disabled) return;
          this.openSubmenu(btn, item.items, { focus: false });
        });
        btn.addEventListener("click", (e) => {
          // Keep the menu open and show the submenu.
          e.preventDefault();
          if (btn.disabled) return;
          this.openSubmenu(btn, item.items, { focus: true });
        });
      } else if (item.type === "item") {
        if (opts.level === "menu") {
          btn.addEventListener("mouseenter", () => {
            // Moving onto a non-submenu item should close any open submenu.
            this.closeSubmenu();
          });
        }
        btn.addEventListener("click", () => {
          if (btn.disabled) return;
          this.close();
          void Promise.resolve(item.onSelect()).catch((err) => {
            console.error("Context menu action failed:", err);
          });
        });
      }

      container.appendChild(btn);
    }
  }

  private openSubmenu(parent: HTMLButtonElement, items: ContextMenuItem[], opts: { focus: boolean }): void {
    if (!this.isShown) return;
    if (parent.disabled) return;

    // Rebuild submenu each time so it stays in sync with the parent items.
    this.closeSubmenu();

    parent.classList.add("context-menu__item--submenu-open");
    parent.setAttribute("aria-expanded", "true");
    this.submenuParent = parent;

    const submenu = document.createElement("div");
    submenu.className = "context-menu__submenu";
    submenu.setAttribute("role", "menu");
    submenu.setAttribute("aria-orientation", "vertical");
    submenu.tabIndex = -1;
    submenu.addEventListener("contextmenu", (e) => e.preventDefault());

    this.buildMenuContents(submenu, items, { level: "submenu" });

    this.overlay.appendChild(submenu);
    this.submenu = submenu;

    const parentRect = parent.getBoundingClientRect();
    submenu.style.left = `${parentRect.right}px`;
    submenu.style.top = `${parentRect.top}px`;

    // Position after layout.
    const rect = submenu.getBoundingClientRect();
    const margin = 8;
    let left = parentRect.right;
    let top = parentRect.top;

    if (left + rect.width + margin > window.innerWidth) {
      left = parentRect.left - rect.width;
    }
    if (top + rect.height + margin > window.innerHeight) {
      top = window.innerHeight - rect.height - margin;
    }

    left = Math.max(margin, Math.min(left, window.innerWidth - rect.width - margin));
    top = Math.max(margin, Math.min(top, window.innerHeight - rect.height - margin));

    submenu.style.left = `${left}px`;
    submenu.style.top = `${top}px`;

    if (opts.focus) {
      this.focusFirstItem(submenu);
    }
  }

  private closeSubmenu(): void {
    this.submenu?.remove();
    this.submenu = null;

    if (this.submenuParent) {
      this.submenuParent.classList.remove("context-menu__item--submenu-open");
      this.submenuParent.setAttribute("aria-expanded", "false");
      this.submenuParent = null;
    }
  }

  private positionMenu(x: number, y: number): void {
    // Set initial position (measured from the pointer).
    this.menu.style.left = `${x}px`;
    this.menu.style.top = `${y}px`;

    const rect = this.menu.getBoundingClientRect();
    const margin = 8;
    let left = x;
    let top = y;

    if (left + rect.width + margin > window.innerWidth) {
      left = x - rect.width;
    }
    if (top + rect.height + margin > window.innerHeight) {
      top = y - rect.height;
    }

    // Clamp to viewport.
    left = Math.max(margin, Math.min(left, window.innerWidth - rect.width - margin));
    top = Math.max(margin, Math.min(top, window.innerHeight - rect.height - margin));

    this.menu.style.left = `${left}px`;
    this.menu.style.top = `${top}px`;
  }
}
