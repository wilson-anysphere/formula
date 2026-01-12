export type ContextMenuSeparator = { type: "separator" };

export type ContextMenuActionItem = {
  type: "item";
  label: string;
  enabled?: boolean;
  shortcut?: string;
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
  items: ContextMenuItem[];
};

export type ContextMenuItem = ContextMenuActionItem | ContextMenuSubmenuItem | ContextMenuSeparator;

export type ContextMenuOpenOptions = {
  x: number;
  y: number;
  items: ContextMenuItem[];
};

export class ContextMenu {
  private readonly overlay: HTMLDivElement;
  private readonly menu: HTMLDivElement;
  private submenu: HTMLDivElement | null = null;
  private submenuParent: HTMLButtonElement | null = null;
  private isShown = false;
  private readonly onClose: (() => void) | null;
  private keydownListener: ((e: KeyboardEvent) => void) | null = null;

  constructor(options: { onClose?: () => void } = {}) {
    this.onClose = options.onClose ?? null;

    const overlay = document.createElement("div");
    overlay.dataset.testid = "context-menu";
    overlay.style.position = "fixed";
    overlay.style.inset = "0";
    overlay.style.display = "none";
    overlay.style.zIndex = "900";
    // Use a transparent overlay so outside clicks can dismiss the menu without
    // triggering underlying UI clicks.
    overlay.style.background = "transparent";

    const menu = document.createElement("div");
    menu.style.position = "absolute";
    menu.style.display = "flex";
    menu.style.flexDirection = "column";
    menu.style.minWidth = "220px";
    menu.style.maxWidth = "360px";
    menu.style.maxHeight = "calc(100vh - 16px)";
    menu.style.overflowY = "auto";
    menu.style.padding = "6px";
    menu.style.borderRadius = "10px";
    menu.style.border = "1px solid var(--border)";
    menu.style.background = "var(--dialog-bg)";
    menu.style.boxShadow = "var(--dialog-shadow)";
    menu.style.zIndex = "901";

    overlay.appendChild(menu);
    document.body.appendChild(overlay);

    overlay.addEventListener("click", (e) => {
      if (!this.isShown) return;
      const target = e.target as Node | null;
      if (!target) return;
      if (menu.contains(target)) return;
      if (this.submenu && this.submenu.contains(target)) return;
      this.close();
    });

    this.overlay = overlay;
    this.menu = menu;
  }

  open({ x, y, items }: ContextMenuOpenOptions): void {
    this.close();
    this.isShown = true;

    this.menu.replaceChildren();
    this.closeSubmenu();

    for (const item of items) {
      if (item.type === "separator") {
        const sep = document.createElement("div");
        sep.setAttribute("role", "separator");
        sep.style.height = "1px";
        sep.style.margin = "6px 6px";
        sep.style.background = "var(--border)";
        this.menu.appendChild(sep);
        continue;
      }

      const enabled = item.enabled ?? true;

      const btn = document.createElement("button");
      btn.type = "button";
      btn.disabled = !enabled;

      btn.style.display = "flex";
      btn.style.alignItems = "center";
      btn.style.justifyContent = "space-between";
      btn.style.gap = "16px";
      btn.style.width = "100%";
      btn.style.textAlign = "left";
      btn.style.padding = "8px 10px";
      btn.style.borderRadius = "8px";
      btn.style.border = "1px solid transparent";
      btn.style.background = "transparent";
      btn.style.color = enabled ? "var(--text-primary)" : "var(--text-secondary)";
      btn.style.cursor = enabled ? "pointer" : "default";

      const label = document.createElement("span");
      label.textContent = item.label;
      label.style.flex = "1";
      label.style.minWidth = "0";
      label.style.overflow = "hidden";
      label.style.textOverflow = "ellipsis";
      label.style.whiteSpace = "nowrap";
      btn.appendChild(label);

      if (item.shortcut) {
        const shortcut = document.createElement("span");
        shortcut.setAttribute("aria-hidden", "true");
        shortcut.textContent = item.shortcut;
        shortcut.style.color = "var(--text-secondary)";
        shortcut.style.fontSize = "12px";
        shortcut.style.flex = "none";
        btn.appendChild(shortcut);
      }

      if (item.type === "submenu") {
        const arrow = document.createElement("span");
        arrow.textContent = "›";
        arrow.style.color = "var(--text-secondary)";
        arrow.style.fontSize = "14px";
        arrow.style.flex = "none";
        btn.appendChild(arrow);
      }

      btn.addEventListener("mousedown", (e) => {
        // Prevent focus from moving off the grid before we close/execute.
        e.preventDefault();
      });

      if (item.type === "submenu") {
        const openSub = () => {
          if (!enabled) return;
          this.openSubmenu(btn, item.items);
        };
        btn.addEventListener("mouseenter", openSub);
        btn.addEventListener("click", (e) => {
          e.preventDefault();
          openSub();
        });
      } else {
        btn.addEventListener("click", () => {
          if (!enabled) return;
          this.close();
          void Promise.resolve(item.onSelect()).catch((err) => {
            console.error("Context menu action failed:", err);
          });
        });
      }

      btn.addEventListener("mouseenter", () => {
        if (!enabled) return;
        if (item.type !== "submenu") {
          // Moving onto a non-submenu item should close any open submenu.
          this.closeSubmenu();
        }
        btn.style.background = "var(--bg-hover)";
        btn.style.borderColor = "var(--border)";
      });
      btn.addEventListener("mouseleave", () => {
        // If this is the active submenu parent, keep it highlighted while the
        // submenu is visible.
        if (this.submenu && this.submenuParent === btn) return;
        btn.style.background = "transparent";
        btn.style.borderColor = "transparent";
      });

      this.menu.appendChild(btn);
    }

    this.overlay.style.display = "block";
    this.positionMenu(x, y);

    this.keydownListener = (e) => {
      if (e.key !== "Escape") return;
      e.preventDefault();
      this.close();
    };
    window.addEventListener("keydown", this.keydownListener, true);
  }

  close(): void {
    if (!this.isShown) return;
    this.isShown = false;
    this.overlay.style.display = "none";
    this.menu.replaceChildren();
    this.closeSubmenu();

    if (this.keydownListener) {
      window.removeEventListener("keydown", this.keydownListener, true);
      this.keydownListener = null;
    }

    try {
      this.onClose?.();
    } catch {
      // ignore
    }
  }

  private openSubmenu(parent: HTMLButtonElement, items: ContextMenuItem[]): void {
    if (!this.isShown) return;

    if (this.submenuParent && this.submenuParent !== parent) {
      // Clear hover state of the previous parent.
      this.submenuParent.style.background = "transparent";
      this.submenuParent.style.borderColor = "transparent";
    }

    // Rebuild submenu each time so it stays in sync with the parent items.
    this.closeSubmenu({ keepParent: true });
    this.submenuParent = parent;

    const submenu = document.createElement("div");
    submenu.style.position = "absolute";
    submenu.style.display = "flex";
    submenu.style.flexDirection = "column";
    submenu.style.minWidth = "200px";
    submenu.style.maxWidth = "360px";
    submenu.style.maxHeight = "calc(100vh - 16px)";
    submenu.style.overflowY = "auto";
    submenu.style.padding = "6px";
    submenu.style.borderRadius = "10px";
    submenu.style.border = "1px solid var(--border)";
    submenu.style.background = "var(--dialog-bg)";
    submenu.style.boxShadow = "var(--dialog-shadow)";
    submenu.style.zIndex = "902";

    for (const item of items) {
      if (item.type === "separator") {
        const sep = document.createElement("div");
        sep.setAttribute("role", "separator");
        sep.style.height = "1px";
        sep.style.margin = "6px 6px";
        sep.style.background = "var(--border)";
        submenu.appendChild(sep);
        continue;
      }

      // Only one-level submenus are supported; if a submenu item itself contains
      // a submenu, render it as disabled text.
      const enabled = item.enabled ?? true;

      const btn = document.createElement("button");
      btn.type = "button";
      btn.disabled = !enabled || item.type === "submenu";

      btn.style.display = "flex";
      btn.style.alignItems = "center";
      btn.style.justifyContent = "space-between";
      btn.style.gap = "16px";
      btn.style.width = "100%";
      btn.style.textAlign = "left";
      btn.style.padding = "8px 10px";
      btn.style.borderRadius = "8px";
      btn.style.border = "1px solid transparent";
      btn.style.background = "transparent";
      btn.style.color = btn.disabled ? "var(--text-secondary)" : "var(--text-primary)";
      btn.style.cursor = btn.disabled ? "default" : "pointer";

      const label = document.createElement("span");
      label.textContent = item.label;
      label.style.flex = "1";
      label.style.minWidth = "0";
      label.style.overflow = "hidden";
      label.style.textOverflow = "ellipsis";
      label.style.whiteSpace = "nowrap";
      btn.appendChild(label);

      if (item.shortcut) {
        const shortcut = document.createElement("span");
        shortcut.textContent = item.shortcut;
        shortcut.style.color = "var(--text-secondary)";
        shortcut.style.fontSize = "12px";
        shortcut.style.flex = "none";
        btn.appendChild(shortcut);
      }

      btn.addEventListener("mousedown", (e) => {
        e.preventDefault();
      });

      if (item.type === "item") {
        btn.addEventListener("click", () => {
          if (!enabled) return;
          this.close();
          void Promise.resolve(item.onSelect()).catch((err) => {
            console.error("Context menu action failed:", err);
          });
        });
      }

      btn.addEventListener("mouseenter", () => {
        if (btn.disabled) return;
        btn.style.background = "var(--bg-hover)";
        btn.style.borderColor = "var(--border)";
      });
      btn.addEventListener("mouseleave", () => {
        btn.style.background = "transparent";
        btn.style.borderColor = "transparent";
      });

      submenu.appendChild(btn);
    }

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
  }

  private closeSubmenu(options: { keepParent?: boolean } = {}): void {
    this.submenu?.remove();
    this.submenu = null;

    if (this.submenuParent && !options.keepParent) {
      this.submenuParent.style.background = "transparent";
      this.submenuParent.style.borderColor = "transparent";
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
