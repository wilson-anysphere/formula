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

export type ContextMenuItem = ContextMenuActionItem | ContextMenuSeparator;

export type ContextMenuOpenOptions = {
  x: number;
  y: number;
  items: ContextMenuItem[];
};

export class ContextMenu {
  private readonly overlay: HTMLDivElement;
  private readonly menu: HTMLDivElement;
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
      this.close();
    });

    this.overlay = overlay;
    this.menu = menu;
  }

  open({ x, y, items }: ContextMenuOpenOptions): void {
    this.close();
    this.isShown = true;

    this.menu.replaceChildren();

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
        shortcut.textContent = item.shortcut;
        shortcut.style.color = "var(--text-secondary)";
        shortcut.style.fontSize = "12px";
        shortcut.style.flex = "none";
        btn.appendChild(shortcut);
      }

      btn.addEventListener("mousedown", (e) => {
        // Prevent focus from moving off the grid before we close/execute.
        e.preventDefault();
      });

      btn.addEventListener("click", () => {
        if (!enabled) return;
        this.close();
        void Promise.resolve(item.onSelect()).catch((err) => {
          console.error("Context menu action failed:", err);
        });
      });

      btn.addEventListener("mouseenter", () => {
        if (!enabled) return;
        btn.style.background = "var(--bg-hover)";
        btn.style.borderColor = "var(--border)";
      });
      btn.addEventListener("mouseleave", () => {
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

