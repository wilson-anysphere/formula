import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";

type ExtensionHostManagerLike = {
  ready: boolean;
  error: unknown;
  subscribe(listener: () => void): () => void;
};

export class CommandPaletteController {
  private readonly commandRegistry: CommandRegistry;
  private readonly ensureExtensionsLoaded: () => Promise<void>;
  private readonly extensionHostManager: ExtensionHostManagerLike | null;
  private readonly syncContributedCommands: (() => void) | null;
  private readonly onClose: (() => void) | null;
  private readonly onCommandError: ((err: unknown) => void) | null;

  private readonly overlay: HTMLDivElement;
  private readonly palette: HTMLDivElement;
  private readonly input: HTMLInputElement;
  private readonly list: HTMLUListElement;
  private readonly footer: HTMLDivElement;

  private paletteQuery = "";
  private paletteSelected = 0;
  private isOpen = false;

  private extensionsLoadPromise: Promise<void> | null = null;
  private extensionsLoadRequested = false;

  private readonly handleDocumentFocusIn = (e: FocusEvent): void => {
    if (this.overlay.hidden) return;
    const target = e.target as Node | null;
    if (!target) return;
    if (this.overlay.contains(target)) return;
    // If focus escapes the modal dialog, bring it back to the input.
    this.input.focus();
  };

  private readonly handleOverlayKeydown = (e: KeyboardEvent): void => {
    if (e.key !== "Tab") return;
    if (this.overlay.hidden) return;

    // Keep at least two tabbable targets (input + list) so we can implement a minimal focus trap.
    const focusable = [this.input, this.list].filter((el) => !el.hasAttribute("disabled"));
    if (focusable.length === 0) return;
    if (focusable.length === 1) {
      e.preventDefault();
      focusable[0]!.focus();
      return;
    }

    const first = focusable[0]!;
    const last = focusable[focusable.length - 1]!;
    const active = document.activeElement as HTMLElement | null;

    if (e.shiftKey) {
      if (!active || active === first) {
        e.preventDefault();
        last.focus();
      }
      return;
    }

    if (!active || active === last) {
      e.preventDefault();
      first.focus();
    }
  };

  private readonly handlePaletteKeydown = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      e.preventDefault();
      this.close();
      return;
    }

    const list = this.filteredCommands();
    if (e.key === "ArrowDown") {
      e.preventDefault();
      this.paletteSelected = list.length === 0 ? 0 : Math.min(list.length - 1, this.paletteSelected + 1);
      this.renderList();
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      this.paletteSelected = list.length === 0 ? 0 : Math.max(0, this.paletteSelected - 1);
      this.renderList();
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      const cmd = list[this.paletteSelected];
      if (!cmd) return;
      this.close();
      this.executeCommand(cmd.commandId);
    }
  };

  constructor(params: {
    commandRegistry: CommandRegistry;
    ensureExtensionsLoaded: () => Promise<void>;
    extensionHostManager?: ExtensionHostManagerLike | null;
    syncContributedCommands?: (() => void) | null;
    placeholder?: string;
    onClose?: (() => void) | null;
    onCommandError?: ((err: unknown) => void) | null;
  }) {
    this.commandRegistry = params.commandRegistry;
    this.ensureExtensionsLoaded = params.ensureExtensionsLoaded;
    this.extensionHostManager = params.extensionHostManager ?? null;
    this.syncContributedCommands = params.syncContributedCommands ?? null;
    this.onClose = params.onClose ?? null;
    this.onCommandError = params.onCommandError ?? null;

    this.overlay = document.createElement("div");
    this.overlay.className = "command-palette-overlay";
    this.overlay.hidden = true;
    this.overlay.setAttribute("role", "dialog");
    this.overlay.setAttribute("aria-modal", "true");

    this.palette = document.createElement("div");
    this.palette.className = "command-palette";
    this.palette.dataset.testid = "command-palette";

    this.input = document.createElement("input");
    this.input.className = "command-palette__input";
    this.input.dataset.testid = "command-palette-input";
    this.input.placeholder = params.placeholder ?? "";

    this.list = document.createElement("ul");
    this.list.className = "command-palette__list";
    this.list.dataset.testid = "command-palette-list";
    // Ensure there's always a second tabbable target for the focus trap.
    this.list.tabIndex = 0;

    this.footer = document.createElement("div");
    this.footer.className = "command-palette__footer";
    this.footer.dataset.testid = "command-palette-footer";

    this.palette.appendChild(this.input);
    this.palette.appendChild(this.list);
    this.palette.appendChild(this.footer);
    this.overlay.appendChild(this.palette);
    document.body.appendChild(this.overlay);

    this.overlay.addEventListener("click", (e) => {
      if (e.target === this.overlay) this.close();
    });
    this.overlay.addEventListener("keydown", this.handleOverlayKeydown);

    this.input.addEventListener("input", () => {
      this.paletteQuery = this.input.value;
      this.paletteSelected = 0;
      this.renderList();
    });

    this.input.addEventListener("keydown", this.handlePaletteKeydown);
    this.list.addEventListener("keydown", this.handlePaletteKeydown);

    // Keep the palette contents in sync with the command registry so newly-loaded
    // extension commands become visible without reopening.
    this.commandRegistry.subscribe(() => {
      if (!this.isOpen) return;
      this.renderList();
    });

    this.extensionHostManager?.subscribe(() => {
      if (!this.isOpen) return;
      this.renderFooter();
    });

    this.renderFooter();
  }

  open(): void {
    if (this.isOpen) {
      this.input.focus();
      this.input.select();
      return;
    }

    this.isOpen = true;
    this.paletteQuery = "";
    this.paletteSelected = 0;
    this.input.value = "";
    this.overlay.hidden = false;
    document.addEventListener("focusin", this.handleDocumentFocusIn);
    this.renderList();
    this.renderFooter();
    this.input.focus();
    this.input.select();

    this.triggerExtensionsLoad();
  }

  close(): void {
    if (!this.isOpen) return;
    this.isOpen = false;
    this.overlay.hidden = true;
    document.removeEventListener("focusin", this.handleDocumentFocusIn);
    this.onClose?.();
  }

  private triggerExtensionsLoad(): void {
    if (!this.extensionHostManager) return;

    if (this.extensionHostManager.ready) {
      this.renderFooter();
      return;
    }

    // Only trigger the extension host load once per app session. The underlying
    // `ensureExtensionsLoaded` is already idempotent, but keeping the controller
    // side-effect-free avoids redundant work on subsequent opens.
    if (this.extensionsLoadPromise) {
      this.extensionsLoadRequested = true;
      this.renderFooter();
      return;
    }

    this.extensionsLoadRequested = true;
    this.renderFooter();

    const promise = this.ensureExtensionsLoaded();
    this.extensionsLoadPromise = promise;
    void promise
      .then(() => {
        // If extensions loaded successfully, ensure the command registry is updated.
        this.syncContributedCommands?.();
      })
      .catch((err) => {
        // Non-blocking: built-in commands should still work.
        this.onCommandError?.(err);
      })
      .finally(() => {
        if (!this.isOpen) return;
        this.renderFooter();
        this.renderList();
      });
  }

  private executeCommand(commandId: string): void {
    void this.commandRegistry.executeCommand(commandId).catch((err) => {
      this.onCommandError?.(err);
    });
  }

  private displayLabel(cmd: CommandContribution): string {
    const category = typeof cmd.category === "string" && cmd.category.trim() !== "" ? cmd.category.trim() : null;
    if (category) return `${category}: ${cmd.title}`;
    return cmd.title;
  }

  private filteredCommands(): CommandContribution[] {
    const list = this.commandRegistry.listCommands();
    const q = this.paletteQuery.trim().toLowerCase();
    if (!q) return list;
    return list.filter((cmd) => {
      const label = this.displayLabel(cmd);
      const haystack = `${label} ${cmd.commandId} ${cmd.category ?? ""}`.toLowerCase();
      return haystack.includes(q);
    });
  }

  private renderFooter(): void {
    this.footer.dataset.variant = "";
    const mgr = this.extensionHostManager;
    if (!mgr) {
      this.footer.textContent = "";
      return;
    }

    if (mgr.error) {
      this.footer.dataset.variant = "warning";
      this.footer.textContent = `Failed to load extensions: ${String((mgr.error as any)?.message ?? mgr.error)}`;
      return;
    }

    if (this.extensionsLoadRequested && !mgr.ready) {
      this.footer.dataset.variant = "loading";
      this.footer.textContent = "Loading extensions…";
      return;
    }

    this.footer.textContent = "";
  }

  private renderList(): void {
    const list = this.filteredCommands();
    if (this.paletteSelected >= list.length) {
      this.paletteSelected = Math.max(0, list.length - 1);
    }
    this.list.replaceChildren();

    for (let i = 0; i < list.length; i += 1) {
      const cmd = list[i]!;
      const li = document.createElement("li");
      li.className = "command-palette__item";
      li.textContent = this.displayLabel(cmd);
      li.setAttribute("aria-selected", i === this.paletteSelected ? "true" : "false");
      li.addEventListener("mousedown", (e) => {
        // Prevent focus leaving the input before we run the command.
        e.preventDefault();
      });
      li.addEventListener("click", () => {
        this.close();
        this.executeCommand(cmd.commandId);
      });
      this.list.appendChild(li);
    }
  }
}
