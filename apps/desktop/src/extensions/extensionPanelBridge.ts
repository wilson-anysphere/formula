import type { PanelRegistry } from "../panels/panelRegistry.js";

type DockSide = "left" | "right" | "bottom";

type HostPanelRecord = {
  id: string;
  title: string;
  html: string;
  icon?: string | null;
  position?: DockSide | null;
  extensionId: string;
};

type ExtensionHostLike = {
  getPanel: (panelId: string) => HostPanelRecord | undefined;
  dispatchPanelMessage: (panelId: string, message: unknown) => void;
  activateView: (viewId: string) => Promise<void>;
};

type LayoutControllerLike = {
  openPanel: (panelId: string) => void;
  closePanel: (panelId: string) => void;
  layout: any;
};

function defaultFloatingRect(): { x: number; y: number; width: number; height: number } {
  return { x: 120, y: 120, width: 520, height: 640 };
}

export class ExtensionPanelBridge {
  private readonly host: ExtensionHostLike;
  private readonly panelRegistry: PanelRegistry;
  private readonly layoutController: LayoutControllerLike;

  private readonly listeners = new Map<string, Set<() => void>>();
  private readonly iframes = new Map<string, HTMLIFrameElement>();
  private readonly windowToPanel = new WeakMap<object, string>();
  private readonly pendingActivations = new Map<string, Promise<void>>();
  private readonly queuedMessages = new Map<string, unknown[]>();

  constructor(params: { host: ExtensionHostLike; panelRegistry: PanelRegistry; layoutController: LayoutControllerLike }) {
    this.host = params.host;
    this.panelRegistry = params.panelRegistry;
    this.layoutController = params.layoutController;

    window.addEventListener("message", this.handleWindowMessage);
  }

  dispose(): void {
    window.removeEventListener("message", this.handleWindowMessage);
    this.listeners.clear();
    this.iframes.clear();
    this.pendingActivations.clear();
    this.queuedMessages.clear();
  }

  getPanelHtml(panelId: string): string {
    return this.host.getPanel(panelId)?.html ?? "";
  }

  subscribe(panelId: string, listener: () => void): () => void {
    const id = String(panelId);
    let set = this.listeners.get(id);
    if (!set) {
      set = new Set();
      this.listeners.set(id, set);
    }
    set.add(listener);
    return () => {
      const s = this.listeners.get(id);
      if (!s) return;
      s.delete(listener);
      if (s.size === 0) this.listeners.delete(id);
    };
  }

  connect(panelId: string, iframe: HTMLIFrameElement): void {
    const id = String(panelId);
    this.iframes.set(id, iframe);
    const win = iframe.contentWindow;
    if (win) this.windowToPanel.set(win as any, id);

    const queued = this.queuedMessages.get(id);
    if (queued && queued.length > 0) {
      this.queuedMessages.delete(id);
      for (const message of queued) {
        this.postMessageToIframe(id, message);
      }
    }

    void this.activateView(id).catch(() => {});
  }

  disconnect(panelId: string, iframe: HTMLIFrameElement): void {
    const id = String(panelId);
    const existing = this.iframes.get(id);
    if (existing !== iframe) return;
    this.iframes.delete(id);
  }

  async activateView(viewId: string): Promise<void> {
    const id = String(viewId);
    const existing = this.pendingActivations.get(id);
    if (existing) return existing;

    const promise = this.host
      .activateView(id)
      .catch(() => {
        // ignore activation failures; the UI will show an empty panel instead.
      })
      .finally(() => {
        this.pendingActivations.delete(id);
      });
    this.pendingActivations.set(id, promise);
    return promise;
  }

  onPanelCreated(panel: HostPanelRecord): void {
    const panelId = String(panel.id);
    const owner = String(panel.extensionId);

    const existing = this.panelRegistry.get(panelId);
    const source = existing?.source;

    if (!existing) {
      this.panelRegistry.registerPanel(
        panelId,
        {
          title: panel.title,
          icon: panel.icon ?? null,
          defaultDock: (panel.position ?? "right") as DockSide,
          defaultFloatingRect: defaultFloatingRect(),
          source: { kind: "extension", extensionId: owner, contributed: false },
        },
        { owner },
      );
    } else if (source?.kind === "extension" && source.extensionId === owner && panel.position) {
      // Respect createPanel({ position }) overrides for contributed panels.
      this.panelRegistry.registerPanel(panelId, { ...existing, defaultDock: panel.position }, { owner, overwrite: true });
    }

    try {
      this.layoutController.openPanel(panelId);
    } catch {
      // ignore layout issues (e.g. missing registry entry)
    }

    this.emit(panelId);
  }

  onPanelHtmlUpdated(panelId: string): void {
    this.emit(panelId);
  }

  onPanelMessage(panelId: string, message: unknown): void {
    const id = String(panelId);
    if (!this.postMessageToIframe(id, message)) {
      const queue = this.queuedMessages.get(id) ?? [];
      queue.push(message);
      this.queuedMessages.set(id, queue);
    }
  }

  onPanelDisposed(panelId: string): void {
    const id = String(panelId);
    const def = this.panelRegistry.get(id);
    const source = def?.source;

    try {
      this.layoutController.closePanel(id);
    } catch {
      // ignore
    }

    if (source?.kind === "extension" && source.contributed === false) {
      this.panelRegistry.unregisterPanel(id, { owner: source.extensionId });
    }

    this.iframes.delete(id);
    this.queuedMessages.delete(id);
    this.emit(id);
  }

  private emit(panelId: string): void {
    const set = this.listeners.get(String(panelId));
    if (!set) return;
    for (const listener of [...set]) {
      try {
        listener();
      } catch {
        // ignore
      }
    }
  }

  private postMessageToIframe(panelId: string, message: unknown): boolean {
    const iframe = this.iframes.get(panelId);
    const win = iframe?.contentWindow;
    if (!win) return false;
    try {
      win.postMessage(message, "*");
      return true;
    } catch {
      return false;
    }
  }

  private handleWindowMessage = (event: MessageEvent): void => {
    const source = event.source as any;
    if (!source) return;
    const panelId = this.windowToPanel.get(source);
    if (!panelId) return;

    try {
      this.host.dispatchPanelMessage(panelId, event.data);
    } catch {
      // ignore
    }
  };
}
