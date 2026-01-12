// Desktop-side glue for the browser extension host.
//
// This is intentionally lightweight: it wires the BrowserExtensionHost runtime
// into the desktop UI (toasts, panels, commands).

import { BrowserExtensionHost } from "../../../../packages/extension-host/src/browser/index.mjs";

type ExtensionHostUiApi = {
  showMessage?: (message: string, type?: string) => Promise<void> | void;
  showQuickPick?: (items: any[], options?: any) => Promise<any>;
  showInputBox?: (options?: any) => Promise<any>;
  onPanelCreated?: (panel: any) => void;
  onPanelHtmlUpdated?: (panelId: string, html: string) => void;
  onPanelMessage?: (panelId: string, message: unknown) => void;
  onPanelDisposed?: (panelId: string) => void;
};

type ClipboardApi = {
  readText: () => Promise<string>;
  writeText: (text: string) => Promise<void>;
};

type TauriClipboardApi = {
  readText?: () => Promise<string>;
  writeText?: (text: string) => Promise<void>;
};

function getTauriClipboardApi(): TauriClipboardApi | null {
  const clipboard = (globalThis as any)?.__TAURI__?.clipboard as TauriClipboardApi | undefined;
  return clipboard ?? null;
}

function createDesktopClipboardApi(uiApi: ExtensionHostUiApi): ClipboardApi {
  // Mirror the BrowserExtensionHost in-memory fallback so extensions always have a usable
  // clipboard even when the real Clipboard API is unavailable/permission-gated.
  let fallbackText = "";
  let warnedWriteFailure = false;
  let warnedReadFailure = false;

  const warnOnce = (kind: "read" | "write", err: unknown) => {
    const alreadyWarned = kind === "write" ? warnedWriteFailure : warnedReadFailure;
    if (alreadyWarned) return;
    if (kind === "write") warnedWriteFailure = true;
    else warnedReadFailure = true;

    const message =
      kind === "write"
        ? "Could not write to the system clipboard. Falling back to an in-memory clipboard for extensions."
        : "Could not read from the system clipboard. Falling back to an in-memory clipboard for extensions.";

    try {
      uiApi.showMessage?.(message, "warning");
    } catch {
      // ignore
    }

    // eslint-disable-next-line no-console
    console.warn(message, err);
  };

  return {
    async readText() {
      const navClipboard = globalThis.navigator?.clipboard;
      let lastError: unknown = null;

      if (typeof navClipboard?.readText === "function") {
        try {
          const text = await navClipboard.readText();
          fallbackText = String(text ?? "");
          return fallbackText;
        } catch (err) {
          lastError = err;
        }
      }

      const tauriClipboard = getTauriClipboardApi();
      if (typeof tauriClipboard?.readText === "function") {
        try {
          const text = await tauriClipboard.readText();
          fallbackText = String(text ?? "");
          return fallbackText;
        } catch (err) {
          lastError = err;
        }
      }

      warnOnce("read", lastError ?? new Error("Clipboard API not available"));
      return fallbackText;
    },

    async writeText(text: string) {
      const value = String(text ?? "");
      fallbackText = value;

      const navClipboard = globalThis.navigator?.clipboard;
      let lastError: unknown = null;
      if (typeof navClipboard?.writeText === "function") {
        try {
          await navClipboard.writeText(value);
          return;
        } catch (err) {
          lastError = err;
        }
      }

      const tauriClipboard = getTauriClipboardApi();
      if (typeof tauriClipboard?.writeText === "function") {
        try {
          await tauriClipboard.writeText(value);
          return;
        } catch (err) {
          lastError = err;
        }
      }

      warnOnce("write", lastError ?? new Error("Clipboard API not available"));
    },
  };
}

export class DesktopExtensionHostManager {
  readonly host: InstanceType<typeof BrowserExtensionHost>;
  private readonly listeners = new Set<() => void>();
  private _ready = false;
  private _error: unknown = null;

  constructor(params: {
    engineVersion: string;
    spreadsheetApi: any;
    uiApi: ExtensionHostUiApi;
    permissionPrompt?: any;
  }) {
    this.host = new BrowserExtensionHost({
      engineVersion: params.engineVersion,
      spreadsheetApi: params.spreadsheetApi,
      uiApi: params.uiApi,
      permissionPrompt: params.permissionPrompt ?? (async () => true),
      clipboardApi: createDesktopClipboardApi(params.uiApi),
    });
  }

  get ready(): boolean {
    return this._ready;
  }

  get error(): unknown {
    return this._error;
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of [...this.listeners]) {
      try {
        listener();
      } catch {
        // ignore
      }
    }
  }

  async loadBuiltInExtensions(): Promise<void> {
    try {
      const manifestUrl = new URL("../../../../extensions/sample-hello/package.json", import.meta.url).toString();
      await this.host.loadExtensionFromUrl(manifestUrl);
      await this.host.startup();
      this._ready = true;
      this.emit();
    } catch (err) {
      this._error = err;
      this._ready = true;
      this.emit();
    }
  }

  getContributedCommands(): any[] {
    return this.host.getContributedCommands();
  }

  getContributedPanels(): any[] {
    return this.host.getContributedPanels();
  }

  getContributedKeybindings(): any[] {
    return this.host.getContributedKeybindings();
  }

  getContributedMenu(menuId: string): any[] {
    return this.host.getContributedMenu(menuId);
  }

  async executeCommand(commandId: string, ...args: any[]): Promise<any> {
    return this.host.executeCommand(commandId, ...args);
  }
}
