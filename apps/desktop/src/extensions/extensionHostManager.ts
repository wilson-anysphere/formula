// Desktop-side glue for the browser extension host.
//
// This is intentionally lightweight: it wires the BrowserExtensionHost runtime
// into the desktop UI (toasts, panels, commands).

import { BrowserExtensionHost } from "../../../../packages/extension-host/src/browser/index.mjs";
import { createDesktopPermissionPrompt } from "./permissionPrompt.js";

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

export class DesktopExtensionHostManager {
  readonly host: InstanceType<typeof BrowserExtensionHost>;
  readonly engineVersion: string;
  private readonly listeners = new Set<() => void>();
  private _ready = false;
  private _error: unknown = null;

  constructor(params: {
    engineVersion: string;
    spreadsheetApi: any;
    clipboardApi?: ClipboardApi;
    uiApi: ExtensionHostUiApi;
    permissionPrompt?: any;
  }) {
    const basePrompt = createDesktopPermissionPrompt();
    const permissionPrompt =
      params.permissionPrompt ??
      (async (req: unknown) => {
        // E2E / debugging hook: allow tests to override permission decisions without
        // threading a prompt implementation through the whole app.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const override = (globalThis as any).__formulaPermissionPrompt;
        if (typeof override === "function") {
          return await override(req);
        }
        // Fall back to the real desktop prompt UI (persists via PermissionManager).
        return await basePrompt(req as any);
      });

    this.engineVersion = String(params.engineVersion ?? "");
    this.host = new BrowserExtensionHost({
      engineVersion: this.engineVersion,
      spreadsheetApi: params.spreadsheetApi,
      clipboardApi: params.clipboardApi,
      uiApi: params.uiApi,
      permissionPrompt,
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
