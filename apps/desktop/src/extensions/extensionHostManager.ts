// Desktop-side glue for the browser extension host.
//
// This is intentionally lightweight: it wires the BrowserExtensionHost runtime
// into the desktop UI (toasts, panels, commands).

import { BrowserExtensionHost } from "../../../../packages/extension-host/src/browser/index.mjs";
import { createDesktopPermissionPrompt } from "./permissionPrompt.js";
import { validateExtensionManifest } from "../../../../packages/extension-host/src/browser/manifest.mjs";

import sampleHelloManifestJson from "../../../../extensions/sample-hello/package.json";
import sampleHelloEntrypointSource from "../../../../extensions/sample-hello/dist/extension.mjs?raw";
import e2eEventsManifestJson from "../../../../extensions/e2e-events/package.json";
import e2eEventsEntrypointSource from "../../../../extensions/e2e-events/dist/extension.mjs?raw";

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

  private _extensionApiModule: { url: string; revoke: () => void } | null = null;
  private readonly _loadedBuiltIns = new Map<string, { mainUrl: string; revoke: () => void }>();

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
    if (this._ready) return;

    let error: unknown = null;

    try {
      await this.loadBuiltInSampleHello();
      await this.loadBuiltInE2eEvents();
    } catch (err) {
      error = err;
    }

    // Always run `startup()` so the Extensions UI can render (even with zero loaded extensions).
    try {
      await this.host.startup();
    } catch (err) {
      error ??= err;
    }

    this._error = error;
    this._ready = true;
    this.emit();
  }

  private async loadBuiltInSampleHello(): Promise<void> {
    const manifest = validateExtensionManifest(sampleHelloManifestJson as any, {
      engineVersion: this.engineVersion,
      enforceEngine: true,
    }) as any;

    const extensionId = `${String(manifest.publisher)}.${String(manifest.name)}`;
    if (this._loadedBuiltIns.has(extensionId)) return;

    // Browser-loaded extensions cannot reliably import `@formula/extension-api` by bare specifier
    // (no import maps in workers, and production builds don't have Vite's `/@fs` rewriting).
    // Mirror the web marketplace loader: provide an in-memory shim module and rewrite imports.
    this._extensionApiModule ??= createModuleUrlFromText(EXTENSION_API_SHIM_SOURCE);

    const rewritten = rewriteEntrypointSource(sampleHelloEntrypointSource, {
      extensionApiUrl: this._extensionApiModule.url,
    });
    const { url: mainUrl, revoke } = createModuleUrlFromText(rewritten);

    await this.host.loadExtension({
      extensionId,
      extensionPath: `builtin://formula/extensions/${extensionId}/`,
      manifest,
      mainUrl,
    });

    this._loadedBuiltIns.set(extensionId, { mainUrl, revoke });
  }

  private async loadBuiltInE2eEvents(): Promise<void> {
    const manifest = validateExtensionManifest(e2eEventsManifestJson as any, {
      engineVersion: this.engineVersion,
      enforceEngine: true,
    }) as any;
    const extensionId = `${String(manifest.publisher)}.${String(manifest.name)}`;
    if (this._loadedBuiltIns.has(extensionId)) return;

    // E2E extension code imports `@formula/extension-api` and must run without Vite's import
    // rewriting in production/preview builds. Use the same shim + rewrite flow as sample-hello.
    this._extensionApiModule ??= createModuleUrlFromText(EXTENSION_API_SHIM_SOURCE);

    const rewritten = rewriteEntrypointSource(e2eEventsEntrypointSource, {
      extensionApiUrl: this._extensionApiModule.url,
    });
    const { url: mainUrl, revoke } = createModuleUrlFromText(rewritten);

    await this.host.loadExtension({
      extensionId,
      extensionPath: `builtin://formula/extensions/${extensionId}/`,
      manifest,
      mainUrl,
    });

    this._loadedBuiltIns.set(extensionId, { mainUrl, revoke });
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

  async getGrantedPermissions(extensionId: string): Promise<any> {
    return this.host.getGrantedPermissions(extensionId);
  }

  async revokePermission(extensionId: string, permission: string): Promise<void> {
    await this.host.revokePermissions(extensionId, [permission]);
    this.emit();
  }

  async resetPermissionsForExtension(extensionId: string): Promise<void> {
    await this.host.resetPermissions(extensionId);
    this.emit();
  }

  async resetAllPermissions(): Promise<void> {
    await this.host.resetAllPermissions();
    this.emit();
  }

  async executeCommand(commandId: string, ...args: any[]): Promise<any> {
    return this.host.executeCommand(commandId, ...args);
  }
}

// -- In-memory module helpers ----------------------------------------------------

// The extension worker (`packages/extension-host/src/browser/extension-worker.mjs`) eagerly imports
// the Formula extension API (workspace source), which initializes the runtime and installs the
// API object on `globalThis[Symbol.for("formula.extensionApi.api")]`.
//
// Built-in/browser-loaded extensions cannot import `@formula/extension-api` by bare specifier in
// production (no Vite dev-server transforms, no import maps in workers). Instead we provide a
// tiny, self-contained ESM shim that re-exports the already-initialized API object.
const EXTENSION_API_SHIM_SOURCE = `
const api = globalThis[Symbol.for("formula.extensionApi.api")];
if (!api) { throw new Error("@formula/extension-api runtime failed to initialize"); }
export const workbook = api.workbook;
export const sheets = api.sheets;
export const cells = api.cells;
export const commands = api.commands;
export const functions = api.functions;
export const dataConnectors = api.dataConnectors;
export const network = api.network;
export const clipboard = api.clipboard;
export const ui = api.ui;
export const storage = api.storage;
export const config = api.config;
export const events = api.events;
export const context = api.context;
export const __setTransport = api.__setTransport;
export const __setContext = api.__setContext;
export const __handleMessage = api.__handleMessage;
`;

function bytesToBase64(bytes: Uint8Array): string {
  // Node test environment fallback.
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  if (typeof btoa === "function") {
    let bin = "";
    for (const b of bytes) bin += String.fromCharCode(b);
    return btoa(bin);
  }
  throw new Error("Base64 encoding is not available in this runtime");
}

function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  return `data:${mime};base64,${bytesToBase64(bytes)}`;
}

function createModuleUrl(bytes: Uint8Array, mime = "text/javascript"): { url: string; revoke: () => void } {
  const isNodeRuntime = typeof process !== "undefined" && typeof (process as any)?.versions?.node === "string";

  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  if (
    !isNodeRuntime &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function" &&
    typeof Blob !== "undefined"
  ) {
    const url = URL.createObjectURL(new Blob([normalized], { type: mime }));
    return { url, revoke: () => URL.revokeObjectURL(url) };
  }

  const url = bytesToDataUrl(normalized, mime);
  return { url, revoke: () => {} };
}

function createModuleUrlFromText(source: string): { url: string; revoke: () => void } {
  const bytes = new TextEncoder().encode(source);
  return createModuleUrl(bytes);
}

function rewriteEntrypointSource(source: string, { extensionApiUrl }: { extensionApiUrl: string }): string {
  const rewritten = String(source ?? "")
    .replace(/from\s+["']@formula\/extension-api["']/g, `from "${extensionApiUrl}"`)
    .replace(/import\s+["']@formula\/extension-api["']/g, `import "${extensionApiUrl}"`)
    .replace(/from\s+["']formula["']/g, `from "${extensionApiUrl}"`)
    .replace(/import\s+["']formula["']/g, `import "${extensionApiUrl}"`);

  const specifiers = new Set<string>();
  const importRe = /\bimport\s+(?:[^"']*?\s+from\s+)?["']([^"']+)["']/g;
  const exportRe = /\bexport\s+(?:\*|\{[^}]*\})\s+from\s+["']([^"']+)["']/g;
  for (const re of [importRe, exportRe]) {
    for (;;) {
      const match = re.exec(rewritten);
      if (!match) break;
      specifiers.add(match[1]);
    }
  }

  for (const specifier of specifiers) {
    // The built-in loader only supports importing verified/bundled code. The only allowed
    // imports are other in-memory modules (blob/data URLs).
    if (!specifier.startsWith("blob:") && !specifier.startsWith("data:")) {
      throw new Error(
        `Unsupported import specifier "${specifier}". Built-in extensions must be bundled as a single-file entrypoint.`,
      );
    }
  }

  return rewritten;
}
