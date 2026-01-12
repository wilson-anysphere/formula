// Desktop-side glue for the browser extension host.
//
// This is intentionally lightweight: it wires the BrowserExtensionHost runtime
// into the desktop UI (toasts, panels, commands).

import { BrowserExtensionHost } from "../../../../packages/extension-host/src/browser/index.mjs";
import { createDesktopPermissionPrompt } from "./permissionPrompt.js";
import { validateExtensionManifest } from "../../../../packages/extension-host/src/browser/manifest.mjs";

import { MarketplaceClient, WebExtensionManager } from "@formula/extension-marketplace";
import { getMarketplaceBaseUrl } from "../panels/marketplace/getMarketplaceBaseUrl.ts";
import { verifyExtensionPackageV2Desktop } from "../panels/marketplace/verifyExtensionPackageV2Desktop.ts";

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

type TaintedRange = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

type ClipboardWriteGuard = (params: {
  extensionId: string;
  taintedRanges: TaintedRange[];
}) => Promise<void> | void;

export class DesktopExtensionHostManager {
  readonly host: InstanceType<typeof BrowserExtensionHost>;
  readonly engineVersion: string;
  private readonly listeners = new Set<() => void>();
  private readonly uiApi: ExtensionHostUiApi;
  private _ready = false;
  private _error: unknown = null;
  private _loadPromise: Promise<void> | null = null;

  private _extensionApiModule: { url: string; revoke: () => void } | null = null;
  private readonly _loadedBuiltIns = new Map<string, { mainUrl: string; revoke: () => void }>();
  private _marketplaceClient: MarketplaceClient | null = null;
  private _marketplaceExtensionManager: WebExtensionManager | null = null;

  constructor(params: {
    engineVersion: string;
    spreadsheetApi: any;
    clipboardApi?: ClipboardApi;
    clipboardWriteGuard?: ClipboardWriteGuard;
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
        // The built-in `formula.e2e-events` extension is only used by our Playwright suite to
        // validate the formula.events bridge. It should never block the desktop UI (or other
        // extensions) on an interactive permission prompt.
        //
        // Auto-accept its permission requests so non-extension-focused e2e tests don't flake
        // when the Extensions panel is opened (lazy host boot) and the extension activates on
        // `onStartupFinished`.
        const extensionId = typeof (req as any)?.extensionId === "string" ? String((req as any).extensionId) : "";
        if (extensionId === "formula.e2e-events") return true;
        // Fall back to the real desktop prompt UI (persists via PermissionManager).
        return await basePrompt(req as any);
      });

    this.engineVersion = String(params.engineVersion ?? "");
    this.uiApi = params.uiApi;
    this.host = new BrowserExtensionHost({
      engineVersion: this.engineVersion,
      spreadsheetApi: params.spreadsheetApi,
      clipboardApi: params.clipboardApi,
      clipboardWriteGuard: params.clipboardWriteGuard,
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

  getMarketplaceClient(): MarketplaceClient {
    if (!this._marketplaceClient) {
      this._marketplaceClient = new MarketplaceClient({ baseUrl: getMarketplaceBaseUrl() });
    }
    return this._marketplaceClient;
  }

  getMarketplaceExtensionManager(): WebExtensionManager {
    if (!this._marketplaceExtensionManager) {
      this._marketplaceExtensionManager = new WebExtensionManager({
        marketplaceClient: this.getMarketplaceClient(),
        host: this.host as any,
        engineVersion: this.engineVersion,
        verifyPackage: verifyExtensionPackageV2Desktop,
      });
    }
    return this._marketplaceExtensionManager;
  }

  async loadBuiltInExtensions(): Promise<void> {
    if (this._ready) return;
    if (this._loadPromise) {
      await this._loadPromise;
      return;
    }

    this._loadPromise = (async () => {
      let error: unknown = null;

      try {
        await this.loadBuiltInSampleHello();
        // The e2e-events helper extension is for Playwright coverage only; do not ship it in
        // production builds.
        if (import.meta.env.DEV) {
          await this.loadBuiltInE2eEvents();
        }
      } catch (err) {
        error = err;
      }
      // Always run `startup()` so the Extensions UI can render (even with zero loaded extensions).
      //
      // Desktop also supports IndexedDB-installed extensions via WebExtensionManager. Use the
      // manager's `loadAllInstalled()` boot helper so `onStartupFinished` + the initial
      // `workbookOpened` event behave consistently for built-in *and* installed extensions.
      try {
        await this.getMarketplaceExtensionManager().loadAllInstalled({
          onExtensionError: ({ id, version, error }) => {
            const message = `Failed to load extension ${id}@${version}: ${String((error as any)?.message ?? error)}`;
            // eslint-disable-next-line no-console
            console.error(`[formula][desktop] ${message}`);
            try {
              void this.uiApi.showMessage?.(message, "error");
            } catch {
              // ignore UI errors
            }
          },
        });
      } catch (err) {
        // Best-effort: surface startup failures but keep going so built-in extensions can run.
        try {
          const msg = String((err as any)?.message ?? err);
          void this.uiApi.showMessage?.(`Failed to load installed extensions: ${msg}`, "error");
        } catch {
          // ignore
        }
        // Fallback: if loading installed extensions fails (IndexedDB unavailable/corrupted),
        // still attempt to start the host so built-in extensions can run.
        try {
          await this.host.startup();
        } catch (startupErr) {
          error ??= startupErr;
        }
      }

      this._error = error;
      this._ready = true;
      this.emit();
    })().finally(() => {
      // If the load flow errors before setting `_ready`, allow a subsequent retry.
      this._loadPromise = null;
    });

    await this._loadPromise;
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

    // The e2e-events extension is an internal Playwright harness that activates on
    // startup and writes its last-seen event payloads via `formula.storage.*`.
    //
    // In e2e/headless runs there is no user to respond to the interactive permission
    // prompt, so pre-grant `storage` to keep the UI unblocked. (The extension has no
    // UI contributions; it only writes diagnostic state for tests.)
    pregrantPermissions(extensionId, ["storage"]);

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

  /**
   * Notify subscribers that extension contributions may have changed (e.g. after
   * installing/uninstalling an extension).
   */
  notifyDidChange(): void {
    this.emit();
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
    // Clears all stored grants for a single extension, forcing the next privileged API
    // call to re-prompt for permissions.
    await this.host.resetPermissions(extensionId);
    // Resetting grants should behave like a fresh install: drop any in-memory registrations that
    // were established under the old permission set so the next activation re-prompts as needed.
    try {
      await this.host.reloadExtension(extensionId);
    } catch {
      // ignore reload failures (extension may not be loaded yet)
    }
    this.emit();
  }

  async resetAllPermissions(): Promise<void> {
    await this.host.resetAllPermissions();
    // Like per-extension reset, reload any already-loaded extensions so their workers are reset
    // and will re-request permissions on next activation.
    try {
      const exts = this.host.listExtensions?.() ?? [];
      for (const ext of exts as any[]) {
        const id = typeof ext?.id === "string" ? ext.id : null;
        if (!id) continue;
        try {
          // eslint-disable-next-line no-await-in-loop
          await this.host.reloadExtension(id);
        } catch {
          // ignore individual failures
        }
      }
    } catch {
      // ignore
    }
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

function isNodeRuntime(): boolean {
  // Avoid relying on Node-only `process.versions` fields (some bundlers/polyfills might provide
  // them inconsistently). We only need a best-effort detector to avoid `blob:` module URLs in
  // Node-based test runners.
  const proc = (globalThis as any).process as any;
  if (!proc || typeof proc !== "object") return false;
  if (proc.release && typeof proc.release === "object" && proc.release.name === "node") return true;
  return typeof proc.version === "string" && proc.version.startsWith("v");
}

function createModuleUrl(bytes: Uint8Array, mime = "text/javascript"): { url: string; revoke: () => void } {
  const nodeRuntime = isNodeRuntime();

  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  if (
    !nodeRuntime &&
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

function pregrantPermissions(extensionId: string, permissions: string[]): void {
  try {
    if (typeof localStorage === "undefined") return;
    const key = "formula.extensionHost.permissions";
    const existing = (() => {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : {};
      } catch {
        return {};
      }
    })();
    const record = { ...(existing?.[extensionId] ?? {}) };
    for (const perm of permissions) {
      record[String(perm)] = true;
    }
    existing[extensionId] = record;
    localStorage.setItem(key, JSON.stringify(existing));
  } catch {
    // ignore
  }
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
