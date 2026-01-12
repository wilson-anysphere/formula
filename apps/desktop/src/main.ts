import { SpreadsheetApp } from "./app/spreadsheetApp";
import type { SheetNameResolver } from "./sheet/sheetNameResolver";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/command-palette.css";
import "./styles/dialogs.css";
import "./styles/workspace.css";
import "./styles/charts-overlay.css";
import "./styles/scrollbars.css";
import "./styles/comments.css";
import "./styles/shell.css";
import "./styles/auditing.css";
import "./styles/format-cells-dialog.css";
import "./styles/context-menu.css";
import "./styles/conflicts.css";
import "./styles/macros-runner.css";

import React from "react";
import { createRoot } from "react-dom/client";

import { SheetTabStrip } from "./sheets/SheetTabStrip";

import { ThemeController } from "./theme/themeController.js";

import { mountRibbon } from "./ribbon/index.js";

import { computeSelectionFormatState } from "./ribbon/selectionFormatState.js";
import { setRibbonUiState } from "./ribbon/ribbonUiState.js";

import type { CellRange as GridCellRange } from "@formula/grid";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import { SecondaryGridView } from "./grid/splitView/secondaryGridView.js";
import { getPanelTitle, panelRegistry, PanelIds } from "./panels/panelRegistry.js";
import { createPanelBodyRenderer } from "./panels/panelBodyRenderer.js";
import { MacroRecorder, generatePythonMacro, generateTypeScriptMacro } from "./macro-recorder/index.js";
import { mountTitlebar } from "./titlebar/mountTitlebar.js";
import {
  renderMacroRunner,
  TauriMacroBackend,
  WebMacroBackend,
  wrapTauriMacroBackendWithUiContext,
  type MacroRunRequest,
  type MacroTrustDecision,
} from "./macros";
import { applyMacroCellUpdates } from "./macros/applyUpdates";
import { fireWorkbookBeforeCloseBestEffort, installVbaEventMacros } from "./macros/event_macros";
import { mountScriptEditorPanel } from "./panels/script-editor/index.js";
import { installUnsavedChangesPrompt } from "./document/index.js";
import type { DocumentController } from "./document/documentController.js";
import { DocumentControllerWorkbookAdapter } from "./scripting/documentControllerWorkbookAdapter.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize } from "./formatting/selectionSizeGuard.js";
import { registerFindReplaceShortcuts, FindReplaceController } from "./panels/find-replace/index.js";
import { t, tWithVars } from "./i18n/index.js";
import { getOpenFileFilters } from "./file_dialog_filters.js";
import { formatRangeAddress, parseRangeAddress } from "@formula/scripting";
import { normalizeFormulaTextOpt } from "@formula/engine";
import type { CollabSession } from "@formula/collab-session";
import * as Y from "yjs";
import { startWorkbookSync } from "./tauri/workbookSync";
import { TauriWorkbookBackend } from "./tauri/workbookBackend";
import * as nativeDialogs from "./tauri/nativeDialogs";
import { shellOpen } from "./tauri/shellOpen";
import { setTrayStatus } from "./tauri/trayStatus";
import { installUpdaterUi } from "./tauri/updaterUi";
import { notify } from "./tauri/notifications";
import { registerAppQuitHandlers, requestAppQuit } from "./tauri/appQuit";
import { checkForUpdatesFromCommandPalette } from "./tauri/updater.js";
import type { WorkbookInfo } from "@formula/workbook-backend";
import { chartThemeFromWorkbookPalette } from "./charts/theme";
import { parseA1Range, splitSheetQualifier } from "../../../packages/search/index.js";
import { refreshDefinedNameSignaturesFromBackend, refreshTableSignaturesFromBackend } from "./power-query/tableSignatures";
import { oauthBroker } from "./power-query/oauthBroker.js";
import {
  DesktopPowerQueryService,
  loadQueriesFromStorage,
  saveQueriesToStorage,
  setDesktopPowerQueryService,
} from "./power-query/service.js";
import { createPowerQueryRefreshStateStore } from "./power-query/refreshStateStore.js";
import { createClipboardProvider } from "./clipboard/platform/provider.js";
import { createDesktopDlpContext } from "./dlp/desktopDlp.js";
import { enforceClipboardCopy } from "./dlp/enforceClipboardCopy.js";
import { showInputBox, showQuickPick, showToast } from "./extensions/ui.js";
import { openFormatCellsDialog } from "./formatting/openFormatCellsDialog.js";
import { DesktopExtensionHostManager } from "./extensions/extensionHostManager.js";
import { ExtensionPanelBridge } from "./extensions/extensionPanelBridge.js";
import { ContextKeyService } from "./extensions/contextKeys.js";
import { resolveMenuItems } from "./extensions/contextMenus.js";
import { CELL_CONTEXT_MENU_ID } from "./extensions/menuIds.js";
import { buildContextMenuModel } from "./extensions/contextMenuModel.js";
import {
  buildCommandKeybindingDisplayIndex,
  getPrimaryCommandKeybindingDisplay,
  type ContributedKeybinding,
} from "./extensions/keybindings.js";
import { KeybindingService } from "./extensions/keybindingService.js";
import { deriveSelectionContextKeys } from "./extensions/selectionContextKeys.js";
import { CommandRegistry } from "./extensions/commandRegistry.js";
import { createCommandPalette } from "./command-palette/index.js";
import { registerBuiltinCommands } from "./commands/registerBuiltinCommands.js";
import { DEFAULT_GRID_LIMITS } from "./selection/selection.js";
import type { GridLimits, Range, SelectionState } from "./selection/types";
import { ContextMenu, type ContextMenuItem } from "./menus/contextMenu.js";
import { getPasteSpecialMenuItems } from "./clipboard/pasteSpecial.js";
import { WorkbookSheetStore, generateDefaultSheetName, validateSheetName } from "./sheets/workbookSheetStore";
import { startSheetStoreDocumentSync } from "./sheets/sheetStoreDocumentSync";
import { rewriteSheetNamesInFormula } from "./workbook/formulaRewrite";
import {
  applyAllBorders,
  applyNumberFormatPreset,
  NUMBER_FORMATS,
  setFillColor,
  setFontColor,
  setFontSize,
  setHorizontalAlign,
  toggleBold,
  toggleItalic,
  toggleUnderline,
  toggleWrap,
  type CellRange,
} from "./formatting/toolbar.js";
import { PageSetupDialog, type CellRange as PrintCellRange, type PageSetup } from "./print/index.js";
import {
  getDefaultSeedStoreStorage,
  readContributedPanelsSeedStore,
  removeSeedPanelsForExtension,
  seedPanelRegistryFromContributedPanelsSeedStore,
  setSeedPanelsForExtension,
} from "./extensions/contributedPanelsSeedStore.js";
import { builtinKeybindings as builtinKeybindingsCatalog } from "./commands/builtinKeybindings.js";

import sampleHelloManifest from "../../../extensions/sample-hello/package.json";
import { purgeLegacyDesktopLLMSettings } from "./ai/llm/desktopLLMClient.js";
import {
  installStartupTimingsListeners,
  markStartupTimeToInteractive,
  reportStartupWebviewLoaded,
} from "./tauri/startupMetrics.js";
import { openExternalHyperlink } from "./hyperlinks/openExternal.js";
import { clampUsedRange, resolveWorkbookLoadLimits } from "./workbook/load/clampUsedRange.js";

// Apply theme + reduced motion settings as early as possible to avoid rendering with
// default tokens before the user's preference is known.
const themeController = new ThemeController();
themeController.start();
window.addEventListener("unload", () => {
  try {
    themeController.stop();
  } catch {
    // Best-effort cleanup; ignore failures during teardown.
  }
});

// Startup performance instrumentation (no-op for web builds).
void (async () => {
  try {
    await installStartupTimingsListeners();
  } catch {
    // Best-effort; instrumentation should never block startup.
  }
  reportStartupWebviewLoaded();
})();

/**
 * SharedArrayBuffer requires cross-origin isolation (COOP/COEP). When we ship a
 * packaged Tauri build without it, Pyodide cannot use its Worker backend.
 *
 * Vite dev/preview servers configure COOP/COEP headers; the packaged app relies
 * on Tauri's asset protocol configuration. Fail loudly (but non-fatally) in
 * production Tauri builds so this doesn't ship silently.
 */
function warnIfMissingCrossOriginIsolationInTauriProd(): void {
  const isTauri = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
  if (!isTauri) return;
  if (!import.meta.env.PROD) return;

  const crossOriginIsolated = globalThis.crossOriginIsolated === true;
  const hasSharedArrayBuffer = typeof (globalThis as any).SharedArrayBuffer !== "undefined";
  if (crossOriginIsolated && hasSharedArrayBuffer) return;

  const details = `crossOriginIsolated=${String(globalThis.crossOriginIsolated)}, SharedArrayBuffer=${
    hasSharedArrayBuffer ? "present" : "missing"
  }`;

  console.error(
    `[formula][desktop] Cross-origin isolation missing in packaged Tauri build (${details}). ` +
      "This breaks SharedArrayBuffer and forces Pyodide to fall back to its main-thread backend. " +
      "Ensure the Tauri asset protocol sets COOP/COEP headers (see apps/desktop/README.md).",
  );

  // Keep it on-screen long enough that developers won't miss it.
  showToast(
    `Cross-origin isolation missing in packaged desktop build (${details}). Pyodide Worker backend disabled.`,
    "error",
    { timeoutMs: 60_000 },
  );
}

warnIfMissingCrossOriginIsolationInTauriProd();

let workbookSheetStore = new WorkbookSheetStore([{ id: "Sheet1", name: "Sheet1", visibility: "visible" }]);
const workbookSheetNames = new Map<string, string>();

function syncWorkbookSheetNamesFromSheetStore(): void {
  workbookSheetNames.clear();
  for (const sheet of workbookSheetStore.listAll()) {
    workbookSheetNames.set(sheet.id, sheet.name);
  }
}

syncWorkbookSheetNamesFromSheetStore();

// Task 13 adds this helper in collab mode. Declare it here so main.ts can
// consume it without taking a hard dependency on collab wiring being present.
interface SpreadsheetApp {
  getCollabSession?: () => CollabSession | null;
}

function installExternalLinkInterceptor(): void {
  if (typeof document === "undefined") return;

  const handler = (event: MouseEvent) => {
      if (event.defaultPrevented) return;

      const target = event.target as Element | null;
      if (!target || typeof target.closest !== "function") return;
      const anchor = target.closest("a[href]") as HTMLAnchorElement | null;
      if (!anchor) return;

      // Don't interfere with download links (e.g. blob exports like PDF).
      if (anchor.hasAttribute("download")) return;

      const href = anchor.getAttribute("href");
      if (typeof href !== "string" || href.trim() === "") return;

      // Only handle absolute URLs. Relative links are treated as internal navigation.
      if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(href)) return;

      // Always block javascript:/data: URL navigations.
      try {
        const parsed = new URL(href);
        const protocol = parsed.protocol.replace(":", "").toLowerCase();
        if (protocol === "javascript" || protocol === "data") {
          event.preventDefault();
          event.stopPropagation();
          return;
        }
      } catch {
        // Ignore invalid URLs.
        return;
      }

      // Only intercept when running under Tauri. In web builds, let the browser handle
      // normal navigation/new-tab behavior.
      const isTauri = Boolean((globalThis as any).__TAURI__);
      if (!isTauri) return;

      // Prevent the webview from navigating away; open through the OS instead.
      event.preventDefault();
      event.stopPropagation();

      void openExternalHyperlink(href, {
        shellOpen,
        confirmUntrustedProtocol: nativeDialogs.confirm,
      }).catch((err) => {
        console.error("Failed to open external link:", err);
      });
  };

  document.addEventListener("click", handler, { capture: true });
  // Middle-click on links in browsers opens a new tab via `auxclick` rather than `click`.
  // Intercept it too so desktop/Tauri always delegates to the OS browser.
  document.addEventListener(
    "auxclick",
    (event) => {
      if (event.button !== 1) return;
      handler(event);
    },
    { capture: true },
  );
}

installExternalLinkInterceptor();

const sheetNameResolver: SheetNameResolver = {
  getSheetNameById(id: string): string | null {
    const key = String(id ?? "").trim();
    if (!key) return null;
    return workbookSheetStore.getName(key) ?? null;
  },
  getSheetIdByName(name: string): string | null {
    const trimmed = String(name ?? "").trim();
    if (!trimmed) return null;
    return workbookSheetStore.resolveIdByName(trimmed) ?? null;
  },
};

// Cursor desktop no longer supports user-provided LLM settings, but legacy builds
// persisted provider + API keys in localStorage. Best-effort purge on startup so
// we don't leave stale secrets behind.
try {
  purgeLegacyDesktopLLMSettings();
} catch {
  // ignore
}

// Exposed to Playwright tests via `window.__formulaExtensionHostManager`.
let extensionHostManagerForE2e: DesktopExtensionHostManager | null = null;

type SheetActivatedEvent = { sheet: { id: string; name: string } };
const sheetActivatedListeners = new Set<(event: SheetActivatedEvent) => void>();

function emitSheetActivated(sheetId: string): void {
  const id = String(sheetId ?? "").trim();
  if (!id) return;
  const event: SheetActivatedEvent = { sheet: { id, name: workbookSheetStore.getName(id) ?? id } };
  for (const listener of [...sheetActivatedListeners]) {
    try {
      listener(event);
    } catch {
      // ignore listener errors
    }
  }
}
// Seed contributed panels early so layout persistence doesn't drop their ids before the
// extension host finishes loading installed extensions.
const contributedPanelsSeedStorage = getDefaultSeedStoreStorage();
if (contributedPanelsSeedStorage) {
  seedPanelRegistryFromContributedPanelsSeedStore(contributedPanelsSeedStorage, panelRegistry, {
    onError: (message, err) => {
      console.error(message, err);
      showToast(message, "error");
    },
  });
}

const sampleHelloExtensionId = `${(sampleHelloManifest as any).publisher}.${(sampleHelloManifest as any).name}`;
for (const panel of (sampleHelloManifest as any)?.contributes?.panels ?? []) {
  try {
    panelRegistry.registerPanel(
      String(panel.id),
      {
        title: String(panel.title ?? panel.id),
        icon: panel.icon ?? null,
        defaultDock: "right",
        defaultFloatingRect: { x: 140, y: 140, width: 520, height: 640 },
        source: { kind: "extension", extensionId: sampleHelloExtensionId, contributed: true },
      },
      { owner: sampleHelloExtensionId },
    );
  } catch (err) {
    console.error("Failed to seed built-in extension panel:", err);
    showToast(`Failed to seed extension panel: ${String(panel?.id ?? "")}`, "error");
  }
}

// Tauri/desktop state (declared early so panel wiring can reference them without TDZ errors).
type TauriInvoke = (cmd: string, args?: any) => Promise<any>;
let tauriBackend: TauriWorkbookBackend | null = null;
let activeWorkbook: WorkbookInfo | null = null;
let pendingBackendSync: Promise<void> = Promise.resolve();
let queuedInvoke: TauriInvoke | null = null;
let workbookSync: ReturnType<typeof startWorkbookSync> | null = null;
let rerenderLayout: (() => void) | null = null;
let vbaEventMacros: ReturnType<typeof installVbaEventMacros> | null = null;
let ribbonLayoutController: LayoutController | null = null;
let ensureExtensionsLoadedRef: (() => Promise<void>) | null = null;
let syncContributedCommandsRef: (() => void) | null = null;
let syncContributedPanelsRef: (() => void) | null = null;
let updateKeybindingsRef: (() => void) | null = null;

function toggleDockPanel(panelId: string): void {
  const controller = ribbonLayoutController;
  if (!controller) return;
  const placement = getPanelPlacement(controller.layout, panelId);
  if (placement.kind === "closed") controller.openPanel(panelId);
  else controller.closePanel(panelId);
}
let handleCloseRequestForRibbon: ((opts: { quit: boolean }) => Promise<void>) | null = null;

function installCollabStatusIndicator(app: unknown, element: HTMLElement): void {
  const abortController = new AbortController();

  const cleanup = (): void => {
    if (abortController.signal.aborted) return;
    abortController.abort();
  };

  window.addEventListener("unload", cleanup, { once: true });

  // If SpreadsheetApp exposes a `destroy()` method, wrap it so collab listeners detach
  // in tests / fast-refresh scenarios.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const anyApp = app as any;
  if (anyApp && typeof anyApp.destroy === "function") {
    const originalDestroy = anyApp.destroy.bind(anyApp) as () => void;
    anyApp.destroy = () => {
      cleanup();
      originalDestroy();
    };
  }

  let currentProvider: unknown = null;
  let providerStatus: string | null = null;
  let providerSynced: boolean | null = null;
  let currentOffline: unknown = null;
  let offlineWaitStarted = false;

  const detachProviderListeners = (provider: unknown): void => {
    if (!provider) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyProvider = provider as any;
    if (typeof anyProvider.off !== "function") return;
    try {
      anyProvider.off("status", onProviderStatus);
    } catch {
      // ignore
    }
    try {
      anyProvider.off("sync", onProviderSync);
    } catch {
      // ignore
    }
  };

  const attachProviderListeners = (provider: unknown): void => {
    if (!provider) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyProvider = provider as any;
    if (typeof anyProvider.on !== "function") return;
    try {
      anyProvider.on("status", onProviderStatus);
    } catch {
      // ignore
    }
    try {
      anyProvider.on("sync", onProviderSync);
    } catch {
      // ignore
    }
  };

  const getSession = (): unknown | null => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const maybeGetter = (app as any)?.getCollabSession as (() => unknown) | undefined;
    if (typeof maybeGetter !== "function") return null;
    try {
      return maybeGetter.call(app) ?? null;
    } catch {
      return null;
    }
  };

  const getDocId = (session: unknown): string => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const s = session as any;
    const direct = s?.docId ?? s?.doc_id ?? s?.id;
    if (typeof direct === "string" && direct.trim() !== "") return direct;
    const guid = s?.doc?.guid;
    if (typeof guid === "string" && guid.trim() !== "") return guid;
    return "unknown";
  };

  const onProviderStatus = (evt: unknown): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const status = (evt as any)?.status;
    providerStatus = typeof status === "string" ? status : null;
    render();
  };

  const onProviderSync = (isSynced: unknown): void => {
    providerSynced = Boolean(isSynced);
    render();
  };

  const render = (): void => {
    if (abortController.signal.aborted) return;

    const session = getSession();
    if (!session) {
      detachProviderListeners(currentProvider);
      currentProvider = null;
      providerStatus = null;
      providerSynced = null;
      currentOffline = null;
      offlineWaitStarted = false;
      element.textContent = "Local";
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const s = session as any;
    const docId = getDocId(session);

    const offline = s?.offline as unknown;
    let offlineLoaded: boolean | null = null;
    if (offline && typeof offline === "object") {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const loaded = (offline as any).isLoaded;
      if (typeof loaded === "boolean") offlineLoaded = loaded;
    }

    const offlineLoading = offlineLoaded === false;
    if (offline !== currentOffline) {
      currentOffline = offline;
      offlineWaitStarted = false;
    }

    if (offlineLoading && offline && typeof offline === "object" && !offlineWaitStarted) {
      offlineWaitStarted = true;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const whenLoaded = (offline as any).whenLoaded as (() => Promise<void>) | undefined;
      if (typeof whenLoaded === "function") {
        void Promise.resolve()
          .then(() => whenLoaded.call(offline))
          .catch(() => {
            // Offline persistence load failures should not crash the UI.
          })
          .finally(() => {
            if (!abortController.signal.aborted) render();
          });
      }
    }

    const provider = (s?.provider as unknown) ?? null;
    if (provider !== currentProvider) {
      detachProviderListeners(currentProvider);
      currentProvider = provider;
      providerStatus = null;
      providerSynced = null;
      attachProviderListeners(currentProvider);
    }

    if (offlineLoading) {
      element.textContent = `${docId} • Loading…`;
      return;
    }

    const connected = (() => {
      if (!currentProvider) return null;
      if (providerStatus === "connected") return true;
      if (providerStatus === "disconnected") return false;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyProvider = currentProvider as any;
      if (typeof anyProvider.wsconnected === "boolean") return anyProvider.wsconnected;
      if (typeof anyProvider.connected === "boolean") return anyProvider.connected;
      return null;
    })();

    const synced = (() => {
      if (!currentProvider) return null;
      if (typeof providerSynced === "boolean") return providerSynced;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyProvider = currentProvider as any;
      if (typeof anyProvider.synced === "boolean") return anyProvider.synced;
      return null;
    })();

    const connectionLabel = currentProvider
      ? connected === true
        ? "Connected"
        : connected === false
          ? "Disconnected"
          : "Connecting…"
      : "Offline";

    const syncLabel = currentProvider ? (synced === true ? "Synced" : "Syncing…") : "Local";

    element.textContent = `${docId} • ${connectionLabel} • ${syncLabel}`;
  };

  abortController.signal.addEventListener("abort", () => {
    window.removeEventListener("unload", cleanup);
    detachProviderListeners(currentProvider);
  });

  render();
}

const gridRoot = document.getElementById("grid");
if (!gridRoot) {
  throw new Error("Missing #grid container");
}

const titlebarRoot = document.getElementById("titlebar-root");
if (!titlebarRoot) {
  throw new Error("Missing #titlebar-root container");
}
const titlebarRootEl = titlebarRoot;

const ribbonRoot = document.getElementById("ribbon");
if (!ribbonRoot) {
  throw new Error("Missing #ribbon container");
}

const formulaBarRoot = document.getElementById("formula-bar");
if (!formulaBarRoot) {
  throw new Error("Missing #formula-bar container");
}

const statusMode = document.querySelector<HTMLElement>('[data-testid="status-mode"]');
const activeCell = document.querySelector<HTMLElement>('[data-testid="active-cell"]');
const selectionRange = document.querySelector<HTMLElement>('[data-testid="selection-range"]');
const activeValue = document.querySelector<HTMLElement>('[data-testid="active-value"]');
const collabStatus = document.querySelector<HTMLElement>('[data-testid="collab-status"]');
const selectionSum = document.querySelector<HTMLElement>('[data-testid="selection-sum"]');
const selectionAverage = document.querySelector<HTMLElement>('[data-testid="selection-avg"]');
const selectionCount = document.querySelector<HTMLElement>('[data-testid="selection-count"]');
const sheetSwitcher = document.querySelector<HTMLSelectElement>('[data-testid="sheet-switcher"]');
const zoomControl = document.querySelector<HTMLSelectElement>('[data-testid="zoom-control"]');
const statusZoom = document.querySelector<HTMLElement>('[data-testid="status-zoom"]');
const sheetPosition = document.querySelector<HTMLElement>('[data-testid="sheet-position"]');
if (
  !activeCell ||
  !selectionRange ||
  !activeValue ||
  !selectionSum ||
  !selectionAverage ||
  !selectionCount ||
  !statusMode ||
  !sheetSwitcher ||
  !zoomControl ||
  !statusZoom ||
  !sheetPosition
) {
  throw new Error("Missing status bar elements");
}
const sheetSwitcherEl = sheetSwitcher;
const zoomControlEl = zoomControl;
const statusZoomEl = statusZoom;
const sheetPositionEl = sheetPosition;

const docIdParam = new URL(window.location.href).searchParams.get("docId");
const docId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam : null;
const workbookId = docId ?? "local-workbook";
const app = new SpreadsheetApp(
  gridRoot,
  { activeCell, selectionRange, activeValue, selectionSum, selectionAverage, selectionCount },
  { formulaBar: formulaBarRoot, workbookId, sheetNameResolver },
);

// Expose a small API for Playwright assertions early so e2e can still attach even if
// optional desktop integrations (e.g. Tauri host wiring) fail during startup.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(app as any).getWorkbookSheetStore = () => workbookSheetStore;

function sharedGridZoomStorageKey(): string {
  // Scope zoom persistence by workbook/session id. For file-backed workbooks this can be
  // swapped to a path-based id if/when the desktop shell exposes it.
  return `formula:shared-grid:zoom:${workbookId}`;
}

function loadPersistedSharedGridZoom(): number | null {
  try {
    const storage = globalThis.localStorage;
    if (!storage) return null;
    const raw = storage.getItem(sharedGridZoomStorageKey());
    if (!raw) return null;
    const value = Number(raw);
    if (!Number.isFinite(value) || value <= 0) return null;
    return value;
  } catch {
    return null;
  }
}

function persistSharedGridZoom(value: number): void {
  try {
    const storage = globalThis.localStorage;
    if (!storage) return;
    storage.setItem(sharedGridZoomStorageKey(), String(value));
  } catch {
    // Ignore storage errors (disabled storage, quota, etc).
  }
}

function persistCurrentSharedGridZoom(): void {
  if (!app.supportsZoom()) return;
  persistSharedGridZoom(app.getZoom());
}

function applyPersistedSharedGridZoom(): void {
  if (!app.supportsZoom()) return;
  const persisted = loadPersistedSharedGridZoom();
  if (persisted == null) return;
  app.setZoom(persisted);
}

// Apply persisted zoom as early as possible so shared-grid renders at the user's
// preferred zoom before they interact with the sheet.
applyPersistedSharedGridZoom();

window.addEventListener("formula:zoom-changed", () => {
  syncZoomControl();
  persistCurrentSharedGridZoom();
});

function getGridLimitsForFormatting(): GridLimits {
  const anyApp = app as any;
  const raw = anyApp.limits ?? { maxRows: 10_000, maxCols: 200 };
  const maxRows = Number.isInteger(raw?.maxRows) && raw.maxRows > 0 ? raw.maxRows : 10_000;
  const maxCols = Number.isInteger(raw?.maxCols) && raw.maxCols > 0 ? raw.maxCols : 200;
  return { maxRows, maxCols };
}

function normalizeSelectionRange(range: Range): { startRow: number; endRow: number; startCol: number; endCol: number } {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function selectionRangesForFormatting(): CellRange[] {
  const limits = getGridLimitsForFormatting();
  const ranges = app.getSelectionRanges();
  if (ranges.length === 0) {
    const cell = app.getActiveCell();
    return [{ start: { row: cell.row, col: cell.col }, end: { row: cell.row, col: cell.col } }];
  }

  // When the UI selection is a full row/column/sheet *within the current grid limits*,
  // expand it to the canonical Excel bounds so DocumentController can use fast layered
  // formatting paths (sheet/row/col style ids) without enumerating every cell.
  return ranges.map((range) => {
    const r = normalizeSelectionRange(range);
    const isFullColBand = r.startRow === 0 && r.endRow === limits.maxRows - 1;
    const isFullRowBand = r.startCol === 0 && r.endCol === limits.maxCols - 1;

    return {
      start: { row: r.startRow, col: r.startCol },
      end: {
        row: isFullColBand ? DEFAULT_GRID_LIMITS.maxRows - 1 : r.endRow,
        col: isFullRowBand ? DEFAULT_GRID_LIMITS.maxCols - 1 : r.endCol,
      },
    };
  });
}

function rgbHexToArgb(rgb: string): string | null {
  if (!/^#[0-9A-Fa-f]{6}$/.test(rgb)) return null;
  // DocumentController formatting expects #AARRGGBB.
  return ["#", "FF", rgb.slice(1)].join("");
}

function applyFormattingToSelection(
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void,
  options: { forceBatch?: boolean } = {},
): void {
  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  const selection = app.getSelectionRanges();
  const limits = getGridLimitsForFormatting();
  const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });

  if (!decision.allowed) {
    showToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
    return;
  }

  const ranges = selectionRangesForFormatting();
  const shouldBatch = Boolean(options.forceBatch) || ranges.length > 1;

  if (shouldBatch) doc.beginBatch({ label });
  let committed = false;
  try {
    fn(doc, sheetId, ranges);
    committed = true;
  } finally {
    if (!shouldBatch) {
      // no-op
    } else if (committed) {
      doc.endBatch();
    } else {
      doc.cancelBatch();
    }
  }
  app.focus();
}

function createHiddenColorInput(): HTMLInputElement {
  const input = document.createElement("input");
  input.type = "color";
  input.tabIndex = -1;
  input.className = "hidden-color-input";
  document.body.appendChild(input);
  return input;
}

const fontColorPicker = createHiddenColorInput();
const fillColorPicker = createHiddenColorInput();

function openColorPicker(
  input: HTMLInputElement,
  label: string,
  apply: (sheetId: string, ranges: CellRange[], argb: string) => void,
): void {
  input.addEventListener(
    "change",
    () => {
      const argb = rgbHexToArgb(input.value);
      if (!argb) return;
      applyFormattingToSelection(label, (_doc, sheetId, ranges) => apply(sheetId, ranges, argb));
    },
    { once: true },
  );
  input.click();
}

const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

function activeCellFontSizePt(): number {
  const sheetId = app.getCurrentSheetId();
  const cell = app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = app.getDocument() as any;
  const state = docAny.getCell?.(sheetId, cell);
  const style = docAny.styleTable?.get?.(state?.styleId ?? 0) ?? {};
  const size = style.font?.size;
  return typeof size === "number" && Number.isFinite(size) && size > 0 ? size : 11;
}

function activeCellNumberFormat(): string | null {
  const sheetId = app.getCurrentSheetId();
  const cell = app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = app.getDocument() as any;
  const format = docAny.getCellFormat?.(sheetId, cell)?.numberFormat;
  return typeof format === "string" && format.trim() ? format : null;
}

function stepFontSize(current: number, direction: "increase" | "decrease"): number {
  const value = Number(current);
  const resolved = Number.isFinite(value) && value > 0 ? value : 11;
  if (direction === "increase") {
    for (const step of FONT_SIZE_STEPS) {
      if (step > resolved + 1e-6) return step;
    }
    return resolved;
  }

  for (let i = FONT_SIZE_STEPS.length - 1; i >= 0; i -= 1) {
    const step = FONT_SIZE_STEPS[i]!;
    if (step < resolved - 1e-6) return step;
  }
  return resolved;
}

function parseDecimalPlaces(format: string): number {
  const dot = format.indexOf(".");
  if (dot === -1) return 0;
  let count = 0;
  for (let i = dot + 1; i < format.length; i++) {
    const ch = format[i];
    if (ch === "0" || ch === "#") count += 1;
    else break;
  }
  return count;
}

function stepDecimalPlacesInNumberFormat(format: string | null, direction: "increase" | "decrease"): string | null {
  const raw = (format ?? "").trim();
  const section = (raw.split(";")[0] ?? "").trim();
  const lower = section.toLowerCase();
  // Avoid trying to manipulate date/time format codes.
  if (lower.includes("m/d/yyyy") || lower.includes("yyyy-mm-dd")) return null;

  const prefix = section.includes("$") ? "$" : "";
  const suffix = section.includes("%") ? "%" : "";
  const useThousands = section.includes(",");
  const decimals = parseDecimalPlaces(section);

  const nextDecimals =
    direction === "increase" ? Math.min(10, decimals + 1) : Math.max(0, decimals - 1);
  if (nextDecimals === decimals) return null;

  const integer = useThousands ? "#,##0" : "0";
  const fraction = nextDecimals > 0 ? `.${"0".repeat(nextDecimals)}` : "";
  return `${prefix}${integer}${fraction}${suffix}`;
}
// Panels persist state keyed by a workbook/document identifier. For file-backed workbooks we use
// their on-disk path; for unsaved sessions we generate a random session id so distinct new
// workbooks don't collide.
let activePanelWorkbookId = workbookId;
if (collabStatus) installCollabStatusIndicator(app, collabStatus);
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();

app.focus();

function openFormatCells(): void {
  openFormatCellsDialog({
    isEditing: () => app.isEditing(),
    getDocument: () => app.getDocument(),
    getSheetId: () => app.getCurrentSheetId(),
    getActiveCell: () => app.getActiveCell(),
    getSelectionRanges: () => app.getSelectionRanges(),
    focusGrid: () => app.focus(),
  });
}

const onUndo = () => {
  app.undo();
  app.focus();
};

const onRedo = () => {
  app.redo();
  app.focus();
};

const titlebarWindowControls = (() => {
  const winApi = (globalThis as any).__TAURI__?.window;
  const hasWindowHandle =
    winApi &&
    (typeof winApi.getCurrentWebviewWindow === "function" ||
      typeof winApi.getCurrentWindow === "function" ||
      typeof winApi.getCurrent === "function" ||
      Boolean(winApi.appWindow));
  if (!hasWindowHandle) return undefined;
  return {
    onClose: () => {
      void hideTauriWindow();
    },
    onMinimize: () => {
      void minimizeTauriWindow();
    },
    onToggleMaximize: () => {
      void toggleTauriWindowMaximize();
    },
  };
})();

function basename(path: string): string {
  const normalized = path.replace(/\\/g, "/").replace(/\/+$/g, "");
  const parts = normalized.split("/");
  return parts[parts.length - 1] || path;
}

function computeTitlebarDocumentName(): string {
  const isTauri = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
  if (!isTauri) return "Untitled";
  if (!activeWorkbook?.path) return "Untitled";
  return basename(activeWorkbook.path);
}

const buildTitlebarProps = () => ({
  documentName: computeTitlebarDocumentName(),
  actions: [
    {
      label: "Save",
      ariaLabel: "Save document",
      onClick: () => {
        if (typeof (globalThis as any).__TAURI__?.core?.invoke !== "function") {
          showToast("Save is only available in the desktop app.");
          return;
        }
        void handleSave().catch((err) => {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        });
      },
    },
    {
      label: "Share",
      ariaLabel: "Share document",
      variant: "primary" as const,
      onClick: () => showToast("Share is not implemented yet."),
    },
  ],
  windowControls: titlebarWindowControls,
  undoRedo: {
    ...app.getUndoRedoState(),
    onUndo,
    onRedo,
  },
});

const titlebar = mountTitlebar(titlebarRootEl, buildTitlebarProps());

const syncTitlebar = () => {
  titlebar.update(buildTitlebarProps());
};

function renderStatusMode(): void {
  statusMode.textContent = app.isEditing() ? "Edit" : "Ready";
}

renderStatusMode();

const unsubscribeTitlebarHistory = app.getDocument().on("history", () => syncTitlebar());
const unsubscribeTitlebarEditState = app.onEditStateChange(() => {
  renderStatusMode();
  syncTitlebar();
});
window.addEventListener("unload", () => {
  unsubscribeTitlebarHistory();
  unsubscribeTitlebarEditState();
  titlebar.dispose();
});

// --- Ribbon selection formatting state ----------------------------------------
// The ribbon UI maintains internal toggle state for user interactions, but we want
// formatting-related controls (Bold/Italic/Underline/Wrap/Align) to reflect the
// current selection's formatting, similar to Excel.
//
// Selection can change at pointer-move frequency while dragging. Keep updates
// responsive by throttling to one computation per animation frame.
let ribbonFormatStateUpdateScheduled = false;
let ribbonFormatStateUpdateRequested = false;

function scheduleRibbonSelectionFormatStateUpdate(): void {
  ribbonFormatStateUpdateRequested = true;
  if (ribbonFormatStateUpdateScheduled) return;
  ribbonFormatStateUpdateScheduled = true;

  requestAnimationFrame(() => {
    ribbonFormatStateUpdateScheduled = false;
    if (!ribbonFormatStateUpdateRequested) return;
    ribbonFormatStateUpdateRequested = false;

    const sheetId = app.getCurrentSheetId();
    const ranges = app.getSelectionRanges();
    const formatState = computeSelectionFormatState(app.getDocument(), sheetId, ranges);
    const isEditing = app.isEditing();

    const pressedById = {
      "home.font.bold": formatState.bold,
      "home.font.italic": formatState.italic,
      "home.font.underline": formatState.underline,
      "home.font.strikethrough": formatState.strikethrough,
      "home.alignment.wrapText": formatState.wrapText,
      "home.alignment.alignLeft": formatState.align === "left",
      "home.alignment.center": formatState.align === "center",
      "home.alignment.alignRight": formatState.align === "right",
      // Keep AutoSave off until a real autosave implementation exists.
      "file.save.autoSave": false,
      "view.show.showFormulas": app.getShowFormulas(),
      "formulas.formulaAuditing.showFormulas": app.getShowFormulas(),
      "view.show.performanceStats": Boolean((app.getGridPerfStats() as any)?.enabled),
      "view.window.split": ribbonLayoutController ? ribbonLayoutController.layout.splitView.direction !== "none" : false,
      "data.queriesConnections.queriesConnections":
        ribbonLayoutController != null &&
        getPanelPlacement(ribbonLayoutController.layout, PanelIds.DATA_QUERIES).kind !== "closed",
      "review.comments.showComments": app.isCommentsPanelVisible(),
    };

    const numberFormatLabel = (() => {
      const format = formatState.numberFormat;
      if (format === "mixed") return "Mixed";
      if (format == null) return "General";
      if (format === NUMBER_FORMATS.currency) return "Currency";
      if (format === NUMBER_FORMATS.percent) return "Percent";
      if (format === NUMBER_FORMATS.date) return "Date";
      return "Custom";
    })();

    const zoomDisabled = !app.supportsZoom();
    const disabledById = {
      ...(isEditing
        ? {
            // Formatting commands are disabled while editing (Excel-style behavior).
            "home.font.bold": true,
            "home.font.italic": true,
            "home.font.underline": true,
            "home.font.strikethrough": true,
            "home.font.fontName": true,
            "home.font.fontSize": true,
            "home.font.increaseFont": true,
            "home.font.decreaseFont": true,
            "home.font.fontColor": true,
            "home.font.fillColor": true,
            "home.font.borders": true,
            "home.alignment.wrapText": true,
            "home.alignment.topAlign": true,
            "home.alignment.middleAlign": true,
            "home.alignment.bottomAlign": true,
            "home.alignment.alignLeft": true,
            "home.alignment.center": true,
            "home.alignment.alignRight": true,
            "home.alignment.orientation": true,
            "home.number.numberFormat": true,
            "home.number.percent": true,
            "home.number.accounting": true,
            "home.number.date": true,
            "home.number.comma": true,
            "home.number.increaseDecimal": true,
            "home.number.decreaseDecimal": true,
          }
        : null),
      // View/zoom controls depend on the current runtime (e.g. shared-grid mode).
      "view.zoom.zoom": zoomDisabled,
      "view.zoom.zoom100": zoomDisabled,
      "view.zoom.zoomToSelection": zoomDisabled,
    };

    setRibbonUiState({
      pressedById,
      labelById: { "home.number.numberFormat": numberFormatLabel },
      disabledById,
    });
  });
}

app.subscribeSelection(() => {
  renderStatusMode();
  scheduleRibbonSelectionFormatStateUpdate();
});
app.getDocument().on("change", () => scheduleRibbonSelectionFormatStateUpdate());
app.onEditStateChange(() => scheduleRibbonSelectionFormatStateUpdate());
window.addEventListener("formula:view-changed", () => scheduleRibbonSelectionFormatStateUpdate());
window.addEventListener("formula:zoom-changed", () => syncZoomControl());
scheduleRibbonSelectionFormatStateUpdate();

function isTextInputTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el) return false;
  const tag = el.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
}

function canRunGridFormattingShortcuts(event: KeyboardEvent): boolean {
  if (event.defaultPrevented) return false;
  if (app.isEditing()) return false;
  if (isTextInputTarget(event.target)) return false;
  return true;
}

window.addEventListener("keydown", (e) => {
  if (!canRunGridFormattingShortcuts(e)) return;
  if (e.repeat) return;

  const primary = e.ctrlKey || e.metaKey;
  if (!primary) return;
  if (e.altKey) return;

  const key = e.key ?? "";
  const keyLower = key.toLowerCase();

  // --- Core formatting (Excel shortcuts) ---------------------------------------
  if (!e.shiftKey) {
    if (keyLower === "b") {
      e.preventDefault();
      applyFormattingToSelection("Bold", (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges));
      return;
    }
    // IMPORTANT: Cmd+I is reserved for the AI sidebar. Only bind italic to Ctrl+I.
    if (keyLower === "i" && e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      applyFormattingToSelection("Italic", (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges));
      return;
    }
    if (keyLower === "u") {
      e.preventDefault();
      applyFormattingToSelection("Underline", (doc, sheetId, ranges) => toggleUnderline(doc, sheetId, ranges));
      return;
    }
  }

  if (e.shiftKey) {
    const preset =
      key === "$" || e.code === "Digit4"
        ? "currency"
        : key === "%" || e.code === "Digit5"
          ? "percent"
          : key === "#" || e.code === "Digit3"
            ? "date"
            : null;

    if (preset) {
      e.preventDefault();
      const label =
        preset === "currency" ? "Currency format" : preset === "percent" ? "Percentage format" : "Date format";
      applyFormattingToSelection(label, (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, preset));
    }
  }
});

const ZOOM_PRESET_VALUES = new Set<number>(
  Array.from(zoomControlEl.querySelectorAll("option"))
    .map((opt) => Number(opt.value))
    .filter((value) => Number.isFinite(value) && value > 0),
);

function getCustomZoomOption(): HTMLOptionElement | null {
  return zoomControlEl.querySelector('option[data-zoom-custom="true"]');
}

function upsertCustomZoomOption(percent: number): void {
  let option = getCustomZoomOption();
  if (!option) {
    option = document.createElement("option");
    option.dataset.zoomCustom = "true";
    zoomControlEl.appendChild(option);
  }
  option.value = String(percent);
  option.textContent = `${percent}%`;
}

function syncZoomControl(): void {
  const percent = Math.round(app.getZoom() * 100);
  if (ZOOM_PRESET_VALUES.has(percent)) {
    getCustomZoomOption()?.remove();
  } else {
    // Ctrl/Cmd+wheel zoom and "zoom to selection" can produce arbitrary percent values.
    // Keep the dropdown stable by updating a single "custom" option instead of
    // appending a new option for every zoom tick.
    upsertCustomZoomOption(percent);
  }
  zoomControlEl.value = String(percent);
  statusZoomEl.textContent = `${percent}%`;
  zoomControlEl.disabled = !app.supportsZoom();
}

syncZoomControl();

zoomControlEl.addEventListener("change", () => {
  const nextPercent = Number(zoomControlEl.value);
  if (!Number.isFinite(nextPercent) || nextPercent <= 0) return;
  app.setZoom(nextPercent / 100);
  syncZoomControl();
  persistCurrentSharedGridZoom();
  app.focus();
});

let powerQueryService: DesktopPowerQueryService | null = null;
let powerQueryServiceWorkbookId: string | null = null;
let stopPowerQueryTrayListener: (() => void) | null = null;
const powerQueryTrayJobs = new Set<string>();
let powerQueryTrayHadError = false;

function updateTrayFromPowerQuery(): void {
  const status = powerQueryTrayJobs.size > 0 ? "syncing" : powerQueryTrayHadError ? "error" : "idle";
  void setTrayStatus(status);
}

function currentPowerQueryWorkbookId(): string {
  return activePanelWorkbookId;
}

function startPowerQueryService(): void {
  stopPowerQueryService();
  const serviceWorkbookId = currentPowerQueryWorkbookId();
  powerQueryServiceWorkbookId = serviceWorkbookId;
  powerQueryService = new DesktopPowerQueryService({
    workbookId: serviceWorkbookId,
    document: app.getDocument(),
    concurrency: 1,
    batchSize: 1024,
  });
  setDesktopPowerQueryService(serviceWorkbookId, powerQueryService);

  stopPowerQueryTrayListener?.();
  powerQueryTrayJobs.clear();
  powerQueryTrayHadError = false;
  updateTrayFromPowerQuery();

  stopPowerQueryTrayListener = powerQueryService.onEvent((evt) => {
    if (!evt || typeof evt !== "object") return;
    const type = (evt as any).type;
    if (type !== "apply:started" && type !== "apply:completed" && type !== "apply:cancelled" && type !== "apply:error") return;
    const jobId = (evt as any).jobId;
    if (typeof jobId !== "string" || jobId.trim() === "") return;

    if (type === "apply:started") {
      powerQueryTrayHadError = false;
      powerQueryTrayJobs.add(jobId);
    } else {
      powerQueryTrayJobs.delete(jobId);
      if (type === "apply:error") powerQueryTrayHadError = true;
    }

    updateTrayFromPowerQuery();
  });
}

function stopPowerQueryService(): void {
  const existingWorkbookId = powerQueryServiceWorkbookId;
  powerQueryServiceWorkbookId = null;
  if (existingWorkbookId) setDesktopPowerQueryService(existingWorkbookId, null);
  stopPowerQueryTrayListener?.();
  stopPowerQueryTrayListener = null;
  powerQueryService?.dispose();
  powerQueryService = null;

  powerQueryTrayJobs.clear();
  powerQueryTrayHadError = false;
  updateTrayFromPowerQuery();
}

startPowerQueryService();
window.addEventListener("unload", () => stopPowerQueryService());

type SelectionRect = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  activeRow: number;
  activeCol: number;
};

function currentSelectionRect(): SelectionRect {
  const sheetId = app.getCurrentSheetId();
  const active = app.getActiveCell();
  const ranges = app.getSelectionRanges();
  const normalize = (range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
    const startRow = Math.min(range.startRow, range.endRow);
    const endRow = Math.max(range.startRow, range.endRow);
    const startCol = Math.min(range.startCol, range.endCol);
    const endCol = Math.max(range.startCol, range.endCol);
    return { startRow, startCol, endRow, endCol };
  };

  const normalizedRanges = ranges.map(normalize);
  const containing =
    normalizedRanges.find(
      (r) =>
        active.row >= r.startRow && active.row <= r.endRow && active.col >= r.startCol && active.col <= r.endCol,
    ) ?? normalizedRanges[0];
  if (containing) {
    return {
      sheetId,
      startRow: containing.startRow,
      startCol: containing.startCol,
      endRow: containing.endRow,
      endCol: containing.endCol,
      activeRow: active.row,
      activeCol: active.col,
    };
  }
  return {
    sheetId,
    startRow: active.row,
    startCol: active.col,
    endRow: active.row,
    endCol: active.col,
    activeRow: active.row,
    activeCol: active.col,
  };
}

let openCommandPalette: (() => void) | null = null;
const commandRegistry = new CommandRegistry();

// Expose for Playwright e2e so tests can execute commands by id without going
// through UI affordances.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaCommandRegistry = commandRegistry;

// --- Sheet tabs (Excel-like multi-sheet UI) -----------------------------------

const sheetTabsRoot = document.getElementById("sheet-tabs");
if (!sheetTabsRoot) {
  throw new Error("Missing #sheet-tabs container");
}
const sheetTabsRootEl = sheetTabsRoot;
// The shell uses `.sheet-bar` (with an inner `.sheet-tabs` strip) for styling.
// Normalize older HTML scaffolds that used `.sheet-tabs` on the container itself.
sheetTabsRootEl.classList.add("sheet-bar");
sheetTabsRootEl.classList.remove("sheet-tabs");

let sheetTabsReactRoot: ReturnType<typeof createRoot> | null = null;
let stopSheetStoreListener: (() => void) | null = null;
let addSheetInFlight = false;

let sheetStoreDocSync: ReturnType<typeof startSheetStoreDocumentSync> | null = null;
let sheetStoreDocSyncStore: WorkbookSheetStore | null = null;

type SheetUiInfo = { id: string; name: string };

function listDocumentSheetIds(): string[] {
  const sheetIds = app.getDocument().getSheetIds();
  return sheetIds.length > 0 ? sheetIds : ["Sheet1"];
}
function shouldRestoreFocusAfterSheetNavigation(): boolean {
  // If the user is currently renaming a sheet tab inline, do not steal focus away
  // from the input (Excel-like).
  const active = document.activeElement;
  if (active instanceof HTMLInputElement && sheetTabsRootEl.contains(active)) {
    return false;
  }

  // If a context menu is open, let it manage focus and restore it when the menu closes.
  // (Some menus intentionally trap focus while open for accessibility.)
  const contextMenu = document.querySelector<HTMLElement>('[data-testid="context-menu"]');
  if (contextMenu && contextMenu.style.display !== "none") {
    return false;
  }

  return true;
}

function restoreFocusAfterSheetNavigation(): void {
  if (!shouldRestoreFocusAfterSheetNavigation()) return;
  app.focusAfterSheetNavigation();
}

function installSheetStoreDocSync(): void {
  // In collab mode, the sheet list is driven by the Yjs workbook schema (`session.sheets`).
  // Avoid reconciling the UI sheet store against DocumentController's lazily-created sheets.
  const session = app.getCollabSession?.() ?? null;
  if (session) {
    sheetStoreDocSync?.dispose();
    sheetStoreDocSync = null;
    sheetStoreDocSyncStore = null;
    return;
  }

  // When `workbookSheetStore` is replaced (workbook open, collab teardown), restart the sync.
  if (sheetStoreDocSync && sheetStoreDocSyncStore === workbookSheetStore) return;

  sheetStoreDocSync?.dispose();
  sheetStoreDocSync = startSheetStoreDocumentSync(
    app.getDocument(),
    workbookSheetStore,
    () => app.getCurrentSheetId(),
    (sheetId) => {
      app.activateSheet(sheetId);
      restoreFocusAfterSheetNavigation();
    },
  );
  sheetStoreDocSyncStore = workbookSheetStore;
}

function coerceCollabSheetField(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  const maybe = value as any;
  if (maybe?.constructor?.name === "YText" && typeof maybe.toString === "function") {
    try {
      return maybe.toString();
    } catch {
      return null;
    }
  }
  return null;
}

function listSheetsFromCollabSession(session: CollabSession): SheetUiInfo[] {
  const out: SheetUiInfo[] = [];
  const seen = new Set<string>();
  const entries = session?.sheets?.toArray?.() ?? [];
  for (const entry of entries) {
    const map: any = entry;
    const id = coerceCollabSheetField(map?.get?.("id") ?? map?.id);
    if (!id) continue;
    const trimmed = id.trim();
    if (!trimmed || seen.has(trimmed)) continue;
    seen.add(trimmed);
    const name = coerceCollabSheetField(map?.get?.("name") ?? map?.name) ?? trimmed;
    out.push({ id: trimmed, name });
  }
  return out.length > 0 ? out : [{ id: "Sheet1", name: "Sheet1" }];
}

function findCollabSheetIndexById(session: CollabSession, sheetId: string): number {
  const query = String(sheetId ?? "").trim();
  if (!query) return -1;
  for (let i = 0; i < session.sheets.length; i += 1) {
    const entry: any = session.sheets.get(i);
    const id = coerceCollabSheetField(entry?.get?.("id") ?? entry?.id);
    if (id && id.trim() === query) return i;
  }
  return -1;
}

function cloneCollabSheetMetaValue(value: unknown): unknown {
  if (value == null) return value;
  if (typeof value !== "object") return value;

  const maybe: any = value;
  if (maybe?.constructor?.name === "YText" && typeof maybe.toString === "function") {
    try {
      return maybe.toString();
    } catch {
      return null;
    }
  }

  // Avoid copying nested Yjs types directly; they can't be re-parented safely.
  const ctor = maybe?.constructor?.name ?? "";
  if (ctor.startsWith("Y") || ctor === "AbstractType") return undefined;

  const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
  if (typeof structuredCloneFn === "function") {
    try {
      return structuredCloneFn(value);
    } catch {
      // Fall through to JSON clone below.
    }
  }

  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function cloneCollabSheetMap(entry: unknown): Y.Map<unknown> {
  const out = new Y.Map<unknown>();
  const map: any = entry;

  if (map && typeof map.forEach === "function") {
    map.forEach((value: unknown, key: string) => {
      const k = String(key ?? "");
      if (!k) return;
      if (k === "id") return;
      if (k === "name") return;
      const cloned = cloneCollabSheetMetaValue(value);
      if (cloned === undefined) return;
      out.set(k, cloned);
    });
  }

  const id = coerceCollabSheetField(map?.get?.("id") ?? map?.id);
  if (id) out.set("id", id.trim());

  const hasName = typeof map?.has === "function" ? Boolean(map.has("name")) : map?.get?.("name") !== undefined;
  if (hasName) {
    const nameRaw = map?.get?.("name") ?? map?.name;
    const name = coerceCollabSheetField(nameRaw);
    if (name != null) out.set("name", name);
  }

  return out;
}

class CollabWorkbookSheetStore extends WorkbookSheetStore {
  constructor(
    private readonly session: CollabSession,
    initialSheets: ConstructorParameters<typeof WorkbookSheetStore>[0],
  ) {
    super(initialSheets);
  }

  override rename(id: string, newName: string): void {
    const before = this.getName(id);
    super.rename(id, newName);
    const after = this.getName(id);
    if (!after || after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;
      entry.set("name", after);
      // This update originated locally; update the cached key so our observer
      // doesn't unnecessarily rebuild the sheet store instance.
      lastCollabSheetsKey = listSheetsFromCollabSession(this.session)
        .map((s) => `${s.id}\u0000${s.name}`)
        .join("|");
    });
  }

  override move(id: string, toIndex: number): void {
    const before = this.listAll().map((s) => s.id).join("|");
    super.move(id, toIndex);
    const after = this.listAll().map((s) => s.id).join("|");
    if (after === before) return;

    this.session.transactLocal(() => {
      const fromIndex = findCollabSheetIndexById(this.session, id);
      if (fromIndex < 0) return;

      const entry: any = this.session.sheets.get(fromIndex);
      if (!entry) return;

      const clone = cloneCollabSheetMap(entry);
      this.session.sheets.delete(fromIndex, 1);
      this.session.sheets.insert(toIndex, [clone as any]);

      // This update originated locally; update the cached key so our observer
      // doesn't unnecessarily rebuild the sheet store instance.
      lastCollabSheetsKey = listSheetsFromCollabSession(this.session)
        .map((s) => `${s.id}\u0000${s.name}`)
        .join("|");
    });
  }
}

function reconcileSheetStoreWithDocument(ids: string[]): void {
  if (ids.length === 0) return;

  const docIdSet = new Set(ids);
  const existing = workbookSheetStore.listAll();
  const existingIdSet = new Set(existing.map((s) => s.id));

  // Add missing sheets (append in document order; UI order remains store-managed).
  let insertAfterId = workbookSheetStore.listAll().at(-1)?.id ?? "";
  for (const id of ids) {
    if (existingIdSet.has(id)) continue;
    workbookSheetStore.addAfter(insertAfterId, { id, name: id });
    existingIdSet.add(id);
    insertAfterId = id;
  }

  // Remove sheets that no longer exist in the document.
  for (const sheet of existing) {
    if (docIdSet.has(sheet.id)) continue;
    try {
      workbookSheetStore.remove(sheet.id);
    } catch {
      // Best-effort: avoid crashing the UI if the sheet store invariants don't
      // allow the removal (e.g. last-sheet guard).
    }
  }
}

let lastCollabSheetsKey = "";

function syncSheetStoreFromCollabSession(session: CollabSession): void {
  const sheets = listSheetsFromCollabSession(session);
  const key = sheets.map((s) => `${s.id}\u0000${s.name}`).join("|");
  if (key === lastCollabSheetsKey) return;
  lastCollabSheetsKey = key;

  try {
    workbookSheetStore = new CollabWorkbookSheetStore(
      session,
      sheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        visibility: "visible",
      })),
    );
  } catch (err) {
    // If collab sheet names are invalid/duplicated (shouldn't happen, but can if a remote
    // client writes bad metadata), fall back to using the stable sheet id as the display name
    // so the UI remains functional.
    console.error("[formula][desktop] Failed to apply collab sheet metadata:", err);
    workbookSheetStore = new CollabWorkbookSheetStore(
      session,
      sheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.id,
        visibility: "visible",
      })),
    );
  }

  // The sheet store instance is replaced whenever collab metadata changes; keep any
  // main.ts listeners (status bar, context keys, etc) subscribed to the latest store.
  installSheetStoreSubscription();
  syncWorkbookSheetNamesFromSheetStore();
}

function listSheetsForUi(): SheetUiInfo[] {
  const visible = workbookSheetStore.listVisible();
  if (visible.length > 0) return visible.map((s) => ({ id: s.id, name: s.name }));
  const ids = listDocumentSheetIds();
  return ids.map((id) => ({ id, name: id }));
}

async function handleAddSheet(): Promise<void> {
  if (addSheetInFlight) return;
  addSheetInFlight = true;
  try {
    const activeId = app.getCurrentSheetId();
    const desiredName = generateDefaultSheetName(workbookSheetStore.listAll());
    const doc = app.getDocument();

    const collabSession = app.getCollabSession?.() ?? null;
    if (collabSession) {
      // In collab mode, the Yjs `session.sheets` array is the authoritative sheet list.
      // Create the new sheet by updating that metadata so it propagates to other clients.
      const existing = listSheetsFromCollabSession(collabSession);
      const existingIds = new Set(existing.map((sheet) => sheet.id));

      const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
      const generateId = () => {
        const uuid = typeof randomUuid === "function" ? randomUuid.call((globalThis as any).crypto) : null;
        return uuid ? `sheet_${uuid}` : `sheet_${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
      };

      let id = generateId();
      for (let i = 0; i < 10 && existingIds.has(id); i += 1) {
        id = generateId();
      }
      while (existingIds.has(id)) {
        id = `${id}_${Math.random().toString(16).slice(2)}`;
      }

      collabSession.transactLocal(() => {
        const sheet = new Y.Map<unknown>();
        sheet.set("id", id);
        sheet.set("name", desiredName);

        // Insert after the active sheet when possible; otherwise append.
        let insertIndex = collabSession.sheets.length;
        for (let i = 0; i < collabSession.sheets.length; i += 1) {
          const entry: any = collabSession.sheets.get(i);
          const entryId = coerceCollabSheetField(entry?.get?.("id") ?? entry?.id)?.trim();
          if (entryId === activeId) {
            insertIndex = i + 1;
            break;
          }
        }

        collabSession.sheets.insert(insertIndex, [sheet as any]);
      });

      // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
      doc.getCell(id, { row: 0, col: 0 });
      doc.markDirty();
      app.activateSheet(id);
      restoreFocusAfterSheetNavigation();
      return;
    }

    const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
    if (typeof baseInvoke === "function") {
      // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
      const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

      // Allow any microtask-batched workbook edits to enqueue before we request a new sheet id.
      await new Promise<void>((resolve) => queueMicrotask(resolve));

      const info = (await invoke("add_sheet", { name: desiredName })) as SheetUiInfo;
      const id = String((info as any)?.id ?? "").trim();
      const name = String((info as any)?.name ?? "").trim();
      if (!id) throw new Error("Backend returned empty sheet id");

      // Backend may adjust the name for uniqueness; trust it.
      workbookSheetStore.addAfter(activeId, { id, name: name || desiredName });

      // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
      doc.getCell(id, { row: 0, col: 0 });
      doc.markDirty();
      app.activateSheet(id);
      restoreFocusAfterSheetNavigation();
      return;
    }

    // Web-only behavior: create a local DocumentController sheet lazily.
    // Until the DocumentController gains first-class sheet metadata, keep `id` and
    // `name` in lockstep for newly-created sheets.
    const newSheetId = desiredName;
    workbookSheetStore.addAfter(activeId, { id: newSheetId, name: desiredName });
    doc.getCell(newSheetId, { row: 0, col: 0 });
    doc.markDirty();
    app.activateSheet(newSheetId);
    restoreFocusAfterSheetNavigation();
  } catch (err) {
    showToast(`Failed to add sheet: ${String((err as any)?.message ?? err)}`, "error");
  } finally {
    addSheetInFlight = false;
  }
}

function renderSheetTabs(): void {
  if (!sheetTabsReactRoot) {
    sheetTabsReactRoot = createRoot(sheetTabsRootEl);
  }

  sheetTabsReactRoot.render(
    React.createElement(SheetTabStrip, {
      store: workbookSheetStore,
      activeSheetId: app.getCurrentSheetId(),
      onActivateSheet: (sheetId: string) => {
        app.activateSheet(sheetId);
        restoreFocusAfterSheetNavigation();
      },
      onAddSheet: handleAddSheet,
      onError: (message: string) => showToast(message, "error"),
    }),
  );
}

function renderSheetPosition(sheets: SheetUiInfo[], activeId: string): void {
  const total = sheets.length;
  const index = sheets.findIndex((sheet) => sheet.id === activeId);
  const position = index >= 0 ? index + 1 : 1;
  sheetPositionEl.textContent = `Sheet ${position} of ${total}`;
}

let syncingSheetUi = false;
let observedCollabSession: CollabSession | null = null;
let collabSheetsObserver: ((events: any, transaction: any) => void) | null = null;
let collabSheetsUnloadHookInstalled = false;

function ensureCollabSheetObserver(): void {
  const session = app.getCollabSession?.() ?? null;
  if (!session) return;
  if (observedCollabSession === session) return;

  if (observedCollabSession && collabSheetsObserver) {
    observedCollabSession.sheets.unobserveDeep(collabSheetsObserver as any);
  }

  observedCollabSession = session;
  collabSheetsObserver = () => {
    syncSheetStoreFromCollabSession(session);
    syncSheetUi();
  };
  session.sheets.observeDeep(collabSheetsObserver as any);

  // Initial sync (collab sheet list may include sheets with no cells, which the
  // DocumentController won't create and therefore wouldn't show up otherwise).
  syncSheetStoreFromCollabSession(session);

  if (!collabSheetsUnloadHookInstalled) {
    collabSheetsUnloadHookInstalled = true;
    window.addEventListener("unload", () => {
      if (observedCollabSession && collabSheetsObserver) {
        observedCollabSession.sheets.unobserveDeep(collabSheetsObserver as any);
      }
      observedCollabSession = null;
      collabSheetsObserver = null;
    });
  }
}

function syncSheetUi(): void {
  if (syncingSheetUi) return;
  syncingSheetUi = true;
  try {
    ensureCollabSheetObserver();

    const session = app.getCollabSession?.() ?? null;
    if (session) {
      // Keep the UI store aligned with the authoritative sheet list in Yjs.
      syncSheetStoreFromCollabSession(session);
    }

    // Keep `workbookSheetNames` in sync so sheet-name consumers (extension API,
    // context keys, etc) reflect collab metadata.
    syncWorkbookSheetNamesFromSheetStore();

    const sheets = listSheetsForUi();
    const activeId = app.getCurrentSheetId();
    if (!sheets.some((sheet) => sheet.id === activeId)) {
      const fallback = sheets[0]?.id ?? null;
      if (fallback) {
        // If the active sheet is removed (eg: via version restore or branch checkout),
        // automatically switch to the first remaining sheet.
        app.activateSheet(fallback);
      }
    }

    const nextActiveId = app.getCurrentSheetId();
    renderSheetTabs();
    renderSheetSwitcher(sheets, nextActiveId);
    renderSheetPosition(sheets, nextActiveId);
  } finally {
    syncingSheetUi = false;
  }
}

function installSheetStoreSubscription(): void {
  stopSheetStoreListener?.();
  stopSheetStoreListener = workbookSheetStore.subscribe(() => {
    // Sheet tab operations (rename/reorder/hide/tab color/etc) are workbook metadata changes
    // that may not touch any cells. Mark the DocumentController dirty so the unsaved-changes
    // prompt stays accurate.
    //
    // Guard against marking dirty during internal UI sync transactions.
    if (!syncingSheetUi) {
      app.getDocument().markDirty();
    }

    syncWorkbookSheetNamesFromSheetStore();
    const sheets = listSheetsForUi();
    const activeId = app.getCurrentSheetId();
    renderSheetSwitcher(sheets, activeId);
    renderSheetPosition(sheets, activeId);
  });
}

{
  installSheetStoreDocSync();
  installSheetStoreSubscription();
  syncSheetUi();
}

// `SpreadsheetApp.restoreDocumentState()` replaces the DocumentController model (including sheet ids).
// Keep the sheet metadata store in sync so tabs/switcher reflect the restored workbook.
const originalRestoreDocumentState = app.restoreDocumentState.bind(app);
app.restoreDocumentState = async (...args: Parameters<SpreadsheetApp["restoreDocumentState"]>): Promise<void> => {
  await originalRestoreDocumentState(...args);

  // `restoreDocumentState()` is used by version restore and workbook open. Ensure the doc->store sync
  // is installed for the current store instance, then force a synchronous reconciliation.
  installSheetStoreDocSync();
  sheetStoreDocSync?.syncNow();
  syncSheetUi();
};

const originalActivateSheet = app.activateSheet.bind(app);
app.activateSheet = (sheetId: string): void => {
  const prevSheet = app.getCurrentSheetId();
  originalActivateSheet(sheetId);
  syncSheetUi();
  const nextSheet = app.getCurrentSheetId();
  if (nextSheet !== prevSheet) emitSheetActivated(nextSheet);
};

const originalActivateCell = app.activateCell.bind(app);
app.activateCell = (...args: Parameters<SpreadsheetApp["activateCell"]>): void => {
  const prevSheet = app.getCurrentSheetId();
  originalActivateCell(...args);
  const nextSheet = app.getCurrentSheetId();
  if (nextSheet !== prevSheet) {
    syncSheetUi();
    emitSheetActivated(nextSheet);
  }
};

const originalSelectRange = app.selectRange.bind(app);
app.selectRange = (...args: Parameters<SpreadsheetApp["selectRange"]>): void => {
  const prevSheet = app.getCurrentSheetId();
  originalSelectRange(...args);
  const nextSheet = app.getCurrentSheetId();
  if (nextSheet !== prevSheet) {
    syncSheetUi();
    emitSheetActivated(nextSheet);
  }
};

// Keep the canvas renderer in sync with programmatic document mutations (e.g. AI tools)
// and re-render when edits create new sheets (DocumentController creates sheets lazily).
app.getDocument().on("change", () => {
  app.refresh();

  // Keep the sheet metadata store aligned with the DocumentController's sheet ids.
  // DocumentController can create sheets lazily (e.g. `setCellValue("Sheet2", ...)`) and
  // `applyState` can remove sheets after emitting its change event, so the sync layer
  // is microtask-debounced (see `sheetStoreDocumentSync.ts`).
  installSheetStoreDocSync();

  // In collab mode, the sheet list is driven by the Yjs workbook schema (`session.sheets`).
  // Avoid reconciling the UI sheet store against DocumentController's lazily-created sheets.
  const session = app.getCollabSession?.() ?? null;
  if (session) {
    // If collab comes online after initial render, sync once to attach the observer
    // and switch the sheet UI over to the Yjs-backed sheet list.
    if (observedCollabSession !== session) syncSheetUi();
    return;
  }
});

// --- Dock layout + persistence (minimal shell wiring for e2e + demos) ----------

const dockLeft = document.getElementById("dock-left");
const dockRight = document.getElementById("dock-right");
const dockBottom = document.getElementById("dock-bottom");
const floatingRoot = document.getElementById("floating-root");
const workspaceRoot = document.getElementById("workspace");
const gridSplit = document.getElementById("grid-split");
const gridSecondary = document.getElementById("grid-secondary");
const gridSplitter = document.getElementById("grid-splitter");

if (
  dockLeft &&
  dockRight &&
  dockBottom &&
  floatingRoot &&
  workspaceRoot &&
  gridSplit &&
  gridSecondary &&
  gridSplitter
) {
  const dockLeftEl = dockLeft;
  const dockRightEl = dockRight;
  const dockBottomEl = dockBottom;
  const floatingRootEl = floatingRoot;
  const workspaceRootEl = workspaceRoot;
  const gridSplitEl = gridSplit;
  const gridSecondaryEl = gridSecondary;
  const gridSplitterEl = gridSplitter;

  // --- Split view secondary pane keyboard shortcuts ---------------------------------
  //
  // The primary grid wires clipboard + delete via SpreadsheetApp.onKeyDown, which only
  // runs when `#grid` is focused. When focus is in the secondary pane we need to map
  // Excel-style shortcuts back into SpreadsheetApp command APIs.
  const isEditableTarget = (target: EventTarget | null): boolean => {
    const el = target as HTMLElement | null;
    if (!el) return false;
    const tag = el.tagName;
    return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
  };

  gridSecondaryEl.addEventListener("keydown", (e) => {
    if (e.defaultPrevented) return;

    // Match SpreadsheetApp guards: never steal shortcuts from active text editing.
    if (isEditableTarget(e.target)) return;
    if (app.isEditing()) return;

    if (e.key === "F2") {
      e.preventDefault();
      app.openCellEditorAtActiveCell();
      return;
    }

    const primary = e.ctrlKey || e.metaKey;
    const key = e.key.toLowerCase();

    if (primary && !e.altKey && !e.shiftKey) {
      if (key === "c") {
        e.preventDefault();
        app.copy();
        return;
      }
      if (key === "x") {
        e.preventDefault();
        app.cut();
        return;
      }
      if (key === "v") {
        e.preventDefault();
        app.paste();
        return;
      }
    }

    if (e.key === "Delete") {
      e.preventDefault();
      app.clearSelection();
    }
  });

  const workspaceManager = new LayoutWorkspaceManager({ storage: localStorage, panelRegistry });
  const layoutController = new LayoutController({
    workbookId,
    workspaceManager,
    primarySheetId: "Sheet1",
    workspaceId: "default",
  });
  ribbonLayoutController = layoutController;

  // Expose layout state for Playwright assertions (e.g. split view persistence).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__layoutController = layoutController;

  let lastAppliedZoom: number | null = null;

  function applyPrimaryPaneZoomFromLayout(): void {
    const zoom = (layoutController.layout as any)?.splitView?.panes?.primary?.zoom;
    const next = typeof zoom === "number" && Number.isFinite(zoom) ? zoom : 1;
    // Avoid redundant work on panel-only layout changes.
    if (lastAppliedZoom == null || Math.abs(lastAppliedZoom - next) > 1e-6) {
      app.setZoom(next);
      lastAppliedZoom = app.getZoom();
    }
    syncZoomControl();
  }

  const panelMounts = new Map<string, { container: HTMLElement; dispose: () => void }>();

  const scriptingWorkbook = new DocumentControllerWorkbookAdapter(app.getDocument(), {
    getActiveSheetName: () => app.getCurrentSheetId(),
    getSelection: () => {
      const ranges = app.getSelectionRanges();
      const first = ranges[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
      return { sheetName: app.getCurrentSheetId(), address: formatRangeAddress(first) };
    },
    setSelection: (sheetName, address) => {
      const range = parseRangeAddress(address);
      if (range.startRow === range.endRow && range.startCol === range.endCol) {
        app.activateCell({ sheetId: sheetName, row: range.startRow, col: range.startCol });
        return;
      }
      app.selectRange({
        sheetId: sheetName,
        range: { startRow: range.startRow, startCol: range.startCol, endRow: range.endRow, endCol: range.endCol },
      });
    },
    onDidMutate: () => {
      // SpreadsheetApp doesn't currently subscribe to DocumentController changes; it re-renders
      // directly after user-initiated edits. Scripts mutate the document out-of-band, so we
      // force a repaint after each script-side mutation.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyApp = app as any;
      anyApp.renderGrid?.();
      anyApp.renderCharts?.();
      anyApp.renderSelection?.();
      anyApp.updateStatus?.();
    },
  });

  const macroRecorder = new MacroRecorder(scriptingWorkbook);

  // SpreadsheetApp selection changes live outside the DocumentController mutation stream. Emit
  // selectionChanged events from the UI so the macro recorder can capture selection steps.
  let lastSelectionKey = "";
  scriptingWorkbook.events.on("selectionChanged", (evt: any) => {
    lastSelectionKey = `${evt.sheetName}:${evt.address}`;
  });
  app.subscribeSelection((selection) => {
    const first = selection.ranges[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
    const address = formatRangeAddress(first);
    const sheetName = app.getCurrentSheetId();
    const key = `${sheetName}:${address}`;
    if (key === lastSelectionKey) return;
    scriptingWorkbook.events.emit("selectionChanged", { sheetName, address });
  });

  function zoneVisible(zone: { panels: string[]; collapsed: boolean }) {
    return zone.panels.length > 0 && !zone.collapsed;
  }

  function applyDockSizes() {
    const layout = layoutController.layout;

    const leftSize = zoneVisible(layout.docks.left) ? layout.docks.left.size : 0;
    const rightSize = zoneVisible(layout.docks.right) ? layout.docks.right.size : 0;
    const bottomSize = zoneVisible(layout.docks.bottom) ? layout.docks.bottom.size : 0;

    workspaceRootEl.style.setProperty("--dock-left-size", `${leftSize}px`);
    workspaceRootEl.style.setProperty("--dock-right-size", `${rightSize}px`);
    workspaceRootEl.style.setProperty("--dock-bottom-size", `${bottomSize}px`);

    dockLeftEl.dataset.hidden = zoneVisible(layout.docks.left) ? "false" : "true";
    dockRightEl.dataset.hidden = zoneVisible(layout.docks.right) ? "false" : "true";
    dockBottomEl.dataset.hidden = zoneVisible(layout.docks.bottom) ? "false" : "true";
  }

  let secondaryGridView: SecondaryGridView | null = null;
  let splitPanePersistTimer: number | null = null;
  let splitPanePersistDirty = false;

  const syncSecondaryGridInteractionMode = () => {
    if (!secondaryGridView) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const formulaBar = (app as any).formulaBar as any;
    const mode = formulaBar?.isFormulaEditing?.() ? "rangeSelection" : "default";
    secondaryGridView.grid.setInteractionMode(mode);
    if (mode === "default") {
      // Ensure we don't leave behind transient formula-range selection overlays when exiting
      // formula editing (e.g. after committing/canceling, even if the last drag happened in
      // the secondary pane).
      secondaryGridView.grid.clearRangeSelection();
    }
  };

  const syncSecondaryGridInteractionModeSoon = () => {
    const schedule =
      typeof queueMicrotask === "function"
        ? queueMicrotask
        : (cb: () => void) => window.setTimeout(cb, 0);
    schedule(() => syncSecondaryGridInteractionMode());
  };

  // Keep the secondary pane interaction mode in sync with the formula bar's state.
  // We use events rather than polling to avoid unnecessary work.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const formulaBar = (app as any).formulaBar as any;
  if (formulaBar?.textarea instanceof HTMLTextAreaElement) {
    formulaBar.textarea.addEventListener("input", () => syncSecondaryGridInteractionMode());
    formulaBar.textarea.addEventListener("focus", () => syncSecondaryGridInteractionMode());
    // `blur` fires before FormulaBarView updates its model state during commit/cancel; sync on a microtask.
    formulaBar.textarea.addEventListener("blur", () => syncSecondaryGridInteractionModeSoon());
  }
  if (formulaBar?.root instanceof HTMLElement) {
    // Clicking fx/commit/cancel can mutate the draft without triggering a textarea input event.
    // Sync on a microtask so we observe the final FormulaBarView state.
    formulaBar.root.addEventListener("click", () => syncSecondaryGridInteractionModeSoon());
  }

  // High-frequency split-pane interactions (scroll/zoom) update the in-memory layout
  // without persisting on every event. Flush to storage on a debounce so we avoid
  // spamming localStorage writes.
  const scheduleSplitPanePersist = (delayMs = 500) => {
    splitPanePersistDirty = true;
    if (splitPanePersistTimer != null) window.clearTimeout(splitPanePersistTimer);
    splitPanePersistTimer = window.setTimeout(() => {
      splitPanePersistTimer = null;
      if (!splitPanePersistDirty) return;
      splitPanePersistDirty = false;
      layoutController.persistNow();
    }, delayMs);
  };

  const flushSplitPanePersist = () => {
    if (splitPanePersistTimer != null) {
      window.clearTimeout(splitPanePersistTimer);
      splitPanePersistTimer = null;
    }
    if (!splitPanePersistDirty) return;
    splitPanePersistDirty = false;
    layoutController.persistNow();
  };

  const persistLayoutNow = () => {
    if (splitPanePersistTimer != null) {
      window.clearTimeout(splitPanePersistTimer);
      splitPanePersistTimer = null;
    }
    splitPanePersistDirty = false;
    layoutController.persistNow();
  };

  const persistPrimaryZoomFromApp = () => {
    const pane = layoutController.layout.splitView.panes.primary;
    const zoom = app.getZoom();
    if (pane.zoom === zoom) return;
    layoutController.setSplitPaneZoom("primary", zoom, { persist: false });
    scheduleSplitPanePersist();
  };

  window.addEventListener("formula:zoom-changed", persistPrimaryZoomFromApp);

  const invalidateSecondaryProvider = () => {
    if (!secondaryGridView) return;
    // Sheet view state (frozen panes + axis overrides) lives in the DocumentController and is
    // independent of cell contents. Even when we reuse the primary grid's provider, we still
    // need to re-apply the current sheet's view state (e.g. when switching sheets).
    secondaryGridView.syncSheetViewFromDocument();
    const sharedProvider = (app as any).sharedProvider ?? null;
    // In shared-grid mode we reuse the primary provider, and SpreadsheetApp already
    // invalidates it on sheet changes / show-formulas toggles. Avoid extra churn.
    if (sharedProvider && secondaryGridView.provider === sharedProvider) return;
    secondaryGridView.provider.invalidateAll();
  };

  // Keep secondary grid in sync with non-DocumentController view changes in legacy mode
  // (sheet switching, show-formulas toggles). Shared-grid mode reuses the primary provider,
  // so the app already handles invalidations there.
  const activateSheetWithSplitSync = app.activateSheet.bind(app);
  app.activateSheet = (sheetId: string): void => {
    activateSheetWithSplitSync(sheetId);
    invalidateSecondaryProvider();
  };

  const activateCellWithSplitSync = app.activateCell.bind(app);
  app.activateCell = (...args: Parameters<SpreadsheetApp["activateCell"]>): void => {
    const target = args[0];
    const prevSheet = app.getCurrentSheetId();
    activateCellWithSplitSync(...args);
    if (target.sheetId && target.sheetId !== prevSheet) invalidateSecondaryProvider();
  };

  const selectRangeWithSplitSync = app.selectRange.bind(app);
  app.selectRange = (...args: Parameters<SpreadsheetApp["selectRange"]>): void => {
    const target = args[0];
    const prevSheet = app.getCurrentSheetId();
    selectRangeWithSplitSync(...args);
    if (target.sheetId && target.sheetId !== prevSheet) invalidateSecondaryProvider();
  };

  const setShowFormulasWithSplitSync = app.setShowFormulas.bind(app);
  app.setShowFormulas = (enabled: boolean): void => {
    setShowFormulasWithSplitSync(enabled);
    invalidateSecondaryProvider();
  };

  // --- Split-view selection synchronization (primary SpreadsheetApp ↔ secondary grid) ---

  const SPLIT_HEADER_ROWS = 1;
  const SPLIT_HEADER_COLS = 1;
  let splitSelectionSyncInProgress = false;
  let lastSplitSelection: SelectionState | null = null;

  function gridRangeFromDocRange(range: Range): GridCellRange {
    return {
      startRow: range.startRow + SPLIT_HEADER_ROWS,
      endRow: range.endRow + SPLIT_HEADER_ROWS + 1,
      startCol: range.startCol + SPLIT_HEADER_COLS,
      endCol: range.endCol + SPLIT_HEADER_COLS + 1,
    };
  }

  function docRangeFromGridRange(range: GridCellRange): Range {
    return {
      startRow: Math.max(0, range.startRow - SPLIT_HEADER_ROWS),
      endRow: Math.max(0, range.endRow - SPLIT_HEADER_ROWS - 1),
      startCol: Math.max(0, range.startCol - SPLIT_HEADER_COLS),
      endCol: Math.max(0, range.endCol - SPLIT_HEADER_COLS - 1),
    };
  }

  function syncPrimarySelectionFromSecondary(): void {
    if (!secondaryGridView) return;
    if (splitSelectionSyncInProgress) return;

    const gridSelection = secondaryGridView.grid.renderer.getSelection();
    const gridRanges = secondaryGridView.grid.renderer.getSelectionRanges();
    const activeIndex = secondaryGridView.grid.renderer.getActiveSelectionIndex();
    if (!gridSelection || gridRanges.length === 0) return;

    splitSelectionSyncInProgress = true;
    try {
      // Prefer syncing via the primary shared-grid instance (when available) so we preserve:
      // - multi-range selection
      // - the shared-grid active cell semantics (mouse-drag keeps the anchor cell active)
      // while still avoiding cross-pane scrolling.
      const primarySharedGrid = (app as any).sharedGrid as
        | { setSelectionRanges?: (ranges: GridCellRange[] | null, opts?: unknown) => void }
        | null;
      if (primarySharedGrid?.setSelectionRanges) {
        primarySharedGrid.setSelectionRanges(gridRanges, {
          activeIndex,
          activeCell: gridSelection,
          scrollIntoView: false,
        });
        return;
      }

      // Fallback (legacy grid mode): SpreadsheetApp does not currently support multi-range
      // or explicit active-cell programmatic selection. Mirror the active range only.
      const activeRange = gridRanges[Math.max(0, Math.min(activeIndex, gridRanges.length - 1))] ?? gridRanges[0];
      if (!activeRange) return;

      const docRange = docRangeFromGridRange(activeRange);

      // Prevent the primary pane from scrolling/focusing when selection is driven from the secondary pane.
      app.selectRange({ range: docRange }, { scrollIntoView: false, focus: false });
    } finally {
      splitSelectionSyncInProgress = false;
    }
  }

  app.subscribeSelection((selection) => {
    lastSplitSelection = selection;
    if (!secondaryGridView) return;
    if (splitSelectionSyncInProgress) return;

    splitSelectionSyncInProgress = true;
    try {
      const ranges = selection.ranges.map((r) => gridRangeFromDocRange(r));
      const activeCell = { row: selection.active.row + SPLIT_HEADER_ROWS, col: selection.active.col + SPLIT_HEADER_COLS };
      secondaryGridView.grid.setSelectionRanges(ranges, {
        activeIndex: selection.activeRangeIndex,
        activeCell,
        // Never cross-scroll panes: selection sync should not disturb the destination pane's scroll.
        scrollIntoView: false,
      });
    } finally {
      splitSelectionSyncInProgress = false;
    }
  });
  // --- Split view primary pane persistence (scroll/zoom) ----------------------
  // --- Split view primary pane persistence (scroll/zoom) ----------------------

  let stopPrimaryScrollSubscription: (() => void) | null = null;
  let stopPrimaryZoomSubscription: (() => void) | null = null;
  let primaryPaneViewportRestored = false;

  const stopPrimarySplitPanePersistence = () => {
    stopPrimaryScrollSubscription?.();
    stopPrimaryScrollSubscription = null;
    stopPrimaryZoomSubscription?.();
    stopPrimaryZoomSubscription = null;
    window.removeEventListener("beforeunload", persistLayoutNow);
  };

  const ensurePrimarySplitPanePersistence = () => {
    if (stopPrimaryScrollSubscription || stopPrimaryZoomSubscription) return;

    stopPrimaryScrollSubscription = app.subscribeScroll((scroll) => {
      if (layoutController.layout.splitView.direction === "none") return;

      const pane = layoutController.layout.splitView.panes.primary;
      if (pane.scrollX === scroll.x && pane.scrollY === scroll.y) return;

      layoutController.setSplitPaneScroll("primary", { scrollX: scroll.x, scrollY: scroll.y }, { persist: false });
      scheduleSplitPanePersist();
    });

    stopPrimaryZoomSubscription = app.subscribeZoom((zoom) => {
      if (layoutController.layout.splitView.direction === "none") return;

      const pane = layoutController.layout.splitView.panes.primary;
      if (pane.zoom === zoom) return;

      layoutController.setSplitPaneZoom("primary", zoom, { persist: false });
      scheduleSplitPanePersist();
    });

    window.addEventListener("beforeunload", persistLayoutNow);
  };

  const restorePrimarySplitPaneViewport = () => {
    const pane = layoutController.layout.splitView.panes.primary;
    if (app.supportsZoom()) {
      app.setZoom(pane.zoom ?? 1);
    }
    app.setScroll(pane.scrollX ?? 0, pane.scrollY ?? 0);
  };

  function renderSplitView() {
    const split = layoutController.layout.splitView;
    const ratio = typeof split.ratio === "number" ? split.ratio : 0.5;
    const clamped = Math.max(0.1, Math.min(0.9, ratio));
    const primaryPct = Math.round(clamped * 1000) / 10;
    const secondaryPct = Math.round((100 - primaryPct) * 10) / 10;
    gridSplitEl.dataset.splitDirection = split.direction;
    gridSplitEl.style.setProperty("--split-primary-size", `${primaryPct}%`);
    gridSplitEl.style.setProperty("--split-secondary-size", `${secondaryPct}%`);

    if (split.direction === "none") {
      secondaryGridView?.destroy();
      secondaryGridView = null;
      stopPrimarySplitPanePersistence();
      primaryPaneViewportRestored = false;
      flushSplitPanePersist();
      gridRoot.dataset.splitActive = "false";
      gridSecondaryEl.dataset.splitActive = "false";
      return;
    }

    if (!primaryPaneViewportRestored) {
      restorePrimarySplitPaneViewport();
      primaryPaneViewportRestored = true;
    }
    ensurePrimarySplitPanePersistence();

    if (!secondaryGridView) {
      const pane = split.panes.secondary;
      const initialScroll = { scrollX: pane.scrollX ?? 0, scrollY: pane.scrollY ?? 0 };
      const initialZoom = pane.zoom ?? 1;

      // Use the same DocumentController / computed value cache as the primary grid so
      // the secondary pane stays live with edits and formula recalculation.
      const anyApp = app as any;
      const limits = anyApp.limits ?? { maxRows: 10_000, maxCols: 200 };
      const rowCount = Number.isInteger(limits.maxRows) ? limits.maxRows + 1 : 10_001;
      const colCount = Number.isInteger(limits.maxCols) ? limits.maxCols + 1 : 201;

      secondaryGridView = new SecondaryGridView({
        container: gridSecondaryEl,
        provider: (app as any).sharedProvider ?? undefined,
        document: app.getDocument(),
        getSheetId: () => app.getCurrentSheetId(),
        rowCount,
        colCount,
        showFormulas: () => app.getShowFormulas(),
        getComputedValue: (cell) => app.getCellComputedValueForSheet(app.getCurrentSheetId(), cell),
        onSelectionChange: () => syncPrimarySelectionFromSecondary(),
        onSelectionRangeChange: () => syncPrimarySelectionFromSecondary(),
        callbacks: {
          onRangeSelectionStart: (range) => (app as any).onSharedRangeSelectionStart(range),
          onRangeSelectionChange: (range) => (app as any).onSharedRangeSelectionChange(range),
          onRangeSelectionEnd: () => (app as any).onSharedRangeSelectionEnd(),
        },
        initialScroll,
        initialZoom,
        persistScroll: (scroll) => {
          const pane = layoutController.layout.splitView.panes.secondary;
          if (pane.scrollX === scroll.scrollX && pane.scrollY === scroll.scrollY) return;
          layoutController.setSplitPaneScroll("secondary", scroll, { persist: false });
          scheduleSplitPanePersist();
        },
        persistZoom: (zoom) => {
          const pane = layoutController.layout.splitView.panes.secondary;
          if (pane.zoom === zoom) return;
          layoutController.setSplitPaneZoom("secondary", zoom, { persist: false });
          scheduleSplitPanePersist();
        },
      });

      // Ensure the secondary selection reflects the current primary selection without
      // affecting either pane's scroll positions.
      if (lastSplitSelection) {
        const selection = lastSplitSelection;
        splitSelectionSyncInProgress = true;
        try {
          const ranges = selection.ranges.map((r) => gridRangeFromDocRange(r));
          const activeCell = { row: selection.active.row + SPLIT_HEADER_ROWS, col: selection.active.col + SPLIT_HEADER_COLS };
          secondaryGridView.grid.setSelectionRanges(ranges, {
            activeIndex: selection.activeRangeIndex,
            activeCell,
            scrollIntoView: false,
          });
        } finally {
          splitSelectionSyncInProgress = false;
        }
      }

      // Ensure the secondary pane respects the current formula bar state (e.g. when enabling split view
      // while already editing a formula).
      syncSecondaryGridInteractionMode();
    }
    const active = split.activePane ?? "primary";
    gridRoot.dataset.splitActive = active === "primary" ? "true" : "false";
    gridSecondaryEl.dataset.splitActive = active === "secondary" ? "true" : "false";
  }

  function panelTitle(panelId: string) {
    return getPanelTitle(panelId);
  }

  function normalizeExtensionCellValue(value: unknown) {
    if (value == null) return null;
    if (typeof value === "string") return value;
    if (typeof value === "number") return value;
    if (typeof value === "boolean") return value;
    if (typeof value === "object" && value && "text" in value && typeof (value as any).text === "string") {
      return String((value as any).text);
    }
    return null;
  }

  function findSheetIdByName(name: string): string | null {
    const query = String(name ?? "").trim();
    if (!query) return null;

    // Prefer matching against display names from the sheet metadata store.
    const resolved = workbookSheetStore.resolveIdByName(query);
    if (resolved) return resolved;

    // Fall back to treating the name as a raw sheet id (eg: in-memory sessions where
    // sheet ids are "Sheet1", "Sheet2", ...).
    if (workbookSheetStore.getById(query)) return query;
    if (listDocumentSheetIds().includes(query)) return query;

    const activeSheetId = app.getCurrentSheetId();
    const activeSheetName = workbookSheetStore.getName(activeSheetId) ?? activeSheetId;
    if (query === activeSheetId || query === activeSheetName) return activeSheetId;

    return null;
  }

  function parseSheetQualifiedRange(ref: string): {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  } {
    const { sheetName, ref: a1Ref } = splitSheetQualifier(ref);
    const sheetId =
      sheetName == null ? app.getCurrentSheetId() : findSheetIdByName(sheetName) ?? null;
    if (!sheetId) {
      throw new Error(`Unknown sheet: ${String(sheetName)}`);
    }

    const { startRow, startCol, endRow, endCol } = parseRangeAddress(a1Ref);
    return { sheetId, startRow, startCol, endRow, endCol };
  }

  const contextKeys = new ContextKeyService();

  let lastSelection: SelectionState | null = null;

  const updateContextKeys = (selection: SelectionState | null = lastSelection) => {
    if (!selection) return;
    const sheetId = app.getCurrentSheetId();
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;
    const active = selection.active;
    const cell = app.getDocument().getCell(sheetId, { row: active.row, col: active.col }) as any;
    const value = normalizeExtensionCellValue(cell?.value ?? null);
    const formula = typeof cell?.formula === "string" ? cell.formula : null;
    const selectionKeys = deriveSelectionContextKeys(selection);

    contextKeys.batch({
      sheetName,
      ...selectionKeys,
      cellHasValue: (value != null && String(value).trim().length > 0) || (formula != null && formula.trim().length > 0),
      commentsPanelVisible: app.isCommentsPanelVisible(),
      cellHasComment: app.activeCellHasComment(),
    });
  };

  app.subscribeSelection((selection) => {
    lastSelection = selection;
    updateContextKeys(selection);
  });
  app.getDocument().on("change", () => updateContextKeys());
  window.addEventListener("formula:comments-panel-visibility-changed", () => updateContextKeys());
  window.addEventListener("formula:comments-changed", () => updateContextKeys());

  type ExtensionSelectionChangedEvent = {
    selection: {
      startRow: number;
      startCol: number;
      endRow: number;
      endCol: number;
      address: string;
      values: Array<Array<string | number | boolean | null>>;
    };
  };

  type ExtensionCellChangedEvent = {
    sheetId: string;
    row: number;
    col: number;
    value: string | number | boolean | null;
  };

  const selectionChangedEventListeners = new Set<(event: ExtensionSelectionChangedEvent) => void>();
  let selectionSubscription: (() => void) | null = null;
  let lastSelectionEventKey = "";
  let lastSelectionEventSheetId = app.getCurrentSheetId();
  let selectionSubscriptionInitialized = false;

  function ensureSelectionSubscription(): void {
    if (selectionSubscription) return;
    selectionSubscription = app.subscribeSelection(() => {
      const rect = currentSelectionRect();
      const range = { startRow: rect.startRow, startCol: rect.startCol, endRow: rect.endRow, endCol: rect.endCol };
      const address = formatRangeAddress(range);
      const key = `${rect.sheetId}:${address}`;
      if (!selectionSubscriptionInitialized) {
        selectionSubscriptionInitialized = true;
        lastSelectionEventKey = key;
        lastSelectionEventSheetId = rect.sheetId;
        return;
      }

      // Mirror the Node/InMemorySpreadsheet semantics: sheet activation emits `sheetActivated`
      // but does not implicitly emit `selectionChanged`.
      if (rect.sheetId !== lastSelectionEventSheetId) {
        lastSelectionEventSheetId = rect.sheetId;
        lastSelectionEventKey = key;
        return;
      }

      if (key === lastSelectionEventKey) return;
      lastSelectionEventKey = key;

      const values: Array<Array<string | number | boolean | null>> = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const row: Array<string | number | boolean | null> = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          const cell = app.getDocument().getCell(rect.sheetId, { row: r, col: c }) as any;
          row.push(normalizeExtensionCellValue(cell?.value ?? null));
        }
        values.push(row);
      }

      const payload: ExtensionSelectionChangedEvent = { selection: { ...range, address, values } };
      for (const listener of [...selectionChangedEventListeners]) {
        try {
          listener(payload);
        } catch {
          // ignore
        }
      }
    });
  }

  const cellChangedEventListeners = new Set<(event: ExtensionCellChangedEvent) => void>();
  let cellSubscription: (() => void) | null = null;
  let cellFlushScheduled = false;
  const pendingCellChanges = new Map<string, ExtensionCellChangedEvent>();

  function ensureCellSubscription(): void {
    if (cellSubscription) return;
    cellSubscription = app.getDocument().on("change", (payload: any) => {
      const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
      if (deltas.length === 0) return;

      for (const delta of deltas) {
        const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : app.getCurrentSheetId();
        const row = Number(delta?.row);
        const col = Number(delta?.col);
        if (!Number.isInteger(row) || row < 0) continue;
        if (!Number.isInteger(col) || col < 0) continue;

        const beforeValue = normalizeExtensionCellValue(delta?.before?.value ?? null);
        const afterValue = normalizeExtensionCellValue(delta?.after?.value ?? null);
        const beforeFormula = typeof delta?.before?.formula === "string" ? delta.before.formula : null;
        const afterFormula = typeof delta?.after?.formula === "string" ? delta.after.formula : null;
        if (beforeValue === afterValue && beforeFormula === afterFormula) continue;

        const event: ExtensionCellChangedEvent = { sheetId, row, col, value: afterValue };
        pendingCellChanges.set(`${sheetId}:${row},${col}`, event);
      }

      if (pendingCellChanges.size === 0) return;
      if (cellFlushScheduled) return;
      cellFlushScheduled = true;
      queueMicrotask(() => {
        cellFlushScheduled = false;
        if (pendingCellChanges.size === 0) return;
        const batch = Array.from(pendingCellChanges.values());
        pendingCellChanges.clear();
        for (const event of batch) {
          for (const listener of [...cellChangedEventListeners]) {
            try {
              listener(event);
            } catch {
              // ignore
            }
          }
        }
      });
    });
  }

  let extensionPanelBridge: ExtensionPanelBridge | null = null;

  type ExtensionClipboardSelectionContext = {
    sheetId: string;
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    timestampMs: number;
  };

  // Extensions can read selection values via `formula.cells.getSelection()` and then write arbitrary
  // text to the system clipboard via `formula.clipboard.writeText()`. SpreadsheetApp's copy/cut
  // handlers already enforce clipboard-copy DLP, but extensions would otherwise bypass it.
  //
  // Track the last selection returned to extensions and enforce clipboard-copy DLP on clipboard
  // writes that occur shortly afterwards.
  const extensionClipboardDlp = createDesktopDlpContext({ documentId: workbookId });
  let lastExtensionSelection: ExtensionClipboardSelectionContext | null = null;

  const clearLastExtensionSelection = () => {
    lastExtensionSelection = null;
  };

  const normalizeSelectionRange = (range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
    const startRow = Math.min(range.startRow, range.endRow);
    const endRow = Math.max(range.startRow, range.endRow);
    const startCol = Math.min(range.startCol, range.endCol);
    const endCol = Math.max(range.startCol, range.endCol);
    return { startRow, startCol, endRow, endCol };
  };

  const recordLastExtensionSelection = (
    sheetId: string,
    range: { startRow: number; startCol: number; endRow: number; endCol: number },
  ) => {
    lastExtensionSelection = { sheetId, ...range, timestampMs: Date.now() };
  };

  // The desktop UI is used both inside the Tauri shell and as a pure-web fallback (e2e, local dev).
  // Only expose real workbook lifecycle operations to extensions when the Tauri bridge is present;
  // otherwise let BrowserExtensionHost fall back to its in-memory stub workbook implementation.
  const hasTauriWorkbookBridge = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const extensionSpreadsheetApi: any = {
    getActiveSheet() {
      const sheetId = app.getCurrentSheetId();
      return { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? sheetId };
    },
    listSheets() {
      const sheets = workbookSheetStore.listAll();
      if (sheets.length > 0) return sheets.map((sheet) => ({ id: sheet.id, name: sheet.name }));
      return [{ id: "Sheet1", name: "Sheet1" }];
    },
    onSelectionChanged(handler: (event: ExtensionSelectionChangedEvent) => void) {
      if (typeof handler !== "function") return () => {};
      selectionChangedEventListeners.add(handler);
      ensureSelectionSubscription();
      return () => selectionChangedEventListeners.delete(handler);
    },
    onCellChanged(handler: (event: ExtensionCellChangedEvent) => void) {
      if (typeof handler !== "function") return () => {};
      cellChangedEventListeners.add(handler);
      ensureCellSubscription();
      return () => cellChangedEventListeners.delete(handler);
    },
    onSheetActivated(handler: (event: SheetActivatedEvent) => void) {
      if (typeof handler !== "function") return () => {};
      sheetActivatedListeners.add(handler);
      return () => sheetActivatedListeners.delete(handler);
    },
    async getSheet(name: string) {
      const sheetId = findSheetIdByName(name);
      if (!sheetId) return undefined;
      return { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? sheetId };
    },
    async activateSheet(name: string) {
      const sheetId = findSheetIdByName(name);
      if (!sheetId) {
        throw new Error(`Unknown sheet: ${String(name)}`);
      }
      app.activateSheet(sheetId);
      restoreFocusAfterSheetNavigation();
      return { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? sheetId };
    },
    async createSheet(name: string) {
      const sheetName = String(name ?? "").trim();
      if (!sheetName) {
        throw new Error("Sheet name must be a non-empty string");
      }
      if (findSheetIdByName(sheetName)) {
        throw new Error(`Sheet already exists: ${sheetName}`);
      }

      const activeId = app.getCurrentSheetId();
      const doc = app.getDocument();
      const validatedName = validateSheetName(sheetName, { sheets: workbookSheetStore.listAll(), ignoreId: null });

      const collabSession = app.getCollabSession?.() ?? null;
      if (collabSession) {
        const normalizedName = validateSheetName(sheetName, { sheets: workbookSheetStore.listAll() });

        const existingIds = new Set(listSheetsFromCollabSession(collabSession).map((sheet) => sheet.id));

        const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
        const generateId = () => {
          const uuid = typeof randomUuid === "function" ? randomUuid.call((globalThis as any).crypto) : null;
          return uuid ? `sheet_${uuid}` : `sheet_${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
        };

        let id = generateId();
        for (let i = 0; i < 10 && existingIds.has(id); i += 1) {
          id = generateId();
        }
        while (existingIds.has(id)) {
          id = `${id}_${Math.random().toString(16).slice(2)}`;
        }

        collabSession.transactLocal(() => {
          const sheet = new Y.Map<unknown>();
          sheet.set("id", id);
          sheet.set("name", normalizedName);

          const activeIdx = findCollabSheetIndexById(collabSession, activeId);
          const insertIndex = activeIdx >= 0 ? activeIdx + 1 : collabSession.sheets.length;
          collabSession.sheets.insert(insertIndex, [sheet as any]);
        });

        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
        doc.markDirty();
        app.activateSheet(id);
        restoreFocusAfterSheetNavigation();
        return { id, name: normalizedName };
      }

      const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
      if (typeof baseInvoke === "function") {
        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we request a new sheet id.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        const info = (await invoke("add_sheet", { name: validatedName })) as SheetUiInfo;
        const id = String((info as any)?.id ?? "").trim();
        const resolvedName = String((info as any)?.name ?? "").trim();
        if (!id) throw new Error("Backend returned empty sheet id");

        // Backend may adjust the name for uniqueness; trust it.
        workbookSheetStore.addAfter(activeId, { id, name: resolvedName || validatedName });
        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
        doc.markDirty();
        app.activateSheet(id);
        restoreFocusAfterSheetNavigation();
        const storedName = workbookSheetStore.getName(id);
        return { id, name: storedName ?? (resolvedName || validatedName || id) };
      }

      // Web-only behavior: create a local DocumentController sheet lazily.
      // Until the DocumentController gains first-class sheet metadata, keep `id` and
      // `name` in lockstep for newly-created sheets.
      const newSheetId = validatedName;
      workbookSheetStore.addAfter(activeId, { id: newSheetId, name: validatedName });
      doc.getCell(newSheetId, { row: 0, col: 0 });
      doc.markDirty();
      app.activateSheet(newSheetId);
      restoreFocusAfterSheetNavigation();
      const storedName = workbookSheetStore.getName(newSheetId);
      return { id: newSheetId, name: storedName ?? validatedName };
    },
    async renameSheet(_oldName: string, _newName: string) {
      const oldName = String(_oldName ?? "");
      const sheetId = findSheetIdByName(oldName);
      if (!sheetId) {
        throw new Error(`Unknown sheet: ${oldName}`);
      }

      const oldDisplayName = workbookSheetStore.getName(sheetId) ?? sheetId;
      const normalizedNewName = validateSheetName(String(_newName ?? ""), {
        sheets: workbookSheetStore.listAll(),
        ignoreId: sheetId,
      });

      // No-op rename; preserve the same return shape as other sheet APIs.
      if (oldDisplayName === normalizedNewName) {
        return { id: sheetId, name: oldDisplayName };
      }

      const collabSession = app.getCollabSession?.() ?? null;
      if (!collabSession) {
        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke === "function") {
          // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
          const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

          // Allow any microtask-batched workbook edits to enqueue before we rename.
          await new Promise<void>((resolve) => queueMicrotask(resolve));

          await invoke("rename_sheet", { sheet_id: sheetId, name: normalizedNewName });
        }
      }

      // Update UI metadata first so follow-up operations (eg: `getActiveSheet`) observe the new name.
      workbookSheetStore.rename(sheetId, normalizedNewName);
      syncSheetUi();
      updateContextKeys();

      // Rewrite existing formulas that reference the old sheet name (Excel-style behavior).
      const doc = app.getDocument();
      const rewrittenInputs: Array<{ sheetId: string; row: number; col: number; value: null; formula: string }> = [];
      for (const id of doc.getSheetIds()) {
        doc.forEachCellInSheet(id, ({ row, col, cell }: any) => {
          const formula = typeof cell?.formula === "string" ? cell.formula : null;
          if (!formula) return;
          const rewritten = rewriteSheetNamesInFormula(formula, oldDisplayName, normalizedNewName);
          if (rewritten === formula) return;
          rewrittenInputs.push({ sheetId: id, row, col, value: null, formula: rewritten });
        });
      }

      if (rewrittenInputs.length > 0) {
        doc.beginBatch({ label: "Rename sheet" });
        let committed = false;
        try {
          doc.setCellInputs(rewrittenInputs, { label: "Rename sheet", source: "extension" });
          committed = true;
        } finally {
          if (committed) doc.endBatch();
          else doc.cancelBatch();
        }
      }

      return { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? normalizedNewName };
    },
    async deleteSheet(name: string) {
      const sheetId = findSheetIdByName(name);
      if (!sheetId) {
        throw new Error(`Unknown sheet: ${String(name)}`);
      }

      const doc = app.getDocument();
      const wasActive = app.getCurrentSheetId() === sheetId;

      // Update sheet metadata first to enforce workbook invariants (e.g. last-sheet guard).
      workbookSheetStore.remove(sheetId);

      // DocumentController doesn't expose first-class sheet deletion yet, but the sheet map
      // is authoritative for `getSheetIds()` which drives UI reconciliation.
      try {
        doc.model?.sheets?.delete?.(sheetId);
      } catch {
        // ignore
      }

      if (wasActive) {
        const next =
          workbookSheetStore.listVisible().at(0)?.id ??
          workbookSheetStore.listAll().at(0)?.id ??
          app.getCurrentSheetId();
        if (next && next !== sheetId) {
          app.activateSheet(next);
        }
      }

      // Nudge downstream observers: deleting a sheet via `doc.model.sheets.delete` doesn't
      // emit a DocumentController change event, but our UI and caches are reconciled on change.
      try {
        const nudgeSheetId = app.getCurrentSheetId();
        const before = doc.getCell(nudgeSheetId, { row: 0, col: 0 });
        doc.applyExternalDeltas(
          [
            {
              sheetId: nudgeSheetId,
              row: 0,
              col: 0,
              before,
              after: before,
            },
          ],
          { recalc: false, source: "Extension deleteSheet" },
        );
      } catch {
        // ignore
      }

      syncSheetUi();
      updateContextKeys();
    },
    async getSelection() {
      const sheetId = app.getCurrentSheetId();
      const range = normalizeSelectionRange(
        app.getSelectionRanges()[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      );
      recordLastExtensionSelection(sheetId, range);
      const values: Array<Array<string | number | boolean | null>> = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const row: Array<string | number | boolean | null> = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          const cell = app.getDocument().getCell(sheetId, { row: r, col: c }) as any;
          row.push(normalizeExtensionCellValue(cell?.value ?? null));
        }
        values.push(row);
      }
      return { ...range, values };
    },
    async getCell(row: number, col: number) {
      const sheetId = app.getCurrentSheetId();
      const cell = app.getDocument().getCell(sheetId, { row, col }) as any;
      return normalizeExtensionCellValue(cell?.value ?? null);
    },
    async setCell(row: number, col: number, value: unknown) {
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, { row, col }, value);
    },
    async getRange(ref: string) {
      const { sheetId, startRow, startCol, endRow, endCol } = parseSheetQualifiedRange(ref);
      const values: Array<Array<string | number | boolean | null>> = [];
      for (let r = startRow; r <= endRow; r++) {
        const row: Array<string | number | boolean | null> = [];
        for (let c = startCol; c <= endCol; c++) {
          const cell = app.getDocument().getCell(sheetId, { row: r, col: c }) as any;
          row.push(normalizeExtensionCellValue(cell?.value ?? null));
        }
        values.push(row);
      }
      return { startRow, startCol, endRow, endCol, values };
    },
    async setRange(ref: string, values: unknown[][]) {
      const { sheetId, startRow, startCol, endRow, endCol } = parseSheetQualifiedRange(ref);
      const expectedRows = endRow - startRow + 1;
      const expectedCols = endCol - startCol + 1;

      if (!Array.isArray(values) || values.length !== expectedRows) {
        throw new Error(
          `Range values must be a ${expectedRows}x${expectedCols} array (got ${Array.isArray(values) ? values.length : 0} rows)`,
        );
      }

      const inputs: Array<{ sheetId: string; row: number; col: number; value: unknown; formula: null }> = [];

      for (let r = 0; r < expectedRows; r++) {
        const rowValues = values[r];
        if (!Array.isArray(rowValues) || rowValues.length !== expectedCols) {
          throw new Error(
            `Range values must be a ${expectedRows}x${expectedCols} array (row ${r} has ${Array.isArray(rowValues) ? rowValues.length : 0} cols)`,
          );
        }
        for (let c = 0; c < expectedCols; c++) {
          inputs.push({
            sheetId,
            row: startRow + r,
            col: startCol + c,
            value: rowValues[c],
            formula: null,
          });
        }
      }

      app.getDocument().setCellInputs(inputs, { label: "Extension setRange" });
    },
  };

  if (hasTauriWorkbookBridge) {
    extensionSpreadsheetApi.getActiveWorkbook = async () => {
      const sheetId = app.getCurrentSheetId();
      const activeSheet = { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? sheetId };
      const sheets = workbookSheetStore.listAll().map((sheet) => ({ id: sheet.id, name: sheet.name }));

      const path =
        activeWorkbook?.path ??
        activeWorkbook?.origin_path ??
        // If no backend workbook is active, treat this as an unsaved session.
        null;

      const name = (() => {
        const pick = typeof path === "string" && path.trim() !== "" ? path : null;
        if (!pick) return "Workbook";
        return pick.split(/[/\\]/).pop() ?? "Workbook";
      })();

      return { name, path, sheets, activeSheet };
    };

    extensionSpreadsheetApi.openWorkbook = async (path: string) => {
      await openWorkbookFromPath(String(path), { notifyExtensions: false, throwOnCancel: true });
    };

    extensionSpreadsheetApi.createWorkbook = async () => {
      await handleNewWorkbook({ notifyExtensions: false, throwOnCancel: true });
    };

    extensionSpreadsheetApi.saveWorkbook = async () => {
      await handleSave({ notifyExtensions: false, throwOnCancel: true });
    };

    extensionSpreadsheetApi.saveWorkbookAs = async (path: string) => {
      await handleSaveAsPath(String(path), { notifyExtensions: false });
    };

    extensionSpreadsheetApi.closeWorkbook = async () => {
      await handleNewWorkbook({ notifyExtensions: false, throwOnCancel: true });
    };
  }

  const clipboardProviderPromise = createClipboardProvider();

  const extensionHostManager = new DesktopExtensionHostManager({
    engineVersion: "1.0.0",
    spreadsheetApi: extensionSpreadsheetApi,
    clipboardApi: {
      readText: async () => {
        const provider = await clipboardProviderPromise;
        const { text } = await provider.read();
        return text ?? "";
      },
      writeText: async (text: string) => {
        const selection = lastExtensionSelection;
        if (selection && Date.now() - selection.timestampMs <= 2000) {
          enforceClipboardCopy({
            documentId: extensionClipboardDlp.documentId,
            sheetId: selection.sheetId,
            range: {
              start: { row: selection.startRow, col: selection.startCol },
              end: { row: selection.endRow, col: selection.endCol },
            },
            classificationStore: extensionClipboardDlp.classificationStore,
            policy: extensionClipboardDlp.policy,
          });
        }

        const provider = await clipboardProviderPromise;
        await provider.write({ text: String(text ?? "") });
        clearLastExtensionSelection();
      },
    },
    uiApi: {
      showMessage: async (message: string, type?: string) => {
        showToast(String(message ?? ""), (type as any) ?? "info");
      },
      showInputBox: async (options: any) => showInputBox(options),
      showQuickPick: async (items: any[], options: any) => showQuickPick(items, options),
      onPanelCreated: (panel: any) => extensionPanelBridge?.onPanelCreated(panel),
      onPanelHtmlUpdated: (panelId: string) => extensionPanelBridge?.onPanelHtmlUpdated(panelId),
      onPanelMessage: (panelId: string, message: unknown) => extensionPanelBridge?.onPanelMessage(panelId, message),
      onPanelDisposed: (panelId: string) => extensionPanelBridge?.onPanelDisposed(panelId),
    },
  });

  extensionHostManagerForE2e = extensionHostManager;

  registerBuiltinCommands({
    commandRegistry,
    app,
    layoutController,
    focusAfterSheetNavigation: restoreFocusAfterSheetNavigation,
    getVisibleSheetIds: () => workbookSheetStore.listVisible().map((sheet) => sheet.id),
    ensureExtensionsLoaded: () => ensureExtensionsLoadedRef?.() ?? Promise.resolve(),
    onExtensionsLoaded: () => {
      updateKeybindingsRef?.();
      syncContributedCommandsRef?.();
      syncContributedPanelsRef?.();
    },
  });

  const commandCategoryFormat = t("commandCategory.format");

  commandRegistry.registerBuiltinCommand(
    "format.toggleBold",
    t("command.format.toggleBold"),
    () =>
      applyFormattingToSelection(
        t("command.format.toggleBold"),
        (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleItalic",
    t("command.format.toggleItalic"),
    () =>
      applyFormattingToSelection(
        t("command.format.toggleItalic"),
        (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.toggleUnderline",
    t("command.format.toggleUnderline"),
    () =>
      applyFormattingToSelection(
        t("command.format.toggleUnderline"),
        (doc, sheetId, ranges) => toggleUnderline(doc, sheetId, ranges),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.numberFormat.currency",
    t("command.format.numberFormat.currency"),
    () =>
      applyFormattingToSelection(
        t("command.format.numberFormat.currency"),
        (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.numberFormat.percent",
    t("command.format.numberFormat.percent"),
    () =>
      applyFormattingToSelection(
        t("command.format.numberFormat.percent"),
        (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.numberFormat.date",
    t("command.format.numberFormat.date"),
    () =>
      applyFormattingToSelection(
        t("command.format.numberFormat.date"),
        (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"),
        { forceBatch: true },
      ),
    { category: commandCategoryFormat },
  );

  commandRegistry.registerBuiltinCommand(
    "format.openFormatCells",
    t("command.format.openFormatCells"),
    async () => {
      type Choice = "general" | "currency" | "percent" | "date";
      const choice = await showQuickPick<Choice>(
        [
          { label: "General", description: "Clear number format", value: "general" },
          { label: "Currency", description: NUMBER_FORMATS.currency, value: "currency" },
          { label: "Percent", description: NUMBER_FORMATS.percent, value: "percent" },
          { label: "Date", description: NUMBER_FORMATS.date, value: "date" },
        ],
        { placeHolder: "Format Cells" },
      );
      if (!choice) return;

      const patch =
        choice === "general"
          ? { numberFormat: null }
          : {
              numberFormat: NUMBER_FORMATS[choice],
            };

      applyFormattingToSelection(
        "Format Cells",
        (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, patch, { label: "Format Cells" });
          }
        },
        { forceBatch: true },
      );
    },
    { category: commandCategoryFormat },
  );

  extensionPanelBridge = new ExtensionPanelBridge({
    host: extensionHostManager.host as any,
    panelRegistry,
    layoutController: layoutController as any,
  });

  // Loading extensions spins up an additional Worker for each workbook tab. That can be
  // expensive in e2e + low-spec environments, so defer actually loading extensions until
  // the user opens the Extensions panel or invokes an extension command.
  let extensionsLoadPromise: Promise<void> | null = null;
  const ensureExtensionsLoaded = async () => {
    if (extensionHostManager.ready) return;
    if (!extensionsLoadPromise) {
      extensionsLoadPromise = extensionHostManager.loadBuiltInExtensions();
    }
    await extensionsLoadPromise;
  };

  const executeExtensionCommand = async (commandId: string, ...args: any[]) => {
    try {
      await ensureExtensionsLoaded();
      syncContributedCommands();
      await commandRegistry.executeCommand(commandId, ...args);
    } catch (err) {
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    }
  };

  const executeCommand = (commandId: string) => {
    const cmd = commandRegistry.getCommand(commandId);
    if (cmd?.source.kind === "builtin") {
      void commandRegistry.executeCommand(commandId).catch((err) => {
        showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
      });
      return;
    }
    executeExtensionCommand(commandId);
  };

  const openExtensionPanel = (panelId: string) => {
    void (async () => {
      await ensureExtensionsLoaded();
      syncContributedPanels();
      layoutController.openPanel(panelId);
      await extensionPanelBridge?.activateView(panelId);
    })().catch((err) => {
      showToast(`Failed to open panel: ${String((err as any)?.message ?? err)}`, "error");
    });
  };

  // Built-in keyboard shortcuts (Excel-compatible formatting).
  window.addEventListener(
    "keydown",
    (e) => {
      if (e.defaultPrevented) return;
      if (e.repeat) return;
      if (app.isEditing()) return;
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;

      const primary = e.ctrlKey || e.metaKey;
      if (!primary || e.altKey) return;

      const key = e.key ?? "";
      const keyLower = key.toLowerCase();

      // Font style toggles.
      if (!e.shiftKey && keyLower === "b") {
        e.preventDefault();
        e.stopPropagation();
        executeCommand("format.toggleBold");
        return;
      }
      // IMPORTANT: Cmd+I is reserved for the AI sidebar. Only bind italic to Ctrl+I.
      if (!e.shiftKey && keyLower === "i" && e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        e.stopPropagation();
        executeCommand("format.toggleItalic");
        return;
      }
      if (!e.shiftKey && keyLower === "u") {
        e.preventDefault();
        e.stopPropagation();
        executeCommand("format.toggleUnderline");
        return;
      }

      // Number formats.
      if (!e.shiftKey && (key === "1" || e.code === "Digit1")) {
        e.preventDefault();
        e.stopPropagation();
        executeCommand("format.openFormatCells");
        return;
      }
      if (e.shiftKey) {
        const preset =
          key === "$" || e.code === "Digit4"
            ? "currency"
            : key === "%" || e.code === "Digit5"
              ? "percent"
              : key === "#" || e.code === "Digit3"
                ? "date"
                : null;

        if (preset) {
          e.preventDefault();
          e.stopPropagation();
          executeCommand(`format.numberFormat.${preset}`);
          return;
        }
      }
    },
    true,
  );

  // Keybindings: central dispatch with built-in precedence over extensions.
  const commandKeybindingDisplayIndex = new Map<string, string[]>();
  let lastLoadedExtensionIds = new Set<string>();

  const platform = /Mac|iPhone|iPad|iPod/.test(navigator.platform) ? "mac" : "other";

  // Keybindings used for UI surfaces (command palette, context menu shortcut hints).
  // Prefer using `./commands/builtinKeybindings.ts` for new bindings.
  const builtinKeybindingHints = builtinKeybindingsCatalog;

  const syncContributedCommands = () => {
    if (!extensionHostManager.ready || extensionHostManager.error) return;
    try {
      commandRegistry.setExtensionCommands(
        extensionHostManager.getContributedCommands(),
        (commandId, ...args) => extensionHostManager.executeCommand(commandId, ...args),
      );
    } catch (err) {
      showToast(`Failed to register extension commands: ${String((err as any)?.message ?? err)}`, "error");
    }
  };

  const syncContributedPanels = () => {
    if (!extensionHostManager.ready || extensionHostManager.error) return;
    const contributed = extensionHostManager.getContributedPanels() as Array<{ extensionId: string; id: string; title: string; icon?: string | null }>;
    const contributedIds = new Set(contributed.map((p) => p.id));

    // Update the synchronous seed store from loaded extension manifests.
    if (contributedPanelsSeedStorage) {
      try {
        const extensions = extensionHostManager.host.listExtensions?.() ?? [];
        const currentLoadedExtensionIds = new Set<string>();
        for (const ext of extensions as any[]) {
          const id = typeof ext?.id === "string" ? ext.id : null;
          if (!id) continue;
          currentLoadedExtensionIds.add(id);
          const panels = Array.isArray(ext?.manifest?.contributes?.panels) ? ext.manifest.contributes.panels : [];
          setSeedPanelsForExtension(contributedPanelsSeedStorage, id, panels, {
            onError: (message) => {
              console.error(message);
              showToast(message, "error");
            },
          });
        }

        // Uninstall/unload: when an extension disappears from the runtime, remove its
        // contributed panel seeds so layouts stop preserving stale ids on future restarts.
        for (const prevId of lastLoadedExtensionIds) {
          if (currentLoadedExtensionIds.has(prevId)) continue;
          removeSeedPanelsForExtension(contributedPanelsSeedStorage, prevId);
        }
        lastLoadedExtensionIds = currentLoadedExtensionIds;
      } catch (err) {
        console.error("Failed to update contributed panel seed store:", err);
      }
    }

    // Remove contributed panels that are no longer installed (no longer present in the seed store).
    const seededPanels = contributedPanelsSeedStorage ? readContributedPanelsSeedStore(contributedPanelsSeedStorage) : {};
    const keepIds = new Set<string>([...contributedIds, ...Object.keys(seededPanels)]);

    for (const panel of contributed) {
      const seeded = seededPanels[panel.id];
      try {
        panelRegistry.registerPanel(
          panel.id,
          {
            title: panel.title,
            icon: (panel as any).icon ?? seeded?.icon ?? null,
            defaultDock: seeded?.defaultDock ?? "right",
            defaultFloatingRect: { x: 140, y: 140, width: 520, height: 640 },
            source: { kind: "extension", extensionId: panel.extensionId, contributed: true },
          },
          { owner: panel.extensionId, overwrite: true },
        );
      } catch (err) {
        console.error("Failed to register extension panel:", err);
        showToast(`Failed to register extension panel: ${panel.id}`, "error");
      }
    }

    for (const id of panelRegistry.listPanelIds()) {
      const def = panelRegistry.get(id) as any;
      const source = def?.source;
      if (source?.kind !== "extension" || source.contributed !== true) continue;
      if (keepIds.has(id)) continue;
      panelRegistry.unregisterPanel(id, { owner: source.extensionId });
    }
  };

  const keybindingService = new KeybindingService({
    commandRegistry,
    contextKeys,
    platform,
    onBeforeExecuteCommand: async (_commandId, source) => {
      if (source.kind !== "extension") return;
      await ensureExtensionsLoaded();
      syncContributedCommands();
    },
    onCommandError: (commandId, err) => {
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    },
  });
  keybindingService.setBuiltinKeybindings(builtinKeybindingHints);
  // Bubble-phase listener so SpreadsheetApp can `preventDefault()` first.
  keybindingService.installWindowListener(window, { capture: false });

  const updateKeybindings = () => {
    commandKeybindingDisplayIndex.clear();
    const contributed =
      extensionHostManager.ready && !extensionHostManager.error
        ? (extensionHostManager.getContributedKeybindings() as ContributedKeybinding[])
        : [];
    const nextKeybindingsIndex = buildCommandKeybindingDisplayIndex({
      platform,
      builtin: builtinKeybindingsCatalog,
      contributed,
    });
    for (const [commandId, bindings] of nextKeybindingsIndex.entries()) {
      commandKeybindingDisplayIndex.set(commandId, bindings);
    }
    keybindingService.setExtensionKeybindings(contributed);
  };

  const activateOpenExtensionPanels = () => {
    if (!extensionHostManager.ready || extensionHostManager.error) return;
    if (!extensionPanelBridge) return;
    const layout = layoutController.layout;
    const openPanelIds = [
      ...layout.docks.left.panels,
      ...layout.docks.right.panels,
      ...layout.docks.bottom.panels,
      ...Object.keys(layout.floating),
    ];
    for (const panelId of openPanelIds) {
      const def = panelRegistry.get(panelId) as any;
      if (def?.source?.kind !== "extension") continue;
      void extensionPanelBridge.activateView(panelId);
    }
  };

  extensionHostManager.subscribe(() => {
    // Update keybinding hints before commands emit registry change events so the command
    // palette renders shortcuts immediately after extensions finish loading.
    updateKeybindings();
    syncContributedCommands();
    syncContributedPanels();
    activateOpenExtensionPanels();
  });
  syncContributedCommands();
  syncContributedPanels();
  updateKeybindings();

  // Expose extension-loading hooks so the ribbon can lazily open the Extensions panel.
  ensureExtensionsLoadedRef = ensureExtensionsLoaded;
  syncContributedCommandsRef = syncContributedCommands;
  syncContributedPanelsRef = syncContributedPanels;
  updateKeybindingsRef = updateKeybindings;

  const contextMenu = new ContextMenu({
    onClose: () => {
      // Best-effort: return focus to the grid after closing.
      app.focus();
    },
  });

  const executeBuiltinCommand = (commandId: string, ...args: any[]) => {
    void commandRegistry.executeCommand(commandId, ...args).catch((err) => {
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    });
  };

  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.platform);
  const primaryShortcut = (key: string) => (isMac ? `⌘${key}` : `Ctrl+${key}`);
  const primaryShiftShortcut = (key: string) => (isMac ? `⌘⇧${key}` : `Ctrl+Shift+${key}`);

  const buildGridContextMenuItems = (): ContextMenuItem[] => {
    const undoRedo = app.getUndoRedoState();
    const undoLabelText = typeof undoRedo.undoLabel === "string" ? undoRedo.undoLabel.trim() : "";
    const redoLabelText = typeof undoRedo.redoLabel === "string" ? undoRedo.redoLabel.trim() : "";
    const undoLabel = undoLabelText
      ? tWithVars("menu.undoWithLabel", { label: undoLabelText })
      : t("command.edit.undo");
    const redoLabel = redoLabelText
      ? tWithVars("menu.redoWithLabel", { label: redoLabelText })
      : t("command.edit.redo");
    const allowEditCommands = !app.isEditing();

    const menuItems: ContextMenuItem[] = [
      {
        type: "item",
        label: undoLabel,
        enabled: undoRedo.canUndo,
        shortcut: getPrimaryCommandKeybindingDisplay("edit.undo", commandKeybindingDisplayIndex) ?? primaryShortcut("Z"),
        onSelect: () => executeBuiltinCommand("edit.undo"),
      },
      {
        type: "item",
        label: redoLabel,
        enabled: undoRedo.canRedo,
        shortcut:
          getPrimaryCommandKeybindingDisplay("edit.redo", commandKeybindingDisplayIndex) ?? (isMac ? "⇧⌘Z" : "Ctrl+Y"),
        onSelect: () => executeBuiltinCommand("edit.redo"),
      },
      { type: "separator" },
      {
        type: "item",
        label: t("clipboard.cut"),
        shortcut: getPrimaryCommandKeybindingDisplay("clipboard.cut", commandKeybindingDisplayIndex) ?? primaryShortcut("X"),
        onSelect: () => executeBuiltinCommand("clipboard.cut"),
      },
      {
        type: "item",
        label: t("clipboard.copy"),
        shortcut: getPrimaryCommandKeybindingDisplay("clipboard.copy", commandKeybindingDisplayIndex) ?? primaryShortcut("C"),
        onSelect: () => executeBuiltinCommand("clipboard.copy"),
      },
      {
        type: "item",
        label: t("clipboard.paste"),
        shortcut: getPrimaryCommandKeybindingDisplay("clipboard.paste", commandKeybindingDisplayIndex) ?? primaryShortcut("V"),
        onSelect: () => executeBuiltinCommand("clipboard.paste"),
      },
      { type: "separator" },
      {
        type: "item",
        label: t("clipboard.pasteSpecial.title"),
        shortcut:
          getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", commandKeybindingDisplayIndex) ??
          (isMac ? "⇧⌘V" : "Ctrl+Shift+V"),
        onSelect: () => executeBuiltinCommand("clipboard.pasteSpecial"),
      },
      { type: "separator" },
      {
        type: "item",
        label: t("menu.clearContents"),
        enabled: (() => {
          if (!allowEditCommands) return false;

          const isSingleCell = contextKeys.get("isSingleCell") === true;
          const activeHasValue = contextKeys.get("cellHasValue") === true;
          if (isSingleCell) return activeHasValue;
          if (activeHasValue) return true;

          // More accurate (but still sparse/efficient) check for multi-cell selections:
          // scan the sheet's sparse cell map and see if any non-empty cell falls within
          // the current selection ranges.
          //
          // This avoids scanning potentially huge rectangular selections.
          const sheetId = app.getCurrentSheetId();
          const doc: any = app.getDocument() as any;
          const sheetModel = doc?.model?.sheets?.get?.(sheetId) ?? null;
          const cells: Map<string, any> | null = sheetModel?.cells ?? null;
          if (!cells || cells.size === 0) return false;

          const ranges = app.getSelectionRanges();
          if (ranges.length === 0) return false;
          const normalized = ranges.map((range) => ({
            startRow: Math.min(range.startRow, range.endRow),
            endRow: Math.max(range.startRow, range.endRow),
            startCol: Math.min(range.startCol, range.endCol),
            endCol: Math.max(range.startCol, range.endCol),
          }));

          for (const [key, cell] of cells.entries()) {
            if (!cell || (cell.value == null && cell.formula == null)) continue;
            const [rowStr, colStr] = String(key).split(",", 2);
            const row = Number(rowStr);
            const col = Number(colStr);
            if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
            for (const range of normalized) {
              if (row < range.startRow || row > range.endRow) continue;
              if (col < range.startCol || col > range.endCol) continue;
              return true;
            }
          }

          return false;
        })(),
        shortcut:
          getPrimaryCommandKeybindingDisplay("edit.clearContents", commandKeybindingDisplayIndex) ?? (isMac ? "⌫" : "Del"),
        onSelect: () => executeBuiltinCommand("edit.clearContents"),
      },
      { type: "separator" },
      {
        type: "submenu",
        label: t("menu.format"),
        items: [
          {
            type: "item",
            label: t("command.format.toggleBold"),
            shortcut: getPrimaryCommandKeybindingDisplay("format.toggleBold", commandKeybindingDisplayIndex) ?? primaryShortcut("B"),
            onSelect: () => executeBuiltinCommand("format.toggleBold"),
          },
          {
            type: "item",
            label: t("command.format.toggleItalic"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.toggleItalic", commandKeybindingDisplayIndex) ?? primaryShortcut("I"),
            onSelect: () => executeBuiltinCommand("format.toggleItalic"),
          },
          {
            type: "item",
            label: t("command.format.toggleUnderline"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.toggleUnderline", commandKeybindingDisplayIndex) ??
              primaryShortcut("U"),
            onSelect: () => executeBuiltinCommand("format.toggleUnderline"),
          },
          { type: "separator" },
          {
            type: "item",
            label: t("command.format.numberFormat.currency"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.numberFormat.currency", commandKeybindingDisplayIndex) ??
              primaryShiftShortcut("$"),
            onSelect: () => executeBuiltinCommand("format.numberFormat.currency"),
          },
          {
            type: "item",
            label: t("command.format.numberFormat.percent"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.numberFormat.percent", commandKeybindingDisplayIndex) ??
              primaryShiftShortcut("%"),
            onSelect: () => executeBuiltinCommand("format.numberFormat.percent"),
          },
          {
            type: "item",
            label: t("command.format.numberFormat.date"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.numberFormat.date", commandKeybindingDisplayIndex) ??
              primaryShiftShortcut("#"),
            onSelect: () => executeBuiltinCommand("format.numberFormat.date"),
          },
          { type: "separator" },
          {
            type: "item",
            label: t("command.format.openFormatCells"),
            shortcut:
              getPrimaryCommandKeybindingDisplay("format.openFormatCells", commandKeybindingDisplayIndex) ?? primaryShortcut("1"),
            onSelect: () => executeBuiltinCommand("format.openFormatCells"),
          },
        ],
      },
      { type: "separator" },
      {
        type: "item",
        label: t("menu.addComment"),
        shortcut: getPrimaryCommandKeybindingDisplay("comments.addComment", commandKeybindingDisplayIndex) ?? undefined,
        onSelect: () => executeBuiltinCommand("comments.addComment"),
      },
      {
        type: "item",
        label: t("command.view.toggleShowFormulas"),
        enabled: allowEditCommands,
        shortcut:
          getPrimaryCommandKeybindingDisplay("view.toggleShowFormulas", commandKeybindingDisplayIndex) ?? primaryShortcut("`"),
        onSelect: () => executeBuiltinCommand("view.toggleShowFormulas"),
      },
      {
        type: "item",
        label: t("command.audit.togglePrecedents"),
        enabled: allowEditCommands,
        shortcut:
          getPrimaryCommandKeybindingDisplay("audit.togglePrecedents", commandKeybindingDisplayIndex) ?? primaryShortcut("["),
        onSelect: () => executeBuiltinCommand("audit.togglePrecedents"),
      },
      {
        type: "item",
        label: t("command.audit.toggleDependents"),
        enabled: allowEditCommands,
        shortcut:
          getPrimaryCommandKeybindingDisplay("audit.toggleDependents", commandKeybindingDisplayIndex) ?? primaryShortcut("]"),
        onSelect: () => executeBuiltinCommand("audit.toggleDependents"),
      },
    ];

    if (!extensionHostManager.ready) {
      // Extensions are loaded lazily. Show a non-blocking placeholder while the host spins up
      // so users understand why extension-contributed items are not immediately visible.
      menuItems.push({ type: "separator" });
      menuItems.push({
        type: "item",
        label: t("contextMenu.extensions.loading"),
        enabled: false,
        onSelect: () => {
          // Disabled placeholder item.
        },
      });
      return menuItems;
    }

    if (extensionHostManager.error) {
      menuItems.push({ type: "separator" });
      menuItems.push({
        type: "item",
        label: t("contextMenu.extensions.failedToLoad"),
        enabled: false,
        onSelect: () => {
          // Disabled error item.
        },
      });
      return menuItems;
    }

    // Ensure command labels are available.
    syncContributedCommands();

    const contributed = resolveMenuItems(
      extensionHostManager.getContributedMenu(CELL_CONTEXT_MENU_ID),
      contextKeys.asLookup(),
    );
    if (contributed.length > 0) {
      menuItems.push({ type: "separator" });
      const model = buildContextMenuModel(contributed, commandRegistry);
      for (const entry of model) {
        if (entry.kind === "separator") {
          menuItems.push({ type: "separator" });
          continue;
        }
        menuItems.push({
          type: "item",
          label: entry.label,
          enabled: entry.enabled,
          shortcut: getPrimaryCommandKeybindingDisplay(entry.commandId, commandKeybindingDisplayIndex) ?? undefined,
          onSelect: () => executeExtensionCommand(entry.commandId),
        });
      }
    }

    return menuItems;
  };

  // While the context menu is open, keep its enabled/disabled state in sync with
  // `ContextKeyService` so `when`-clauses can react to selection changes.
  contextKeys.onDidChange(() => {
    if (!contextMenu.isOpen()) return;
    contextMenu.update(buildGridContextMenuItems());
  });

  let contextMenuSession = 0;

  const openGridContextMenuAtPoint = (x: number, y: number, options: { focusFirst?: boolean } = {}) => {
    const session = (contextMenuSession += 1);
    const focusFirst = options.focusFirst === true;
    contextMenu.open({ x, y, items: buildGridContextMenuItems(), focusFirst });

    // Extensions are lazy-loaded to keep startup light. Right-clicking should still
    // surface extension-contributed context menu items, so load them on-demand and
    // update the menu if it is still open.
    if (!extensionHostManager.ready) {
      void ensureExtensionsLoaded()
        .then(() => {
          if (session !== contextMenuSession) return;
          if (!contextMenu.isOpen()) return;
          if (!extensionHostManager.ready) return;
          contextMenu.update(buildGridContextMenuItems());
          if (focusFirst) contextMenu.focusFirst();
        })
        .catch(() => {
          // Best-effort: keep the context menu functional even if extension loading fails.
          if (session !== contextMenuSession) return;
          if (!contextMenu.isOpen()) return;
          contextMenu.update(buildGridContextMenuItems());
          if (focusFirst) contextMenu.focusFirst();
        });
    }
  };

  gridRoot.addEventListener("contextmenu", (e) => {
    // Always prevent the native context menu; we render our own.
    e.preventDefault();

    const anchorX = e.clientX;
    const anchorY = e.clientY;

    const picked = app.pickCellAtClientPoint(anchorX, anchorY);
    if (picked) {
      // Excel-like behavior: if the user right-clicks outside the current selection,
      // move the active cell to the clicked coordinate. If they right-click within
      // the selection, keep it intact (important for when-clause context keys like
      // `hasSelection`).
      const ranges = app.getSelectionRanges();
      const inSelection = ranges.some((range) => {
        const startRow = Math.min(range.startRow, range.endRow);
        const endRow = Math.max(range.startRow, range.endRow);
        const startCol = Math.min(range.startCol, range.endCol);
        const endCol = Math.max(range.startCol, range.endCol);
        return picked.row >= startRow && picked.row <= endRow && picked.col >= startCol && picked.col <= endCol;
      });
      if (!inSelection) {
        app.activateCell({ row: picked.row, col: picked.col });
      }
    }

    openGridContextMenuAtPoint(anchorX, anchorY);
  });

  const openGridContextMenuAtActiveCell = () => {
    const rect = app.getActiveCellRect();
    if (rect) {
      openGridContextMenuAtPoint(rect.x, rect.y + rect.height, { focusFirst: true });
      return;
    }

    const gridRect = gridRoot.getBoundingClientRect();
    openGridContextMenuAtPoint(gridRect.left + gridRect.width / 2, gridRect.top + gridRect.height / 2, { focusFirst: true });
  };

  window.addEventListener(
    "keydown",
    (e) => {
      if (e.defaultPrevented) return;
      if (isEditableTarget(e.target as HTMLElement | null)) return;

      const shouldOpen = (e.shiftKey && e.key === "F10") || e.key === "ContextMenu" || e.code === "ContextMenu";
      if (!shouldOpen) return;

      e.preventDefault();
      openGridContextMenuAtActiveCell();
    },
    true,
  );

  let macrosBackend: unknown | null = null;
  const getMacrosBackend = () => {
    if (macrosBackend) return macrosBackend as any;
    try {
      macrosBackend = new TauriMacroBackend();
    } catch {
      macrosBackend = new WebMacroBackend({
        getDocumentController: () => app.getDocument(),
        getActiveSheetId: () => app.getCurrentSheetId(),
      });
    }
    return macrosBackend as any;
  };

  const panelBodyRenderer = createPanelBodyRenderer({
    getDocumentController: () => app.getDocument(),
    getActiveSheetId: () => app.getCurrentSheetId(),
    getSearchWorkbook: () => app.getSearchWorkbook(),
    getSelection: () => {
      const selection = currentSelectionRect();
      return {
        sheetId: selection.sheetId,
        range: {
          startRow: selection.startRow,
          startCol: selection.startCol,
          endRow: selection.endRow,
          endCol: selection.endCol,
        },
      };
    },
    workbookId,
    getWorkbookId: () => activePanelWorkbookId,
    getCollabSession: () => (app as any).getCollabSession?.() ?? null,
    invoke:
      typeof (globalThis as any).__TAURI__?.core?.invoke === "function"
        ? (cmd, args) => {
            const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
            const invokeFn = queuedInvoke ?? baseInvoke;
            if (!invokeFn) {
              return Promise.reject(new Error("Tauri invoke API not available"));
            }
            return invokeFn(cmd, args);
          }
        : undefined,
    drainBackendSync,
    getMacroUiContext: () => {
      const selection = currentSelectionRect();
      return {
        sheetId: selection.sheetId,
        activeRow: selection.activeRow ?? selection.startRow,
        activeCol: selection.activeCol ?? selection.startCol,
        selection: {
          startRow: selection.startRow,
          startCol: selection.startCol,
          endRow: selection.endRow,
          endCol: selection.endCol,
        },
      };
    },
    createChart: (spec) => app.addChart(spec),
    panelRegistry,
    extensionPanelBridge: extensionPanelBridge ?? undefined,
    extensionHostManager,
    onExecuteExtensionCommand: executeExtensionCommand,
    onOpenExtensionPanel: openExtensionPanel,
    renderMacrosPanel: (body) => {
      body.textContent = "Loading macros…";
      queueMicrotask(() => {
        try {
          body.replaceChildren();
          body.classList.add("macros-panel");

          const recorderPanel = document.createElement("div");
          recorderPanel.className = "macros-panel__recorder";

          const runnerPanel = document.createElement("div");
          runnerPanel.className = "macros-panel__runner";

          body.appendChild(recorderPanel);
          body.appendChild(runnerPanel);

          const scriptStorageId = activePanelWorkbookId;

          // Prefer a composite backend in desktop builds: run VBA via Tauri (when available),
          // and run modern scripts (TypeScript/Python) via the web backend + local storage.
          const backend = (() => {
            try {
              const baseBackend = new TauriMacroBackend({ invoke: queuedInvoke ?? undefined });
              const tauriBackend = wrapTauriMacroBackendWithUiContext(
                baseBackend,
                () => {
                  const selection = currentSelectionRect();
                  return {
                    sheetId: selection.sheetId,
                    activeRow: selection.activeRow ?? selection.startRow,
                    activeCol: selection.activeCol ?? selection.startCol,
                    selection: {
                      startRow: selection.startRow,
                      startCol: selection.startCol,
                      endRow: selection.endRow,
                      endCol: selection.endCol,
                    },
                  };
                },
                {
                  beforeRunMacro: async () => {
                    // Allow any microtask-batched workbook edits to enqueue before the macro runs
                    // so backend state reflects the latest grid changes.
                    await new Promise<void>((resolve) => queueMicrotask(resolve));
                    await drainBackendSync();
                  },
                }
              );

              const scriptBackend = new WebMacroBackend({
                getDocumentController: () => app.getDocument(),
                getActiveSheetId: () => app.getCurrentSheetId(),
              });

              /** @type {Map<string, "tauri" | "web">} */
              const macroOrigin = new Map<string, "tauri" | "web">();

              return {
                listMacros: async (id: string) => {
                  const [tauriMacros, webMacros] = await Promise.all([
                    baseBackend.listMacros(id).catch(() => []),
                    scriptBackend.listMacros(scriptStorageId).catch(() => []),
                  ]);

                  macroOrigin.clear();
                  for (const macro of tauriMacros) macroOrigin.set(macro.id, "tauri");
                  for (const macro of webMacros) {
                    if (!macroOrigin.has(macro.id)) macroOrigin.set(macro.id, "web");
                  }

                  const byId = new Map<string, any>();
                  for (const macro of webMacros) byId.set(macro.id, macro);
                  for (const macro of tauriMacros) byId.set(macro.id, macro);
                  return [...byId.values()].sort((a, b) => String(a.name).localeCompare(String(b.name)));
                },
                getMacroSecurityStatus: (id: string) => baseBackend.getMacroSecurityStatus(id),
                setMacroTrust: (id: string, decision: MacroTrustDecision) => baseBackend.setMacroTrust(id, decision),
                runMacro: async (request: MacroRunRequest) => {
                  const origin = macroOrigin.get(request.macroId);
                  if (origin === "web") {
                    return scriptBackend.runMacro({ ...request, workbookId: scriptStorageId });
                  }

                  return tauriBackend.runMacro(request);
                },
              };
            } catch {
              return getMacrosBackend();
            }
           })();
 
           const refreshRunner = () =>
             renderMacroRunner(runnerPanel, backend, workbookId, {
               onApplyUpdates: async (updates) => {
                  if (vbaEventMacros) {
                    await vbaEventMacros.applyMacroUpdates(updates, { label: "Run macro" });
                    return;
                  }

                  const doc = app.getDocument();
                  doc.beginBatch({ label: "Run macro" });
                  let committed = false;
                  try {
                    applyMacroCellUpdates(doc, updates);
                    committed = true;
                  } finally {
                    if (committed) doc.endBatch();
                    else doc.cancelBatch();
                  }
                  app.refresh();
                  await app.whenIdle();
                  app.refresh();
                },
              });

           const title = document.createElement("div");
           title.textContent = "Macro Recorder";
          title.className = "macros-panel__title";
          recorderPanel.appendChild(title);

          const status = document.createElement("div");
          status.className = "macros-panel__status";
          recorderPanel.appendChild(status);

          const buttons = document.createElement("div");
          buttons.className = "macros-panel__buttons";
          recorderPanel.appendChild(buttons);

          const startButton = document.createElement("button");
          startButton.type = "button";
          startButton.textContent = "Start Recording";
          buttons.appendChild(startButton);

          const stopButton = document.createElement("button");
          stopButton.type = "button";
          stopButton.textContent = "Stop Recording";
          buttons.appendChild(stopButton);

          const copyTsButton = document.createElement("button");
          copyTsButton.type = "button";
          copyTsButton.textContent = "Copy TypeScript";
          buttons.appendChild(copyTsButton);

          const copyPyButton = document.createElement("button");
          copyPyButton.type = "button";
          copyPyButton.textContent = "Copy Python";
          buttons.appendChild(copyPyButton);

          const openScriptEditorButton = document.createElement("button");
          openScriptEditorButton.type = "button";
          openScriptEditorButton.textContent = "Open in Script Editor";
          buttons.appendChild(openScriptEditorButton);

          const saveButton = document.createElement("button");
          saveButton.type = "button";
          saveButton.textContent = "Save as Macro…";
          buttons.appendChild(saveButton);

          const meta = document.createElement("div");
          meta.className = "macros-panel__meta";
          recorderPanel.appendChild(meta);

          const preview = document.createElement("pre");
          preview.className = "macros-panel__preview";
          recorderPanel.appendChild(preview);

          const copyText = async (text: string) => {
            try {
              const provider = await clipboardProviderPromise;
              await provider.write({ text });
            } catch {
              // Fall back to execCommand in case Clipboard API permissions are unavailable.
              // This is best-effort; ignore failures.
            }

            // Clipboard provider writes are best-effort, and in web contexts clipboard access
            // can still be permission-gated. Keep an execCommand fallback when not running
            // under Tauri to preserve "Copy" behavior in dev/preview browsers.
            const isTauri = typeof (globalThis as any).__TAURI__ !== "undefined";
            if (isTauri) return;

            let textarea: HTMLTextAreaElement | null = null;
            try {
              textarea = document.createElement("textarea");
              textarea.value = text;
              textarea.className = "macros-panel__clipboard-textarea";
              document.body.appendChild(textarea);
              textarea.select();
              document.execCommand("copy");
            } catch {
              // ignore
            } finally {
              textarea?.remove();
            }
          };

          const openScriptEditor = (code: string) => {
            const placement = getPanelPlacement(layoutController.layout, PanelIds.SCRIPT_EDITOR);
            if (placement.kind === "closed") layoutController.openPanel(PanelIds.SCRIPT_EDITOR);
            const dispatch = () => {
              window.dispatchEvent(new CustomEvent("formula:script-editor:set-code", { detail: { code } }));
            };
            // Allow the layout to mount the panel before we attempt to populate the editor.
            requestAnimationFrame(() => requestAnimationFrame(dispatch));
          };

          const storedMacroKey = (id: string) => `formula:macros:${id}`;
          const saveScriptsToStorage = (macroName: string, scripts: { ts: string; py: string }) => {
            try {
              const storage = globalThis.localStorage;
              if (!storage) return;
              const key = storedMacroKey(scriptStorageId);
              const raw = storage.getItem(key);
              const existing = raw ? JSON.parse(raw) : [];
              const list = Array.isArray(existing) ? existing.filter((m) => m && typeof m === "object") : [];
              const baseId = `recorded-${Date.now()}`;
              list.push(
                {
                  id: `${baseId}-ts`,
                  name: `${macroName} (TS)`,
                  language: "typescript",
                  module: "recorded",
                  code: scripts.ts,
                },
                {
                  id: `${baseId}-py`,
                  name: `${macroName} (Python)`,
                  language: "python",
                  module: "recorded",
                  code: scripts.py,
                },
              );
              storage.setItem(key, JSON.stringify(list));
            } catch {
              // Ignore storage errors (disabled storage, quota, etc).
            }
          };

          const currentActions = () => macroRecorder.getOptimizedActions();

          const updateRecorderUi = () => {
            status.textContent = macroRecorder.recording ? "Recording…" : "Not recording";
            const raw = macroRecorder.getRawActions();
            const optimized = macroRecorder.getOptimizedActions();
            meta.textContent = `Recorded: ${raw.length} steps · Optimized: ${optimized.length} steps`;
            preview.textContent = optimized.length ? JSON.stringify(optimized, null, 2) : "(no recorded actions)";
          };

          startButton.onclick = () => {
            macroRecorder.start();
            updateRecorderUi();
          };

          stopButton.onclick = () => {
            macroRecorder.stop();
            updateRecorderUi();
          };

          copyTsButton.onclick = async () => {
            const actions = currentActions();
            if (actions.length === 0) return;
            await copyText(generateTypeScriptMacro(actions));
          };

          copyPyButton.onclick = async () => {
            const actions = currentActions();
            if (actions.length === 0) return;
            await copyText(generatePythonMacro(actions));
          };

          openScriptEditorButton.onclick = () => {
            const actions = currentActions();
            if (actions.length === 0) return;
            openScriptEditor(generateTypeScriptMacro(actions));
          };

          saveButton.onclick = async () => {
            const actions = currentActions();
            if (actions.length === 0) return;
            const name = window.prompt("Macro name:", "Recorded Macro");
            if (!name) return;
            saveScriptsToStorage(name, { ts: generateTypeScriptMacro(actions), py: generatePythonMacro(actions) });
            await refreshRunner();
          };

          updateRecorderUi();

          void refreshRunner().catch((err) => {
            runnerPanel.textContent = `Failed to load macros: ${String(err)}`;
          });
        } catch (err) {
          body.textContent = `Macros backend not available: ${String(err)}`;
        }
      });
    },
  });

  function renderPanelBody(panelId: string, body: HTMLDivElement) {
    if (panelId === PanelIds.SCRIPT_EDITOR) {
      let mount = panelMounts.get(panelId);
      if (!mount) {
        const container = document.createElement("div");
        container.className = "dock-panel__mount";
        const mounted = mountScriptEditorPanel({ workbook: scriptingWorkbook, container });
        mount = { container, dispose: mounted.dispose };
        panelMounts.set(panelId, mount);
      }

      body.replaceChildren();
      body.appendChild(mount.container);
      return;
    }

    if (panelId === PanelIds.PYTHON) {
      let mount = panelMounts.get(panelId);
      if (!mount) {
        const container = document.createElement("div");
        container.className = "dock-panel__mount";
        container.textContent = "Loading Python runtime…";

        let disposed = false;
        mount = {
          container,
          dispose: () => {
            disposed = true;
            container.innerHTML = "";
          },
        };
        panelMounts.set(panelId, mount);

        import("./panels/python/pythonPanelMount.js")
          .then(({ mountPythonPanel }) => {
            if (disposed) return;
             const mounted = mountPythonPanel({
               doc: app.getDocument(),
               container,
               workbookId,
               invoke: queuedInvoke ?? undefined,
               drainBackendSync,
               getActiveSheetId: () => app.getCurrentSheetId(),
               getSelection: () => {
                 const ranges = app.getSelectionRanges();
                 const first = ranges[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
                return {
                  sheet_id: app.getCurrentSheetId(),
                  start_row: first.startRow,
                  start_col: first.startCol,
                  end_row: first.endRow,
                  end_col: first.endCol,
                };
              },
              setSelection: (selection) => {
                if (!selection || !selection.sheet_id) return;
                if (selection.start_row === selection.end_row && selection.start_col === selection.end_col) {
                  app.activateCell({ sheetId: selection.sheet_id, row: selection.start_row, col: selection.start_col });
                  return;
                }
                app.selectRange({
                  sheetId: selection.sheet_id,
                  range: {
                    startRow: selection.start_row,
                    startCol: selection.start_col,
                    endRow: selection.end_row,
                    endCol: selection.end_col,
                  },
                });
              },
            });
            mount!.dispose = mounted.dispose;
          })
          .catch((err) => {
            if (disposed) return;
            container.textContent = `Failed to load Python runtime: ${String(err)}`;
          });
      }

      body.replaceChildren();
      body.appendChild(mount.container);
      return;
    }

    const panelDef = panelRegistry.get(panelId) as any;
    if (panelDef?.source?.kind === "extension" && !extensionHostManager.ready) {
      // If the user persisted an extension panel in their layout, ensure the extension host is
      // loaded so the view can activate and populate its webview.
      void ensureExtensionsLoaded();
    }

    panelBodyRenderer.renderPanelBody(panelId, body);
  }

  function openPanelIds() {
    const layout = layoutController.layout;
    return [
      ...layout.docks.left.panels,
      ...layout.docks.right.panels,
      ...layout.docks.bottom.panels,
      ...Object.keys(layout.floating),
    ];
  }

  const SVG_NS = "http://www.w3.org/2000/svg";

  function iconSvg(kind: "dock-left" | "dock-right" | "dock-bottom" | "float" | "close") {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("width", "16");
    svg.setAttribute("height", "16");
    svg.setAttribute("fill", "none");
    svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", "1.5");
    svg.setAttribute("stroke-linecap", "round");
    svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("aria-hidden", "true");

    const path = (d: string) => {
      const el = document.createElementNS(SVG_NS, "path");
      el.setAttribute("d", d);
      svg.appendChild(el);
    };

    switch (kind) {
      case "dock-left": {
        path("M3 3v10");
        path("M12 8H5");
        path("M8 5L5 8l3 3");
        break;
      }
      case "dock-right": {
        path("M13 3v10");
        path("M4 8H11");
        path("M8 5l3 3-3 3");
        break;
      }
      case "dock-bottom": {
        path("M3 13H13");
        path("M8 4V11");
        path("M5 8l3 3 3-3");
        break;
      }
      case "float": {
        // Two overlapping rectangles.
        path("M3 5h8v8H3z");
        path("M5 3h8v8H5z");
        break;
      }
      case "close": {
        path("M4 4l8 8");
        path("M12 4l-8 8");
        break;
      }
      default: {
        // Exhaustive guard.
        const _exhaustive: never = kind;
        return _exhaustive;
      }
    }

    return svg;
  }

  function renderDock(el: HTMLElement, zone: { panels: string[]; active: string | null }, currentSide: "left" | "right" | "bottom") {
    el.replaceChildren();
    if (zone.panels.length === 0) return;

    const active = zone.active ?? zone.panels[0];
    if (!active) return;

    function focusDockTab(panelId: string): void {
      const testId = `dock-tab-${panelId}`;
      for (const tab of el.querySelectorAll<HTMLButtonElement>(".dock-panel__tab")) {
        if (tab.dataset.testid === testId) {
          tab.focus();
          // Ensure the newly-focused tab is visible when the tab strip overflows.
          tab.scrollIntoView({ block: "nearest", inline: "nearest" });
          return;
        }
      }
    }

    function dockTabDomId(panelId: string): string {
      return `dock-tab-${encodeURIComponent(panelId)}`;
    }

    function dockTabPanelDomId(panelId: string): string {
      return `dock-tabpanel-${encodeURIComponent(panelId)}`;
    }

    const panel = document.createElement("div");
    panel.className = "dock-panel";
    panel.dataset.testid = `panel-${active}`;
    if (active === PanelIds.AI_CHAT) panel.dataset.testid = "panel-aiChat";

    const header = document.createElement("div");
    header.className = "dock-panel__header";

    const titleOrTabs = (() => {
      if (zone.panels.length <= 1) {
        const title = document.createElement("div");
        title.className = "dock-panel__title";
        title.textContent = panelTitle(active);
        return title;
      }

      const tabs = document.createElement("div");
      tabs.className = "dock-panel__tabs";
      tabs.setAttribute("role", "tablist");
      tabs.setAttribute("aria-label", "Docked panels");
      tabs.setAttribute("aria-orientation", "horizontal");

      for (const panelId of zone.panels) {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "dock-panel__tab";
        tab.dataset.testid = `dock-tab-${panelId}`;
        tab.id = dockTabDomId(panelId);
        tab.textContent = panelTitle(panelId);
        tab.title = tab.textContent ?? "";
        tab.setAttribute("role", "tab");
        tab.setAttribute("aria-selected", panelId === active ? "true" : "false");
        tab.setAttribute("aria-controls", dockTabPanelDomId(panelId));
        tab.tabIndex = panelId === active ? 0 : -1;
        tab.addEventListener("click", (e) => {
          e.preventDefault();
          if (panelId === active) return;
          layoutController.activateDockedPanel(panelId, currentSide);
          // `activateDockedPanel` triggers a synchronous re-render via the layout controller
          // change listener. Focus the newly-rendered tab so keyboard users can continue to
          // navigate the tab strip after clicking.
          focusDockTab(panelId);
        });
        tab.addEventListener("keydown", (e) => {
          const key = e.key;
          if (key !== "ArrowLeft" && key !== "ArrowRight" && key !== "Home" && key !== "End") return;
          e.preventDefault();

          const ids = zone.panels;
          const idx = ids.indexOf(panelId);
          if (idx < 0 || ids.length === 0) return;

          let nextIdx = idx;
          if (key === "Home") nextIdx = 0;
          else if (key === "End") nextIdx = ids.length - 1;
          else if (key === "ArrowLeft") nextIdx = (idx - 1 + ids.length) % ids.length;
          else nextIdx = (idx + 1) % ids.length;

          const nextId = ids[nextIdx];
          if (!nextId) return;

          layoutController.activateDockedPanel(nextId, currentSide);
          focusDockTab(nextId);
        });
        tabs.appendChild(tab);
      }

      return tabs;
    })();

    const controls = document.createElement("div");
    controls.className = "dock-panel__controls";

    function iconButton(label: string, testId: string, icon: SVGElement, onClick: () => void) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "dock-panel__control";
      btn.title = label;
      btn.setAttribute("aria-label", label);
      btn.appendChild(icon);
      btn.dataset.testid = testId;
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        onClick();
      });
      return btn;
    }

    // Dock controls. (Only left is required for the e2e smoke test, but we wire all sides.)
    if (currentSide !== "left") {
      controls.appendChild(
        iconButton("Dock left", active === PanelIds.AI_CHAT ? "dock-ai-panel-left" : "dock-panel-left", iconSvg("dock-left"), () => {
          layoutController.dockPanel(active, "left");
        }),
      );
    }

    if (currentSide !== "right") {
      controls.appendChild(
        iconButton("Dock right", "dock-panel-right", iconSvg("dock-right"), () => {
          layoutController.dockPanel(active, "right");
        }),
      );
    }

    if (currentSide !== "bottom") {
      controls.appendChild(
        iconButton("Dock bottom", "dock-panel-bottom", iconSvg("dock-bottom"), () => {
          layoutController.dockPanel(active, "bottom");
        }),
      );
    }

    controls.appendChild(
      iconButton("Float", "float-panel", iconSvg("float"), () => {
        const rect = (panelRegistry.get(active) as any)?.defaultFloatingRect ?? { x: 80, y: 80, width: 420, height: 560 };
        layoutController.floatPanel(active, rect);
      }),
    );

    controls.appendChild(
      iconButton("Close", active === PanelIds.AI_CHAT ? "close-ai-panel" : "close-panel", iconSvg("close"), () => {
        layoutController.closePanel(active);
      }),
    );

    header.appendChild(titleOrTabs);
    header.appendChild(controls);

    panel.appendChild(header);

    if (zone.panels.length > 1) {
      // Keep a tabpanel element for each open panel so `aria-controls` references always
      // point at an element in the DOM (inactive tabpanels remain empty + hidden).
      for (const panelId of zone.panels) {
        const body = document.createElement("div");
        body.className = "dock-panel__body";
        body.id = dockTabPanelDomId(panelId);
        body.setAttribute("role", "tabpanel");
        body.setAttribute("aria-labelledby", dockTabDomId(panelId));
        if (panelId !== active) {
          body.hidden = true;
        } else {
          renderPanelBody(active, body);
        }
        panel.appendChild(body);
      }
    } else {
      const body = document.createElement("div");
      body.className = "dock-panel__body";
      renderPanelBody(active, body);
      panel.appendChild(body);
    }

    el.appendChild(panel);
  }

  function renderFloating() {
    floatingRootEl.replaceChildren();
    const layout = layoutController.layout;

    type FloatingRect = { x: number; y: number; width: number; height: number };
    for (const [panelId, rect] of Object.entries(layout.floating) as Array<[string, FloatingRect]>) {
      const panel = document.createElement("div");
      panel.className = "floating-panel";
      panel.style.left = `${rect.x}px`;
      panel.style.top = `${rect.y}px`;
      panel.style.width = `${rect.width}px`;
      panel.style.height = `${rect.height}px`;

      const inner = document.createElement("div");
      inner.className = "dock-panel";

      const header = document.createElement("div");
      header.className = "dock-panel__header";

      const title = document.createElement("div");
      title.className = "dock-panel__title";
      title.textContent = panelTitle(panelId);

      const controls = document.createElement("div");
      controls.className = "dock-panel__controls";

      const dockLeftBtn = document.createElement("button");
      dockLeftBtn.type = "button";
      dockLeftBtn.className = "dock-panel__control";
      dockLeftBtn.title = "Dock left";
      dockLeftBtn.setAttribute("aria-label", "Dock left");
      dockLeftBtn.appendChild(iconSvg("dock-left"));
      dockLeftBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "left"));

      const dockRightBtn = document.createElement("button");
      dockRightBtn.type = "button";
      dockRightBtn.className = "dock-panel__control";
      dockRightBtn.title = "Dock right";
      dockRightBtn.setAttribute("aria-label", "Dock right");
      dockRightBtn.appendChild(iconSvg("dock-right"));
      dockRightBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "right"));

      const dockBottomBtn = document.createElement("button");
      dockBottomBtn.type = "button";
      dockBottomBtn.className = "dock-panel__control";
      dockBottomBtn.title = "Dock bottom";
      dockBottomBtn.setAttribute("aria-label", "Dock bottom");
      dockBottomBtn.appendChild(iconSvg("dock-bottom"));
      dockBottomBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "bottom"));

      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.className = "dock-panel__control";
      closeBtn.title = "Close";
      closeBtn.setAttribute("aria-label", "Close");
      closeBtn.appendChild(iconSvg("close"));
      closeBtn.addEventListener("click", () => layoutController.closePanel(panelId));

      controls.appendChild(dockLeftBtn);
      controls.appendChild(dockRightBtn);
      controls.appendChild(dockBottomBtn);
      controls.appendChild(closeBtn);

      header.appendChild(title);
      header.appendChild(controls);

      const body = document.createElement("div");
      body.className = "dock-panel__body";
      renderPanelBody(panelId, body);

      inner.appendChild(header);
      inner.appendChild(body);

      panel.appendChild(inner);
      floatingRootEl.appendChild(panel);
    }
  }

  function renderLayout() {
    applyPrimaryPaneZoomFromLayout();
    applyDockSizes();
    renderSplitView();
    renderDock(dockLeftEl, layoutController.layout.docks.left, "left");
    renderDock(dockRightEl, layoutController.layout.docks.right, "right");
    renderDock(dockBottomEl, layoutController.layout.docks.bottom, "bottom");
    renderFloating();

    const openPanels = openPanelIds();
    panelBodyRenderer.cleanup(openPanels);
    const openSet = new Set(openPanels);
    for (const [panelId, mount] of panelMounts.entries()) {
      if (openSet.has(panelId)) continue;
      mount.dispose();
      mount.container.remove();
      panelMounts.delete(panelId);
    }
  }

  const setActivePane = (pane: "primary" | "secondary") => {
    if (layoutController.layout.splitView.activePane === pane) return;
    layoutController.setActiveSplitPane(pane);
  };

  // Minimal UX: clicking/focusing a pane marks it active so the split-view outline is stable.
  gridRoot.addEventListener("pointerdown", () => setActivePane("primary"), { capture: true });
  gridRoot.addEventListener("focusin", () => setActivePane("primary"));
  gridSecondaryEl.addEventListener("pointerdown", () => setActivePane("secondary"), { capture: true });
  gridSecondaryEl.addEventListener("focusin", () => setActivePane("secondary"));

  // Splitter drag: pointermove adjusts ratio (0.1–0.9) based on cursor position.
  gridSplitterEl.addEventListener("pointerdown", (event) => {
    const direction = layoutController.layout.splitView.direction;
    if (direction === "none") return;
    if (event.button !== 0) return;
    event.preventDefault();

    const pointerId = event.pointerId;
    gridSplitterEl.setPointerCapture(pointerId);

    const updateRatio = (clientX: number, clientY: number) => {
      const rect = gridSplitEl.getBoundingClientRect();
      const size = direction === "vertical" ? rect.width : rect.height;
      if (size <= 0) return;
      const offset = direction === "vertical" ? clientX - rect.left : clientY - rect.top;
      const ratio = Math.max(0.1, Math.min(0.9, offset / size));
      // Dragging the splitter updates very frequently; avoid persisting on every event.
      layoutController.setSplitRatio(ratio, { persist: false });
    };

    updateRatio(event.clientX, event.clientY);

    const onMove = (move: PointerEvent) => {
      if (move.pointerId !== pointerId) return;
      move.preventDefault();
      updateRatio(move.clientX, move.clientY);
    };

    const onUp = (up: PointerEvent) => {
      if (up.pointerId !== pointerId) return;
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
      try {
        gridSplitterEl.releasePointerCapture(pointerId);
      } catch {
      // Ignore capture release errors.
      }

      // Persist the final ratio (without emitting an additional change event).
      persistLayoutNow();
    };

    window.addEventListener("pointermove", onMove, { passive: false });
    window.addEventListener("pointerup", onUp, { passive: false });
    window.addEventListener("pointercancel", onUp, { passive: false });
  });

  // --- Command palette -----------------------------------------------------------

  commandRegistry.registerBuiltinCommand(
    "checkForUpdates",
    t("commandPalette.command.checkForUpdates"),
    () => {
      void checkForUpdatesFromCommandPalette().catch((err) => {
        console.error("Failed to check for updates:", err);
        showToast(
          tWithVars("updater.checkFailedWithMessage", { message: String((err as any)?.message ?? err) }),
          "error",
          { timeoutMs: 10_000 },
        );
      });
    },
    { category: t("commandCategory.help") },
  );
  if (import.meta.env.DEV) {
    commandRegistry.registerBuiltinCommand(
      "debugShowSystemNotification",
      t("command.debugShowSystemNotification"),
      () => {
        void notify({ title: "Formula", body: "This is a test system notification." });
      },
      { category: t("commandCategory.debug") },
    );
  }

  const commandPalette = createCommandPalette({
    commandRegistry,
    contextKeys,
    keybindingIndex: commandKeybindingDisplayIndex,
    ensureExtensionsLoaded,
    onCloseFocus: () => app.focus(),
    placeholder: t("commandPalette.placeholder"),
    onSelectFunction: (name) => {
      const template = `=${name}()`;
      app.insertIntoFormulaBar(template, { focus: true, cursorOffset: template.length - 1 });
    },
    goTo: {
      workbook: app.getSearchWorkbook(),
      getCurrentSheetName: () => app.getCurrentSheetId(),
      onGoTo: (parsed) => {
        const { range } = parsed;
        if (range.startRow === range.endRow && range.startCol === range.endCol) {
          app.activateCell({ sheetId: parsed.sheetName, row: range.startRow, col: range.startCol });
        } else {
          app.selectRange({ sheetId: parsed.sheetName, range });
        }
      },
    },
  });

  openCommandPalette = commandPalette.open;

  // `registerBuiltinCommands(...)` wires this as a no-op so the Tauri shell can own
  // opening the palette. Override it in the browser/desktop UI so keybinding dispatch
  // through `CommandRegistry.executeCommand(...)` works as well.
  commandRegistry.registerBuiltinCommand(
    "workbench.showCommandPalette",
    t("command.workbench.showCommandPalette"),
    () => commandPalette.open(),
    {
      category: t("commandCategory.navigation"),
      icon: null,
      description: t("commandDescription.workbench.showCommandPalette"),
      keywords: ["command palette", "commands"],
    },
  );

  // Paste Special… (Ctrl/Cmd+Shift+V)
  window.addEventListener("keydown", (e) => {
    if (e.defaultPrevented) return;
    if (e.repeat) return;
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || !e.shiftKey || e.altKey) return;
    if (e.key !== "V" && e.key !== "v") return;

    const target = (e.target instanceof HTMLElement ? e.target : null) ?? (document.activeElement as HTMLElement | null);
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return;
    }

    e.preventDefault();
    executeBuiltinCommand("clipboard.pasteSpecial");
  });
  layoutController.on("change", () => {
    renderLayout();
    scheduleRibbonSelectionFormatStateUpdate();
  });
  rerenderLayout = () => {
    renderLayout();
    scheduleRibbonSelectionFormatStateUpdate();
  };
  renderLayout();
  scheduleRibbonSelectionFormatStateUpdate();

  // Allow panel content to request opening another panel.
  window.addEventListener("formula:open-panel", (evt) => {
    const panelId = (evt as any)?.detail?.panelId;
    if (typeof panelId !== "string") return;
    const placement = getPanelPlacement(layoutController.layout, panelId);
    if (placement.kind === "closed") layoutController.openPanel(panelId);
    else if (placement.kind === "docked") layoutController.activateDockedPanel(panelId, placement.side);
  });
}

const workbook = app.getSearchWorkbook();

function currentSheetDisplayName(): string {
  return sheetNameResolver.getSheetNameById(app.getCurrentSheetId()) ?? app.getCurrentSheetId();
}

function resolveSheetIdFromName(name: string): string | null {
  const trimmed = String(name ?? "").trim();
  if (!trimmed) return null;
  const resolved = sheetNameResolver.getSheetIdByName(trimmed);
  if (resolved) return resolved;

  // Allow addressing sheets by id as a fallback (e.g. older formulas or empty-sheet metadata).
  const needle = trimmed.toLowerCase();
  const docIds = app.getDocument().getSheetIds();
  return docIds.find((id) => id.toLowerCase() === needle) ?? null;
}

const findReplaceController = new FindReplaceController({
  workbook,
  getCurrentSheetName: () => currentSheetDisplayName(),
  getActiveCell: () => {
    const cell = app.getActiveCell();
    return { sheetName: currentSheetDisplayName(), row: cell.row, col: cell.col };
  },
  setActiveCell: ({ sheetName, row, col }) => {
    const sheetId = resolveSheetIdFromName(sheetName);
    if (!sheetId) return;
    app.activateCell({ sheetId, row, col });
  },
  getSelectionRanges: () => app.getSelectionRanges(),
  beginBatch: (opts) => app.getDocument().beginBatch(opts),
  endBatch: () => app.getDocument().endBatch()
});

const { findDialog, replaceDialog, goToDialog } = registerFindReplaceShortcuts({
  controller: findReplaceController,
  workbook,
  getCurrentSheetName: () => currentSheetDisplayName(),
  setActiveCell: ({ sheetName, row, col }) => {
    const sheetId = resolveSheetIdFromName(sheetName);
    if (!sheetId) return;
    app.activateCell({ sheetId, row, col });
  },
  selectRange: ({ sheetName, range }) => {
    const sheetId = resolveSheetIdFromName(sheetName);
    if (!sheetId) return;
    app.selectRange({ sheetId, range });
  }
});

function showDialogAndFocus(dialog: HTMLDialogElement): void {
  if (!dialog.open) {
    dialog.showModal();
  }

  const focusInput = () => {
    const input = dialog.querySelector<HTMLInputElement | HTMLTextAreaElement>("input, textarea");
    if (!input) return;
    input.focus();
    input.select?.();
  };

  requestAnimationFrame(focusInput);
}

commandRegistry.registerBuiltinCommand(
  "edit.find",
  t("command.edit.find"),
  () => showDialogAndFocus(findDialog as any),
  {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.find"),
    keywords: ["find", "search"],
  },
);

commandRegistry.registerBuiltinCommand(
  "edit.replace",
  t("command.edit.replace"),
  () => showDialogAndFocus(replaceDialog as any),
  {
    category: t("commandCategory.editing"),
    icon: null,
    description: t("commandDescription.edit.replace"),
    keywords: ["replace", "find"],
  },
);

commandRegistry.registerBuiltinCommand(
  "navigation.goTo",
  t("command.navigation.goTo"),
  () => showDialogAndFocus(goToDialog as any),
  {
    category: t("commandCategory.navigation"),
    icon: null,
    description: t("commandDescription.navigation.goTo"),
    keywords: ["go to", "goto", "reference", "name box"],
  },
);

function showDesktopOnlyToast(message: string): void {
  showToast(`Desktop-only: ${message}`);
}

function getTauriInvokeForPrint(): TauriInvoke | null {
  const invoke =
    queuedInvoke ?? ((globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined) ?? null;
  if (!invoke) {
    showDesktopOnlyToast("Print/Export is available in the desktop app.");
    return null;
  }
  return invoke;
}

function selectionBoundingBox1Based(): PrintCellRange {
  const ranges = app.getSelectionRanges();
  const active = app.getActiveCell();
  if (ranges.length === 0) {
    return { startRow: active.row + 1, endRow: active.row + 1, startCol: active.col + 1, endCol: active.col + 1 };
  }

  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;

  for (const r of ranges) {
    const startRow0 = Math.min(r.startRow, r.endRow);
    const endRow0 = Math.max(r.startRow, r.endRow);
    const startCol0 = Math.min(r.startCol, r.endCol);
    const endCol0 = Math.max(r.startCol, r.endCol);
    minRow = Math.min(minRow, startRow0);
    minCol = Math.min(minCol, startCol0);
    maxRow = Math.max(maxRow, endRow0);
    maxCol = Math.max(maxCol, endCol0);
  }

  if (!Number.isFinite(minRow) || !Number.isFinite(minCol) || !Number.isFinite(maxRow) || !Number.isFinite(maxCol)) {
    return { startRow: active.row + 1, endRow: active.row + 1, startCol: active.col + 1, endCol: active.col + 1 };
  }

  return { startRow: minRow + 1, endRow: maxRow + 1, startCol: minCol + 1, endCol: maxCol + 1 };
}

function decodeBase64ToBytes(data: string): Uint8Array {
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function downloadBytes(bytes: Uint8Array, filename: string, mime: string): void {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

type TauriPageSetup = {
  orientation: "portrait" | "landscape";
  paper_size: number;
  margins: {
    left: number;
    right: number;
    top: number;
    bottom: number;
    header: number;
    footer: number;
  };
  scaling:
    | { kind: "percent"; percent: number }
    | { kind: "fitTo"; width_pages: number; height_pages: number };
};

function pageSetupFromTauri(raw: any): PageSetup {
  const orientation = raw?.orientation === "landscape" ? "landscape" : "portrait";
  const paperSize = typeof raw?.paper_size === "number" ? raw.paper_size : Number(raw?.paper_size) || 9;
  const marginsRaw = raw?.margins ?? {};
  const margins = {
    left: Number(marginsRaw.left) || 0,
    right: Number(marginsRaw.right) || 0,
    top: Number(marginsRaw.top) || 0,
    bottom: Number(marginsRaw.bottom) || 0,
    header: Number(marginsRaw.header) || 0,
    footer: Number(marginsRaw.footer) || 0,
  };

  const scalingRaw = raw?.scaling ?? {};
  const scaling: PageSetup["scaling"] =
    scalingRaw.kind === "fitTo"
      ? {
          kind: "fitTo",
          widthPages: Number(scalingRaw.width_pages) || 1,
          heightPages: Number(scalingRaw.height_pages) || 0,
        }
      : { kind: "percent", percent: Number(scalingRaw.percent) || 100 };

  return { orientation, paperSize, margins, scaling };
}

function pageSetupToTauri(raw: PageSetup): TauriPageSetup {
  const toU16 = (value: number): number => {
    if (!Number.isFinite(value)) return 0;
    const rounded = Math.round(value);
    if (rounded < 0) return 0;
    if (rounded > 65535) return 65535;
    return rounded;
  };

  const scaling: TauriPageSetup["scaling"] =
    raw.scaling.kind === "fitTo"
      ? {
          kind: "fitTo",
          width_pages: toU16(raw.scaling.widthPages),
          height_pages: toU16(raw.scaling.heightPages),
        }
      : { kind: "percent", percent: toU16(raw.scaling.percent) };

  return {
    orientation: raw.orientation,
    paper_size: toU16(raw.paperSize),
    margins: raw.margins,
    scaling,
  };
}

function showPageSetupDialogModal(args: { initialValue: PageSetup; onChange: (next: PageSetup) => void }): void {
  const dialog = document.createElement("dialog");
  dialog.className = "page-setup-dialog";

  const container = document.createElement("div");
  dialog.appendChild(container);
  document.body.appendChild(dialog);

  const root = createRoot(container);

  const close = () => dialog.close();

  function Wrapper() {
    const [value, setValue] = React.useState<PageSetup>(args.initialValue);

    const handleChange = React.useCallback(
      (next: PageSetup) => {
        setValue(next);
        args.onChange(next);
      },
      [args],
    );

    return React.createElement(PageSetupDialog, { value, onChange: handleChange, onClose: close });
  }

  root.render(React.createElement(Wrapper));

  dialog.addEventListener(
    "close",
    () => {
      root.unmount();
      dialog.remove();
      app.focus();
    },
    { once: true },
  );

  dialog.addEventListener("cancel", (e) => {
    e.preventDefault();
    dialog.close();
  });

  dialog.showModal();
}

async function handleRibbonPageSetup(): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const sheetId = app.getCurrentSheetId();
    const settings = await invoke("get_sheet_print_settings", { sheet_id: sheetId });
    const pageSetup = pageSetupFromTauri((settings as any)?.page_setup);

    showPageSetupDialogModal({
      initialValue: pageSetup,
      onChange: (next) => {
        void invoke("set_sheet_page_setup", {
          sheet_id: sheetId,
          page_setup: pageSetupToTauri(next),
        }).catch((err) => {
          console.error("Failed to set page setup:", err);
          showToast(`Failed to update page setup: ${String(err)}`, "error");
        });
      },
    });
  } catch (err) {
    console.error("Failed to open page setup:", err);
    showToast(`Failed to open page setup: ${String(err)}`, "error");
  }
}

async function handleRibbonSetPrintArea(): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const sheetId = app.getCurrentSheetId();
    const range = selectionBoundingBox1Based();
    await invoke("set_sheet_print_area", {
      sheet_id: sheetId,
      print_area: [
        {
          start_row: range.startRow,
          end_row: range.endRow,
          start_col: range.startCol,
          end_col: range.endCol,
        },
      ],
    });
    app.focus();
  } catch (err) {
    console.error("Failed to set print area:", err);
    showToast(`Failed to set print area: ${String(err)}`, "error");
  }
}

async function handleRibbonClearPrintArea(): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const sheetId = app.getCurrentSheetId();
    await invoke("set_sheet_print_area", { sheet_id: sheetId, print_area: null });
    app.focus();
  } catch (err) {
    console.error("Failed to clear print area:", err);
    showToast(`Failed to clear print area: ${String(err)}`, "error");
  }
}

async function handleRibbonExportPdf(): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    // Best-effort: ensure any pending workbook sync changes are flushed before exporting.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await drainBackendSync();

    const sheetId = app.getCurrentSheetId();
    let range: PrintCellRange = selectionBoundingBox1Based();

    try {
      const settings = await invoke("get_sheet_print_settings", { sheet_id: sheetId });
      const printArea = (settings as any)?.print_area;
      const first = Array.isArray(printArea) ? printArea[0] : null;
      if (first) {
        range = {
          startRow: Number(first.start_row),
          endRow: Number(first.end_row),
          startCol: Number(first.start_col),
          endCol: Number(first.end_col),
        };
      }
    } catch (err) {
      console.warn("Failed to fetch print area settings; exporting selection instead:", err);
    }

    const b64 = await invoke("export_sheet_range_pdf", {
      sheet_id: sheetId,
      range: { start_row: range.startRow, end_row: range.endRow, start_col: range.startCol, end_col: range.endCol },
      col_widths_points: undefined,
      row_heights_points: undefined,
    });

    const bytes = decodeBase64ToBytes(String(b64));
    downloadBytes(bytes, `${sheetId}.pdf`, "application/pdf");
    app.focus();
  } catch (err) {
    console.error("Failed to export PDF:", err);
    showToast(`Failed to export PDF: ${String(err)}`, "error");
  }
}

mountRibbon(ribbonRoot, {
  fileActions: {
    newWorkbook: () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Creating new workbooks is available in the desktop app.");
        return;
      }
      void handleNewWorkbook().catch((err) => {
        console.error("Failed to create workbook:", err);
        showToast(`Failed to create workbook: ${String(err)}`, "error");
      });
    },
    openWorkbook: () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Opening workbooks is available in the desktop app.");
        return;
      }
      void promptOpenWorkbook().catch((err) => {
        console.error("Failed to open workbook:", err);
        showToast(`Failed to open workbook: ${String(err)}`, "error");
      });
    },
    saveWorkbook: () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Saving workbooks is available in the desktop app.");
        return;
      }
      void handleSave().catch((err) => {
        console.error("Failed to save workbook:", err);
        showToast(`Failed to save workbook: ${String(err)}`, "error");
      });
    },
    saveWorkbookAs: () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Save As is available in the desktop app.");
        return;
      }
      void handleSaveAs().catch((err) => {
        console.error("Failed to save workbook:", err);
        showToast(`Failed to save workbook: ${String(err)}`, "error");
      });
    },
    pageSetup: () => {
      void handleRibbonPageSetup().catch((err) => {
        console.error("Failed to open page setup:", err);
        showToast(`Failed to open page setup: ${String(err)}`, "error");
      });
    },
    print: () => {
      const invokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
      if (!invokeAvailable) {
        showDesktopOnlyToast("Print is available in the desktop app.");
        return;
      }
      showToast("Print is not implemented yet. Opening Page Setup…");
      void handleRibbonPageSetup().catch((err) => {
        console.error("Failed to open page setup:", err);
        showToast(`Failed to open page setup: ${String(err)}`, "error");
      });
    },
    closeWindow: () => {
      if (!handleCloseRequestForRibbon) {
        showDesktopOnlyToast("Closing windows is available in the desktop app.");
        return;
      }
      void handleCloseRequestForRibbon({ quit: false }).catch((err) => {
        console.error("Failed to close window:", err);
        showToast(`Failed to close window: ${String(err)}`, "error");
      });
    },
    quit: () => {
      if (!handleCloseRequestForRibbon) {
        showDesktopOnlyToast("Quitting is available in the desktop app.");
        return;
      }
      void handleCloseRequestForRibbon({ quit: true }).catch((err) => {
        console.error("Failed to quit app:", err);
        showToast(`Failed to quit app: ${String(err)}`, "error");
      });
    },
  },
  onToggle: (commandId, pressed) => {
    switch (commandId) {
      case "data.queriesConnections.queriesConnections": {
        const layoutController = ribbonLayoutController;
        if (!layoutController) {
          showToast("Queries panel is not available (layout controller missing).", "error");
          // Ensure the ribbon toggle state reflects the actual panel placement.
          scheduleRibbonSelectionFormatStateUpdate();
          app.focus();
          return;
        }

        if (pressed) layoutController.openPanel(PanelIds.DATA_QUERIES);
        else layoutController.closePanel(PanelIds.DATA_QUERIES);
        app.focus();
        return;
      }
      case "review.comments.showComments":
        if (pressed) app.openCommentsPanel();
        else app.closeCommentsPanel();
        return;
      case "view.show.showFormulas":
        app.setShowFormulas(pressed);
        app.focus();
        return;
      case "formulas.formulaAuditing.showFormulas":
        app.setShowFormulas(pressed);
        app.focus();
        return;
      case "view.show.performanceStats":
        app.setGridPerfStatsEnabled(pressed);
        app.focus();
        return;
      case "view.window.split": {
        if (!ribbonLayoutController) {
          showToast("Split view is not available.");
          return;
        }

        const currentDirection = ribbonLayoutController.layout.splitView.direction;
        if (!pressed) {
          ribbonLayoutController.setSplitDirection("none");
        } else if (currentDirection === "none") {
          ribbonLayoutController.setSplitDirection("vertical", 0.5);
        }

        app.focus();
        return;
      }
      case "home.font.bold":
        applyFormattingToSelection("Bold", (doc, sheetId, ranges) => toggleBold(doc, sheetId, ranges, { next: pressed }));
        return;
      case "home.font.italic":
        applyFormattingToSelection("Italic", (doc, sheetId, ranges) => toggleItalic(doc, sheetId, ranges, { next: pressed }));
        return;
      case "home.font.underline":
        applyFormattingToSelection("Underline", (doc, sheetId, ranges) =>
          toggleUnderline(doc, sheetId, ranges, { next: pressed }),
        );
        return;
      case "home.font.strikethrough":
        applyToSelection("Strikethrough", (sheetId, ranges) => {
          const doc = app.getDocument();
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { font: { strike: pressed } }, { label: "Strikethrough" });
          }
        });
        return;
      case "home.alignment.wrapText":
        applyFormattingToSelection("Wrap", (doc, sheetId, ranges) => toggleWrap(doc, sheetId, ranges, { next: pressed }));
        return;
      default:
        return;
    }
  },
  onCommand: (commandId) => {
    const doc = app.getDocument();

    // Toggle buttons trigger both `onToggle` and `onCommand`. We handle most toggle
    // semantics in `onToggle` (since it provides the `pressed` state). Avoid
    // falling through to the default "unimplemented" toast here.
    if (
      commandId === "home.font.bold" ||
      commandId === "home.font.italic" ||
      commandId === "home.font.underline" ||
      commandId === "home.font.strikethrough" ||
      commandId === "home.alignment.wrapText" ||
      commandId === "view.show.showFormulas" ||
      commandId === "view.show.performanceStats" ||
      commandId === "view.window.split" ||
      commandId === "review.comments.showComments" ||
      commandId === "data.queriesConnections.queriesConnections"
    ) {
      return;
    }

    if (
      commandId === "data.queriesConnections.refreshAll" ||
      commandId === "data.queriesConnections.refreshAll.refresh" ||
      commandId === "data.queriesConnections.refreshAll.refreshAllConnections" ||
      commandId === "data.queriesConnections.refreshAll.refreshAllQueries"
    ) {
      void (async () => {
        const service = powerQueryService;
        if (!service) {
          showToast("Queries service not available");
          return;
        }

        try {
          await service.ready;
        } catch (err) {
          console.error("Power Query service failed to initialize:", err);
          showToast("Queries service not available", "error");
          return;
        }

        const queries = service.getQueries();
        if (!queries.length) {
          showToast("No queries to refresh");
          return;
        }

        try {
          const handle = service.refreshAll();
          await handle.promise;
        } catch (err) {
          console.error("Failed to refresh all queries:", err);
          showToast(`Failed to refresh queries: ${String(err)}`, "error");
        }
      })();
      // Don't wait for the refresh to complete; return focus immediately so long-running
      // refresh jobs don't steal focus later when their promise settles.
      app.focus();
      return;
    }

    const openRibbonPanel = (panelId: string): void => {
      const layoutController = ribbonLayoutController;
      if (!layoutController) {
        showToast("Panels are not available (layout controller missing).", "error");
        return;
      }

      const placement = getPanelPlacement(layoutController.layout, panelId);
      if (placement.kind === "closed" || placement.kind === "docked") {
        layoutController.openPanel(panelId);
        return;
      }

      // Floating panels can be minimized; opening should restore them.
      const floating = layoutController.layout?.floating?.[panelId];
      if (floating?.minimized) {
        layoutController.setFloatingPanelMinimized(panelId, false);
      }
    };

    const openCustomZoomQuickPick = async (): Promise<void> => {
      if (!app.supportsZoom()) return;
      const baseOptions = [25, 50, 75, 100, 125, 150, 200];
      const current = Math.round(app.getZoom() * 100);
      const options = baseOptions.includes(current) ? baseOptions : [current, ...baseOptions];
      const picked = await showQuickPick(
        options.map((value) => ({ label: `${value}%`, value })),
        { placeHolder: "Zoom" },
      );
      if (picked == null) return;
      app.setZoom(picked / 100);
      syncZoomControl();
      app.focus();
    };

    const zoomMenuItemPrefix = "view.zoom.zoom.";
    if (commandId.startsWith(zoomMenuItemPrefix)) {
      const suffix = commandId.slice(zoomMenuItemPrefix.length);
      if (suffix === "custom") {
        void openCustomZoomQuickPick();
        return;
      }

      const percent = Number(suffix);
      if (Number.isFinite(percent) && Number.isInteger(percent) && percent > 0) {
        if (!app.supportsZoom()) return;

        app.setZoom(percent / 100);
        syncZoomControl();
        app.focus();
        return;
      }
    }

    const fontNamePrefix = "home.font.fontName.";
    if (commandId.startsWith(fontNamePrefix)) {
      const preset = commandId.slice(fontNamePrefix.length);
      const fontName = (() => {
        switch (preset) {
          case "calibri":
            return "Calibri";
          case "arial":
            return "Arial";
          case "times":
            return "Times New Roman";
          case "courier":
            return "Courier New";
          default:
            return null;
        }
      })();
      if (!fontName) return;
      applyToSelection("Font", (sheetId, ranges) => {
        for (const range of ranges) {
          doc.setRangeFormat(sheetId, range, { font: { name: fontName } }, { label: "Font" });
        }
      });
      return;
    }

    const fontSizePrefix = "home.font.fontSize.";
    if (commandId.startsWith(fontSizePrefix)) {
      const size = Number(commandId.slice(fontSizePrefix.length));
      if (!Number.isFinite(size) || size <= 0) return;
      applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, size));
      return;
    }

    const fillColorPrefix = "home.font.fillColor.";
    if (commandId.startsWith(fillColorPrefix)) {
      const preset = commandId.slice(fillColorPrefix.length);
      if (preset === "moreColors") {
        openColorPicker(fillColorPicker, "Fill color", (sheetId, ranges, argb) => setFillColor(doc, sheetId, ranges, argb));
        return;
      }
      const argb = (() => {
        switch (preset) {
          case "lightGray":
            return ["#", "FF", "D9D9D9"].join("");
          case "yellow":
            return ["#", "FF", "FFFF00"].join("");
          case "blue":
            return ["#", "FF", "0000FF"].join("");
          case "green":
            return ["#", "FF", "00FF00"].join("");
          case "red":
            return ["#", "FF", "FF0000"].join("");
          default:
            return null;
        }
      })();

      if (preset === "none" || preset === "noFill") {
        applyFormattingToSelection("Fill color", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { fill: null }, { label: "Fill color" });
          }
        });
        return;
      }

      if (argb) {
        applyFormattingToSelection("Fill color", (_doc, sheetId, ranges) => setFillColor(doc, sheetId, ranges, argb));
      }
      return;
    }

    const fontColorPrefix = "home.font.fontColor.";
    if (commandId.startsWith(fontColorPrefix)) {
      const preset = commandId.slice(fontColorPrefix.length);
      if (preset === "moreColors") {
        openColorPicker(fontColorPicker, "Font color", (sheetId, ranges, argb) => setFontColor(doc, sheetId, ranges, argb));
        return;
      }
      const argb = (() => {
        switch (preset) {
          case "black":
            return ["#", "FF", "000000"].join("");
          case "blue":
            return ["#", "FF", "0000FF"].join("");
          case "green":
            return ["#", "FF", "00FF00"].join("");
          case "red":
            return ["#", "FF", "FF0000"].join("");
          default:
            return null;
        }
      })();

      if (preset === "automatic") {
        applyFormattingToSelection("Font color", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { font: { color: null } }, { label: "Font color" });
          }
        });
        return;
      }

      if (argb) {
        applyFormattingToSelection("Font color", (_doc, sheetId, ranges) => setFontColor(doc, sheetId, ranges, argb));
      }
      return;
    }

    const clearPrefix = "home.font.clearFormatting.";
    if (commandId.startsWith(clearPrefix)) {
      const kind = commandId.slice(clearPrefix.length);
      if (kind === "clearFormats") {
        applyFormattingToSelection("Clear formats", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, null, { label: "Clear formats" });
          }
        });
        return;
      }
      if (kind === "clearContents") {
        applyFormattingToSelection("Clear contents", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.clearRange(sheetId, range, { label: "Clear contents" });
          }
        });
        return;
      }
      if (kind === "clearAll") {
        applyFormattingToSelection(
          "Clear all",
          (doc, sheetId, ranges) => {
            for (const range of ranges) {
              doc.clearRange(sheetId, range, { label: "Clear all" });
              doc.setRangeFormat(sheetId, range, null, { label: "Clear all" });
            }
          },
          { forceBatch: true },
        );
        return;
      }
      return;
    }

    const bordersPrefix = "home.font.borders.";
    if (commandId.startsWith(bordersPrefix)) {
      const kind = commandId.slice(bordersPrefix.length);
      const defaultBorderColor = ["#", "FF", "000000"].join("");
      if (kind === "none") {
        applyFormattingToSelection("Borders", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { border: null }, { label: "Borders" });
          }
        });
        return;
      }

      if (kind === "all") {
        applyFormattingToSelection("Borders", (_doc, sheetId, ranges) => applyAllBorders(doc, sheetId, ranges));
        return;
      }

      if (kind === "outside" || kind === "thickBox") {
        const edgeStyle = kind === "thickBox" ? "thick" : "thin";
        const edge = { style: edgeStyle, color: defaultBorderColor };
        applyFormattingToSelection(
          "Borders",
          (doc, sheetId, ranges) => {
            for (const range of ranges) {
              const startRow = range.start.row;
              const endRow = range.end.row;
              const startCol = range.start.col;
              const endCol = range.end.col;

              // Top edge.
              doc.setRangeFormat(
                sheetId,
                { start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } },
                { border: { top: edge } },
                { label: "Borders" },
              );

              // Bottom edge.
              doc.setRangeFormat(
                sheetId,
                { start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } },
                { border: { bottom: edge } },
                { label: "Borders" },
              );

              // Left edge.
              doc.setRangeFormat(
                sheetId,
                { start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } },
                { border: { left: edge } },
                { label: "Borders" },
              );

              // Right edge.
              doc.setRangeFormat(
                sheetId,
                { start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } },
                { border: { right: edge } },
                { label: "Borders" },
              );
            }
          },
          { forceBatch: true },
        );
        return;
      }

      const edge = { style: "thin", color: defaultBorderColor };
      const borderPatch = (() => {
        switch (kind) {
          case "bottom":
            return { border: { bottom: edge } };
          case "top":
            return { border: { top: edge } };
          case "left":
            return { border: { left: edge } };
          case "right":
            return { border: { right: edge } };
          default:
            return null;
        }
      })();

      if (borderPatch) {
        applyFormattingToSelection(
          "Borders",
          (doc, sheetId, ranges) => {
            for (const range of ranges) {
              const startRow = range.start.row;
              const endRow = range.end.row;
              const startCol = range.start.col;
              const endCol = range.end.col;

              const targetRange = (() => {
                switch (kind) {
                  case "bottom":
                    return { start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } };
                  case "top":
                    return { start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } };
                  case "left":
                    return { start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } };
                  case "right":
                    return { start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } };
                  default:
                    return range;
                }
              })();

              doc.setRangeFormat(sheetId, targetRange, borderPatch, { label: "Borders" });
            }
          },
          { forceBatch: true },
        );
      }
      return;
    }

    const numberFormatPrefix = "home.number.numberFormat.";
    if (commandId.startsWith(numberFormatPrefix)) {
      const kind = commandId.slice(numberFormatPrefix.length);
      if (kind === "general") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { numberFormat: null }, { label: "Number format" });
          }
        });
        return;
      }
      if (kind === "number") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { numberFormat: "0.00" }, { label: "Number format" });
          }
        });
        return;
      }
      if (kind === "currency" || kind === "accounting") {
        applyFormattingToSelection("Number format", (_doc, sheetId, ranges) =>
          applyNumberFormatPreset(doc, sheetId, ranges, "currency"),
        );
        return;
      }
      if (kind === "percentage") {
        applyFormattingToSelection("Number format", (_doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"));
        return;
      }
      if (kind === "shortDate" || kind === "longDate") {
        applyFormattingToSelection("Number format", (_doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"));
        return;
      }
      return;
    }

    const accountingPrefix = "home.number.accounting.";
    if (commandId.startsWith(accountingPrefix)) {
      // For now, treat all accounting currency picks as the default currency preset.
      applyFormattingToSelection("Number format", (_doc, sheetId, ranges) =>
        applyNumberFormatPreset(doc, sheetId, ranges, "currency"),
      );
      return;
    }

    switch (commandId) {
      case "review.comments.newComment":
        app.openCommentsPanel();
        return;

      case "file.new.new":
      case "file.new.blankWorkbook": {
        if (!tauriBackend) {
          showDesktopOnlyToast("Creating new workbooks is available in the desktop app.");
          return;
        }
        void handleNewWorkbook().catch((err) => {
          console.error("Failed to create workbook:", err);
          showToast(`Failed to create workbook: ${String(err)}`, "error");
        });
        return;
      }

      case "file.open.open": {
        if (!tauriBackend) {
          showDesktopOnlyToast("Opening workbooks is available in the desktop app.");
          return;
        }
        void promptOpenWorkbook().catch((err) => {
          console.error("Failed to open workbook:", err);
          showToast(`Failed to open workbook: ${String(err)}`, "error");
        });
        return;
      }

      case "file.save.save": {
        if (!tauriBackend) {
          showDesktopOnlyToast("Saving workbooks is available in the desktop app.");
          return;
        }
        void handleSave().catch((err) => {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        });
        return;
      }

      case "file.save.autoSave": {
        showToast("AutoSave is not implemented yet.");
        return;
      }

      case "file.save.saveAs":
      case "file.save.saveAs.copy":
      case "file.save.saveAs.download": {
        if (!tauriBackend) {
          showDesktopOnlyToast("Save As is available in the desktop app.");
          return;
        }
        void handleSaveAs().catch((err) => {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        });
        return;
      }

      case "file.export.createPdf":
      case "file.export.export.pdf":
      case "file.export.changeFileType.pdf": {
        void handleRibbonExportPdf().catch((err) => {
          console.error("Failed to export PDF:", err);
          showToast(`Failed to export PDF: ${String(err)}`, "error");
        });
        return;
      }

      case "file.export.export.csv":
      case "file.export.export.xlsx":
      case "file.export.changeFileType.csv":
      case "file.export.changeFileType.tsv":
      case "file.export.changeFileType.xlsx": {
        showToast("Export is not implemented yet.");
        return;
      }

      case "file.print.pageSetup": {
        void handleRibbonPageSetup().catch((err) => {
          console.error("Failed to open page setup:", err);
          showToast(`Failed to open page setup: ${String(err)}`, "error");
        });
        return;
      }

      case "file.print.print": {
        const invokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
        if (!invokeAvailable) {
          showDesktopOnlyToast("Print is available in the desktop app.");
          return;
        }
        showToast("Print is not implemented yet. Opening Page Setup…");
        void handleRibbonPageSetup().catch((err) => {
          console.error("Failed to open page setup:", err);
          showToast(`Failed to open page setup: ${String(err)}`, "error");
        });
        return;
      }

      case "file.print.printPreview": {
        const invokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
        if (!invokeAvailable) {
          showDesktopOnlyToast("Print Preview is available in the desktop app.");
          return;
        }
        showToast("Print Preview is not implemented yet. Exporting PDF instead…");
        void handleRibbonExportPdf().catch((err) => {
          console.error("Failed to export PDF:", err);
          showToast(`Failed to export PDF: ${String(err)}`, "error");
        });
        return;
      }

      case "file.print.pageSetup.printTitles":
      case "file.print.pageSetup.margins": {
        void handleRibbonPageSetup().catch((err) => {
          console.error("Failed to open page setup:", err);
          showToast(`Failed to open page setup: ${String(err)}`, "error");
        });
        return;
      }

      case "file.options.close": {
        if (handleCloseRequestForRibbon) {
          void handleCloseRequestForRibbon({ quit: false }).catch((err) => {
            console.error("Failed to close window:", err);
            showToast(`Failed to close window: ${String(err)}`, "error");
          });
          return;
        }

        const invokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
        if (!invokeAvailable) {
          showDesktopOnlyToast("Closing windows is available in the desktop app.");
        }
        void hideTauriWindow().catch((err) => {
          console.error("Failed to close window:", err);
          showToast(`Failed to close window: ${String(err)}`, "error");
        });
        return;
      }

      case "home.clipboard.cut":
        void app.clipboardCut();
        app.focus();
        return;
      case "home.clipboard.copy":
        void app.clipboardCopy();
        app.focus();
        return;
      case "home.clipboard.paste.default":
        void app.clipboardPaste();
        app.focus();
        return;
      case "home.clipboard.paste.values":
        void app.clipboardPasteSpecial("values");
        app.focus();
        return;
      case "home.clipboard.paste.formulas":
        void app.clipboardPasteSpecial("formulas");
        app.focus();
        return;
      case "home.clipboard.paste.formats":
        void app.clipboardPasteSpecial("formats");
        app.focus();
        return;
      case "home.clipboard.paste.transpose":
        showToast("Paste Transpose not implemented");
        app.focus();
        return;
      case "home.clipboard.pasteSpecial":
      case "home.clipboard.pasteSpecial.dialog":
        void (async () => {
          const picked = await showQuickPick(
            getPasteSpecialMenuItems().map((item) => ({ label: item.label, value: item })),
            { placeHolder: t("clipboard.pasteSpecial.title") },
          );
          if (picked == null) {
            app.focus();
            return;
          }
          await app.clipboardPasteSpecial(picked.mode);
          app.focus();
        })();
        return;
      case "home.clipboard.pasteSpecial.values":
        void app.clipboardPasteSpecial("values");
        app.focus();
        return;
      case "home.clipboard.pasteSpecial.formulas":
        void app.clipboardPasteSpecial("formulas");
        app.focus();
        return;
      case "home.clipboard.pasteSpecial.formats":
        void app.clipboardPasteSpecial("formats");
        app.focus();
        return;
      case "home.clipboard.pasteSpecial.transpose":
        showToast("Paste Transpose not implemented");
        app.focus();
        return;

      case "view.macros.viewMacros":
      case "view.macros.viewMacros.run":
      case "view.macros.viewMacros.edit":
      case "view.macros.viewMacros.delete":
        openRibbonPanel(PanelIds.MACROS);
        // "Edit…" in Excel normally opens an editor; best-effort surface our Script Editor panel too.
        if (commandId.endsWith(".edit")) {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
        }
        return;

      case "view.macros.recordMacro":
      case "view.macros.recordMacro.stop":
        openRibbonPanel(PanelIds.MACROS);
        return;

      case "view.macros.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. Avoid the default toast.
        return;

      case "developer.code.macros":
      case "developer.code.macros.run":
      case "developer.code.macros.edit":
        openRibbonPanel(PanelIds.MACROS);
        if (commandId.endsWith(".edit")) {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
        }
        return;

      case "developer.code.recordMacro":
      case "developer.code.recordMacro.stop":
      case "developer.code.macroSecurity":
      case "developer.code.macroSecurity.trustCenter":
        openRibbonPanel(PanelIds.MACROS);
        return;

      case "developer.code.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. Avoid the default toast.
        return;

      case "developer.code.visualBasic":
        // Desktop builds expose a VBA migration panel (used as a stand-in for the VBA editor).
        if (typeof (globalThis as any).__TAURI__?.core?.invoke === "function") {
          openRibbonPanel(PanelIds.VBA_MIGRATE);
        } else {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
        }
        return;

      case "formulas.formulaAuditing.tracePrecedents":
        app.clearAuditing();
        app.toggleAuditingPrecedents();
        app.focus();
        return;
      case "formulas.formulaAuditing.traceDependents":
        app.clearAuditing();
        app.toggleAuditingDependents();
        app.focus();
        return;
      case "formulas.formulaAuditing.removeArrows":
        app.clearAuditing();
        app.focus();
        return;
      case "insert.tables.pivotTable":
        ribbonLayoutController?.openPanel(PanelIds.PIVOT_BUILDER);
        window.dispatchEvent(new CustomEvent("pivot-builder:use-selection"));
        return;

      case "home.font.borders":
        // This command is a dropdown with menu items; the top-level command is not expected
        // to fire when the menu is present. Keep this as a fallback.
        applyFormattingToSelection("Borders", (_doc, sheetId, ranges) => applyAllBorders(doc, sheetId, ranges));
        return;
      case "home.font.fontColor":
        openColorPicker(fontColorPicker, "Font color", (sheetId, ranges, argb) =>
          setFontColor(doc, sheetId, ranges, argb),
        );
        return;
      case "home.font.fillColor":
        openColorPicker(fillColorPicker, "Fill color", (sheetId, ranges, argb) =>
          setFillColor(doc, sheetId, ranges, argb),
        );
        return;
      case "home.font.fontSize":
        void (async () => {
          const picked = await showQuickPick(
            [
              { label: "8", value: 8 },
              { label: "9", value: 9 },
              { label: "10", value: 10 },
              { label: "11", value: 11 },
              { label: "12", value: 12 },
              { label: "14", value: 14 },
              { label: "16", value: 16 },
              { label: "18", value: 18 },
              { label: "20", value: 20 },
              { label: "24", value: 24 },
              { label: "28", value: 28 },
              { label: "36", value: 36 },
              { label: "48", value: 48 },
              { label: "72", value: 72 },
            ],
            { placeHolder: "Font size" },
          );
          if (picked == null) return;
          applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, picked));
        })();
        return;

      case "home.font.increaseFont": {
        const current = activeCellFontSizePt();
        const next = stepFontSize(current, "increase");
        if (next !== current) {
          applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
        }
        return;
      }

      case "home.font.decreaseFont": {
        const current = activeCellFontSizePt();
        const next = stepFontSize(current, "decrease");
        if (next !== current) {
          applyFormattingToSelection("Font size", (_doc, sheetId, ranges) => setFontSize(doc, sheetId, ranges, next));
        }
        return;
      }

      case "home.alignment.alignLeft":
        applyFormattingToSelection("Align left", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "left"));
        return;
      case "home.alignment.topAlign":
        applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { vertical: "top" } }, { label: "Vertical align" });
          }
        });
        return;
      case "home.alignment.middleAlign":
        applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            // Spreadsheet vertical alignment uses "center" (Excel/OOXML); the grid maps this to CSS middle.
            doc.setRangeFormat(sheetId, range, { alignment: { vertical: "center" } }, { label: "Vertical align" });
          }
        });
        return;
      case "home.alignment.bottomAlign":
        applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { vertical: "bottom" } }, { label: "Vertical align" });
          }
        });
        return;
      case "home.alignment.center":
        applyFormattingToSelection("Align center", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "center"));
        return;
      case "home.alignment.alignRight":
        applyFormattingToSelection("Align right", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "right"));
        return;
      case "home.alignment.orientation.angleCounterclockwise":
        applyToSelection("Text orientation", (sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 45 } }, { label: "Text orientation" });
          }
        });
        return;
      case "home.alignment.orientation.angleClockwise":
        applyToSelection("Text orientation", (sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -45 } }, { label: "Text orientation" });
          }
        });
        return;
      case "home.alignment.orientation.verticalText":
        applyToSelection("Text orientation", (sheetId, ranges) => {
          for (const range of ranges) {
            // Excel/OOXML uses 255 as a sentinel for vertical text (stacked).
            doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 255 } }, { label: "Text orientation" });
          }
        });
        return;
      case "home.alignment.orientation.rotateUp":
        applyToSelection("Text orientation", (sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 90 } }, { label: "Text orientation" });
          }
        });
        return;
      case "home.alignment.orientation.rotateDown":
        applyToSelection("Text orientation", (sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -90 } }, { label: "Text orientation" });
          }
        });
        return;
      case "home.alignment.orientation.formatCellAlignment":
        openFormatCells();
        return;

      case "home.number.percent":
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "percent"));
        return;
      case "home.number.accounting":
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "currency"));
        return;
      case "home.number.date":
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => applyNumberFormatPreset(doc, sheetId, ranges, "date"));
        return;
      case "home.number.comma":
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { numberFormat: "#,##0.00" }, { label: "Number format" });
          }
        });
        return;
      case "home.number.increaseDecimal": {
        const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(), "increase");
        if (!next) return;
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
          }
        });
        return;
      }
      case "home.number.decreaseDecimal": {
        const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(), "decrease");
        if (!next) return;
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          for (const range of ranges) {
            doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
          }
        });
        return;
      }
      case "home.number.formatCells":
      case "home.number.moreFormats.formatCells":
      case "home.cells.format.formatCells":
        openFormatCells();
        return;
      case "home.editing.findSelect.find":
        showDialogAndFocus(findDialog);
        return;
      case "home.editing.findSelect.replace":
        showDialogAndFocus(replaceDialog);
        return;
      case "home.editing.findSelect.goTo":
        showDialogAndFocus(goToDialog);
        return;
      case "pageLayout.pageSetup.pageSetupDialog":
        void handleRibbonPageSetup();
        return;
      case "pageLayout.printArea.setPrintArea":
        void handleRibbonSetPrintArea();
        return;
      case "pageLayout.printArea.clearPrintArea":
        void handleRibbonClearPrintArea();
        return;
      case "pageLayout.export.exportPdf":
        void handleRibbonExportPdf();
        return;
      case "view.window.freezePanes.freezePanes":
        app.freezePanes();
        app.focus();
        return;
      case "view.window.freezePanes.freezeTopRow":
        app.freezeTopRow();
        app.focus();
        return;
      case "view.window.freezePanes.freezeFirstColumn":
        app.freezeFirstColumn();
        app.focus();
        return;
      case "view.window.freezePanes.unfreeze":
        app.unfreezePanes();
        app.focus();
        return;
      case "view.appearance.theme.system":
        themeController.setThemePreference("system");
        app.focus();
        return;
      case "view.appearance.theme.light":
        themeController.setThemePreference("light");
        app.focus();
        return;
      case "view.appearance.theme.dark":
        themeController.setThemePreference("dark");
        app.focus();
        return;
      case "view.appearance.theme.highContrast":
        themeController.setThemePreference("high-contrast");
        app.focus();
        return;

      // --- Debug / dev controls migrated from the legacy status bar ---------------
      // Keep these command ids stable because Playwright e2e depends on their `data-testid`s.
      case "open-panel-ai-chat":
        toggleDockPanel(PanelIds.AI_CHAT);
        return;
      case "open-panel-ai-audit":
        toggleDockPanel(PanelIds.AI_AUDIT);
        return;
      case "open-data-queries-panel":
        toggleDockPanel(PanelIds.DATA_QUERIES);
        return;
      case "open-macros-panel":
        toggleDockPanel(PanelIds.MACROS);
        return;
      case "open-script-editor-panel":
        toggleDockPanel(PanelIds.SCRIPT_EDITOR);
        return;
      case "open-python-panel":
        toggleDockPanel(PanelIds.PYTHON);
        return;
      case "open-marketplace-panel": {
        // Marketplace uses the extension host runtime for install/load actions; ensure it is started.
        const ensure = ensureExtensionsLoadedRef?.();
        if (ensure) {
          void ensure.catch(() => {
            // Best-effort; opening the panel should still work even if extension load fails.
          });
        }
        toggleDockPanel(PanelIds.MARKETPLACE);
        return;
      }
      case "open-extensions-panel": {
        // Extensions are lazy-loaded to keep startup light. Opening the Extensions panel
        // should trigger the host to load + sync contributed panels/commands.
        void ensureExtensionsLoadedRef?.()
          .then(() => {
            updateKeybindingsRef?.();
            syncContributedCommandsRef?.();
            syncContributedPanelsRef?.();
          })
          .catch(() => {
            // ignore; panel open/close should still work
          });
        toggleDockPanel(PanelIds.EXTENSIONS);
        return;
      }
      case "open-vba-migrate-panel":
        toggleDockPanel(PanelIds.VBA_MIGRATE);
        return;
      case "open-comments-panel":
        app.toggleCommentsPanel();
        return;
      case "audit-precedents":
        app.toggleAuditingPrecedents();
        app.focus();
        return;
      case "audit-dependents":
        app.toggleAuditingDependents();
        app.focus();
        return;
      case "audit-transitive":
        app.toggleAuditingTransitive();
        app.focus();
        return;
      case "split-vertical":
        ribbonLayoutController?.setSplitDirection("vertical", 0.5);
        return;
      case "split-horizontal":
        ribbonLayoutController?.setSplitDirection("horizontal", 0.5);
        return;
      case "split-none":
        ribbonLayoutController?.setSplitDirection("none", 0.5);
        return;
      case "freeze-panes":
        app.freezePanes();
        app.focus();
        return;
      case "freeze-top-row":
        app.freezeTopRow();
        app.focus();
        return;
      case "freeze-first-column":
        app.freezeFirstColumn();
        app.focus();
        return;
      case "unfreeze-panes":
        app.unfreezePanes();
        app.focus();
        return;
      case "view.zoom.zoom100":
        app.setZoom(1);
        syncZoomControl();
        app.focus();
        return;
      case "view.zoom.zoomToSelection":
        app.zoomToSelection();
        syncZoomControl();
        app.focus();
        return;
      case "view.zoom.zoom":
        void openCustomZoomQuickPick();
        return;
      default:
        if (commandId.startsWith("file.")) {
          showToast(`File command not implemented: ${commandId}`);
          return;
        }
        showToast(`Ribbon: ${commandId}`);
        return;
    }
  },
});
// In Yjs-backed collaboration mode the workbook is continuously persisted, but
// DocumentController's `isDirty` flips to true on essentially every local/remote
// change (including `applyExternalDeltas`). That makes the browser/Tauri
// beforeunload "unsaved changes" prompt effectively permanent and incorrect.
//
// SpreadsheetApp may attach collaboration support asynchronously, so we check
// `getCollabSession()` at prompt time instead of only once at startup.
const collabAwareDirtyController = {
  get isDirty(): boolean {
    if (isCollabModeActive()) return false;
    return app.getDocument().isDirty;
  },
};

installUnsavedChangesPrompt(window, collabAwareDirtyController);

function renderSheetSwitcher(sheets: { id: string; name: string }[], activeId: string) {
  sheetSwitcherEl.replaceChildren();
  for (const sheet of sheets) {
    const option = document.createElement("option");
    option.value = sheet.id;
    option.textContent = sheet.name;
    sheetSwitcherEl.appendChild(option);
  }
  sheetSwitcherEl.value = activeId;
}

sheetSwitcherEl.addEventListener("change", () => {
  app.activateSheet(sheetSwitcherEl.value);
  restoreFocusAfterSheetNavigation();
});

type TauriListen = (event: string, handler: (event: any) => void) => Promise<() => void>;

function getTauriListen(): TauriListen {
  const listen = (globalThis as any).__TAURI__?.event?.listen as TauriListen | undefined;
  if (!listen) {
    throw new Error("Tauri event API not available");
  }
  return listen;
}

type TauriEmit = (event: string, payload?: any) => Promise<void> | void;

function getTauriEmit(): TauriEmit | null {
  const emit = (globalThis as any).__TAURI__?.event?.emit as TauriEmit | undefined;
  return emit ?? null;
}

type TauriDialogOpen = (options?: Record<string, unknown>) => Promise<string | string[] | null>;
type TauriDialogSave = (options?: Record<string, unknown>) => Promise<string | null>;

function getTauriDialog(): { open: TauriDialogOpen; save: TauriDialogSave } {
  const dialog = (globalThis as any).__TAURI__?.dialog;
  const open = dialog?.open as TauriDialogOpen | undefined;
  const save = dialog?.save as TauriDialogSave | undefined;
  if (!open || !save) {
    throw new Error("Tauri dialog API not available");
  }
  return { open, save };
}

function getTauriWindowHandle(): any {
  const winApi = (globalThis as any).__TAURI__?.window;
  if (!winApi) {
    throw new Error("Tauri window API not available");
  }

  // Tauri v2 exposes window handles via helper functions; keep this flexible since
  // we intentionally avoid a hard dependency on `@tauri-apps/api`.
  const handle =
    (typeof winApi.getCurrentWebviewWindow === "function" ? winApi.getCurrentWebviewWindow() : null) ??
    (typeof winApi.getCurrentWindow === "function" ? winApi.getCurrentWindow() : null) ??
    (typeof winApi.getCurrent === "function" ? winApi.getCurrent() : null) ??
    winApi.appWindow ??
    null;

  if (!handle) {
    throw new Error("Tauri window handle not available");
  }
  return handle;
}

async function hideTauriWindow(): Promise<void> {
  try {
    const win = getTauriWindowHandle();
    if (typeof win.hide === "function") {
      await win.hide();
      return;
    }
    if (typeof win.close === "function") {
      await win.close();
      return;
    }
  } catch {
    // Ignore window API errors and fall back to the browser call below.
  }
  // Best-effort fallback; browsers may ignore this, but the call is harmless.
  window.close();
}

async function minimizeTauriWindow(): Promise<void> {
  try {
    const win = getTauriWindowHandle();
    if (typeof win.minimize === "function") {
      await win.minimize();
      return;
    }
    // Older/newer Tauri variants may expose a setter-style API.
    if (typeof win.setMinimized === "function") {
      await win.setMinimized(true);
    }
  } catch {
    // Best-effort; ignore window API failures.
  }
}

async function toggleTauriWindowMaximize(): Promise<void> {
  try {
    const win = getTauriWindowHandle();
    if (typeof win.toggleMaximize === "function") {
      await win.toggleMaximize();
      return;
    }

    const maximize = typeof win.maximize === "function" ? (win.maximize as () => Promise<void> | void).bind(win) : null;
    const unmaximize =
      typeof win.unmaximize === "function" ? (win.unmaximize as () => Promise<void> | void).bind(win) : null;

    if (maximize && unmaximize && typeof win.isMaximized === "function") {
      const isMaximized = await win.isMaximized();
      if (isMaximized) await unmaximize();
      else await maximize();
      return;
    }

    // Best-effort fallback when we can't query current state.
    if (maximize) {
      try {
        await maximize();
        return;
      } catch {
        // Fall through to unmaximize below.
      }
    }
    if (unmaximize) {
      await unmaximize();
    }
  } catch {
    // Best-effort; ignore window API failures.
  }
}

function encodeDocumentSnapshot(snapshot: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(snapshot));
}

function normalizeSheetList(info: WorkbookInfo): { id: string; name: string }[] {
  const sheets = Array.isArray(info.sheets) ? info.sheets : [];
  return sheets
    .map((s) => ({ id: String((s as any).id ?? ""), name: String((s as any).name ?? (s as any).id ?? "") }))
    .filter((s) => s.id.trim() !== "");
}

function isCollabModeActive(): boolean {
  try {
    return typeof (app as any).getCollabSession === "function" && Boolean((app as any).getCollabSession());
  } catch {
    return false;
  }
}

async function confirmDiscardDirtyState(actionLabel: string): Promise<boolean> {
  const doc = app.getDocument();
  if (!doc.isDirty) return true;
  if (isCollabModeActive()) return true;
  return nativeDialogs.confirm(`You have unsaved changes. Discard them and ${actionLabel}?`);
}

function queueBackendOp<T>(op: () => Promise<T>): Promise<T> {
  const result = pendingBackendSync.then(op);
  pendingBackendSync = result.then(() => undefined).catch((err) => {
    console.error("Failed to sync workbook changes to host:", err);
  });
  return result;
}

async function drainBackendSync(): Promise<void> {
  // `pendingBackendSync` is a growing promise chain. While awaiting it, more work can
  // be appended (e.g. a microtask-batched `set_cell` series). Loop until the chain
  // stabilizes so we don't interleave new workbook opens/saves with stale edits.
  while (true) {
    const current = pendingBackendSync;
    await current;
    if (pendingBackendSync === current) return;
  }
}

function randomSessionId(prefix: string): string {
  const randomUuid = (globalThis as any).crypto?.randomUUID;
  if (typeof randomUuid === "function") return `${prefix}:${randomUuid.call((globalThis as any).crypto)}`;
  return `${prefix}:${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;
}

async function computeWorkbookSignature(info: WorkbookInfo): Promise<string> {
  const basePath =
    typeof info.path === "string" && info.path.trim() !== ""
      ? info.path
      : typeof info.origin_path === "string" && info.origin_path.trim() !== ""
        ? info.origin_path
        : null;

  if (!basePath) {
    // Unsaved/new workbook: scope signatures to the current app session only.
    return randomSessionId("workbook");
  }

  const invoke = queuedInvoke ?? ((globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined);
  if (typeof invoke !== "function") return basePath;

  try {
    const stat = await invoke("stat_file", { path: basePath });
    const mtimeMs = (stat as any)?.mtimeMs ?? (stat as any)?.mtime_ms ?? null;
    const sizeBytes = (stat as any)?.sizeBytes ?? (stat as any)?.size_bytes ?? null;
    if (typeof mtimeMs === "number" && typeof sizeBytes === "number") {
      return `${basePath}:${mtimeMs}:${sizeBytes}`;
    }
  } catch {
    // Fall back to the path-only signature below.
  }

  return basePath;
}

async function loadWorkbookIntoDocument(info: WorkbookInfo): Promise<void> {
  if (!tauriBackend) {
    throw new Error("Workbook backend not available");
  }

  const doc = app.getDocument();
  const workbookSignaturePromise = computeWorkbookSignature(info);
  const sheets = normalizeSheetList(info);
  if (sheets.length === 0) {
    throw new Error("Workbook contains no sheets");
  }

  // We're about to replace the sheet metadata store; restart the doc->store sync once
  // `restoreDocumentState()` has applied the workbook snapshot.
  sheetStoreDocSync?.dispose();
  sheetStoreDocSync = null;
  sheetStoreDocSyncStore = null;

  workbookSheetStore = new WorkbookSheetStore(
    sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      visibility: "visible",
    })),
  );
  syncWorkbookSheetNamesFromSheetStore();
  installSheetStoreSubscription();

  const CHUNK_ROWS = 200;
  const { maxRows: MAX_ROWS, maxCols: MAX_COLS } = resolveWorkbookLoadLimits({
    queryString: typeof window !== "undefined" ? window.location.search : "",
    env: {
      ...((import.meta as any).env ?? {}),
      ...(((globalThis as any).process?.env as Record<string, unknown> | undefined) ?? {}),
    },
  });

  const snapshotSheets: Array<{ id: string; cells: any[] }> = [];
  let truncated = false;

  for (const sheet of sheets) {
    const cells: Array<{ row: number; col: number; value: unknown | null; formula: string | null; format: null }> = [];

    const usedRange = await tauriBackend.getSheetUsedRange(sheet.id);
    if (!usedRange) {
      snapshotSheets.push({ id: sheet.id, cells });
      continue;
    }

    const { startRow, endRow, startCol, endCol, truncatedRows, truncatedCols } = clampUsedRange(usedRange, {
      maxRows: MAX_ROWS,
      maxCols: MAX_COLS,
    });
    if (truncatedRows || truncatedCols) truncated = true;

    if (startRow > endRow || startCol > endCol) {
      snapshotSheets.push({ id: sheet.id, cells });
      continue;
    }

    for (let chunkStartRow = startRow; chunkStartRow <= endRow; chunkStartRow += CHUNK_ROWS) {
      const chunkEndRow = Math.min(endRow, chunkStartRow + CHUNK_ROWS - 1);
      const range = await tauriBackend.getRange({
        sheetId: sheet.id,
        startRow: chunkStartRow,
        startCol,
        endRow: chunkEndRow,
        endCol
      });

      const rows = Array.isArray(range?.values) ? range.values : [];

      for (let r = 0; r < rows.length; r++) {
        const rowValues = Array.isArray(rows[r]) ? rows[r] : [];
        for (let c = 0; c < rowValues.length; c++) {
          const cell = rowValues[c] as any;
          const formula = typeof cell?.formula === "string" ? cell.formula : null;
          const value = cell?.value ?? null;
          if (formula == null && value == null) continue;

          cells.push({
            row: chunkStartRow + r,
            col: startCol + c,
            value: formula != null ? null : value,
            formula,
            format: null
          });
        }
      }
    }

    snapshotSheets.push({ id: sheet.id, cells });
  }

  if (truncated) {
    const message = `Workbook is larger than the current load limit; only the first ${MAX_ROWS} rows and ${MAX_COLS} columns were loaded.`;
    console.warn(message);
    showToast(message, "warning");
  }

  const snapshot = encodeDocumentSnapshot({ schemaVersion: 1, sheets: snapshotSheets });
  const workbookSignature = await workbookSignaturePromise;
  // Reset Power Query table signatures before applying the snapshot so any
  // in-flight query executions cannot reuse cached table results from a
  // previously-opened workbook.
  refreshTableSignaturesFromBackend(doc, [], { workbookSignature });
  refreshDefinedNameSignaturesFromBackend(doc, [], { workbookSignature });
  await app.restoreDocumentState(snapshot);

  // Refresh workbook metadata (defined names + tables) used by the name box and
  // AI completion. This is separate from the cell snapshot that populates the
  // DocumentController.
  workbook.clearSchema();

  const [definedNames, tables] = await Promise.all([
    tauriBackend.listDefinedNames().catch(() => []),
    tauriBackend.listTables().catch(() => []),
  ]);

  const normalizedTables = tables.map((table) => {
    const rawSheetId = typeof (table as any)?.sheet_id === "string" ? String((table as any).sheet_id) : "";
    const sheet_id = rawSheetId ? workbookSheetStore.resolveIdByName(rawSheetId) ?? rawSheetId : rawSheetId;
    return { ...(table as any), sheet_id };
  });
  refreshTableSignaturesFromBackend(doc, normalizedTables as any, { workbookSignature });
  const normalizedDefinedNames = definedNames.map((entry) => {
    const refers_to = typeof (entry as any)?.refers_to === "string" ? String((entry as any).refers_to) : "";
    const { sheetName: explicitSheetName } = splitSheetQualifier(refers_to);
    const sheetIdFromRef = explicitSheetName ? workbookSheetStore.resolveIdByName(explicitSheetName) ?? explicitSheetName : null;
    const rawScopeSheet = typeof (entry as any)?.sheet_id === "string" ? String((entry as any).sheet_id) : null;
    const sheetIdFromScope = rawScopeSheet ? workbookSheetStore.resolveIdByName(rawScopeSheet) ?? rawScopeSheet : null;
    return { ...(entry as any), sheet_id: sheetIdFromScope ?? sheetIdFromRef };
  });
  refreshDefinedNameSignaturesFromBackend(doc, normalizedDefinedNames as any, { workbookSignature });

  for (const entry of definedNames) {
    const name = typeof (entry as any)?.name === "string" ? String((entry as any).name) : "";
    const refersTo =
      typeof (entry as any)?.refers_to === "string" ? String((entry as any).refers_to) : "";
    if (!name || !refersTo) continue;

    const { sheetName: explicitSheetName, ref } = splitSheetQualifier(refersTo);
    const sheetIdFromRef = explicitSheetName ? workbookSheetStore.resolveIdByName(explicitSheetName) ?? explicitSheetName : null;
    const rawScopeSheet = typeof (entry as any)?.sheet_id === "string" ? String((entry as any).sheet_id) : null;
    const sheetIdFromScope = rawScopeSheet ? workbookSheetStore.resolveIdByName(rawScopeSheet) ?? rawScopeSheet : null;
    const sheetId = sheetIdFromRef ?? sheetIdFromScope;
    if (!sheetId) continue;
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;

    let range: { startRow: number; endRow: number; startCol: number; endCol: number } | null = null;
    try {
      range = parseA1Range(ref);
    } catch {
      range = null;
    }
    if (!range) continue;

    workbook.defineName(name, { sheetName, range });
  }

  for (const table of normalizedTables) {
    const name = typeof (table as any)?.name === "string" ? String((table as any).name) : "";
    const sheetId = typeof (table as any)?.sheet_id === "string" ? String((table as any).sheet_id) : "";
    const sheetName = sheetId ? workbookSheetStore.getName(sheetId) ?? sheetId : "";
    const columns = Array.isArray((table as any)?.columns) ? (table as any).columns.map(String) : [];
    if (!name || !sheetName || columns.length === 0) continue;

    workbook.addTable({
      name,
      sheetName,
      startRow: Number((table as any).start_row) || 0,
      startCol: Number((table as any).start_col) || 0,
      endRow: Number((table as any).end_row) || 0,
      endCol: Number((table as any).end_col) || 0,
      columns,
    });
  }

  // Update chart series colors to reflect the workbook's theme palette (if available).
  try {
    const palette = await tauriBackend.getWorkbookThemePalette();
    app.setChartTheme(chartThemeFromWorkbookPalette(palette));
  } catch {
    app.setChartTheme(chartThemeFromWorkbookPalette(null));
  }

  doc.markSaved();

  const firstSheetId = sheets[0].id;
  renderSheetSwitcher(sheets, firstSheetId);
  app.activateSheet(firstSheetId);
  app.activateCell({ sheetId: firstSheetId, row: 0, col: 0 });
  app.refresh();
}

async function openWorkbookFromPath(
  path: string,
  options: { notifyExtensions?: boolean; throwOnCancel?: boolean } = {},
): Promise<void> {
  if (!tauriBackend) return;
  if (typeof path !== "string" || path.trim() === "") return;
  const ok = await confirmDiscardDirtyState("open another workbook");
  if (!ok) {
    if (options.throwOnCancel) {
      throw new Error("Open workbook cancelled");
    }
    return;
  }

  const hadActiveWorkbook = activeWorkbook != null;
  const previousPanelWorkbookId = activePanelWorkbookId;
  vbaEventMacros?.dispose();
  vbaEventMacros = null;

  stopPowerQueryService();
  try {
    // Allow any microtask-batched workbook edits to enqueue into the backend queue,
    // then drain the queue fully before swapping the workbook state.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await drainBackendSync();

    if (hadActiveWorkbook && queuedInvoke) {
      try {
        await fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        // Applying macro updates may schedule backend sync in a microtask; drain it to avoid
        // interleaving stale edits with the next workbook load.
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await drainBackendSync();
      } catch (err) {
        console.warn("Workbook_BeforeClose event macro failed:", err);
      }
    }

    workbookSync?.stop();
    workbookSync = null;

    activeWorkbook = await tauriBackend.openWorkbook(path);
    await loadWorkbookIntoDocument(activeWorkbook);
    if (options.notifyExtensions !== false) {
      try {
        // Prefer broadcasting the workbook snapshot as computed by the extension host, which will
        // incorporate any `spreadsheetApi.getActiveWorkbook()` metadata (name/path) instead of
        // relying on the host stub's filename inference.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const host = (extensionHostManagerForE2e as any)?.host;
        if (host && typeof host._getActiveWorkbook === "function" && typeof host._broadcastEvent === "function") {
          const workbook = await host._getActiveWorkbook();
          host._broadcastEvent("workbookOpened", { workbook });
        } else {
          extensionHostManagerForE2e?.host.openWorkbook(activeWorkbook.path ?? activeWorkbook.origin_path ?? path);
        }
      } catch {
        // Ignore extension host errors; workbook open should still succeed.
      }
    }
    activePanelWorkbookId = activeWorkbook.path ?? activeWorkbook.origin_path ?? path;
    syncTitlebar();
    startPowerQueryService();
    rerenderLayout?.();

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });

    if (queuedInvoke) {
      vbaEventMacros = installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
    }
  } catch (err) {
    activePanelWorkbookId = previousPanelWorkbookId;
    // If we were unable to swap workbooks, restore syncing for the previously-active
    // workbook so edits remain persistable.
    if (hadActiveWorkbook) {
      workbookSync = startWorkbookSync({
        document: app.getDocument(),
        engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
      });
      if (queuedInvoke) {
        vbaEventMacros = installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
      }
    }
    startPowerQueryService();
    throw err;
  }
}

async function promptOpenWorkbook(): Promise<void> {
  if (!tauriBackend) return;
  const { open } = getTauriDialog();
  const selection = await open({
    multiple: false,
    filters: getOpenFileFilters(),
  });

  const path = Array.isArray(selection) ? selection[0] : selection;
  if (typeof path !== "string" || path.trim() === "") return;
  await openWorkbookFromPath(path);
}

async function copyPowerQueryPersistence(fromWorkbookId: string, toWorkbookId: string): Promise<void> {
  if (!fromWorkbookId || !toWorkbookId) return;
  if (fromWorkbookId === toWorkbookId) return;

  // Query definitions are persisted inside the workbook file (`xl/formula/power-query.xml`),
  // but we still mirror them to localStorage during the migration window. Copy the mirror when
  // the workbook id changes (e.g. Save As from an unsaved session) so UI state remains stable.
  try {
    const queries = loadQueriesFromStorage(fromWorkbookId);
    if (queries.length > 0) saveQueriesToStorage(toWorkbookId, queries);
  } catch {
    // Ignore storage failures (disabled storage, quota, etc).
  }

  // Scheduled refresh metadata is persisted via the RefreshStateStore abstraction.
  try {
    const fromStore = createPowerQueryRefreshStateStore({ workbookId: fromWorkbookId });
    const toStore = createPowerQueryRefreshStateStore({ workbookId: toWorkbookId });
    const state = await fromStore.load();
    if (state && Object.keys(state).length > 0) {
      await toStore.save(state);
    }
  } catch {
    // Best-effort: persistence should never block saving.
  }
}

async function handleSave(options: { notifyExtensions?: boolean; throwOnCancel?: boolean } = {}): Promise<void> {
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (!workbookSync) return;

  if (!activeWorkbook.path) {
    await handleSaveAs(options);
    return;
  }

  if (options.notifyExtensions !== false) {
    try {
      extensionHostManagerForE2e?.host.saveWorkbook();
    } catch {
      // Ignore extension host errors; save should still succeed.
    }
  }
  await workbookSync.markSaved();
}

async function handleSaveAs(
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean; throwOnCancel?: boolean } = {},
): Promise<void> {
  if (!tauriBackend) return;
  if (!activeWorkbook) return;

  const previousPanelWorkbookId = options.previousPanelWorkbookId ?? activePanelWorkbookId;
  const { save } = getTauriDialog();
  const path = await save({
    filters: [
      { name: t("fileDialog.filters.excelWorkbook"), extensions: ["xlsx"] },
      { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] },
    ],
  });
  if (!path) {
    if (options.throwOnCancel) {
      throw new Error("Save cancelled");
    }
    return;
  }

  await handleSaveAsPath(path, { previousPanelWorkbookId, notifyExtensions: options.notifyExtensions });
}

async function handleSaveAsPath(
  path: string,
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean } = {},
): Promise<void> {
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (typeof path !== "string" || path.trim() === "") return;

  const previousPanelWorkbookId = options.previousPanelWorkbookId ?? activePanelWorkbookId;

  // Ensure any pending microtask-batched workbook edits are flushed before saving.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await drainBackendSync();
  if (options.notifyExtensions !== false) {
    try {
      extensionHostManagerForE2e?.host.saveWorkbookAs(path);
    } catch {
      // Ignore extension host errors; save should still succeed.
    }
  }
  if (queuedInvoke) {
    await queuedInvoke("save_workbook", { path });
  } else {
    await tauriBackend.saveWorkbook(path);
  }
  activeWorkbook = { ...activeWorkbook, path };
  app.getDocument().markSaved();
  syncTitlebar();

  await copyPowerQueryPersistence(previousPanelWorkbookId, path);
  activePanelWorkbookId = path;
  startPowerQueryService();
  rerenderLayout?.();
}

async function handleNewWorkbook(options: { notifyExtensions?: boolean; throwOnCancel?: boolean } = {}): Promise<void> {
  if (!tauriBackend) return;
  const ok = await confirmDiscardDirtyState("create a new workbook");
  if (!ok) {
    if (options.throwOnCancel) {
      throw new Error("Create workbook cancelled");
    }
    return;
  }

  const hadActiveWorkbook = activeWorkbook != null;
  const previousPanelWorkbookId = activePanelWorkbookId;
  const nextPanelWorkbookId = randomSessionId("workbook");
  vbaEventMacros?.dispose();
  vbaEventMacros = null;

  stopPowerQueryService();
  try {
    // Allow any microtask-batched workbook edits to enqueue into the backend queue,
    // then drain the queue fully before replacing the backend workbook state.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await drainBackendSync();

    if (hadActiveWorkbook && queuedInvoke) {
      try {
        await fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await drainBackendSync();
      } catch (err) {
        console.warn("Workbook_BeforeClose event macro failed:", err);
      }
    }

    workbookSync?.stop();
    workbookSync = null;

    activeWorkbook = await tauriBackend.newWorkbook();
    await loadWorkbookIntoDocument(activeWorkbook);
    if (options.notifyExtensions !== false) {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const host = (extensionHostManagerForE2e as any)?.host;
        if (host && typeof host._getActiveWorkbook === "function" && typeof host._broadcastEvent === "function") {
          const workbook = await host._getActiveWorkbook();
          host._broadcastEvent("workbookOpened", { workbook });
        } else {
          extensionHostManagerForE2e?.host.openWorkbook(activeWorkbook.path ?? activeWorkbook.origin_path);
        }
      } catch {
        // Ignore extension host errors; new workbook should still succeed.
      }
    }
    activePanelWorkbookId = nextPanelWorkbookId;
    syncTitlebar();
    startPowerQueryService();
    rerenderLayout?.();

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });

    if (queuedInvoke) {
      vbaEventMacros = installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
    }
  } catch (err) {
    activePanelWorkbookId = previousPanelWorkbookId;
    if (hadActiveWorkbook) {
      workbookSync = startWorkbookSync({
        document: app.getDocument(),
        engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
      });
      if (queuedInvoke) {
        vbaEventMacros = installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
      }
    }
    startPowerQueryService();
    throw err;
  }
}

try {
  tauriBackend = new TauriWorkbookBackend();
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (invoke) {
    queuedInvoke = (cmd, args) => queueBackendOp(() => invoke(cmd, args));
    // Expose the queued invoke so other subsystems (e.g. Power Query table reads)
    // can sequence behind pending workbook writes from `startWorkbookSync`.
    (globalThis as any).__FORMULA_WORKBOOK_INVOKE__ = queuedInvoke;
    (globalThis as any).__formulaQueuedInvoke = queuedInvoke;
  }

  // Ensure the tray indicator starts in a known-good state once the desktop backend is available.
  void setTrayStatus("idle");
  window.addEventListener("unload", () => {
    vbaEventMacros?.dispose();
    workbookSync?.stop();
  });

  // Tauri v2 event permissions are scoped in `apps/desktop/src-tauri/capabilities/main.json`.
  // If you add a new `listen(...)` (Rust -> JS) or `emit(...)` (JS -> Rust) call here, you MUST
  // update the corresponding allowlist there (and the `eventPermissions.vitest.ts` guardrail test)
  // or the call will fail with a permissions error.
  const listen = getTauriListen();
  const emit = getTauriEmit();
  let pendingOpenFiles: Promise<void> = Promise.resolve();

  const queueOpenWorkbook = (path: string) => {
    pendingOpenFiles = pendingOpenFiles.then(async () => {
      try {
        await openWorkbookFromPath(path);
      } catch (err) {
        console.error("Failed to open workbook:", err);
        void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
      }
    });
  };

  const updaterUiListeners = installUpdaterUi(listen);

  registerAppQuitHandlers({
    isDirty: () => app.getDocument().isDirty && !isCollabModeActive(),
    runWorkbookBeforeClose: async () => {
      if (!queuedInvoke) return;
      await fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
    },
    drainBackendSync,
    quitApp: async () => {
      if (!invoke) {
        window.close();
        return;
      }
      // Exit the desktop shell. The backend command hard-exits the process so this promise
      // will never resolve in the success path.
      await invoke("quit_app");
    },
    restartApp: async () => {
      if (!invoke) {
        window.close();
        return;
      }
      // Restart/exit using Tauri-managed shutdown semantics so updater installs can complete
      // without relying on capability-gated process relaunch APIs. Like `quit_app`, this promise
      // is expected to never resolve on success because the process terminates shortly after the
      // command is invoked.
      // without relying on capability-gated process relaunch APIs.
      try {
        await invoke("restart_app");
      } catch (err) {
        // Older builds may not expose `restart_app`; fall back to `quit_app` so the user
        // can still install the update and relaunch manually.
        console.warn("Failed to restart app; falling back to quit:", err);
        await invoke("quit_app");
      }
    },
  });

  // OAuth PKCE redirect capture:
  // The Rust host emits `oauth-redirect` when a deep-link/protocol handler is invoked
  // (e.g. `formula://oauth/callback?...`). Resolve the pending broker redirect without
  // requiring a manual copy/paste step.
  const oauthRedirectListener = listen("oauth-redirect", (event) => {
    const redirectUrl = (event as any)?.payload;
    if (typeof redirectUrl !== "string" || redirectUrl.trim() === "") return;
    oauthBroker.observeRedirect(redirectUrl);
  });

  // Signal that we're ready to receive (and flush any queued) oauth-redirect events.
  void oauthRedirectListener
    .then(() => {
      if (!emit) return;
      return Promise.resolve(emit("oauth-redirect-ready"));
    })
    .catch((err) => {
      console.error("Failed to signal oauth redirect readiness:", err);
    });

  // When the Rust host receives a close request, it asks the frontend to flush any pending
  // workbook-sync operations and to sync macro UI context before it runs `Workbook_BeforeClose`.
  void listen("close-prep", async (event) => {
    const token = (event as any)?.payload;
    if (typeof token !== "string" || token.trim() === "") return;

    try {
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      await drainBackendSync();

      const invokeForContext = queuedInvoke ?? invoke;
      if (invokeForContext) {
        const nonNegativeInt = (value: unknown) => {
          const num = typeof value === "number" ? value : Number(value);
          if (!Number.isFinite(num)) return 0;
          const floored = Math.floor(num);
          if (!Number.isSafeInteger(floored) || floored < 0) return 0;
          return floored;
        };

        const selection = currentSelectionRect();
        const active_row = nonNegativeInt(selection.activeRow);
        const active_col = nonNegativeInt(selection.activeCol);
        await invokeForContext("set_macro_ui_context", {
          workbook_id: workbookId,
          sheet_id: selection.sheetId,
          active_row,
          active_col,
          selection: {
            start_row: nonNegativeInt(selection.startRow),
            start_col: nonNegativeInt(selection.startCol),
            end_row: nonNegativeInt(selection.endRow),
            end_col: nonNegativeInt(selection.endCol),
          },
        });
      }
    } catch (err) {
      console.error("Failed to prepare close request:", err);
    } finally {
      if (emit) {
        try {
          await emit("close-prep-done", token);
        } catch (err) {
          console.error("Failed to acknowledge close request:", err);
        }
      }
    }
  });

  const openFileListener = listen("open-file", (event) => {
    const payload = (event as any)?.payload;
    if (!Array.isArray(payload)) return;
    const paths = payload.filter((p) => typeof p === "string" && p.trim() !== "");
    if (paths.length === 0) return;

    // Serialize opens to avoid overlapping prompts / backend state swaps.
    for (const path of paths) {
      queueOpenWorkbook(path);
    }
  });

  // Signal that we're ready to receive (and flush any queued) open-file requests.
  void openFileListener
    .then(() => {
      if (!emit) return;
      return Promise.resolve(emit("open-file-ready"));
    })
    .catch((err) => {
      console.error("Failed to signal open-file readiness:", err);
    });

  void listen("file-dropped", async (event) => {
    const paths = (event as any)?.payload;
    const first = Array.isArray(paths) ? paths[0] : null;
    if (typeof first !== "string" || first.trim() === "") return;
    queueOpenWorkbook(first);
  });

  void listen("tray-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("tray-new", () => {
    void handleNewWorkbook().catch((err) => {
      console.error("Failed to create workbook:", err);
      void nativeDialogs.alert(`Failed to create workbook: ${String(err)}`);
    });
  });

  void listen("tray-quit", () => {
    void requestAppQuit().catch((err) => {
      console.error("Failed to quit app:", err);
    });
  });

  // Native menu bar integration (desktop shell emits `menu-*` events).
  void listen("menu-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("menu-new", () => {
    void handleNewWorkbook().catch((err) => {
      console.error("Failed to create workbook:", err);
      void nativeDialogs.alert(`Failed to create workbook: ${String(err)}`);
    });
  });

  void listen("menu-save", () => {
    void handleSave().catch((err) => {
      console.error("Failed to save workbook:", err);
      void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
    });
  });

  void listen("menu-save-as", () => {
    void handleSaveAs().catch((err) => {
      console.error("Failed to save workbook:", err);
      void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
    });
  });

  void listen("menu-export-pdf", () => {
    void handleRibbonExportPdf().catch((err) => {
      console.error("Failed to export PDF:", err);
      showToast(`Failed to export PDF: ${String(err)}`, "error");
    });
  });

  void listen("menu-close-window", () => {
    void handleCloseRequest({ quit: false }).catch((err) => {
      console.error("Failed to close window:", err);
    });
  });

  void listen("menu-quit", () => {
    void requestAppQuit().catch((err) => {
      console.error("Failed to quit app:", err);
    });
  });

  const isMac = /Mac|iPhone|iPad|iPod/i.test(navigator.platform);
  const primaryModifiers = () => ({ ctrlKey: !isMac, metaKey: isMac });
  const getTextEditingTarget = (): HTMLElement | null => {
    const target = document.activeElement as HTMLElement | null;
    if (!target) return null;
    const tag = target.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return target;
    return null;
  };
  const tryExecCommand = (command: string, value?: string): boolean => {
    try {
      return document.execCommand(command, false, value);
    } catch {
      return false;
    }
  };
  const readClipboardTextBestEffort = async (): Promise<string | null> => {
    const navigatorClipboard = (globalThis as any)?.navigator?.clipboard;
    const readText = navigatorClipboard?.readText as (() => Promise<string>) | undefined;
    if (typeof readText === "function") {
      try {
        const text = await readText.call(navigatorClipboard);
        if (typeof text === "string") return text;
      } catch {
        // ignore and fall through
      }
    }

    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ((cmd: string, args?: any) => Promise<any>) | undefined;
    if (typeof invoke === "function") {
      for (const cmd of ["clipboard_read", "read_clipboard"]) {
        try {
          const payload = await invoke(cmd);
          const text = (payload as any)?.text;
          if (typeof text === "string") return text;
        } catch {
          // try next
        }
      }
    }

    const legacyReadText = (globalThis as any).__TAURI__?.clipboard?.readText as (() => Promise<string>) | undefined;
    if (typeof legacyReadText === "function") {
      try {
        const text = await legacyReadText();
        if (typeof text === "string") return text;
      } catch {
        // ignore
      }
    }

    return null;
  };
  const writeClipboardTextBestEffort = async (text: string): Promise<boolean> => {
    const navigatorClipboard = (globalThis as any)?.navigator?.clipboard;
    const writeText = navigatorClipboard?.writeText as ((text: string) => Promise<void>) | undefined;
    if (typeof writeText === "function") {
      try {
        await writeText.call(navigatorClipboard, text);
        return true;
      } catch {
        // ignore and fall through
      }
    }

    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ((cmd: string, args?: any) => Promise<any>) | undefined;
    if (typeof invoke === "function") {
      // Prefer the multi-format clipboard bridge (`clipboard_write`) when available.
      try {
        await invoke("clipboard_write", { payload: { text } });
        return true;
      } catch {
        // fall through
      }
      try {
        await invoke("write_clipboard", { text });
        return true;
      } catch {
        // fall through
      }
    }

    const legacyWriteText = (globalThis as any).__TAURI__?.clipboard?.writeText as ((text: string) => Promise<void>) | undefined;
    if (typeof legacyWriteText === "function") {
      try {
        await legacyWriteText(text);
        return true;
      } catch {
        // ignore
      }
    }

    return false;
  };
  const dispatchSpreadsheetShortcut = (key: string, opts: { shift?: boolean; alt?: boolean } = {}) => {
    const { ctrlKey, metaKey } = primaryModifiers();
    const e = new KeyboardEvent("keydown", {
      key,
      ctrlKey,
      metaKey,
      shiftKey: Boolean(opts.shift),
      altKey: Boolean(opts.alt),
      bubbles: true,
      cancelable: true,
    });
    gridRoot.dispatchEvent(e);
    app.focus();
  };

  void listen("menu-undo", () => {
    const target = getTextEditingTarget();
    if (target) {
      tryExecCommand("undo");
      return;
    }
    app.undo();
    app.focus();
  });
  void listen("menu-redo", () => {
    const target = getTextEditingTarget();
    if (target) {
      tryExecCommand("redo");
      return;
    }
    app.redo();
    app.focus();
  });
  void listen("menu-cut", () => {
    const target = getTextEditingTarget();
    if (target) {
      if (tryExecCommand("cut")) return;

      // Best-effort fallback for WebViews that block execCommand cut/copy.
      void (async () => {
        const selectedText = (() => {
          if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
            const start = target.selectionStart ?? 0;
            const end = target.selectionEnd ?? start;
            return start !== end ? target.value.slice(start, end) : "";
          }
          if (target.isContentEditable) {
            return window.getSelection()?.toString() ?? "";
          }
          return "";
        })();

        if (!selectedText) return;
        const ok = await writeClipboardTextBestEffort(selectedText);
        if (!ok) return;

        if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
          const start = target.selectionStart ?? 0;
          const end = target.selectionEnd ?? start;
          if (start !== end) {
            try {
              target.setRangeText("", start, end, "end");
            } catch {
              target.value = target.value.slice(0, start) + target.value.slice(end);
            }
            target.dispatchEvent(new Event("input", { bubbles: true }));
            target.dispatchEvent(new Event("change", { bubbles: true }));
          }
        } else if (target.isContentEditable) {
          try {
            window.getSelection()?.deleteFromDocument();
          } catch {
            // ignore
          }
        }
      })();
      return;
    }
    void app.cutToClipboard();
  });
  void listen("menu-copy", () => {
    const target = getTextEditingTarget();
    if (target) {
      if (tryExecCommand("copy")) return;

      // Best-effort fallback for WebViews that block execCommand cut/copy.
      void (async () => {
        const selectedText = (() => {
          if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
            const start = target.selectionStart ?? 0;
            const end = target.selectionEnd ?? start;
            return start !== end ? target.value.slice(start, end) : "";
          }
          if (target.isContentEditable) {
            return window.getSelection()?.toString() ?? "";
          }
          return "";
        })();

        if (!selectedText) return;
        await writeClipboardTextBestEffort(selectedText);
      })();
      return;
    }
    void app.copyToClipboard();
  });
  void listen("menu-paste", () => {
    const target = getTextEditingTarget();
    if (target) {
      if (tryExecCommand("paste")) return;

      // WebViews often block `execCommand("paste")`. Fall back to reading clipboard text
      // and inserting it at the current selection.
      void (async () => {
        const text = await readClipboardTextBestEffort();
        if (!text) return;

        if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
          const start = target.selectionStart ?? target.value.length;
          const end = target.selectionEnd ?? start;
          try {
            target.setRangeText(text, start, end, "end");
          } catch {
            target.value = target.value.slice(0, start) + text + target.value.slice(end);
          }
          target.dispatchEvent(new Event("input", { bubbles: true }));
          target.dispatchEvent(new Event("change", { bubbles: true }));
          return;
        }

        if (target.isContentEditable) {
          // Prefer execCommand insertText when available; it preserves undo history.
          if (tryExecCommand("insertText", text)) return;
          try {
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return;
            sel.deleteFromDocument();
            const range = sel.getRangeAt(0);
            const node = document.createTextNode(text);
            range.insertNode(node);
            range.setStartAfter(node);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);
          } catch {
            // ignore
          }
        }
      })();
      return;
    }
    void app.pasteFromClipboard();
  });
  void listen("menu-select-all", () => {
    const target = getTextEditingTarget();
    if (target) {
      if (tryExecCommand("selectAll")) return;
      if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
        try {
          target.select();
        } catch {
          // ignore
        }
      } else if (target.isContentEditable) {
        try {
          const range = document.createRange();
          range.selectNodeContents(target);
          const sel = window.getSelection();
          sel?.removeAllRanges();
          sel?.addRange(range);
        } catch {
          // ignore
        }
      }
      return;
    }
    dispatchSpreadsheetShortcut("a");
  });

  const zoomStepPercent = 10;
  const applyMenuZoom = (nextPercent: number) => {
    if (!app.supportsZoom()) return;
    if (!Number.isFinite(nextPercent) || nextPercent <= 0) return;
    const clamped = Math.min(500, Math.max(10, Math.round(nextPercent)));
    app.setZoom(clamped / 100);
    syncZoomControl();
    app.focus();
  };

  void listen("menu-zoom-in", () => {
    applyMenuZoom(Math.round(app.getZoom() * 100) + zoomStepPercent);
  });
  void listen("menu-zoom-out", () => {
    applyMenuZoom(Math.round(app.getZoom() * 100) - zoomStepPercent);
  });
  void listen("menu-zoom-reset", () => {
    applyMenuZoom(100);
  });

  void listen("menu-about", () => {
    showToast("Formula Desktop", "info");
  });

  // Some desktop builds trigger update checks directly from the Rust menu/tray handlers (and
  // still emit `menu-check-updates` for frontend bookkeeping). Track the last time we saw a
  // manual update-check event so we can avoid invoking a duplicate update check from JS.
  let lastManualUpdateCheckEventAtMs = 0;
  const recordManualUpdateCheckEvent = (event: unknown) => {
    const payload = (event as any)?.payload;
    if (payload?.source !== "manual") return;
    lastManualUpdateCheckEventAtMs = Date.now();
  };
  void listen("update-check-started", recordManualUpdateCheckEvent);
  void listen("update-check-already-running", recordManualUpdateCheckEvent);

  void listen("menu-check-updates", () => {
    // Keep a stable menu event id; the actual update UX is driven by the
    // `update-check-*` events emitted by the Rust updater wrapper.
    const suppressDuplicateWindowMs = 250;
    const fallbackDelayMs = 50;

    // If a manual update check was just kicked off by the backend, don't start another one.
    if (Date.now() - lastManualUpdateCheckEventAtMs < suppressDuplicateWindowMs) return;

    window.setTimeout(() => {
      if (Date.now() - lastManualUpdateCheckEventAtMs < suppressDuplicateWindowMs) return;
      lastManualUpdateCheckEventAtMs = Date.now();
      void checkForUpdatesFromCommandPalette("manual").catch((err) => {
        console.error("Failed to check for updates:", err);
        showToast(
          tWithVars("updater.checkFailedWithMessage", { message: String((err as any)?.message ?? err) }),
          "error",
          { timeoutMs: 10_000 },
        );
      });
    }, fallbackDelayMs);
  });

  // Updater UI (toasts / dialogs / focus management) is handled by `installUpdaterUi(...)`.
  void listen("shortcut-quick-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("shortcut-command-palette", () => {
    openCommandPalette?.();
  });

  // Updater events can fire very early (e.g. a fast startup update check). `listen()` is async,
  // so we wait for registration before signaling readiness to the backend. The Rust host will
  // defer the startup update check until it receives `updater-ui-ready`.
  void updaterUiListeners
    .then(() => {
      if (!emit) return;
      return Promise.resolve(emit("updater-ui-ready"));
    })
    .catch((err) => {
      console.error("Failed to install updater listeners or signal updater-ui-ready:", err);
    });
  let closeInFlight = false;
  type RawCellUpdate = {
    sheet_id: string;
    row: number;
    col: number;
    value: unknown | null;
    formula: string | null;
    display_value?: string;
  };

  function valuesEqual(a: unknown, b: unknown): boolean {
    if (a === b) return true;
    if (a == null || b == null) return false;
    if (typeof a !== "object" || typeof b !== "object") return false;
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch {
      return false;
    }
  }

  function inputEquals(before: any, after: any): boolean {
    return valuesEqual(before?.value ?? null, after?.value ?? null) && (before?.formula ?? null) === (after?.formula ?? null);
  }

  function normalizeCloseMacroUpdates(raw: unknown): Array<{
    sheetId: string;
    row: number;
    col: number;
    value: unknown | null;
    formula: string | null;
    displayValue: string;
  }> {
    if (!Array.isArray(raw) || raw.length === 0) return [];
    const out: Array<{
      sheetId: string;
      row: number;
      col: number;
      value: unknown | null;
      formula: string | null;
      displayValue: string;
    }> = [];
    for (const u of raw as RawCellUpdate[]) {
      const sheetId = typeof u?.sheet_id === "string" ? u.sheet_id.trim() : "";
      const row = Number((u as any)?.row);
      const col = Number((u as any)?.col);
      if (!sheetId) continue;
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      out.push({
        sheetId,
        row,
        col,
        value: (u as any).value ?? null,
        formula: typeof (u as any).formula === "string" ? (u as any).formula : null,
        displayValue: String((u as any).display_value ?? ""),
      });
    }
    return out;
  }

  async function handleCloseRequest({
    quit,
    beforeCloseUpdates,
    closeToken,
  }: {
    quit: boolean;
    beforeCloseUpdates?: unknown;
    closeToken?: string;
  }): Promise<void> {
    if (closeInFlight) return;
    closeInFlight = true;
    try {
      // The Rust host runs `Workbook_BeforeClose` when the user clicks the native window close
      // button (and then emits `close-requested` with any macro-driven updates). Other close
      // entry points (tray/menu) are handled entirely in the frontend.
      const shouldRunBeforeCloseMacro = Boolean(queuedInvoke) && (quit || (!quit && !closeToken));
      if (shouldRunBeforeCloseMacro && queuedInvoke) {
        try {
          // Best-effort Workbook_BeforeClose for tray/menu close flows (no prompt).
          await fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        } catch (err) {
          console.warn("Workbook_BeforeClose event macro failed:", err);
        }
      }

      if (!quit && beforeCloseUpdates) {
        const normalized = normalizeCloseMacroUpdates(beforeCloseUpdates);
        if (normalized.length > 0) {
          if (vbaEventMacros) {
            await vbaEventMacros.applyMacroUpdates(normalized, { label: "Workbook_BeforeClose" });
          } else {
            const doc = app.getDocument();
            if (typeof (doc as any).applyExternalDeltas === "function" && typeof (doc as any).getCell === "function") {
              const deltas: any[] = [];
              for (const update of normalized) {
                const before = (doc as any).getCell(update.sheetId, { row: update.row, col: update.col });
                const formula = update.formula == null ? null : normalizeFormulaTextOpt(update.formula);
                const value = formula ? null : (update.value ?? null);
                const after = { value, formula, styleId: before?.styleId ?? 0 };
                if (inputEquals(before, after)) continue;
                deltas.push({ sheetId: update.sheetId, row: update.row, col: update.col, before, after });
              }
              if (deltas.length > 0) {
                (doc as any).applyExternalDeltas(deltas, { source: "macro" });
              }
            } else {
              doc.beginBatch({ label: "Workbook_BeforeClose" });
              let committed = false;
              try {
                applyMacroCellUpdates(doc, normalized);
                committed = true;
              } finally {
                if (committed) doc.endBatch();
                else doc.cancelBatch();
              }
            }
            app.refresh();
            await app.whenIdle();
            app.refresh();
          }
        }
      }

      const doc = app.getDocument();
      if (doc.isDirty && !isCollabModeActive()) {
        const discard = await nativeDialogs.confirm(t("prompt.unsavedChangesDiscardConfirm"));
        if (!discard) return;
      }

      if (!quit) {
        await hideTauriWindow();
        return;
      }

      // Best-effort flush of any macro-driven workbook edits before exiting.
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      await drainBackendSync();
      if (!invoke) {
        window.close();
        return;
      }
      // Exit the desktop shell. The backend command hard-exits the process so this promise
      // will never resolve in the success path.
      await invoke("quit_app");
    } catch (err) {
      console.error("Failed to handle close request:", err);
    } finally {
      if (closeToken && emit) {
        try {
          await emit("close-handled", closeToken);
        } catch (err) {
          console.error("Failed to signal close handled:", err);
        }
      }
      closeInFlight = false;
    }
  }

  handleCloseRequestForRibbon = handleCloseRequest;

  void listen("close-requested", async (event) => {
    const payload = (event as any)?.payload;
    const token = typeof payload?.token === "string" ? String(payload.token) : undefined;
    await handleCloseRequest({ quit: false, beforeCloseUpdates: payload?.updates, closeToken: token });
  });

  window.addEventListener("keydown", (e) => {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return;

    const keyLower = e.key.toLowerCase();
    const shift = e.shiftKey;

    if (!shift && keyLower === "n") {
      e.preventDefault();
      void handleNewWorkbook().catch((err) => {
        console.error("Failed to create workbook:", err);
        void nativeDialogs.alert(`Failed to create workbook: ${String(err)}`);
      });
      return;
    }

    if (!shift && keyLower === "o") {
      e.preventDefault();
      void promptOpenWorkbook().catch((err) => {
        console.error("Failed to open workbook:", err);
        void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
      });
      return;
    }

    if (!shift && keyLower === "w") {
      e.preventDefault();
      void handleCloseRequest({ quit: false }).catch((err) => {
        console.error("Failed to close window:", err);
      });
      return;
    }

    if (!shift && keyLower === "q") {
      e.preventDefault();
      void requestAppQuit().catch((err) => {
        console.error("Failed to quit app:", err);
      });
      return;
    }

    if (keyLower === "s") {
      if (shift) {
        e.preventDefault();
        void handleSaveAs().catch((err) => {
          console.error("Failed to save workbook:", err);
          void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
        });
        return;
      }
      e.preventDefault();
      void handleSave().catch((err) => {
        console.error("Failed to save workbook:", err);
        void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
      });
    }
  });
} catch {
  // Not running under Tauri; desktop host integration is unavailable.
}

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaExtensionHostManager = extensionHostManagerForE2e;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaExtensionHost = extensionHostManagerForE2e?.host ?? null;

// Time-to-interactive instrumentation (best-effort, no-op for web builds).
void markStartupTimeToInteractive({ whenIdle: () => app.whenIdle() }).catch(() => {});
