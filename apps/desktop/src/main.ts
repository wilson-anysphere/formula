import { SpreadsheetApp } from "./app/spreadsheetApp";
import type { SheetNameResolver } from "./sheet/sheetNameResolver";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/command-palette.css";
import "./styles/dialogs.css";
import "./styles/extensions-ui.css";
import "./styles/extensions.css";
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
import "./styles/script-editor.css";
import "./styles/python-panel.css";

import React from "react";
import { createRoot } from "react-dom/client";

import { SheetTabStrip } from "./sheets/SheetTabStrip";

import { ThemeController } from "./theme/themeController.js";

import { mountRibbon } from "./ribbon/index.js";

import { computeSelectionFormatState } from "./ribbon/selectionFormatState.js";
import { setRibbonUiState } from "./ribbon/ribbonUiState.js";

import type { CellRange as GridCellRange } from "@formula/grid";

import { rewriteDocumentFormulasForSheetDelete, rewriteDocumentFormulasForSheetRename } from "./sheets/sheetFormulaRewrite";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import { SecondaryGridView } from "./grid/splitView/secondaryGridView.js";
import { resolveDesktopGridMode } from "./grid/shared/desktopGridMode.js";
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
import { FORMULA_RELEASES_URL, installUpdaterUi } from "./tauri/updaterUi";
import { installOpenFileIpc } from "./tauri/openFileIpc";
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
import { assertExtensionRangeWithinLimits } from "./extensions/rangeSizeGuard.js";
import { openFormatCellsDialog } from "./formatting/openFormatCellsDialog.js";
import { parseCollabShareLink, serializeCollabShareLink } from "./sharing/collabLink.js";
import { saveCollabConnectionForWorkbook, loadCollabConnectionForWorkbook } from "./sharing/collabConnectionStore.js";
import { loadCollabToken, storeCollabToken } from "./sharing/collabTokenStore.js";
import { DesktopExtensionHostManager } from "./extensions/extensionHostManager.js";
import { ExtensionPanelBridge } from "./extensions/extensionPanelBridge.js";
import { ContextKeyService } from "./extensions/contextKeys.js";
import { resolveMenuItems } from "./extensions/contextMenus.js";
import { CELL_CONTEXT_MENU_ID, COLUMN_CONTEXT_MENU_ID, CORNER_CONTEXT_MENU_ID, ROW_CONTEXT_MENU_ID } from "./extensions/menuIds.js";
import { buildContextMenuModel } from "./extensions/contextMenuModel.js";
import { getPrimaryCommandKeybindingDisplay, type ContributedKeybinding } from "./extensions/keybindings.js";
import { KeybindingService } from "./extensions/keybindingService.js";
import { deriveSelectionContextKeys } from "./extensions/selectionContextKeys.js";
import { CommandRegistry } from "./extensions/commandRegistry.js";
import { createCommandPalette, installCommandPaletteRecentsTracking } from "./command-palette/index.js";
import { registerBuiltinCommands } from "./commands/registerBuiltinCommands.js";
import { DEFAULT_GRID_LIMITS } from "./selection/selection.js";
import type { GridLimits, Range, SelectionState } from "./selection/types";
import { ContextMenu, type ContextMenuItem } from "./menus/contextMenu.js";
import { getPasteSpecialMenuItems } from "./clipboard/pasteSpecial.js";
import {
  WorkbookSheetStore,
  generateDefaultSheetName,
  validateSheetName,
  type SheetVisibility,
  type TabColor,
} from "./sheets/workbookSheetStore";
import {
  CollabWorkbookSheetStore,
  computeCollabSheetsKey,
  findCollabSheetIndexById,
  listSheetsFromCollabSession,
  type CollabSheetsKeyRef,
} from "./sheets/collabWorkbookSheetStore";
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
import { DLP_ACTION } from "../../../packages/security/dlp/src/actions.js";
import { maxClassification, DEFAULT_CLASSIFICATION } from "../../../packages/security/dlp/src/classification.js";
import { DlpViolationError } from "../../../packages/security/dlp/src/errors.js";
import { evaluatePolicy, DLP_DECISION } from "../../../packages/security/dlp/src/policyEngine.js";
import { effectiveRangeClassification } from "../../../packages/security/dlp/src/selectors.js";

import sampleHelloManifest from "../../../extensions/sample-hello/package.json";
import { purgeLegacyDesktopLLMSettings } from "./ai/llm/desktopLLMClient.js";
import {
  installStartupTimingsListeners,
  markStartupTimeToInteractive,
  reportStartupWebviewLoaded,
} from "./tauri/startupMetrics.js";
import { openExternalHyperlink } from "./hyperlinks/openExternal.js";
import {
  clampUsedRange,
  DEFAULT_DESKTOP_LOAD_MAX_COLS,
  DEFAULT_DESKTOP_LOAD_MAX_ROWS,
  resolveWorkbookLoadChunkRows,
  resolveWorkbookLoadLimits,
  WORKBOOK_LOAD_CHUNK_ROWS_STORAGE_KEY,
  WORKBOOK_LOAD_MAX_COLS_STORAGE_KEY,
  WORKBOOK_LOAD_MAX_ROWS_STORAGE_KEY,
} from "./workbook/load/clampUsedRange.js";
import { warnIfWorkbookLoadTruncated, type WorkbookLoadTruncation } from "./workbook/load/truncationWarning.js";
import { exportDocumentRangeToCsv } from "./import-export/csv/export.js";

// Best-effort: older desktop builds persisted provider selection + API keys in localStorage.
// Cursor desktop no longer supports user-provided keys; proactively delete stale secrets on startup.
try {
  purgeLegacyDesktopLLMSettings();
} catch {
  // ignore
}

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

function getWorkbookLoadLimits(): { maxRows: number; maxCols: number; chunkRows: number } {
  const overrides = (() => {
    try {
      const storage = globalThis.localStorage;
      if (!storage) return null;
      return {
        maxRows: storage.getItem(WORKBOOK_LOAD_MAX_ROWS_STORAGE_KEY),
        maxCols: storage.getItem(WORKBOOK_LOAD_MAX_COLS_STORAGE_KEY),
        chunkRows: storage.getItem(WORKBOOK_LOAD_CHUNK_ROWS_STORAGE_KEY),
      };
    } catch {
      // Ignore storage errors (disabled storage, etc).
      return null;
    }
  })();

  const queryString = typeof window !== "undefined" ? window.location.search : "";
  const env = {
    ...((import.meta as any).env ?? {}),
    ...(((globalThis as any).process?.env as Record<string, unknown> | undefined) ?? {}),
  };

  const limits = resolveWorkbookLoadLimits({
    queryString,
    env,
    overrides,
  });

  const chunkRows = resolveWorkbookLoadChunkRows({
    queryString,
    env,
    override: overrides?.chunkRows,
  });

  return { ...limits, chunkRows };
}
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

// Exposed to Playwright tests via `window.__formulaExtensionHostManager`.
let extensionHostManagerForE2e: DesktopExtensionHostManager | null = null;
let sharedContextMenu: ContextMenu | null = null;
// Split-view secondary pane wiring is initialized lazily once the layout scaffolding is present.
// Keep a module-scoped reference so file commands (open/save/close) can commit edits from the
// secondary pane without crashing in runtimes where split view is not initialized
// (e.g. Playwright, minimal DOM).
let secondaryGridView: SecondaryGridView | null = null;

type SheetActivatedEvent = { sheet: { id: string; name: string } };
const sheetActivatedListeners = new Set<(event: SheetActivatedEvent) => void>();

type ExtensionWorkbookLifecycleEvent = {
  workbook: {
    name: string;
    path: string | null;
    sheets: Array<{ id: string; name: string }>;
    activeSheet: { id: string; name: string };
  };
};

const workbookOpenedEventListeners = new Set<(event: ExtensionWorkbookLifecycleEvent) => void>();
const beforeSaveEventListeners = new Set<(event: ExtensionWorkbookLifecycleEvent) => void>();

function emitWorkbookOpenedForExtensions(workbook: ExtensionWorkbookLifecycleEvent["workbook"]): void {
  const event: ExtensionWorkbookLifecycleEvent = { workbook };
  for (const listener of [...workbookOpenedEventListeners]) {
    try {
      listener(event);
    } catch {
      // ignore
    }
  }
}

function emitBeforeSaveForExtensions(workbook: ExtensionWorkbookLifecycleEvent["workbook"]): void {
  const event: ExtensionWorkbookLifecycleEvent = { workbook };
  for (const listener of [...beforeSaveEventListeners]) {
    try {
      listener(event);
    } catch {
      // ignore
    }
  }
}

function getWorkbookSnapshotForExtensions(options: { pathOverride?: string | null } = {}): ExtensionWorkbookLifecycleEvent["workbook"] {
  const sheetId = app.getCurrentSheetId();
  const activeSheet = { id: sheetId, name: workbookSheetStore.getName(sheetId) ?? sheetId };

  const storedSheets = workbookSheetStore.listAll();
  const sheets =
    storedSheets.length > 0 ? storedSheets.map((sheet) => ({ id: sheet.id, name: sheet.name })) : [{ id: "Sheet1", name: "Sheet1" }];

  const rawPath =
    Object.prototype.hasOwnProperty.call(options, "pathOverride")
      ? options.pathOverride
      : activeWorkbook?.path ?? activeWorkbook?.origin_path ?? null;
  const trimmedPath = typeof rawPath === "string" ? rawPath.trim() : "";
  const path = trimmedPath ? trimmedPath : null;

  const name = (() => {
    const pick = typeof path === "string" && path.trim() !== "" ? path : null;
    if (!pick) return "Workbook";
    return pick.split(/[/\\]/).pop() ?? "Workbook";
  })();

  return { name, path, sheets, activeSheet };
}

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
let activeMacroRecorder: MacroRecorder | null = null;
type MacrosPanelFocusTarget = "runner-select" | "runner-run" | "runner-trust-center" | "recorder-start" | "recorder-stop";
let pendingMacrosPanelFocus: MacrosPanelFocusTarget | null = null;
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
  let hasEverSynced = false;
  let providerPollTimer: number | null = null;

  const stopProviderPoll = (): void => {
    if (providerPollTimer == null) return;
    globalThis.clearTimeout(providerPollTimer);
    providerPollTimer = null;
  };
  let currentPersistenceSession: unknown = null;
  let localPersistenceLoaded: boolean | null = null;
  let localPersistenceWaitStarted = false;

  const startProviderPoll = (): void => {
    if (providerPollTimer != null) return;
    // Re-arm on each tick only if polling is still needed. This avoids a race where
    // `render()` decides polling is no longer required but the polling tick would
    // otherwise reschedule itself unconditionally.
    const tick = (): void => {
      if (abortController.signal.aborted) {
        stopProviderPoll();
        return;
      }
      providerPollTimer = null;
      render();
      // `render()` will call `startProviderPoll()` again if it still needs polling.
    };
    providerPollTimer = globalThis.setTimeout(tick, 1000) as unknown as number;
  };

  const setIndicatorText = (
    text: string,
    meta: { mode?: string; conn?: string; sync?: string; docId?: string } = {},
  ): void => {
    element.textContent = text;
    element.title = text;
    if (meta.mode) element.dataset.collabMode = meta.mode;
    else delete element.dataset.collabMode;
    if (meta.conn) element.dataset.collabConn = meta.conn;
    else delete element.dataset.collabConn;
    if (meta.sync) element.dataset.collabSync = meta.sync;
    else delete element.dataset.collabSync;
    if (meta.docId) element.dataset.collabDocId = meta.docId;
    else delete element.dataset.collabDocId;
  };

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
    // Only subscribe when we can also detach later. If `off` is unavailable, fall back
    // to polling (see `startProviderPoll`) to avoid leaking listeners.
    if (typeof anyProvider.on !== "function" || typeof anyProvider.off !== "function") return;
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
    if (providerSynced) hasEverSynced = true;
    render();
  };

  const render = (): void => {
    if (abortController.signal.aborted) return;

    // Best-effort hint for offline mode (Playwright `setOffline` toggles this).
    const networkOnline =
      typeof navigator !== "undefined" && typeof navigator.onLine === "boolean" ? navigator.onLine : true;

    const session = getSession();
    if (!session) {
      detachProviderListeners(currentProvider);
      stopProviderPoll();
      currentProvider = null;
      providerStatus = null;
      providerSynced = null;
      hasEverSynced = false;
      currentPersistenceSession = null;
      localPersistenceLoaded = null;
      localPersistenceWaitStarted = false;
      setIndicatorText("Local", { mode: "local" });
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const s = session as any;
    const docId = getDocId(session);

    const hasLocalPersistence = Boolean(s?.persistence);
    if (session !== currentPersistenceSession) {
      currentPersistenceSession = session;
      localPersistenceLoaded = hasLocalPersistence ? false : null;
      localPersistenceWaitStarted = false;
    }

    const persistenceLoading = hasLocalPersistence && localPersistenceLoaded !== true;
    if (persistenceLoading && !localPersistenceWaitStarted) {
      localPersistenceWaitStarted = true;
      const sessionForWait = session;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const whenLoaded = (s as any).whenLocalPersistenceLoaded as (() => Promise<void>) | undefined;
      if (typeof whenLoaded === "function") {
        void Promise.resolve()
          .then(() => whenLoaded.call(s))
          .catch(() => {
            // Local persistence load failures should not crash the UI; CollabSession
            // can still operate in online mode.
          })
          .finally(() => {
            if (currentPersistenceSession !== sessionForWait) return;
            localPersistenceLoaded = true;
            if (!abortController.signal.aborted) render();
          });
      } else {
        // No explicit signal; treat persistence as ready.
        localPersistenceLoaded = true;
      }
    }

    const provider = (s?.provider as unknown) ?? null;
    if (provider !== currentProvider) {
      detachProviderListeners(currentProvider);
      stopProviderPoll();
      currentProvider = provider;
      providerStatus = null;
      providerSynced = null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      hasEverSynced = Boolean((currentProvider as any)?.synced);
      attachProviderListeners(currentProvider);
    }

    if (persistenceLoading) {
      setIndicatorText(`${docId} • Loading…`, { mode: "collab", conn: "loading", sync: "loading", docId });
      return;
    }

    // No provider: offline-only/local collab session.
    if (!currentProvider) {
      stopProviderPoll();
      setIndicatorText(`${docId} • Offline`, { mode: "collab", conn: "offline", sync: "offline", docId });
      return;
    }

    if (!networkOnline) {
      // Browser/network reports offline; show a clear offline status regardless of provider state.
      stopProviderPoll();
      setIndicatorText(`${docId} • Offline`, { mode: "collab", conn: "offline", sync: "offline", docId });
      return;
    }

    // Provider exists. If it doesn't support event subscriptions, fall back to polling
    // so the UI eventually reflects connection/sync changes.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const canSubscribe = typeof (currentProvider as any)?.on === "function" && typeof (currentProvider as any)?.off === "function";
    if (!canSubscribe) startProviderPoll();
    else stopProviderPoll();

    const synced = (() => {
      if (typeof providerSynced === "boolean") return providerSynced;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyProvider = currentProvider as any;
      if (typeof anyProvider.synced === "boolean") return anyProvider.synced;
      return false;
    })();

    if (synced) hasEverSynced = true;

    const connected = (() => {
      if (providerStatus === "connected") return true;
      if (providerStatus === "disconnected") return false;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyProvider = currentProvider as any;
      if (typeof anyProvider.wsconnected === "boolean") return anyProvider.wsconnected;
      if (typeof anyProvider.connected === "boolean") return anyProvider.connected;
      return false;
    })();

    const connecting = (() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const anyProvider = currentProvider as any;
      if (typeof anyProvider.wsconnecting === "boolean") return anyProvider.wsconnecting;

      // y-websocket reports `status: "disconnected"` both before the first connect and
      // when reconnecting. Treat "not yet synced" as Connecting to avoid flashing
      // Disconnected at startup.
      if (!hasEverSynced && !connected) return true;
      if (!hasEverSynced && providerStatus === "disconnected") return true;
      return false;
    })();

    const connectionLabel = connected ? "Connected" : connecting ? "Connecting…" : "Disconnected";

    const syncLabel = connected ? (synced ? "Synced" : "Syncing…") : connecting ? "Syncing…" : hasEverSynced ? "Not synced" : "Syncing…";

    setIndicatorText(`${docId} • ${connectionLabel} • ${syncLabel}`, {
      mode: "collab",
      conn: connected ? "connected" : connecting ? "connecting" : "disconnected",
      sync: connected ? (synced ? "synced" : "syncing") : connecting ? "syncing" : hasEverSynced ? "unsynced" : "syncing",
      docId,
    });
  };

  const onNetworkChange = (): void => render();
  if (typeof window !== "undefined") {
    window.addEventListener("online", onNetworkChange);
    window.addEventListener("offline", onNetworkChange);
  }

  abortController.signal.addEventListener("abort", () => {
    window.removeEventListener("unload", cleanup);
    window.removeEventListener("online", onNetworkChange);
    window.removeEventListener("offline", onNetworkChange);
    detachProviderListeners(currentProvider);
    stopProviderPoll();
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

if (!document.getElementById("ribbon")) {
  throw new Error("Missing #ribbon container");
}
const ribbonReactRoot = document.getElementById("ribbon-react-root");
if (!ribbonReactRoot) {
  throw new Error("Missing #ribbon-react-root container");
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
const openVersionHistoryPanelButton = document.querySelector<HTMLButtonElement>('[data-testid="open-version-history-panel"]');
const openBranchManagerPanelButton = document.querySelector<HTMLButtonElement>('[data-testid="open-branch-manager-panel"]');
const openMarketplacePanelButton = document.querySelector<HTMLButtonElement>(
  '[data-testid="open-marketplace-panel-statusbar"]',
);
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
  !sheetPosition ||
  !openVersionHistoryPanelButton ||
  !openBranchManagerPanelButton ||
  !openMarketplacePanelButton
) {
  throw new Error("Missing status bar elements");
}
const sheetSwitcherEl = sheetSwitcher;
const zoomControlEl = zoomControl;
const statusZoomEl = statusZoom;
const sheetPositionEl = sheetPosition;
const openVersionHistoryPanelButtonEl = openVersionHistoryPanelButton;
const openBranchManagerPanelButtonEl = openBranchManagerPanelButton;
const openMarketplacePanelButtonEl = openMarketplacePanelButton;

// Collaboration panels should be accessible via always-visible status bar buttons.
openVersionHistoryPanelButtonEl.addEventListener("click", (e) => {
  e.preventDefault();
  toggleDockPanel(PanelIds.VERSION_HISTORY);
});
openBranchManagerPanelButtonEl.addEventListener("click", (e) => {
  e.preventDefault();
  toggleDockPanel(PanelIds.BRANCH_MANAGER);
});
openMarketplacePanelButtonEl.addEventListener("click", (e) => {
  e.preventDefault();
  toggleDockPanel(PanelIds.MARKETPLACE);
});

const docIdParam = new URL(window.location.href).searchParams.get("docId");
const docId = typeof docIdParam === "string" && docIdParam.trim() !== "" ? docIdParam : null;
const workbookId = docId ?? "local-workbook";
const legacyGridLimits = (() => {
  if (resolveDesktopGridMode() !== "legacy") return undefined;
  // `getWorkbookLoadLimits` also resolves snapshot chunking controls; strip it down to
  // `GridLimits` so SpreadsheetApp doesn't accidentally retain unrelated fields.
  const { maxRows, maxCols } = getWorkbookLoadLimits();
  return { maxRows, maxCols };
})();
const app = new SpreadsheetApp(
  gridRoot,
  { activeCell, selectionRange, activeValue, selectionSum, selectionAverage, selectionCount },
  {
    formulaBar: formulaBarRoot,
    workbookId,
    sheetNameResolver,
    // Legacy renderer uses a smaller grid by default for performance reasons. When users
    // explicitly raise the workbook load limits (via query params/env), align the legacy
    // grid limits so the loaded data is reachable.
    limits: legacyGridLimits,
  },
);

// Expose a small API for Playwright assertions early so e2e can still attach even if
// optional desktop integrations (e.g. Tauri host wiring) fail during startup.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(app as any).getWorkbookSheetStore = () => workbookSheetStore;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__workbookSheetStore = workbookSheetStore;

// Panels persist state keyed by a workbook/document identifier. For file-backed workbooks we use
// their on-disk path; for unsaved sessions we generate a random session id so distinct new
// workbooks don't collide.
let activePanelWorkbookId = workbookId;

function sharedGridZoomStorageKey(): string {
  // Scope zoom persistence by workbook/session id. For file-backed workbooks this can be
  // swapped to a path-based id if/when the desktop shell exposes it.
  return `formula:shared-grid:zoom:${activePanelWorkbookId}`;
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

function applyPersistedSharedGridZoom(options: { resetIfMissing?: boolean } = {}): void {
  if (!app.supportsZoom()) return;
  const persisted = loadPersistedSharedGridZoom();
  if (persisted == null) {
    if (options.resetIfMissing) {
      const current = app.getZoom();
      if (Math.abs(current - 1) > 1e-6) app.setZoom(1);
    }
    return;
  }
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
  const raw = anyApp.limits ?? { maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS, maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS };
  const maxRows =
    Number.isInteger(raw?.maxRows) && raw.maxRows > 0 ? raw.maxRows : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
  const maxCols =
    Number.isInteger(raw?.maxCols) && raw.maxCols > 0 ? raw.maxCols : DEFAULT_DESKTOP_LOAD_MAX_COLS;
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
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
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
  let applied = true;
  try {
    const result = fn(doc, sheetId, ranges);
    if (result === false) applied = false;
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
  if (!applied) {
    try {
      showToast("Formatting could not be applied to the full selection. Try selecting fewer cells/rows.", "warning");
    } catch {
      // `showToast` requires a #toast-root; unit tests don't always include it.
    }
  }
  app.focus();
}

function createHiddenColorInput(): HTMLInputElement {
  const input = document.createElement("input");
  input.type = "color";
  input.tabIndex = -1;
  input.className = "hidden-color-input shell-hidden-input";
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
  // Avoid `addEventListener({ once: true })` here.
  //
  // `<input type="color">` does *not* emit a `change` event when the user cancels the native
  // picker. If we used `addEventListener`, we'd accumulate listeners across cancels and the next
  // successful pick would apply formatting multiple times (multiple history entries).
  input.onchange = () => {
    input.onchange = null;
    const argb = rgbHexToArgb(input.value);
    if (!argb) return;
    applyFormattingToSelection(label, (_doc, sheetId, ranges) => apply(sheetId, ranges, argb));
  };
  input.click();
}

const FONT_SIZE_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

function activeCellFontSizePt(): number {
  const sheetId = app.getCurrentSheetId();
  const cell = app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = app.getDocument() as any;
  const effectiveSize = docAny.getCellFormat?.(sheetId, cell)?.font?.size;
  const state = docAny.getCell?.(sheetId, cell);
  const style = docAny.styleTable?.get?.(state?.styleId ?? 0) ?? {};
  const size = typeof effectiveSize === "number" ? effectiveSize : style.font?.size;
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

function activeCellIndentLevel(): number {
  const sheetId = app.getCurrentSheetId();
  const cell = app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = app.getDocument() as any;
  const raw = docAny.getCellFormat?.(sheetId, cell)?.alignment?.indent;
  const value = typeof raw === "number" ? raw : typeof raw === "string" && raw.trim() !== "" ? Number(raw) : 0;
  return Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0;
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

  const currencyMatch = /[$€£¥]/.exec(section);
  const prefix = currencyMatch?.[0] ?? "";
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
if (collabStatus) installCollabStatusIndicator(app, collabStatus);
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();

app.focus();

// Split-view secondary pane mounts its own CellEditorOverlay. SpreadsheetApp's `isEditing()`
// only reflects its *primary* editor/formula bar/inline edit state, so track the secondary
// editor separately for global UI state (status bar, shortcut gating, etc).
let splitViewSecondaryIsEditing = false;

const isSpreadsheetEditing = (): boolean => app.isEditing() || splitViewSecondaryIsEditing;

function openFormatCells(): void {
  openFormatCellsDialog({
    isEditing: () => isSpreadsheetEditing(),
    getDocument: () => app.getDocument(),
    getSheetId: () => app.getCurrentSheetId(),
    getActiveCell: () => app.getActiveCell(),
    getSelectionRanges: () => app.getSelectionRanges(),
    getGridLimits: () => getGridLimitsForFormatting(),
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

function getSharingWorkbookKeys(): string[] {
  // The SpreadsheetApp constructor only knows about `workbookId`, so store under that
  // key to allow best-effort reconnects after reload.
  //
  // When running under Tauri we also store under the workbook file path (when known)
  // so future wiring (or other subsystems) can key by path without losing metadata.
  const keys = new Set<string>();
  const primary = String(workbookId ?? "").trim();
  if (primary) keys.add(primary);
  const path = typeof activeWorkbook?.path === "string" ? activeWorkbook.path.trim() : "";
  if (path) keys.add(path);
  return Array.from(keys);
}

function loadSharingCollabConnection(): { wsUrl: string; docId: string } | null {
  for (const key of getSharingWorkbookKeys()) {
    const stored = loadCollabConnectionForWorkbook({ workbookKey: key });
    if (stored) return { wsUrl: stored.wsUrl, docId: stored.docId };
  }
  return null;
}

function saveSharingCollabConnection(wsUrl: string, docId: string): void {
  for (const key of getSharingWorkbookKeys()) {
    saveCollabConnectionForWorkbook({ workbookKey: key, wsUrl, docId });
  }
}

function generateCollabDocId(): string {
  const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
  if (typeof randomUuid === "function") {
    try {
      return randomUuid.call((globalThis as any).crypto);
    } catch {
      // Fall through to pseudo-random below.
    }
  }
  return `doc_${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
}

// Lazy-load the clipboard provider to avoid import-order issues in the main entrypoint.
let clipboardProviderPromise: Promise<import("./clipboard/index.js").ClipboardProvider> | null = null;
async function getClipboardProvider(): Promise<import("./clipboard/index.js").ClipboardProvider> {
  if (!clipboardProviderPromise) {
    clipboardProviderPromise = import("./clipboard/index.js").then((mod) => mod.createClipboardProvider());
  }
  return clipboardProviderPromise;
}

async function copyTextToClipboard(text: string): Promise<boolean> {
  const value = String(text ?? "");
  if (!value) return false;

  try {
    // If the environment exposes a Clipboard API (or we're running under Tauri), prefer the shared
    // provider so native clipboard fallbacks work consistently.
    const canUseProvider = Boolean((globalThis as any).__TAURI__) || Boolean(globalThis.navigator?.clipboard);
    if (canUseProvider) {
      const provider = await getClipboardProvider();
      await provider.write({ text: value });
      return true;
    }
  } catch {
    // Fall through to execCommand fallback.
  }

  // Best-effort fallback: execCommand("copy").
  try {
    const textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "true");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    textarea.style.top = "0";
    document.body.appendChild(textarea);
    textarea.select();
    const ok = document.execCommand("copy");
    textarea.remove();
    return ok;
  } catch {
    return false;
  }
}

function baseAppUrlForSharing(): string {
  return `${window.location.origin}${window.location.pathname}`;
}

function getCurrentCollabConnectionFromUrlOrStore(): { wsUrl: string; docId: string } | null {
  // Prefer the current URL (in collab mode, wsUrl/docId are in query params).
  const fromUrl = parseCollabShareLink(window.location.href, { baseUrl: baseAppUrlForSharing() });
  if (fromUrl) return { wsUrl: fromUrl.wsUrl, docId: fromUrl.docId };

  // Fall back to persisted metadata for this workbook.
  return loadSharingCollabConnection();
}

async function promptForWsUrl(): Promise<string | null> {
  const stored = loadSharingCollabConnection();
  const defaultValue = stored?.wsUrl ?? "ws://127.0.0.1:1234";
  const wsUrl = await showInputBox({
    prompt: "Sync server WebSocket URL (ws://…)",
    value: defaultValue,
    placeHolder: "ws://127.0.0.1:1234",
  });
  if (wsUrl == null) return null;
  const trimmed = wsUrl.trim();
  if (!trimmed) return null;
  return trimmed;
}

async function promptForToken(): Promise<string | null> {
  const token = await showInputBox({
    prompt: "Sync token (never share publicly)",
    value: "dev-token",
    placeHolder: "dev-token / JWT",
  });
  if (token == null) return null;
  const trimmed = token.trim();
  if (!trimmed) return null;
  return trimmed;
}

async function copyCurrentCollabLink(): Promise<void> {
  const connection = getCurrentCollabConnectionFromUrlOrStore();
  if (!connection) {
    showToast("Not connected to collaboration yet.", "warning");
    return;
  }

  const token = loadCollabToken({ wsUrl: connection.wsUrl, docId: connection.docId });
  const link = serializeCollabShareLink(
    { wsUrl: connection.wsUrl, docId: connection.docId, ...(token ? { token } : {}) },
    { baseUrl: baseAppUrlForSharing() },
  );

  const copied = await copyTextToClipboard(link);
  if (copied) showToast("Copied collaboration link to clipboard.");
  else {
    showToast("Failed to copy collaboration link. Showing it for manual copy…", "warning", { timeoutMs: 6_000 });
    // Show the link in a modal so the user can manually copy it without printing
    // the token to the console.
    await showInputBox({ prompt: "Collaboration link", value: link });
  }
}

async function startCollaborationAndCopyLink(): Promise<void> {
  const wsUrl = await promptForWsUrl();
  if (!wsUrl) return;
  const token = await promptForToken();
  if (!token) return;

  const docId = generateCollabDocId();

  // Persist non-secret metadata so reopening the workbook can attempt to reconnect.
  saveSharingCollabConnection(wsUrl, docId);
  // Store token only for this session (no localStorage persistence).
  storeCollabToken({ wsUrl, docId, token });

  const link = serializeCollabShareLink({ wsUrl, docId, token }, { baseUrl: baseAppUrlForSharing() });
  const copied = await copyTextToClipboard(link);
  if (copied) showToast("Copied collaboration link. Reloading into collaboration mode…");
  else showToast("Reloading into collaboration mode…");

  // Navigate to the collab link. `SpreadsheetApp` will stash the token in sessionStorage
  // and scrub it from the URL on load.
  window.location.assign(link);
}

async function joinCollaborationFromLinkOrToken(): Promise<void> {
  const input = await showInputBox({
    prompt: "Paste a collaboration link (or a token)",
    value: "",
    placeHolder: "http://…/?collab=1&wsUrl=…&docId=…#token=…",
  });
  if (input == null) return;

  const parsed = parseCollabShareLink(input, { baseUrl: baseAppUrlForSharing() });
  if (parsed) {
    // Persist non-secret connection metadata + store token for this session.
    saveSharingCollabConnection(parsed.wsUrl, parsed.docId);
    if (parsed.token) storeCollabToken({ wsUrl: parsed.wsUrl, docId: parsed.docId, token: parsed.token });
    window.location.assign(
      serializeCollabShareLink(
        { wsUrl: parsed.wsUrl, docId: parsed.docId, ...(parsed.token ? { token: parsed.token } : {}) },
        { baseUrl: baseAppUrlForSharing() },
      ),
    );
    return;
  }

  // Treat the input as a raw token and ask for the remaining fields.
  const token = input.trim();
  if (!token) return;

  const wsUrl = await promptForWsUrl();
  if (!wsUrl) return;

  const docId = await showInputBox({ prompt: "Document ID (docId)", value: "", placeHolder: "my-doc-id" });
  if (docId == null) return;
  const trimmedDocId = docId.trim();
  if (!trimmedDocId) return;

  saveSharingCollabConnection(wsUrl, trimmedDocId);
  storeCollabToken({ wsUrl, docId: trimmedDocId, token });
  window.location.assign(
    serializeCollabShareLink({ wsUrl, docId: trimmedDocId, token }, { baseUrl: baseAppUrlForSharing() }),
  );
}

async function handleShareClick(): Promise<void> {
  const collabSession = app.getCollabSession?.() ?? null;

  const items: Array<{ label: string; value: string; description?: string }> = [];
  if (collabSession) {
    items.push({ label: "Copy collaboration link", value: "copy", description: "Share this document with others" });
  } else {
    items.push({ label: "Start collaboration", value: "start", description: "Create a new shared document" });
  }
  items.push({ label: "Join collaboration link…", value: "join", description: "Open a shared document" });

  const action = await showQuickPick(items, { placeHolder: "Collaboration" });
  if (!action) return;

  if (action === "copy") {
    await copyCurrentCollabLink();
    return;
  }
  if (action === "start") {
    await startCollaborationAndCopyLink();
    return;
  }
  if (action === "join") {
    await joinCollaborationFromLinkOrToken();
    return;
  }
}

const buildTitlebarProps = () => ({
  documentName: computeTitlebarDocumentName(),
  actions: [
    {
      id: "save",
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
      id: "share",
      label: "Share",
      ariaLabel: "Share document",
      variant: "primary" as const,
      onClick: () => {
        void handleShareClick().catch(() => {
          // Do not log tokens (or links that may contain tokens).
          console.error("Share failed");
          showToast("Sharing failed.", "error");
        });
      },
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
  statusMode.textContent = isSpreadsheetEditing() ? "Edit" : "Ready";
}

renderStatusMode();

const unsubscribeTitlebarHistory = app.getDocument().on("history", () => syncTitlebar());
// In collaboration mode, undo/redo state is driven by the Yjs undo manager rather than the
// DocumentController history stack. Ensure the titlebar stays in sync even when the
// DocumentController does not emit history events for every change.
const unsubscribeTitlebarChange = app.getDocument().on("change", () => syncTitlebar());
const unsubscribeTitlebarEditState = app.onEditStateChange(() => {
  renderStatusMode();
  syncTitlebar();
});
window.addEventListener("unload", () => {
  unsubscribeTitlebarHistory();
  unsubscribeTitlebarChange();
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
    const isEditing = isSpreadsheetEditing();
    const perfStats = app.getGridPerfStats() as any;
    const perfStatsSupported = perfStats != null;
    const perfStatsEnabled = Boolean(perfStats?.enabled);

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
      "view.show.performanceStats": perfStatsEnabled,
      "view.window.split": ribbonLayoutController ? ribbonLayoutController.layout.splitView.direction !== "none" : false,
      "data.queriesConnections.queriesConnections":
        ribbonLayoutController != null &&
        getPanelPlacement(ribbonLayoutController.layout, PanelIds.DATA_QUERIES).kind !== "closed",
      "review.comments.showComments": app.isCommentsPanelVisible(),
    };

    const numberFormatLabel = (() => {
      const format = formatState.numberFormat;
      if (format === "mixed") return "Mixed";

      const normalized = typeof format === "string" ? format.trim() : "";
      if (!normalized || normalized.toLowerCase() === "general") return "General";

      const compact = normalized.toLowerCase().replace(/\s+/g, "");

      if (/^[$€£¥]#,##0\\.00$/.test(normalized)) return "Currency";
      if (compact === NUMBER_FORMATS.currency.toLowerCase()) return "Currency";
      if (compact.includes("m/d/yyyy") || compact.includes("yyyy-mm-dd")) return "Date";
      if (compact === NUMBER_FORMATS.date.toLowerCase()) return "Date";
      if (/^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact)) return "Time";
      if (compact.includes("%")) return "Percent";
      if (compact === "#,##0.00") return "Comma";
      if (compact === "0.00") return "Number";
      if (compact.includes("e")) return "Scientific";
      if (compact.includes("/")) return "Fraction";
      if (compact === "@") return "Text";

      return "Custom";
    })();

    const themePreferenceLabel = (() => {
      const preference = themeController.getThemePreference();
      switch (preference) {
        case "system":
          return "System";
        case "light":
          return "Light";
        case "dark":
          return "Dark";
        case "high-contrast":
          return "High Contrast";
        default:
          return "System";
      }
    })();

    const labelById: Record<string, string> = {
      "home.number.numberFormat": numberFormatLabel,
      "view.appearance.theme": `Theme: ${themePreferenceLabel}`,
    };

    // Font dropdown labels should reflect the current selection (Excel-style).
    // Keep defaults ("Font"/"Size") when we don't have an explicit value, but show "Mixed"
    // when the selection contains multiple values.
    if (formatState.fontName === "mixed") {
      labelById["home.font.fontName"] = "Mixed";
    } else if (typeof formatState.fontName === "string") {
      labelById["home.font.fontName"] = formatState.fontName;
    }

    if (formatState.fontSize === "mixed") {
      labelById["home.font.fontSize"] = "Mixed";
    } else if (typeof formatState.fontSize === "number") {
      labelById["home.font.fontSize"] = String(formatState.fontSize);
    }

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
            "home.font.clearFormatting": true,
            "home.alignment.wrapText": true,
            "home.alignment.topAlign": true,
            "home.alignment.middleAlign": true,
            "home.alignment.bottomAlign": true,
            "home.alignment.alignLeft": true,
            "home.alignment.center": true,
            "home.alignment.alignRight": true,
            "home.alignment.orientation": true,
            "home.alignment.increaseIndent": true,
            "home.alignment.decreaseIndent": true,
            "home.number.numberFormat": true,
            "home.number.moreFormats": true,
            "home.number.percent": true,
            "home.number.accounting": true,
            "home.number.date": true,
            "home.number.comma": true,
            "home.number.increaseDecimal": true,
            "home.number.decreaseDecimal": true,
            "home.number.formatCells": true,
          }
        : null),
      // View/zoom controls depend on the current runtime (e.g. shared-grid mode).
      "view.show.performanceStats": !perfStatsSupported,
      "view.zoom.zoom": zoomDisabled,
      "view.zoom.zoom100": zoomDisabled,
      "view.zoom.zoomToSelection": zoomDisabled,
    };

    setRibbonUiState({
      pressedById,
      labelById,
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
scheduleRibbonSelectionFormatStateUpdate();

function isTextInputTarget(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el) return false;
  const tag = el.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
}

function canRunGridFormattingShortcuts(event: KeyboardEvent): boolean {
  if (event.defaultPrevented) return false;
  if (isSpreadsheetEditing()) return false;
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
    if (keyLower === "i") {
      // Only handle Ctrl+I here. We intentionally avoid binding Cmd+I in this legacy
      // formatting handler so the global KeybindingService can own it via builtin
      // keybindings (and we avoid conflicts with host-level shortcuts).
      if (!e.ctrlKey || e.metaKey) return;
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
const disposeCommandPaletteRecentsTracking = (() => {
  try {
    return installCommandPaletteRecentsTracking(commandRegistry, localStorage);
  } catch {
    // In extremely restricted environments (or future Node runtimes with throwing `localStorage` accessors),
    // fall back to no-op tracking.
    return () => {};
  }
})();
window.addEventListener("unload", () => {
  disposeCommandPaletteRecentsTracking();
});

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
// During `DocumentController.applyState` restores, the sheet store is re-ordered to match
// the restored sheet order. Avoid feeding those intermediate store moves back into the
// DocumentController while the restore is still in progress.
let suppressDocReorderFromStore = false;

let sheetStoreDocSync: ReturnType<typeof startSheetStoreDocumentSync> | null = null;
let sheetStoreDocSyncStore: WorkbookSheetStore | null = null;

type SheetStoreSnapshot = {
  order: string[];
  byId: Map<string, { name: string; visibility: SheetVisibility; tabColor?: TabColor }>;
};

let lastSheetStoreSnapshot: SheetStoreSnapshot | null = null;

type SheetUiInfo = { id: string; name: string; visibility?: SheetVisibility; tabColor?: TabColor };

function emitSheetMetadataChanged(): void {
  if (typeof window === "undefined") return;
  try {
    window.dispatchEvent(new CustomEvent("formula:sheet-metadata-changed"));
  } catch {
    // ignore
  }
}

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
  //
  // Note: the desktop has multiple context menu surfaces (grid, sheet tabs, etc). Prefer
  // checking for any visible ContextMenu overlay instead of special-casing a specific test id.
  const openContextMenu = document.querySelector<HTMLElement>(".context-menu-overlay:not([hidden])");
  if (openContextMenu) {
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
    {
      withStoreMutations: (fn) => {
        const prev = syncingSheetUi;
        syncingSheetUi = true;
        try {
          return fn();
        } finally {
          syncingSheetUi = prev;
        }
      },
    },
  );
  sheetStoreDocSyncStore = workbookSheetStore;
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

const collabSheetsKeyRef: CollabSheetsKeyRef = { value: "" };
let collabSheetsSession: CollabSession | null = null;

function syncSheetStoreFromCollabSession(session: CollabSession): void {
  const sheets = listSheetsFromCollabSession(session);
  const key = computeCollabSheetsKey(sheets);
  // Avoid rebuilding the sheet store unless the Yjs sheet list actually changed, *and*
  // the current store instance is already backed by the same CollabSession.
  if (key === collabSheetsKeyRef.value && collabSheetsSession === session && workbookSheetStore instanceof CollabWorkbookSheetStore) {
    return;
  }
  collabSheetsKeyRef.value = key;
  collabSheetsSession = session;

  try {
    workbookSheetStore = new CollabWorkbookSheetStore(session, sheets, collabSheetsKeyRef);
  } catch (err) {
    // If collab sheet names are invalid/duplicated (shouldn't happen, but can if a remote
    // client writes bad metadata), fall back to using the stable sheet id as the display name
    // so the UI remains functional.
    console.error("[formula][desktop] Failed to apply collab sheet metadata:", err);
    workbookSheetStore = new CollabWorkbookSheetStore(
      session,
      sheets.map((sheet) => ({ ...sheet, name: sheet.id })),
      collabSheetsKeyRef,
    );
  }

  // The sheet store instance is replaced whenever collab metadata changes; keep any
  // main.ts listeners (status bar, context keys, etc) subscribed to the latest store.
  installSheetStoreSubscription();
  emitSheetMetadataChanged();
}

function listSheetsForUi(): SheetUiInfo[] {
  const visible = workbookSheetStore.listVisible();
  // Only expose visible sheets to UI affordances like the sheet switcher. Hidden/veryHidden
  // sheets should not be directly activatable via dropdowns.
  return visible.map((s) => ({ id: s.id, name: s.name }));
}

async function handleAddSheet(): Promise<void> {
  if (addSheetInFlight) return;
  addSheetInFlight = true;
  try {
    const activeId = app.getCurrentSheetId();
    const allSheets = workbookSheetStore.listAll();
    const desiredName = generateDefaultSheetName(allSheets);
    const activeIndex = allSheets.findIndex((sheet) => sheet.id === activeId);
    const insertIndex = activeIndex >= 0 ? activeIndex + 1 : allSheets.length;
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
        sheet.set("visibility", "visible");

        // Insert after the active sheet when possible; otherwise append.
        const activeIndex = findCollabSheetIndexById(collabSession, activeId);
        const insertIndex = activeIndex >= 0 ? activeIndex + 1 : collabSession.sheets.length;

        collabSession.sheets.insert(insertIndex, [sheet as any]);
      });

      // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
      doc.getCell(id, { row: 0, col: 0 });
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

      const info = (await invoke("add_sheet", { name: desiredName, after_sheet_id: activeId, index: insertIndex })) as SheetUiInfo;
      const id = String((info as any)?.id ?? "").trim();
      const name = String((info as any)?.name ?? "").trim();
      if (!id) throw new Error("Backend returned empty sheet id");

      // Backend may adjust the name for uniqueness; trust it.
      workbookSheetStore.addAfter(activeId, { id, name: name || desiredName });

      // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
      doc.getCell(id, { row: 0, col: 0 });
      app.activateSheet(id);
      restoreFocusAfterSheetNavigation();
      return;
    }

    // Web-only behavior: create a local DocumentController sheet lazily.
    // (In the desktop-web test harness there is no workbook backend, so we synthesize a stable
    // sheet id locally.)
    const newSheetId = desiredName;
    workbookSheetStore.addAfter(activeId, { id: newSheetId, name: desiredName });
    doc.getCell(newSheetId, { row: 0, col: 0 });
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
      onPersistSheetRename: async (sheetId: string, name: string) => {
        // In collab mode, sheet metadata changes are persisted to the shared Yjs document.
        // Skip the local workbook backend.
        const session = app.getCollabSession?.() ?? null;
        if (session) return;

        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke !== "function") return;

        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we request a rename.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        await invoke("rename_sheet", { sheet_id: sheetId, name });
      },
      onPersistSheetDelete: async (sheetId: string) => {
        // In collab mode, sheet metadata changes are persisted to the shared Yjs document.
        // Skip the local workbook backend.
        const session = app.getCollabSession?.() ?? null;
        if (session) return;

        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke !== "function") return;

        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we request deletion.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        await invoke("delete_sheet", { sheet_id: sheetId });
      },
      onPersistSheetVisibility: async (sheetId: string, visibility: SheetVisibility) => {
        // In collab mode, sheet metadata changes are persisted to the shared Yjs document.
        // Skip the local workbook backend.
        const session = app.getCollabSession?.() ?? null;
        if (session) return;

        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke !== "function") return;

        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we request a visibility change.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        await invoke("set_sheet_visibility", { sheet_id: sheetId, visibility });
      },
      onPersistSheetTabColor: async (sheetId: string, tabColor: TabColor | undefined) => {
        // In collab mode, sheet metadata changes are persisted to the shared Yjs document.
        // Skip the local workbook backend.
        const session = app.getCollabSession?.() ?? null;
        if (session) return;

        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke !== "function") return;

        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we request a tab color update.
        await new Promise<void>((resolve) => queueMicrotask(resolve));

        await invoke("set_sheet_tab_color", { sheet_id: sheetId, tab_color: tabColor ?? null });
      },
      onSheetsReordered: () => restoreFocusAfterSheetNavigation(),
      onSheetRenamed: ({ oldName, newName }) => {
        try {
          rewriteDocumentFormulasForSheetRename(app.getDocument(), oldName, newName);
        } catch (err) {
          showToast(`Failed to update formulas after rename: ${String((err as any)?.message ?? err)}`, "error");
        }
        // Renaming sheets affects workbook metadata but does not produce cell deltas, so
        // mark the document dirty explicitly for unsaved-changes prompts.
        app.getDocument().markDirty();
      },
      onSheetDeleted: ({ sheetId, name }) => {
        const doc = app.getDocument() as any;
        try {
          rewriteDocumentFormulasForSheetDelete(doc, name);
        } catch (err) {
          showToast(`Failed to update formulas after delete: ${String((err as any)?.message ?? err)}`, "error");
        }
      },
      onSheetMoved: async ({ sheetId, toIndex }) => {
        const collabSession = app.getCollabSession?.() ?? null;
        if (collabSession) return;

        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke !== "function") return;

        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

        // Allow any microtask-batched workbook edits to enqueue before we move the sheet.
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await invoke("move_sheet", { sheet_id: sheetId, to_index: toIndex });
      },
      onError: (message: string) => showToast(message, "error"),
    }),
  );
}

function renderSheetPosition(sheets: SheetUiInfo[], activeId: string): void {
  const total = sheets.length;
  if (total === 0) {
    // This should not normally happen (Excel disallows hiding the last visible sheet),
    // but guard so the UI doesn't render an impossible "Sheet 1 of 0" state if the
    // workbook metadata becomes inconsistent (e.g. corrupt/remote data).
    sheetPositionEl.textContent = "Sheet 0 of 0";
    return;
  }
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
  if (!session) {
    // Collab session can be torn down without a full page unload (e.g. workbook close).
    // Ensure we detach any previously registered observers so we don't leak listeners or
    // reference a destroyed Y.Doc.
    if (observedCollabSession && collabSheetsObserver) {
      try {
        observedCollabSession.sheets.unobserveDeep(collabSheetsObserver as any);
      } catch {
        // ignore
      }
    }
    observedCollabSession = null;
    collabSheetsObserver = null;
    collabSheetsKeyRef.value = "";
    collabSheetsSession = null;
    return;
  }
  if (observedCollabSession === session) return;

  if (observedCollabSession && collabSheetsObserver) {
    observedCollabSession.sheets.unobserveDeep(collabSheetsObserver as any);
  }

  observedCollabSession = session;
  collabSheetsObserver = () => {
    // `session.sheets` also stores per-sheet view state + formatting metadata that can change
    // frequently (row/col sizing, range-run formatting, etc). The sheet tab/switcher UI only
    // needs to update when the sheet *list* metadata changes (id/name/order/visibility/tabColor).
    const key = computeCollabSheetsKey(listSheetsFromCollabSession(session));
    if (key === collabSheetsKeyRef.value && collabSheetsSession === session && workbookSheetStore instanceof CollabWorkbookSheetStore) {
      return;
    }
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
  // Reset the store snapshot whenever we (re)install the subscription (workbook open, collab teardown, etc).
  const captureSnapshot = (): SheetStoreSnapshot => {
    const sheets = workbookSheetStore.listAll();
    const byId = new Map<string, { name: string; visibility: SheetVisibility; tabColor?: TabColor }>();
    for (const sheet of sheets) {
      byId.set(sheet.id, { name: sheet.name, visibility: sheet.visibility, tabColor: sheet.tabColor });
    }
    return { order: sheets.map((s) => s.id), byId };
  };

  lastSheetStoreSnapshot = captureSnapshot();
  stopSheetStoreListener = workbookSheetStore.subscribe(() => {
    const prevSnapshot = lastSheetStoreSnapshot;
    const nextSnapshot = captureSnapshot();
    lastSheetStoreSnapshot = nextSnapshot;

    // Keep DocumentController in sync with sheet UI store mutations so sheet-tab operations are
    // undoable via the existing Ctrl+Z/Ctrl+Y stack.
    //
    // Guard against applying these updates during internal sheet UI sync transactions
    // (doc -> store, collab observer updates, etc) so we don't create feedback loops.
    if (!syncingSheetUi && prevSnapshot) {
      const session = app.getCollabSession?.() ?? null;
      if (!session) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const doc: any = app.getDocument();

        const tabColorEqual = (a: TabColor | undefined, b: TabColor | undefined): boolean => {
          if (a === b) return true;
          if (!a || !b) return !a && !b;
          return (
            (a.rgb ?? null) === (b.rgb ?? null) &&
            (a.theme ?? null) === (b.theme ?? null) &&
            (a.indexed ?? null) === (b.indexed ?? null) &&
            (a.tint ?? null) === (b.tint ?? null) &&
            (a.auto ?? null) === (b.auto ?? null)
          );
        };

        const prevOrder = prevSnapshot.order;
        const nextOrder = nextSnapshot.order;
        const prevById = prevSnapshot.byId;
        const nextById = nextSnapshot.byId;

        const added = nextOrder.filter((id) => !prevById.has(id));
        const removed = prevOrder.filter((id) => !nextById.has(id));

        // Renames/deletes can be paired with formula rewrites (sheet-qualified refs).
        // Keep the doc batch open through the end of the current task so any synchronous
        // follow-up edits (e.g. `rewriteDocumentFormulasForSheetRename`) become a single undo step.
        //
        // Note: We intentionally don't do this for reorder/hide/tabColor since those don't
        // require paired document mutations.
        const hasRename = nextOrder.some((sheetId) => {
          if (added.includes(sheetId)) return false;
          if (removed.includes(sheetId)) return false;
          const before = prevById.get(sheetId);
          const after = nextById.get(sheetId);
          return Boolean(before && after && before.name !== after.name);
        });
        const shouldBatchSheetMeta = removed.length > 0 || hasRename;
        if (shouldBatchSheetMeta && typeof doc.beginBatch === "function" && typeof doc.endBatch === "function") {
          const label = removed.length > 0 ? "Delete Sheet" : "Rename Sheet";
          let batchStarted = false;
          try {
            doc.beginBatch({ label });
            batchStarted = true;
          } catch {
            // ignore
          }
          if (batchStarted) {
            queueMicrotask(() => {
              try {
                doc.endBatch();
              } catch {
                // ignore
              }
            });
          }
        }

        // Add sheets (undoable).
        for (const sheetId of added) {
          const meta = nextById.get(sheetId);
          if (!meta) continue;
          const idx = nextOrder.indexOf(sheetId);
          let insertAfterId: string | null = null;
          for (let j = idx - 1; j >= 0; j -= 1) {
            const candidate = nextOrder[j];
            if (candidate && prevById.has(candidate)) {
              insertAfterId = candidate;
              break;
            }
          }
          try {
            doc.addSheet({ sheetId, name: meta.name, insertAfterId });
          } catch {
            // ignore
          }
        }

        // Delete sheets (undoable).
        //
        // Important: if the currently-active UI sheet is being deleted, switch the app to a
        // remaining sheet *before* mutating the DocumentController. SpreadsheetApp reacts to
        // DocumentController change events by reading from the active sheet; if that active
        // sheet id has just been deleted, a read can recreate the sheet (DocumentController
        // sheets are created lazily on access), effectively undoing the delete.
        if (removed.length > 0) {
          const activeId = app.getCurrentSheetId();
          if (activeId && removed.includes(activeId)) {
            const fallback = nextOrder[0] ?? null;
            if (fallback && fallback !== activeId) {
              app.activateSheet(fallback);
              restoreFocusAfterSheetNavigation();
            }
          }
        }
        for (const sheetId of removed) {
          try {
            doc.deleteSheet(sheetId);
          } catch {
            // ignore
          }
        }

        // Reorder sheets (undoable).
        if (!suppressDocReorderFromStore && added.length === 0 && removed.length === 0) {
          const prevKey = prevOrder.join("|");
          const nextKey = nextOrder.join("|");
          if (prevKey && nextKey && prevKey !== nextKey) {
            try {
              doc.reorderSheets(nextOrder, { mergeKey: "sheet-tabs-reorder" });
            } catch {
              // ignore
            }
          }
        }

        // Metadata changes (undoable).
        for (const sheetId of nextOrder) {
          if (added.includes(sheetId)) continue;
          if (removed.includes(sheetId)) continue;
          const before = prevById.get(sheetId);
          const after = nextById.get(sheetId);
          if (!before || !after) continue;

          if (before.name !== after.name) {
            try {
              doc.renameSheet(sheetId, after.name);
            } catch {
              // ignore
            }
          }

          if (before.visibility !== after.visibility) {
            try {
              if (before.visibility === "visible" && after.visibility === "hidden") {
                doc.hideSheet(sheetId);
              } else if (before.visibility === "visible" && after.visibility === "veryHidden") {
                doc.setSheetVisibility(sheetId, after.visibility, { label: "Hide Sheet" });
              } else if (before.visibility !== "visible" && after.visibility === "visible") {
                doc.unhideSheet(sheetId);
              } else {
                doc.setSheetVisibility(sheetId, after.visibility);
              }
            } catch {
              // ignore
            }
          }

          if (!tabColorEqual(before.tabColor, after.tabColor)) {
            try {
              doc.setSheetTabColor(sheetId, after.tabColor);
            } catch {
              // ignore
            }
          }
        }
      }
    }
    const sheets = listSheetsForUi();
    const activeId = app.getCurrentSheetId();

    // If the current active sheet becomes hidden, ensure we actually switch the app to a
    // visible sheet (not just the dropdown value). This keeps the sheet switcher + grid
    // consistent even when sheet visibility changes are initiated outside `syncSheetUi()`.
    if (!sheets.some((sheet) => sheet.id === activeId)) {
      const fallback = sheets[0]?.id ?? null;
      if (fallback) {
        app.activateSheet(fallback);
        restoreFocusAfterSheetNavigation();
        return;
      }
    }

    renderSheetSwitcher(sheets, activeId);
    renderSheetPosition(sheets, activeId);
    emitSheetMetadataChanged();
  });
}

{
  installSheetStoreDocSync();
  installSheetStoreSubscription();
  syncSheetUi();
}

// Excel-like keyboard navigation: Ctrl/Cmd+PgUp/PgDn cycles through sheets.
//
// This must follow the UI sheet store ordering + visibility (WorkbookSheetStore),
// not `DocumentController.getSheetIds()`, because the DocumentController does not
// track the user-visible tab order and can create sheets lazily.
//
// Note: we intentionally handle this in the capture phase so we can prevent
// downstream handlers (e.g. browser defaults / other listeners) from consuming
// the shortcut before we can switch sheets.
window.addEventListener(
  "keydown",
  (e) => {
    if (e.defaultPrevented) return;
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return;
    if (e.shiftKey || e.altKey) return;
    if (e.key !== "PageUp" && e.key !== "PageDown") return;

    const target = e.target as EventTarget | null;
    if (target instanceof HTMLElement) {
      const tabList = target.closest?.("#sheet-tabs .sheet-tabs");
      if (tabList) {
        // Let the sheet tab strip handle shortcuts when focus is on a tab.
        return;
      }

      // Never steal the shortcut from text inputs / contenteditable surfaces.
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) {
        return;
      }
    }

    // Ctrl/Cmd+PgUp/PgDn should not switch sheets while editing (cell editor, formula bar, inline edit,
    // or split view secondary editing). Let the focused editor/input surface handle the event.
    if (isSpreadsheetEditing()) return;

    const ordered = workbookSheetStore.listVisible().map((sheet) => sheet.id);
    if (ordered.length === 0) return;

    const current = app.getCurrentSheetId();
    const idx = ordered.indexOf(current);
    if (idx === -1) {
      // Current sheet is no longer visible (should be rare; typically we auto-fallback
      // elsewhere). Treat Ctrl/Cmd+PgUp/PgDn as a "jump to first visible sheet".
      const first = ordered[0];
      if (!first) return;
      e.preventDefault();
      e.stopPropagation();
      app.activateSheet(first);
      restoreFocusAfterSheetNavigation();
      return;
    }

    e.preventDefault();
    e.stopPropagation();

    const commandId = e.key === "PageUp" ? "workbook.previousSheet" : "workbook.nextSheet";
    // Prefer the command registry so the shortcut shares the same behavior as other
    // keybinding surfaces (focus restore hook, analytics, etc). Fall back to a direct
    // sheet switch if commands haven't been registered yet (early startup).
    void commandRegistry.executeCommand(commandId).catch(() => {
      const delta = e.key === "PageUp" ? -1 : 1;
      const next = ordered[(idx + delta + ordered.length) % ordered.length];
      if (!next || next === current) return;
      app.activateSheet(next);
      restoreFocusAfterSheetNavigation();
    });
  },
  { capture: true },
);

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
app.getDocument().on("change", (payload: any) => {
  if (payload?.source === "applyState") {
    suppressDocReorderFromStore = true;
    queueMicrotask(() => {
      suppressDocReorderFromStore = false;
    });
  }
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
  gridSecondaryEl.addEventListener("keydown", (e) => {
    if (e.defaultPrevented) return;

    // Match SpreadsheetApp guards: never steal shortcuts from active text editing.
    if (isTextInputTarget(e.target)) return;
    if (isSpreadsheetEditing()) return;

    // Excel semantics: Shift+F2 opens the comments panel.
    if (e.key === "F2" && e.shiftKey) {
      e.preventDefault();
      app.openCommentsPanel();
      return;
    }

    // Cancel an in-progress fill handle drag (matches primary-grid Escape semantics).
    if (e.key === "Escape") {
      if (secondaryGridView?.grid.cancelFillHandleDrag()) {
        e.preventDefault();
        return;
      }
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

  // Shared-grid zoom is persisted separately (see `sharedGridZoomStorageKey`) so it can
  // survive quick reloads even if layout persistence (which is debounced) hasn't flushed
  // yet. When booting we need to hydrate the layout's primary-pane zoom from the persisted
  // value; otherwise `renderLayout()` would restore the stale layout zoom (typically 1)
  // and clobber the user's setting.
  if (app.supportsZoom()) {
    const persistedZoom = loadPersistedSharedGridZoom();
    if (persistedZoom != null) {
      const currentLayoutZoom = layoutController.layout?.splitView?.panes?.primary?.zoom;
      if (
        typeof currentLayoutZoom !== "number" ||
        !Number.isFinite(currentLayoutZoom) ||
        Math.abs(currentLayoutZoom - persistedZoom) > 1e-6
      ) {
        layoutController.setSplitPaneZoom("primary", persistedZoom, { persist: false });
      }
    }
  }

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
    sheetNameResolver,
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
  activeMacroRecorder = macroRecorder;

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

  let splitPanePersistTimer: number | null = null;
  let splitPanePersistDirty = false;

  const syncSecondaryGridReferenceHighlights = () => {
    if (!secondaryGridView) return;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const formulaBar = (app as any).formulaBar as any;
    const model = formulaBar?.model as any;
    const highlights = typeof model?.referenceHighlights === "function" ? model.referenceHighlights() : null;

    if (!Array.isArray(highlights) || highlights.length === 0) {
      secondaryGridView.grid.renderer.setReferenceHighlights(null);
      return;
    }

    const sheetId = app.getCurrentSheetId();

    // Split-view secondary pane always uses a shared-grid renderer with a 1x1 frozen
    // header row/col (row/column labels), even when the primary pane is in legacy mode.
    const headerRows = 1;
    const headerCols = 1;

    const gridHighlights = highlights
      .filter((h: any) => {
        const sheet = h?.range?.sheet;
        if (!sheet) return true;
        const resolved = app.getSheetIdByName(String(sheet));
        if (!resolved) return false;
        return resolved.toLowerCase() === sheetId.toLowerCase();
      })
      .map((h: any) => {
        const range = h.range;
        const startRow = Math.min(range.startRow, range.endRow);
        const endRow = Math.max(range.startRow, range.endRow);
        const startCol = Math.min(range.startCol, range.endCol);
        const endCol = Math.max(range.startCol, range.endCol);

        const gridRange: GridCellRange = {
          startRow: startRow + headerRows,
          endRow: endRow + headerRows + 1,
          startCol: startCol + headerCols,
          endCol: endCol + headerCols + 1,
        };

        return { range: gridRange, color: String(h.color), active: Boolean(h.active) };
      });

    secondaryGridView.grid.renderer.setReferenceHighlights(gridHighlights.length > 0 ? gridHighlights : null);
  };

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

    syncSecondaryGridReferenceHighlights();
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
    formulaBar.textarea.addEventListener("keyup", () => syncSecondaryGridInteractionMode());
    formulaBar.textarea.addEventListener("select", () => syncSecondaryGridInteractionMode());
    // `blur` fires before FormulaBarView updates its model state during commit/cancel; sync on a microtask.
    formulaBar.textarea.addEventListener("blur", () => syncSecondaryGridInteractionModeSoon());
  }
  if (formulaBar?.root instanceof HTMLElement) {
    // Clicking fx/commit/cancel can mutate the draft without triggering a textarea input event.
    // Sync on a microtask so we observe the final FormulaBarView state.
    formulaBar.root.addEventListener("click", () => syncSecondaryGridInteractionModeSoon());
  }

  // Programmatic formula bar commit/cancel (e.g. File → Save calling `commitPendingEditsForCommand()`)
  // may not produce textarea blur/input events. Subscribe to SpreadsheetApp edit-state changes so we
  // always leave split-view range-selection mode when formula editing ends.
  const unsubscribeSplitViewEditStateSync = app.onEditStateChange(() => syncSecondaryGridInteractionModeSoon());
  window.addEventListener("unload", () => unsubscribeSplitViewEditStateSync());

  // Range selection insertion updates the formula bar draft + reference highlights via
  // `FormulaBarView.begin/updateRangeSelection()`, which are programmatic (no textarea `input`
  // events). When split view is active, mirror those changing highlights onto the secondary pane
  // so the visual outline stays live during drags initiated from *either* pane (legacy or shared).
  if (formulaBar && !(formulaBar as any).__formulaSplitViewRangeSelectionHighlightSyncPatched) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (formulaBar as any).__formulaSplitViewRangeSelectionHighlightSyncPatched = true;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyBar = formulaBar as any;
    const wrap = (methodName: "beginRangeSelection" | "updateRangeSelection" | "endRangeSelection") => {
      const original = anyBar?.[methodName];
      if (typeof original !== "function") return;
      anyBar[methodName] = (...args: any[]) => {
        const result = original.apply(anyBar, args);
        syncSecondaryGridReferenceHighlights();
        return result;
      };
    };

    wrap("beginRangeSelection");
    wrap("updateRangeSelection");
    wrap("endRangeSelection");
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
    // SecondaryGridView debounces persistence of scroll/zoom into the layout. Flush
    // any pending values first so the final viewport state is not lost on reload.
    secondaryGridView?.flushPersistence();
    if (splitPanePersistTimer != null) {
      window.clearTimeout(splitPanePersistTimer);
      splitPanePersistTimer = null;
    }
    splitPanePersistDirty = false;
    layoutController.persistNow();
  };

  // Always flush any in-memory (persist:false) layout updates on unload.
  //
  // Split view interactions update layout at high frequency and rely on debounced
  // persistence. Persisting here ensures we don't lose the user's final scroll/zoom
  // if they close/reload immediately after interacting.
  window.addEventListener("beforeunload", persistLayoutNow);

  const persistPrimaryZoomFromApp = () => {
    const pane = layoutController.layout.splitView.panes.primary;
    const zoom = app.getZoom();
    if (pane.zoom === zoom) return;
    layoutController.setSplitPaneZoom("primary", zoom, { persist: false, emit: false });
    scheduleSplitPanePersist();
  };

  window.addEventListener("formula:zoom-changed", persistPrimaryZoomFromApp);

  const invalidateSecondaryProvider = () => {
    if (!secondaryGridView) return;
    // Sheet view state (frozen panes + axis overrides) lives in the DocumentController and is
    // independent of cell contents. Even when we reuse the primary grid's provider, we still
    // need to re-apply the current sheet's view state (e.g. when switching sheets).
    secondaryGridView.syncSheetViewFromDocument();
    // Reference highlights are derived from the formula bar draft; recompute when the active sheet
    // changes so the secondary pane matches the primary pane behavior (only show highlights that
    // belong to the currently-visible sheet).
    syncSecondaryGridReferenceHighlights();
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
  };

  const ensurePrimarySplitPanePersistence = () => {
    if (stopPrimaryScrollSubscription || stopPrimaryZoomSubscription) return;

    stopPrimaryScrollSubscription = app.subscribeScroll((scroll) => {
      if (layoutController.layout.splitView.direction === "none") return;

      const pane = layoutController.layout.splitView.panes.primary;
      if (pane.scrollX === scroll.x && pane.scrollY === scroll.y) return;

      layoutController.setSplitPaneScroll("primary", { scrollX: scroll.x, scrollY: scroll.y }, { persist: false, emit: false });
      scheduleSplitPanePersist();
    });

    stopPrimaryZoomSubscription = app.subscribeZoom((zoom) => {
      if (layoutController.layout.splitView.direction === "none") return;

      const pane = layoutController.layout.splitView.panes.primary;
      if (pane.zoom === zoom) return;

      layoutController.setSplitPaneZoom("primary", zoom, { persist: false, emit: false });
      scheduleSplitPanePersist();
    });

  };

  const restorePrimarySplitPaneViewport = () => {
    const pane = layoutController.layout.splitView.panes.primary;
    if (app.supportsZoom()) {
      app.setZoom(pane.zoom ?? 1);
    }
    app.setScroll(pane.scrollX ?? 0, pane.scrollY ?? 0);
  };

  let lastSplitPrimarySizeCss: string | null = null;
  let lastSplitSecondarySizeCss: string | null = null;
  const applySplitRatioCss = (ratio: number): void => {
    const clamped = Math.max(0.1, Math.min(0.9, ratio));
    const primaryPct = Math.round(clamped * 1000) / 10;
    const secondaryPct = Math.round((100 - primaryPct) * 10) / 10;
    const primaryCss = `${primaryPct}%`;
    const secondaryCss = `${secondaryPct}%`;
    // Updating CSS vars can trigger grid layout work; skip redundant writes when rounding
    // yields the same percentages (e.g. during tiny pointer moves).
    if (primaryCss !== lastSplitPrimarySizeCss) {
      lastSplitPrimarySizeCss = primaryCss;
      gridSplitEl.style.setProperty("--split-primary-size", primaryCss);
    }
    if (secondaryCss !== lastSplitSecondarySizeCss) {
      lastSplitSecondarySizeCss = secondaryCss;
      gridSplitEl.style.setProperty("--split-secondary-size", secondaryCss);
    }
  };

  function renderSplitView() {
    const split = layoutController.layout.splitView;
    const ratio = typeof split.ratio === "number" ? split.ratio : 0.5;
    if (gridSplitEl.dataset.splitDirection !== split.direction) {
      gridSplitEl.dataset.splitDirection = split.direction;
    }
    applySplitRatioCss(ratio);

    if (split.direction === "none") {
      secondaryGridView?.destroy();
      secondaryGridView = null;
      if (splitViewSecondaryIsEditing) {
        splitViewSecondaryIsEditing = false;
        renderStatusMode();
        syncTitlebar();
        scheduleRibbonSelectionFormatStateUpdate();
      }
      stopPrimarySplitPanePersistence();
      primaryPaneViewportRestored = false;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__formulaSecondaryGrid = null;
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
      // Split view is two panes over the *same* active sheet (Excel-style). Keep the
      // secondary pane sheet in lockstep with SpreadsheetApp's current sheet so:
      // - selection sync stays correct
      // - in-place edits / fill commits in the secondary pane apply to the visible sheet
      // Note: Layout state may persist a `sheetId` per pane for future multi-sheet UX, but
      // the current desktop UI sheet tabs always drive `app.getCurrentSheetId()`.
      const getSecondarySheetId = () => app.getCurrentSheetId();

      // Use the same DocumentController / computed value cache as the primary grid so
      // the secondary pane stays live with edits and formula recalculation.
      const anyApp = app as any;
      const limits = anyApp.limits ?? { maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS, maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS };
      const rowCount = Number.isInteger(limits.maxRows) ? limits.maxRows + 1 : DEFAULT_DESKTOP_LOAD_MAX_ROWS + 1;
      const colCount = Number.isInteger(limits.maxCols) ? limits.maxCols + 1 : DEFAULT_DESKTOP_LOAD_MAX_COLS + 1;

      secondaryGridView = new SecondaryGridView({
        container: gridSecondaryEl,
        provider: (app as any).sharedProvider ?? undefined,
        document: app.getDocument(),
        getSheetId: getSecondarySheetId,
        rowCount,
        colCount,
        showFormulas: () => app.getShowFormulas(),
        getComputedValue: (cell) => app.getCellComputedValueForSheet(getSecondarySheetId(), cell),
        onRequestRefresh: () => app.refresh(),
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
          layoutController.setSplitPaneScroll("secondary", scroll, { persist: false, emit: false });
          scheduleSplitPanePersist();
        },
        persistZoom: (zoom) => {
          const pane = layoutController.layout.splitView.panes.secondary;
          if (pane.zoom === zoom) return;
          layoutController.setSplitPaneZoom("secondary", zoom, { persist: false, emit: false });
          scheduleSplitPanePersist();
        },
        onEditStateChange: (isEditing) => {
          if (splitViewSecondaryIsEditing === isEditing) return;
          splitViewSecondaryIsEditing = isEditing;
          renderStatusMode();
          syncTitlebar();
          scheduleRibbonSelectionFormatStateUpdate();
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
 
    // Expose for Playwright (secondary pane autofill).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__formulaSecondaryGrid = secondaryGridView.grid;
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
  type GridArea = "cell" | "rowHeader" | "colHeader" | "corner";
  let currentGridArea: GridArea = "cell";

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
      gridArea: currentGridArea,
      isRowHeader: currentGridArea === "rowHeader",
      isColHeader: currentGridArea === "colHeader",
      isCorner: currentGridArea === "corner",
    });
  };

  app.subscribeSelection((selection) => {
    lastSelection = selection;
    updateContextKeys(selection);
  });
  app.getDocument().on("change", () => updateContextKeys());
  window.addEventListener("formula:comments-panel-visibility-changed", () => updateContextKeys());
  window.addEventListener("formula:comments-changed", () => updateContextKeys());
  window.addEventListener("formula:sheet-metadata-changed", () => updateContextKeys());

  type ExtensionSelectionChangedEvent = {
    sheetId: string;
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
  // Track whether the selection has *ever* changed after the initial subscription.
  //
  // Some extension-host clipboard DLP enforcement is keyed off the user changing the selection
  // (which is also when the host will broadcast `selectionChanged` events and therefore expose
  // cell data to extensions). We intentionally treat the initial selection as "not yet exposed"
  // so extensions that haven't read any cells (and haven't observed any selectionChanged events)
  // can still write to the clipboard.
  let extensionClipboardSelectionChanged = false;

  const recordLastExtensionSelection = (
    _sheetId: string,
    _range: { startRow: number; startCol: number; endRow: number; endCol: number },
  ): void => {
    void _sheetId;
    void _range;
  };

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
      extensionClipboardSelectionChanged = true;

      const values: Array<Array<string | number | boolean | null>> = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const row: Array<string | number | boolean | null> = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          const cell = app.getDocument().getCell(rect.sheetId, { row: r, col: c }) as any;
          row.push(normalizeExtensionCellValue(cell?.value ?? null));
        }
        values.push(row);
      }

      const payload: ExtensionSelectionChangedEvent = { sheetId: rect.sheetId, selection: { ...range, address, values } };
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

  // Extensions can access spreadsheet data via `formula.cells.*` and `formula.events.*` and then write
  // arbitrary text to the system clipboard via `formula.clipboard.writeText()`. SpreadsheetApp's
  // copy/cut handlers already enforce clipboard-copy DLP, but extensions would otherwise bypass it.
  //
  // The BrowserExtensionHost runtime tracks which ranges have been exposed to each extension
  // (via `cells.getCell/getRange/getSelection` and event payloads) and forwards them through the
  // `clipboardWriteGuard` hook so the desktop adapter can apply the same DLP policy.
  const extensionClipboardDlp = createDesktopDlpContext({ documentId: workbookId });

  const normalizeSelectionRange = (range: { startRow: number; startCol: number; endRow: number; endCol: number }) => {
    const startRow = Math.min(range.startRow, range.endRow);
    const endRow = Math.max(range.startRow, range.endRow);
    const startCol = Math.min(range.startCol, range.endCol);
    const endCol = Math.max(range.startCol, range.endCol);
    return { startRow, startCol, endRow, endCol };
  };

  const enforceExtensionClipboardDlpForRange = (params: {
    sheetId: string;
    range: { startRow: number; startCol: number; endRow: number; endCol: number };
  }) => {
    enforceClipboardCopy({
      documentId: extensionClipboardDlp.documentId,
      sheetId: params.sheetId,
      range: {
        start: { row: params.range.startRow, col: params.range.startCol },
        end: { row: params.range.endRow, col: params.range.endCol },
      },
      classificationStore: extensionClipboardDlp.classificationStore,
      policy: extensionClipboardDlp.policy,
    });
  };

  const extensionClipboardWriteGuard = async (params: { extensionId: string; taintedRanges: any[] }) => {
    try {
      // Conservatively apply selection-based clipboard DLP only after the user has changed the
      // selection at least once. This matches the moment when the host will have broadcast at
      // least one `selectionChanged` event (and therefore could have exposed selection values
      // to extension workers, which can observe raw message events).
      if (extensionClipboardSelectionChanged) {
        const selection = getClipboardDlpSelection();
        enforceExtensionClipboardDlpForRange({ sheetId: selection.sheetId, range: selection.range });
      }

      const taintedRanges = Array.isArray(params.taintedRanges) ? params.taintedRanges : [];
      for (const raw of taintedRanges) {
        if (!raw || typeof raw !== "object") continue;
        const sheetId = typeof (raw as any).sheetId === "string" ? String((raw as any).sheetId) : "";
        if (!sheetId) continue;
        const startRow = Number((raw as any).startRow);
        const startCol = Number((raw as any).startCol);
        const endRow = Number((raw as any).endRow);
        const endCol = Number((raw as any).endCol);
        if (![startRow, startCol, endRow, endCol].every((v) => Number.isFinite(v))) continue;

        enforceExtensionClipboardDlpForRange({
          sheetId,
          range: normalizeSelectionRange({
            startRow: Math.trunc(startRow),
            startCol: Math.trunc(startCol),
            endRow: Math.trunc(endRow),
            endCol: Math.trunc(endCol),
          }),
        });
      }
    } catch (err) {
      const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
      if (isDlpViolation) {
        try {
          const message =
            typeof (err as any)?.message === "string" && (err as any).message.trim()
              ? String((err as any).message)
              : "Clipboard copy blocked by data loss prevention policy.";
          showToast(message, "error");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
      }
      throw err;
    }
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
    onWorkbookOpened(handler: (event: ExtensionWorkbookLifecycleEvent) => void) {
      if (typeof handler !== "function") return () => {};
      workbookOpenedEventListeners.add(handler);
      return () => workbookOpenedEventListeners.delete(handler);
    },
    onBeforeSave(handler: (event: ExtensionWorkbookLifecycleEvent) => void) {
      if (typeof handler !== "function") return () => {};
      beforeSaveEventListeners.add(handler);
      return () => beforeSaveEventListeners.delete(handler);
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
          sheet.set("visibility", "visible");

          const activeIdx = findCollabSheetIndexById(collabSession, activeId);
          const insertIndex = activeIdx >= 0 ? activeIdx + 1 : collabSession.sheets.length;
          collabSession.sheets.insert(insertIndex, [sheet as any]);
        });

        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
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

        const sheets = workbookSheetStore.listAll();
        const activeIndex = sheets.findIndex((sheet) => sheet.id === activeId);
        const insertIndex = activeIndex >= 0 ? activeIndex + 1 : sheets.length;

        const info = (await invoke("add_sheet", { name: validatedName, after_sheet_id: activeId, index: insertIndex })) as SheetUiInfo;
        const id = String((info as any)?.id ?? "").trim();
        const resolvedName = String((info as any)?.name ?? "").trim();
        if (!id) throw new Error("Backend returned empty sheet id");

        // Backend may adjust the name for uniqueness; trust it.
        workbookSheetStore.addAfter(activeId, { id, name: resolvedName || validatedName });
        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
        app.activateSheet(id);
        restoreFocusAfterSheetNavigation();
        const storedName = workbookSheetStore.getName(id);
        return { id, name: storedName ?? (resolvedName || validatedName || id) };
      }

      // Web-only behavior: create a local DocumentController sheet lazily.
      // (In the desktop-web test harness there is no workbook backend, so we synthesize a stable
      // sheet id locally.)
      const newSheetId = validatedName;
      workbookSheetStore.addAfter(activeId, { id: newSheetId, name: validatedName });
      doc.getCell(newSheetId, { row: 0, col: 0 });
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
        doc.setCellInputs(rewrittenInputs, { label: "Rename sheet", source: "extension" });
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
      const deletedName = workbookSheetStore.getName(sheetId) ?? sheetId;

      // In collab mode, sheet operations are persisted through the shared Yjs document.
      // Avoid mutating the local Tauri workbook backend (if present).
      const collabSession = app.getCollabSession?.() ?? null;
      if (!collabSession) {
        const baseInvoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
        if (typeof baseInvoke === "function") {
          // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
          const invoke = queuedInvoke ?? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args)));

          // Allow any microtask-batched workbook edits to enqueue before we request deletion.
          await new Promise<void>((resolve) => queueMicrotask(resolve));
          await invoke("delete_sheet", { sheet_id: sheetId });
        }
      }

      // Update sheet metadata to enforce workbook invariants (e.g. last-sheet guard) and drive UI
      // reconciliation. In desktop mode we perform the backend delete first so a failure does not
      // leave the frontend store out of sync with the host workbook.
      workbookSheetStore.remove(sheetId);
      if (wasActive) {
        const next =
          workbookSheetStore.listVisible().at(0)?.id ??
          workbookSheetStore.listAll().at(0)?.id ??
          app.getCurrentSheetId();
        if (next && next !== sheetId) {
          app.activateSheet(next);
        }
      }

      try {
        rewriteDocumentFormulasForSheetDelete(doc as any, deletedName);
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
      assertExtensionRangeWithinLimits(range, { label: "Selection" });
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
      app.getDocument().setCellValue(sheetId, { row, col }, value, { source: "extension" });
    },
    async getRange(ref: string) {
      const { sheetId, startRow, startCol, endRow, endCol } = parseSheetQualifiedRange(ref);
      assertExtensionRangeWithinLimits({ startRow, startCol, endRow, endCol });
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
      assertExtensionRangeWithinLimits({ startRow, startCol, endRow, endCol });
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

      app.getDocument().setCellInputs(inputs, { label: "Extension setRange", source: "extension" });
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
      // When saving a workbook that already has a file path, the BrowserExtensionHost
      // emits `beforeSave` directly for the calling extension API request.
      //
      // When the workbook has no path, `handleSave()` will prompt for a Save As target.
      // In that scenario we let the desktop save flow notify extensions once a path is
      // chosen (via `handleSaveAsPath()` emitting the `beforeSave` event) so:
      //   - cancelling the dialog does not emit `beforeSave`
      //   - the `beforeSave` event includes the final path selected by the user
      const notifyExtensions = !activeWorkbook?.path;
      await handleSave({ notifyExtensions, throwOnCancel: true });
    };

    extensionSpreadsheetApi.saveWorkbookAs = async (path: string) => {
      await handleSaveAsPath(String(path), { notifyExtensions: false });
    };

    extensionSpreadsheetApi.closeWorkbook = async () => {
      // Model closing the current workbook (Excel-like) as swapping to a fresh blank workbook.
      // Use a close-specific prompt/cancel message to keep UX consistent for extensions.
      await handleNewWorkbook({
        notifyExtensions: false,
        throwOnCancel: true,
        actionLabel: "close this workbook",
        cancelMessage: "Close workbook cancelled",
      });
    };
  }

  const clipboardProviderPromise = createClipboardProvider();

  const extensionHostManager = new DesktopExtensionHostManager({
    engineVersion: "1.0.0",
    spreadsheetApi: extensionSpreadsheetApi,
    clipboardWriteGuard: extensionClipboardWriteGuard,
    clipboardApi: {
      readText: async () => {
        const provider = await clipboardProviderPromise;
        const { text } = await provider.read();
        return text ?? "";
      },
      writeText: async (text: string) => {
        // Extensions can write arbitrary text to the system clipboard via `formula.clipboard.writeText()`.
        // Enforce clipboard-copy DLP against the current UI selection (active-cell fallback) so
        // extensions cannot bypass SpreadsheetApp's `clipboard.copy` policy enforcement.
        try {
          const sheetId = app.getCurrentSheetId();
          const active = app.getActiveCell();
          const selection = app.getSelectionRanges()[0];
          const range = normalizeSelectionRange(
            selection ?? { startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col },
          );
          enforceExtensionClipboardDlpForRange({ sheetId, range });
        } catch (err) {
          const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
          if (isDlpViolation) {
            const message =
              typeof (err as any)?.message === "string" && (err as any).message.trim()
                ? String((err as any).message)
                : "Clipboard copy blocked by data loss prevention policy.";
            try {
              showToast(message, "error");
            } catch {
              // `showToast` requires a #toast-root; unit tests don't always include it.
            }
            // Convert to a generic Error so the extension command fails without relying
            // on the host-side DlpViolationError type.
            throw new Error(message);
          }
          throw err;
        }
        const provider = await clipboardProviderPromise;
        await provider.write({ text: String(text ?? "") });
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, patch, { label: "Format Cells" });
            if (ok === false) applied = false;
          }
          return applied;
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
      extensionsLoadPromise = (async () => {
        await extensionHostManager.loadBuiltInExtensions();
      })();
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

  // Keybindings: central dispatch with built-in precedence over extensions.
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
      try {
        const placement = getPanelPlacement(layoutController.layout, id);
        if (placement.kind !== "closed") {
          layoutController.closePanel(id);
        }
      } catch {
        // ignore
      }
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
  const commandKeybindingDisplayIndex = keybindingService.getCommandKeybindingDisplayIndex();
  // Bubble-phase listener so SpreadsheetApp can `preventDefault()` first.
  keybindingService.installWindowListener(window, { capture: false });

  const updateKeybindings = () => {
    const contributed =
      extensionHostManager.ready && !extensionHostManager.error
        ? (extensionHostManager.getContributedKeybindings() as ContributedKeybinding[])
        : [];
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

  // Marketplace installs (WebExtensionManager) can load/unload extensions directly into the shared
  // BrowserExtensionHost. When that happens we need to resync contributed commands/panels/keybindings
  // so the desktop UI surfaces the new contributions without requiring a reload.
  window.addEventListener("formula:extensions-changed", () => {
    // Preserve the desktop "lazy-load extensions" behavior: do not start the extension host
    // just because something changed in IndexedDB.
    //
    // If extensions are already loading (Extensions panel opened / command executed) or the host
    // is already ready, then sync any newly-installed extensions into the runtime and refresh
    // contributions.
    if (!extensionHostManager.ready && !extensionsLoadPromise) return;

    void ensureExtensionsLoaded()
      .then(async () => {
        try {
          // Pick up any installs that happened after the initial `loadAllInstalled()` pass.
          await extensionHostManager.getMarketplaceExtensionManager().loadAllInstalled();
        } catch {
          // ignore
        }
      })
      .then(() => {
        updateKeybindings();
        syncContributedCommands();
        syncContributedPanels();
        activateOpenExtensionPanels();
      })
      .catch(() => {
        // ignore
      });
  });

  const contextMenu = new ContextMenu({
    onClose: () => {
      // Reset the "where was the context menu opened" context keys when the menu closes so
      // keybindings / when-clauses don't keep matching header-specific items.
      currentGridArea = "cell";
      updateContextKeys();

      // Best-effort: restore focus to the appropriate editing surface after closing.
      // (Typically the grid, but preserve formula-bar focus during range-selection workflows.)
      app.focusAfterSheetNavigation();
    },
  });
  sharedContextMenu = contextMenu;

  const executeBuiltinCommand = (commandId: string, ...args: any[]) => {
    void commandRegistry.executeCommand(commandId, ...args).catch((err) => {
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    });
  };

  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.platform);
  const primaryShortcut = (key: string) => (isMac ? `⌘${key}` : `Ctrl+${key}`);
  const primaryShiftShortcut = (key: string) => (isMac ? `⌘⇧${key}` : `Ctrl+Shift+${key}`);

  type GridHitTest = { area: GridArea; row: number | null; col: number | null };

  const gridAreaForSelection = (selection: SelectionState | null = lastSelection): GridArea => {
    const type = selection?.type;
    if (type === "row") return "rowHeader";
    if (type === "column") return "colHeader";
    if (type === "all") return "corner";
    return "cell";
  };

  const hitTestGridAreaAtClientPoint = (clientX: number, clientY: number): GridHitTest => {
    const anyApp = app as any;
    const rootRect = gridRoot.getBoundingClientRect();
    const x = clientX - rootRect.left;
    const y = clientY - rootRect.top;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return { area: "cell", row: null, col: null };

    // Shared-grid mode: row/col headers are rendered as real cells (headerRows/headerCols = 1),
    // so rely on the renderer's hit-test instead of legacy pixel thresholds.
    const sharedGrid = (anyApp.sharedGrid ?? null) as
      | { renderer?: { pickCellAt?: (x: number, y: number) => { row: number; col: number } | null } }
      | null;
    const selectionCanvas = (anyApp.selectionCanvas ?? null) as HTMLElement | null;
    if (sharedGrid?.renderer && typeof sharedGrid.renderer.pickCellAt === "function" && selectionCanvas) {
      const canvasRect = selectionCanvas.getBoundingClientRect();
      const vx = clientX - canvasRect.left;
      const vy = clientY - canvasRect.top;
      if (Number.isFinite(vx) && Number.isFinite(vy) && vx >= 0 && vy >= 0 && vx <= canvasRect.width && vy <= canvasRect.height) {
        const picked = sharedGrid.renderer.pickCellAt(vx, vy);
        if (picked) {
          const headerRows = 1;
          const headerCols = 1;
          if (picked.row < headerRows && picked.col < headerCols) return { area: "corner", row: null, col: null };
          if (picked.col < headerCols) return { area: "rowHeader", row: picked.row - headerRows, col: null };
          if (picked.row < headerRows) return { area: "colHeader", row: null, col: picked.col - headerCols };
          return { area: "cell", row: picked.row - headerRows, col: picked.col - headerCols };
        }
      }
    }

    const rowHeaderWidth = Number.isFinite(anyApp.rowHeaderWidth) ? Number(anyApp.rowHeaderWidth) : 0;
    const colHeaderHeight = Number.isFinite(anyApp.colHeaderHeight) ? Number(anyApp.colHeaderHeight) : 0;

    const inRowHeader = x < rowHeaderWidth;
    const inColHeader = y < colHeaderHeight;
    if (inRowHeader && inColHeader) return { area: "corner", row: null, col: null };

    if (inRowHeader || inColHeader) {
      let cell: { row: number; col: number } | null = null;
      if (typeof anyApp.cellFromPoint === "function") {
        try {
          cell = anyApp.cellFromPoint(x, y) as { row: number; col: number } | null;
        } catch {
          cell = null;
        }
      }
      if (!cell) {
        const picked = app.pickCellAtClientPoint(clientX, clientY);
        if (picked) cell = { row: picked.row, col: picked.col };
      }
      if (inRowHeader) return { area: "rowHeader", row: cell?.row ?? null, col: null };
      return { area: "colHeader", row: null, col: cell?.col ?? null };
    }

    // Task 40 hit-test: if we're over a sheet cell, treat this as a cell-area click.
    const pickedCell = app.pickCellAtClientPoint(clientX, clientY);
    if (pickedCell) return { area: "cell", row: pickedCell.row, col: pickedCell.col };

    return { area: "cell", row: null, col: null };
  };

  const hitTestSplitGridAreaAtClientPoint = (clientX: number, clientY: number): GridHitTest => {
    if (!secondaryGridView) return { area: "cell", row: null, col: null };
    const renderer = secondaryGridView.grid?.renderer as { pickCellAt?: (x: number, y: number) => { row: number; col: number } | null } | null;
    if (!renderer?.pickCellAt) return { area: "cell", row: null, col: null };
    const selectionCanvas = gridSecondaryEl.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    const canvasRect = (selectionCanvas ?? gridSecondaryEl).getBoundingClientRect();
    const vx = clientX - canvasRect.left;
    const vy = clientY - canvasRect.top;
    if (!Number.isFinite(vx) || !Number.isFinite(vy)) return { area: "cell", row: null, col: null };
    if (vx < 0 || vy < 0 || vx > canvasRect.width || vy > canvasRect.height) return { area: "cell", row: null, col: null };
    const picked = renderer.pickCellAt(vx, vy);
    if (!picked) return { area: "cell", row: null, col: null };
    const headerRows = SPLIT_HEADER_ROWS;
    const headerCols = SPLIT_HEADER_COLS;
    if (picked.row < headerRows && picked.col < headerCols) return { area: "corner", row: null, col: null };
    if (picked.col < headerCols) return { area: "rowHeader", row: picked.row - headerRows, col: null };
    if (picked.row < headerRows) return { area: "colHeader", row: null, col: picked.col - headerCols };
    return { area: "cell", row: picked.row - headerRows, col: picked.col - headerCols };
  };

  const selectedRowIndices = (): number[] => {
    const rows = new Set<number>();
    for (const range of app.getSelectionRanges()) {
      const r = normalizeSelectionRange(range);
      for (let row = r.startRow; row <= r.endRow; row += 1) rows.add(row);
    }
    return [...rows].sort((a, b) => a - b);
  };

  const selectedColIndices = (): number[] => {
    const cols = new Set<number>();
    for (const range of app.getSelectionRanges()) {
      const r = normalizeSelectionRange(range);
      for (let col = r.startCol; col <= r.endCol; col += 1) cols.add(col);
    }
    return [...cols].sort((a, b) => a - b);
  };

  const applyRowHeight = async () => {
    const input = await showInputBox({ prompt: "Row Height", placeHolder: "Enter a row height (px)" });
    if (input == null) return;
    const height = Number(input);
    if (!Number.isFinite(height) || height <= 0) {
      showToast("Row height must be a positive number.", "error");
      return;
    }

    const sheetId = app.getCurrentSheetId();
    const rows = selectedRowIndices();
    if (rows.length === 0) return;

    const doc = app.getDocument();
    doc.beginBatch({ label: "Row Height" });
    for (const row of rows) {
      doc.setRowHeight(sheetId, row, height, { label: "Row Height" });
    }
    doc.endBatch();
  };

  const applyColWidth = async () => {
    const input = await showInputBox({ prompt: "Column Width", placeHolder: "Enter a column width (px)" });
    if (input == null) return;
    const width = Number(input);
    if (!Number.isFinite(width) || width <= 0) {
      showToast("Column width must be a positive number.", "error");
      return;
    }

    const sheetId = app.getCurrentSheetId();
    const cols = selectedColIndices();
    if (cols.length === 0) return;

    const doc = app.getDocument();
    doc.beginBatch({ label: "Column Width" });
    for (const col of cols) {
      doc.setColWidth(sheetId, col, width, { label: "Column Width" });
    }
    doc.endBatch();
  };

  const buildGridContextMenuItems = (): ContextMenuItem[] => {
    const allowEditCommands = !isSpreadsheetEditing();

    let menuItems: ContextMenuItem[] = [];
    if (currentGridArea === "rowHeader") {
      menuItems = [
        { type: "item", label: "Row Height…", enabled: allowEditCommands, onSelect: applyRowHeight },
        {
          type: "item",
          label: "Hide",
          enabled: allowEditCommands,
          onSelect: () => {
            showToast("Hide row is not implemented yet.", "info");
          },
        },
        {
          type: "item",
          label: "Unhide",
          enabled: allowEditCommands,
          onSelect: () => {
            showToast("Unhide row is not implemented yet.", "info");
          },
        },
      ];
    } else if (currentGridArea === "colHeader") {
      menuItems = [
        { type: "item", label: "Column Width…", enabled: allowEditCommands, onSelect: applyColWidth },
        {
          type: "item",
          label: "Hide",
          enabled: allowEditCommands,
          onSelect: () => {
            showToast("Hide column is not implemented yet.", "info");
          },
        },
        {
          type: "item",
          label: "Unhide",
          enabled: allowEditCommands,
          onSelect: () => {
            showToast("Unhide column is not implemented yet.", "info");
          },
        },
      ];
    } else if (currentGridArea === "corner") {
      menuItems = [
        {
          type: "item",
          label: "Select All",
          onSelect: () => {
            const limits = getGridLimitsForFormatting();
            app.selectRange({ range: { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: limits.maxCols - 1 } });
          },
        },
      ];
    } else {
      const undoRedo = app.getUndoRedoState();
      const undoLabelText = typeof undoRedo.undoLabel === "string" ? undoRedo.undoLabel.trim() : "";
      const redoLabelText = typeof undoRedo.redoLabel === "string" ? undoRedo.redoLabel.trim() : "";
      const undoLabel = undoLabelText ? tWithVars("menu.undoWithLabel", { label: undoLabelText }) : t("command.edit.undo");
      const redoLabel = redoLabelText ? tWithVars("menu.redoWithLabel", { label: redoLabelText }) : t("command.edit.redo");

      menuItems = [
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
            const normalized = ranges.map(normalizeSelectionRange);

            for (const [key, cell] of cells.entries()) {
              if (!cell) continue;
              const value = normalizeExtensionCellValue(cell.value ?? null);
              const formula = typeof cell.formula === "string" ? cell.formula : null;
              const cellHasValue =
                (value != null && String(value).trim().length > 0) || (formula != null && formula.trim().length > 0);
              if (!cellHasValue) continue;
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
        {
          type: "item",
          label: t("command.ai.inlineEdit"),
          enabled: allowEditCommands,
          shortcut: getPrimaryCommandKeybindingDisplay("ai.inlineEdit", commandKeybindingDisplayIndex) ?? primaryShortcut("K"),
          onSelect: () => executeBuiltinCommand("ai.inlineEdit"),
        },
        { type: "separator" },
        {
          type: "submenu",
          label: t("menu.format"),
          items: [
            {
              type: "item",
              label: t("command.format.toggleBold"),
              shortcut:
                getPrimaryCommandKeybindingDisplay("format.toggleBold", commandKeybindingDisplayIndex) ?? primaryShortcut("B"),
              onSelect: () => executeBuiltinCommand("format.toggleBold"),
            },
            {
              type: "item",
              label: t("command.format.toggleItalic"),
              shortcut:
                getPrimaryCommandKeybindingDisplay("format.toggleItalic", commandKeybindingDisplayIndex) ??
                (isMac ? "⌘I" : "Ctrl+I"),
              onSelect: () => executeBuiltinCommand("format.toggleItalic"),
            },
            {
              type: "item",
              label: t("command.format.toggleUnderline"),
              shortcut:
                getPrimaryCommandKeybindingDisplay("format.toggleUnderline", commandKeybindingDisplayIndex) ?? primaryShortcut("U"),
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
    }

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

    // Extensions can contribute context menu items for different grid areas:
    // - cell/context (default), row/context, column/context, corner/context
    const menuId =
      currentGridArea === "rowHeader"
        ? ROW_CONTEXT_MENU_ID
        : currentGridArea === "colHeader"
          ? COLUMN_CONTEXT_MENU_ID
          : currentGridArea === "corner"
            ? CORNER_CONTEXT_MENU_ID
            : CELL_CONTEXT_MENU_ID;
    const contributed = resolveMenuItems(extensionHostManager.getContributedMenu(menuId), contextKeys.asLookup());
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

  const openGridContextMenuAtPoint = (x: number, y: number) => {
    const session = (contextMenuSession += 1);
    contextMenu.open({ x, y, items: buildGridContextMenuItems() });

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
        })
        .catch(() => {
          // Best-effort: keep the context menu functional even if extension loading fails.
          if (session !== contextMenuSession) return;
          if (!contextMenu.isOpen()) return;
          contextMenu.update(buildGridContextMenuItems());
        });
    }
  };

  gridRoot.addEventListener("contextmenu", (e) => {
    // Always prevent the native context menu; we render our own.
    e.preventDefault();

    const anchorX = e.clientX;
    const anchorY = e.clientY;

    const hit = hitTestGridAreaAtClientPoint(anchorX, anchorY);
    currentGridArea = hit.area;

    const limits = getGridLimitsForFormatting();
    const ranges = app.getSelectionRanges();
    const normalizedRanges = ranges.map(normalizeSelectionRange);

    // Excel-like behavior: right-click updates selection differently depending on the area.
    if (hit.area === "cell" && hit.row != null && hit.col != null) {
      // Move the active cell only when right-clicking outside the current selection.
      const inSelection = normalizedRanges.some(
        (range) =>
          hit.row! >= range.startRow && hit.row! <= range.endRow && hit.col! >= range.startCol && hit.col! <= range.endCol,
      );
      if (!inSelection) {
        app.activateCell({ row: hit.row, col: hit.col });
      }
    } else if (hit.area === "rowHeader" && hit.row != null) {
      const row = hit.row;
      const alreadySelected = normalizedRanges.some(
        (range) => range.startCol === 0 && range.endCol === limits.maxCols - 1 && row >= range.startRow && row <= range.endRow,
      );
      if (!alreadySelected) {
        app.selectRange({ range: { startRow: row, endRow: row, startCol: 0, endCol: limits.maxCols - 1 } });
      }
    } else if (hit.area === "colHeader" && hit.col != null) {
      const col = hit.col;
      const alreadySelected = normalizedRanges.some(
        (range) => range.startRow === 0 && range.endRow === limits.maxRows - 1 && col >= range.startCol && col <= range.endCol,
      );
      if (!alreadySelected) {
        app.selectRange({ range: { startRow: 0, endRow: limits.maxRows - 1, startCol: col, endCol: col } });
      }
    } else if (hit.area === "corner") {
      const alreadySelected = normalizedRanges.some(
        (range) =>
          range.startRow === 0 &&
          range.endRow === limits.maxRows - 1 &&
          range.startCol === 0 &&
          range.endCol === limits.maxCols - 1,
      );
      if (!alreadySelected) {
        app.selectRange({ range: { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: limits.maxCols - 1 } });
      }
    }

    updateContextKeys();

    openGridContextMenuAtPoint(anchorX, anchorY);
  });

  gridSecondaryEl.addEventListener("contextmenu", (e) => {
    // Always prevent the native context menu; we render our own.
    e.preventDefault();

    const anchorX = e.clientX;
    const anchorY = e.clientY;

    const hit = hitTestSplitGridAreaAtClientPoint(anchorX, anchorY);
    currentGridArea = hit.area;

    const limits = getGridLimitsForFormatting();
    const ranges = app.getSelectionRanges();
    const normalizedRanges = ranges.map(normalizeSelectionRange);

    // Mirror primary-grid right-click semantics, but do not steal focus/scroll from the other pane.
    const selectionOpts = { scrollIntoView: false, focus: false } as const;

    if (hit.area === "cell" && hit.row != null && hit.col != null) {
      const inSelection = normalizedRanges.some(
        (range) =>
          hit.row! >= range.startRow && hit.row! <= range.endRow && hit.col! >= range.startCol && hit.col! <= range.endCol,
      );
      if (!inSelection) {
        app.activateCell({ row: hit.row, col: hit.col }, selectionOpts);
      }
    } else if (hit.area === "rowHeader" && hit.row != null) {
      const row = hit.row;
      const alreadySelected = normalizedRanges.some(
        (range) => range.startCol === 0 && range.endCol === limits.maxCols - 1 && row >= range.startRow && row <= range.endRow,
      );
      if (!alreadySelected) {
        app.selectRange({ range: { startRow: row, endRow: row, startCol: 0, endCol: limits.maxCols - 1 } }, selectionOpts);
      }
    } else if (hit.area === "colHeader" && hit.col != null) {
      const col = hit.col;
      const alreadySelected = normalizedRanges.some(
        (range) => range.startRow === 0 && range.endRow === limits.maxRows - 1 && col >= range.startCol && col <= range.endCol,
      );
      if (!alreadySelected) {
        app.selectRange({ range: { startRow: 0, endRow: limits.maxRows - 1, startCol: col, endCol: col } }, selectionOpts);
      }
    } else if (hit.area === "corner") {
      const alreadySelected = normalizedRanges.some(
        (range) =>
          range.startRow === 0 &&
          range.endRow === limits.maxRows - 1 &&
          range.startCol === 0 &&
          range.endCol === limits.maxCols - 1,
      );
      if (!alreadySelected) {
        app.selectRange(
          { range: { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: limits.maxCols - 1 } },
          selectionOpts,
        );
      }
    }

    updateContextKeys();

    openGridContextMenuAtPoint(anchorX, anchorY);
  });

  const openGridContextMenuAtActiveCell = () => {
    currentGridArea = gridAreaForSelection();
    updateContextKeys();

    // In split-view mode, anchor to the active pane so keyboard-invoked context menus
    // (Shift+F10 / ContextMenu key) open next to the focused grid.
    const split = layoutController.layout.splitView;
    if (split.direction !== "none" && (split.activePane ?? "primary") === "secondary" && secondaryGridView) {
      const active = app.getActiveCell();
      const gridRow = active.row + SPLIT_HEADER_ROWS;
      const gridCol = active.col + SPLIT_HEADER_COLS;
      const rect = secondaryGridView.grid.getCellRect(gridRow, gridCol);
      if (rect) {
        const selectionCanvas = gridSecondaryEl.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
        const rootRect = (selectionCanvas ?? gridSecondaryEl).getBoundingClientRect();
        openGridContextMenuAtPoint(rootRect.left + rect.x, rootRect.top + rect.y + rect.height);
        return;
      }

      const gridRect = gridSecondaryEl.getBoundingClientRect();
      openGridContextMenuAtPoint(gridRect.left + gridRect.width / 2, gridRect.top + gridRect.height / 2);
      return;
    }

    const rect = app.getActiveCellRect();
    if (rect) {
      openGridContextMenuAtPoint(rect.x, rect.y + rect.height);
      return;
    }

    const gridRect = gridRoot.getBoundingClientRect();
    openGridContextMenuAtPoint(gridRect.left + gridRect.width / 2, gridRect.top + gridRect.height / 2);
  };

  window.addEventListener(
    "keydown",
    (e) => {
      if (e.defaultPrevented) return;
      if (isTextInputTarget(e.target)) return;

      const shouldOpen = (e.shiftKey && e.key === "F10") || e.key === "ContextMenu" || e.code === "ContextMenu";
      if (!shouldOpen) return;

      // If *any* context menu is already open, let it manage focus/keyboard handling.
      // This avoids opening the grid context menu while another menu (e.g. sheet tabs)
      // is active.
      const openContextMenu = document.querySelector<HTMLElement>(".context-menu-overlay:not([hidden])");
      if (openContextMenu) return;

      // When focus is on a sheet tab, Shift+F10 / ContextMenu should behave like Excel and
      // open the sheet-tab context menu instead of the active-cell context menu.
      // (The tab strip is responsible for handling these keys.)
      const target = e.target as HTMLElement | null;
      if (target?.closest?.('#sheet-tabs button[role="tab"]')) return;

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
        sheetNameResolver,
      });
    }
    return macrosBackend as any;
  };

  const panelBodyRenderer = createPanelBodyRenderer({
    getDocumentController: () => app.getDocument(),
    getActiveSheetId: () => app.getCurrentSheetId(),
    getSearchWorkbook: () => app.getSearchWorkbook(),
    sheetNameResolver,
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
    ensureExtensionsLoaded,
    onExecuteExtensionCommand: executeExtensionCommand,
    onOpenExtensionPanel: openExtensionPanel,
    onSyncExtensions: () => {
      syncContributedCommands();
      syncContributedPanels();
      updateKeybindings();
      activateOpenExtensionPanels();
    },
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
                sheetNameResolver,
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
 
           const focusMacrosPanelElement = (el: HTMLElement | null) => {
             if (!el) return;
             try {
               el.focus();
             } catch {
               // Best-effort: ignore focus errors (e.g. element not focusable in a headless environment).
             }
           };

           const refreshRunner = async () => {
             await renderMacroRunner(runnerPanel, backend, workbookId, {
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

             const focusTarget = pendingMacrosPanelFocus;
             if (!focusTarget) return;

             if (focusTarget === "runner-select") {
               focusMacrosPanelElement(runnerPanel.querySelector<HTMLElement>('[data-testid="macro-runner-select"]'));
               pendingMacrosPanelFocus = null;
               return;
             }

             if (focusTarget === "runner-run") {
               const runButton = runnerPanel.querySelector<HTMLButtonElement>('[data-testid="macro-runner-run"]');
               if (runButton && !runButton.disabled) {
                 focusMacrosPanelElement(runButton);
               } else {
                 focusMacrosPanelElement(runnerPanel.querySelector<HTMLElement>('[data-testid="macro-runner-select"]'));
               }
               pendingMacrosPanelFocus = null;
               return;
             }

             if (focusTarget === "runner-trust-center") {
               const trustCenterButton = runnerPanel.querySelector<HTMLButtonElement>('[data-testid="macro-runner-trust-center"]');
               if (trustCenterButton && !trustCenterButton.disabled) {
                 focusMacrosPanelElement(trustCenterButton);
               } else {
                 focusMacrosPanelElement(runnerPanel.querySelector<HTMLElement>('[data-testid="macro-runner-select"]'));
               }
               pendingMacrosPanelFocus = null;
             }
           };

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
	          startButton.dataset["testid"] = "macros-recorder-start";
	          startButton.type = "button";
	          startButton.textContent = "Start Recording";
	          buttons.appendChild(startButton);

	          const stopButton = document.createElement("button");
	          stopButton.dataset["testid"] = "macros-recorder-stop";
	          stopButton.type = "button";
	          stopButton.textContent = "Stop Recording";
	          buttons.appendChild(stopButton);

	          if (pendingMacrosPanelFocus === "recorder-start") {
	            focusMacrosPanelElement(startButton);
	            pendingMacrosPanelFocus = null;
	          } else if (pendingMacrosPanelFocus === "recorder-stop") {
	            focusMacrosPanelElement(stopButton);
	            pendingMacrosPanelFocus = null;
	          }

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
            const name = await showInputBox({ prompt: "Macro name:", value: "Recorded Macro" });
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
        container.className = "panel-body__container";
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
        container.className = "panel-body__container";
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
    controls.setAttribute("role", "toolbar");
    controls.setAttribute("aria-label", "Panel controls");

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
      controls.setAttribute("role", "toolbar");
      controls.setAttribute("aria-label", "Panel controls");

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

  // When split view is active, many command handlers (menus, ribbon actions, keybinding commands)
  // call `app.focus()` to restore keyboard focus to the spreadsheet. In split-view mode, restore
  // focus to the *active* pane so shortcuts invoked from the secondary pane don't steal focus back
  // to the primary grid.
  const focusPrimaryGrid = app.focus.bind(app);
  (app as any).focus = () => {
    const split = layoutController.layout.splitView;
    if (split.direction !== "none" && (split.activePane ?? "primary") === "secondary") {
      try {
        (gridSecondaryEl as any).focus?.({ preventScroll: true });
      } catch {
        gridSecondaryEl.focus?.();
      }
      return;
    }
    focusPrimaryGrid();
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

    const rect = gridSplitEl.getBoundingClientRect();
    let latestClientX = event.clientX;
    let latestClientY = event.clientY;
    let lastRatio: number | null = null;
    let rafHandle: number | null = null;

    const applyRatio = (emit: boolean) => {
      const size = direction === "vertical" ? rect.width : rect.height;
      if (size <= 0) return;
      const offset = direction === "vertical" ? latestClientX - rect.left : latestClientY - rect.top;
      const ratio = Math.max(0.1, Math.min(0.9, offset / size));
      // Dragging the splitter updates very frequently; keep updates cheap:
      // - update the in-memory layout without emitting (avoids full renderLayout() churn)
      // - update CSS variables directly so the UI stays live while dragging
      // - rAF-throttle updates so we don't structuredClone+normalize on every pointermove
      if (!emit && lastRatio != null && Math.abs(lastRatio - ratio) < 1e-6) return;
      lastRatio = ratio;
      layoutController.setSplitRatio(ratio, { persist: false, emit });
      applySplitRatioCss(ratio);
    };

    const scheduleDragApply = () => {
      if (rafHandle != null) return;
      rafHandle = window.requestAnimationFrame(() => {
        rafHandle = null;
        applyRatio(false);
      });
    };

    // Apply the initial click point immediately (no need to wait for rAF).
    applyRatio(false);

    const onMove = (move: PointerEvent) => {
      if (move.pointerId !== pointerId) return;
      move.preventDefault();
      latestClientX = move.clientX;
      latestClientY = move.clientY;
      scheduleDragApply();
    };

    const onUp = (up: PointerEvent) => {
      if (up.pointerId !== pointerId) return;
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
      if (rafHandle != null) {
        window.cancelAnimationFrame(rafHandle);
        rafHandle = null;
      }
      try {
        gridSplitterEl.releasePointerCapture(pointerId);
      } catch {
      // Ignore capture release errors.
      }

      // Emit a final layout change (one-time) and persist it so split ratio restores after reload.
      latestClientX = up.clientX;
      latestClientY = up.clientY;
      applyRatio(true);
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
      // `parseGoTo` expects sheet names (what users type, and what formulas use),
      // so we provide the current sheet display name and resolve back to a stable id
      // at execution time.
      getCurrentSheetName: () => workbookSheetStore.getName(app.getCurrentSheetId()) ?? app.getCurrentSheetId(),
      onGoTo: (parsed) => {
        const sheetId = resolveSheetIdFromName(parsed.sheetName);
        if (!sheetId) return;
        const { range } = parsed;
        if (range.startRow === range.endRow && range.startCol === range.endCol) {
          app.activateCell({ sheetId, row: range.startRow, col: range.startCol });
        } else {
          app.selectRange({ sheetId, range });
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

  // Allow stable ids to pass through when the sheet metadata store recognizes them
  // (even if the DocumentController hasn't materialized the sheet yet).
  if (sheetNameResolver.getSheetNameById(trimmed)) return trimmed;

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

const findReplaceDialogs = [findDialog, replaceDialog, goToDialog] as HTMLDialogElement[];

function showExclusiveFindReplaceDialog(dialog: HTMLDialogElement): void {
  for (const other of findReplaceDialogs) {
    if (other === dialog) continue;
    try {
      if (other.open) other.close();
    } catch {
      // ignore
    }
  }
  showDialogAndFocus(dialog);
}

commandRegistry.registerBuiltinCommand(
  "edit.find",
  t("command.edit.find"),
  () => showExclusiveFindReplaceDialog(findDialog as any),
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
  () => showExclusiveFindReplaceDialog(replaceDialog as any),
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
  () => showExclusiveFindReplaceDialog(goToDialog as any),
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

function selectionBoundingBox0Based(): CellRange {
  const range1 = selectionBoundingBox1Based();
  return {
    start: { row: range1.startRow - 1, col: range1.startCol - 1 },
    end: { row: range1.endRow - 1, col: range1.endCol - 1 },
  };
}

function sanitizeFilename(raw: string): string {
  const cleaned = String(raw ?? "")
    // Windows-reserved + generally-illegal filename characters.
    .replace(/[\\/:*?"<>|]+/g, "_")
    .trim();
  return cleaned || "export";
}

function handleExportDelimitedText(args: { delimiter: string; extension: string; mime: string; label: string }): void {
  try {
    const sheetId = app.getCurrentSheetId();
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;
    const doc = app.getDocument();
    const limits = getGridLimitsForFormatting();

    const active = app.getActiveCell();
    const activeCellFallback: CellRange = {
      start: { row: active.row, col: active.col },
      end: { row: active.row, col: active.col },
    };

    let range = selectionBoundingBox0Based();

    // Exporting full-row/full-column/full-sheet selections at Excel scale would attempt to
    // materialize millions/billions of empty cells. Clip those band selections to the used
    // range when possible, and otherwise fall back to exporting the active cell.
    const isFullHeight = range.start.row === 0 && range.end.row === limits.maxRows - 1;
    const isFullWidth = range.start.col === 0 && range.end.col === limits.maxCols - 1;
    if (isFullHeight || isFullWidth) {
      const used = doc.getUsedRange(sheetId);
      if (used) {
        const startRow = Math.max(range.start.row, used.startRow);
        const endRow = Math.min(range.end.row, used.endRow);
        const startCol = Math.max(range.start.col, used.startCol);
        const endCol = Math.min(range.end.col, used.endCol);
        const clipped =
          startRow <= endRow && startCol <= endCol
            ? { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } }
            : null;
        range = clipped ?? activeCellFallback;
      } else {
        range = activeCellFallback;
      }
    }

    const csv = exportDocumentRangeToCsv(doc, sheetId, range, { delimiter: args.delimiter });
    const bytes = new TextEncoder().encode(csv);
    downloadBytes(bytes, `${sanitizeFilename(sheetName)}.${args.extension}`, args.mime);
    app.focus();
  } catch (err) {
    console.error(`Failed to export ${args.label}:`, err);
    showToast(`Failed to export ${args.label}: ${String(err)}`, "error");
  }
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

mountRibbon(ribbonReactRoot, {
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
    versionHistory: () => {
      toggleDockPanel(PanelIds.VERSION_HISTORY);
    },
    branchManager: () => {
      toggleDockPanel(PanelIds.BRANCH_MANAGER);
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
        applyFormattingToSelection("Strikethrough", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { font: { strike: pressed } }, { label: "Strikethrough" });
            if (ok === false) applied = false;
          }
          return applied;
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
    const executeBuiltinCommand = (builtinId: string, ...args: any[]) => {
      void commandRegistry.executeCommand(builtinId, ...args).catch((err) => {
        showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
      });
    };

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
      commandId === "formulas.formulaAuditing.showFormulas" ||
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

        const startedAtMs = Date.now();
        const queryCount = queries.length;

        const shouldNotifyInBackground = (): boolean => {
          try {
            if (typeof document === "undefined") return false;
            // Only notify when the app is not focused / not visible (user likely switched away).
            if ((document as any).hidden) return true;
            if (typeof document.hasFocus === "function") return !document.hasFocus();
            return false;
          } catch {
            return false;
          }
        };

        try {
          const handle = service.refreshAll();
          await handle.promise;
          const elapsedMs = Date.now() - startedAtMs;
          // Avoid spamming notifications for extremely fast refreshes; only notify when the
          // user likely switched away or when the refresh took long enough to be meaningful.
          if (shouldNotifyInBackground() || elapsedMs >= 5_000) {
            const noun = queryCount === 1 ? "query" : "queries";
            void notify({ title: "Power Query refresh complete", body: `Refreshed ${queryCount} ${noun}.` });
          }
        } catch (err) {
          console.error("Failed to refresh all queries:", err);
          showToast(`Failed to refresh queries: ${String(err)}`, "error");
          if (shouldNotifyInBackground()) {
            void notify({ title: "Power Query refresh failed", body: "One or more queries failed to refresh." });
          }
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
      // Always call openPanel so we activate docked panels and also trigger a layout
      // re-render even when the panel is already floating (useful for commands like
      // Record Macro that update panel-local state).
      layoutController.openPanel(panelId);

      // Floating panels can be minimized; opening should restore them.
      if (placement.kind === "floating" && layoutController.layout?.floating?.[panelId]?.minimized) {
        layoutController.setFloatingPanelMinimized(panelId, false);
      }
    };

    const focusScriptEditorPanel = () => {
      // Script editor is mounted via panel rendering; wait a frame (or two) so the
      // DOM nodes exist before focusing.
      requestAnimationFrame(() => requestAnimationFrame(() => {
        const el =
          document.querySelector<HTMLElement>('[data-testid="script-editor-code"]') ??
          document.querySelector<HTMLElement>('[data-testid="script-editor-run"]');
        try {
          el?.focus();
        } catch {
          // Best-effort: ignore focus errors (e.g. element not focusable in headless environments).
        }
      }));
    };

    const focusVbaMigratePanel = () => {
      // The VBA migrate panel is a React mount; give it a couple of frames to render.
      requestAnimationFrame(() => requestAnimationFrame(() => {
        const el =
          document.querySelector<HTMLElement>('[data-testid="vba-entrypoint"]') ??
          document.querySelector<HTMLElement>('button[data-testid^="vba-module-"]') ??
          document.querySelector<HTMLElement>('[data-testid="vba-refresh"]');
        try {
          el?.focus();
        } catch {
          // Best-effort: ignore focus errors (e.g. element not focusable in headless environments).
        }
      }));
    };

    const openCustomZoomQuickPick = async (): Promise<void> => {
      if (!app.supportsZoom()) return;
      // Keep the custom zoom picker aligned with the shared-grid zoom clamp
      // (currently 25%–400%, Excel-style).
      const baseOptions = [25, 50, 75, 100, 125, 150, 200, 400];
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
      applyFormattingToSelection("Font", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { font: { name: fontName } }, { label: "Font" });
          if (ok === false) applied = false;
        }
        return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { fill: null }, { label: "Fill color" });
            if (ok === false) applied = false;
          }
          return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { font: { color: null } }, { label: "Font color" });
            if (ok === false) applied = false;
          }
          return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear formats" });
            if (ok === false) applied = false;
          }
          return applied;
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
            let applied = true;
            for (const range of ranges) {
              doc.clearRange(sheetId, range, { label: "Clear all" });
              const ok = doc.setRangeFormat(sheetId, range, null, { label: "Clear all" });
              if (ok === false) applied = false;
            }
            return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { border: null }, { label: "Borders" });
            if (ok === false) applied = false;
          }
          return applied;
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
            let applied = true;
            const applyBorder = (targetRange: CellRange, patch: Record<string, any> | null) => {
              const ok = doc.setRangeFormat(sheetId, targetRange, patch, { label: "Borders" });
              if (ok === false) applied = false;
            };
            for (const range of ranges) {
              const startRow = range.start.row;
              const endRow = range.end.row;
              const startCol = range.start.col;
              const endCol = range.end.col;

              // Top edge.
              applyBorder({ start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } }, { border: { top: edge } });

              // Bottom edge.
              applyBorder({ start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } }, { border: { bottom: edge } });

              // Left edge.
              applyBorder({ start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } }, { border: { left: edge } });

              // Right edge.
              applyBorder({ start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } }, { border: { right: edge } });
            }
            return applied;
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
            let applied = true;
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

              const ok = doc.setRangeFormat(sheetId, targetRange, borderPatch, { label: "Borders" });
              if (ok === false) applied = false;
            }
            return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: null }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      if (kind === "number") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "0.00" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
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
      if (kind === "time") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "h:mm:ss" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      if (kind === "fraction") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "# ?/?" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      if (kind === "scientific") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "0.00E+00" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      if (kind === "text") {
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "@" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      return;
    }

    const accountingPrefix = "home.number.accounting.";
    if (commandId.startsWith(accountingPrefix)) {
      const currency = commandId.slice(accountingPrefix.length);
      const symbol = (() => {
        switch (currency) {
          case "eur":
            return "€";
          case "gbp":
            return "£";
          case "jpy":
            return "¥";
          case "usd":
          default:
            return "$";
        }
      })();

      applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
        let applied = true;
        for (const range of ranges) {
          const ok = doc.setRangeFormat(sheetId, range, { numberFormat: `${symbol}#,##0.00` }, { label: "Number format" });
          if (ok === false) applied = false;
        }
        return applied;
      });
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

      case "file.info.manageWorkbook.versions": {
        toggleDockPanel(PanelIds.VERSION_HISTORY);
        return;
      }
      case "file.info.manageWorkbook.branches": {
        toggleDockPanel(PanelIds.BRANCH_MANAGER);
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
      case "file.export.changeFileType.csv":
        handleExportDelimitedText({ delimiter: ",", extension: "csv", mime: "text/csv", label: "CSV" });
        return;
      case "file.export.changeFileType.tsv":
        handleExportDelimitedText({
          delimiter: "\t",
          extension: "tsv",
          mime: "text/tab-separated-values",
          label: "TSV",
        });
        return;
      case "file.export.export.xlsx":
      case "file.export.changeFileType.xlsx": {
        if (!tauriBackend) {
          showDesktopOnlyToast("Exporting workbooks is available in the desktop app.");
          return;
        }
        void handleSaveAs().catch((err) => {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        });
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
      case "home.clipboard.paste":
      case "home.clipboard.paste.default":
        void app.clipboardPaste();
        // Ribbon dropdown menu items restore focus to the trigger button after dispatching the
        // command. Defer grid focus so keyboard shortcuts continue to work even if the clipboard
        // operation is a no-op (e.g. blocked clipboard permissions).
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.paste.values":
        void app.clipboardPasteSpecial("values");
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.paste.formulas":
        void app.clipboardPasteSpecial("formulas");
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.paste.formats":
        void app.clipboardPasteSpecial("formats");
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.paste.transpose":
        showToast("Paste Transpose not implemented");
        queueMicrotask(() => app.focus());
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
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.pasteSpecial.formulas":
        void app.clipboardPasteSpecial("formulas");
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.pasteSpecial.formats":
        void app.clipboardPasteSpecial("formats");
        queueMicrotask(() => app.focus());
        return;
      case "home.clipboard.pasteSpecial.transpose":
        showToast("Paste Transpose not implemented");
        queueMicrotask(() => app.focus());
        return;

      case "view.macros.viewMacros":
      case "view.macros.viewMacros.run":
      case "view.macros.viewMacros.edit":
      case "view.macros.viewMacros.delete":
        // Clear any previously-requested focus so that edit/Visual Basic actions don't
        // get focus stolen by an earlier async macro runner render.
        if (commandId.endsWith(".edit")) pendingMacrosPanelFocus = null;
        if (commandId === "view.macros.viewMacros") pendingMacrosPanelFocus = "runner-select";
        if (commandId.endsWith(".run")) pendingMacrosPanelFocus = "runner-run";
        if (commandId.endsWith(".delete")) pendingMacrosPanelFocus = "runner-select";
        openRibbonPanel(PanelIds.MACROS);
        // "Edit…" in Excel normally opens an editor; best-effort surface our Script Editor panel too.
        if (commandId.endsWith(".edit")) {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
          focusScriptEditorPanel();
        }
        return;

      case "view.macros.recordMacro":
        pendingMacrosPanelFocus = "recorder-stop";
        activeMacroRecorder?.start();
        openRibbonPanel(PanelIds.MACROS);
        return;
      case "view.macros.recordMacro.stop":
        pendingMacrosPanelFocus = "recorder-start";
        activeMacroRecorder?.stop();
        openRibbonPanel(PanelIds.MACROS);
        return;

      case "view.macros.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. Avoid the default toast.
        return;

      case "developer.code.macros":
      case "developer.code.macros.run":
      case "developer.code.macros.edit":
        // Clear any previously-requested focus so that edit/Visual Basic actions don't
        // get focus stolen by an earlier async macro runner render.
        if (commandId.endsWith(".edit")) pendingMacrosPanelFocus = null;
        if (commandId === "developer.code.macros") pendingMacrosPanelFocus = "runner-select";
        if (commandId.endsWith(".run")) pendingMacrosPanelFocus = "runner-run";
        openRibbonPanel(PanelIds.MACROS);
        if (commandId.endsWith(".edit")) {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
          focusScriptEditorPanel();
        }
        return;

      case "developer.code.macroSecurity":
      case "developer.code.macroSecurity.trustCenter":
        pendingMacrosPanelFocus = "runner-trust-center";
        openRibbonPanel(PanelIds.MACROS);
        return;

      case "developer.code.recordMacro":
        pendingMacrosPanelFocus = "recorder-stop";
        activeMacroRecorder?.start();
        openRibbonPanel(PanelIds.MACROS);
        return;
      case "developer.code.recordMacro.stop":
        pendingMacrosPanelFocus = "recorder-start";
        activeMacroRecorder?.stop();
        openRibbonPanel(PanelIds.MACROS);
        return;

      case "developer.code.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. Avoid the default toast.
        return;

      case "developer.code.visualBasic":
        pendingMacrosPanelFocus = null;
        // Desktop builds expose a VBA migration panel (used as a stand-in for the VBA editor).
        if (typeof (globalThis as any).__TAURI__?.core?.invoke === "function") {
          openRibbonPanel(PanelIds.VBA_MIGRATE);
          focusVbaMigratePanel();
        } else {
          openRibbonPanel(PanelIds.SCRIPT_EDITOR);
          focusScriptEditorPanel();
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "top" } }, { label: "Vertical align" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.middleAlign":
        applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            // Spreadsheet vertical alignment uses "center" (Excel/OOXML); the grid maps this to CSS middle.
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "center" } }, { label: "Vertical align" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.bottomAlign":
        applyFormattingToSelection("Vertical align", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { vertical: "bottom" } }, { label: "Vertical align" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.center":
        applyFormattingToSelection("Align center", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "center"));
        return;
      case "home.alignment.alignRight":
        applyFormattingToSelection("Align right", (doc, sheetId, ranges) => setHorizontalAlign(doc, sheetId, ranges, "right"));
        return;

      case "home.alignment.increaseIndent": {
        const current = activeCellIndentLevel();
        const next = Math.min(250, current + 1);
        if (next === current) return;
        applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }

      case "home.alignment.decreaseIndent": {
        const current = activeCellIndentLevel();
        const next = Math.max(0, current - 1);
        if (next === current) return;
        applyFormattingToSelection("Indent", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { indent: next } }, { label: "Indent" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }

      case "home.alignment.orientation.angleCounterclockwise":
        applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 45 } }, { label: "Text orientation" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.orientation.angleClockwise":
        applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -45 } }, { label: "Text orientation" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.orientation.verticalText":
        applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            // Excel/OOXML uses 255 as a sentinel for vertical text (stacked).
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 255 } }, { label: "Text orientation" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.orientation.rotateUp":
        applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: 90 } }, { label: "Text orientation" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.alignment.orientation.rotateDown":
        applyFormattingToSelection("Text orientation", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { alignment: { textRotation: -90 } }, { label: "Text orientation" });
            if (ok === false) applied = false;
          }
          return applied;
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
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: "#,##0.00" }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      case "home.number.increaseDecimal": {
        const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(), "increase");
        if (!next) return;
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      case "home.number.decreaseDecimal": {
        const next = stepDecimalPlacesInNumberFormat(activeCellNumberFormat(), "decrease");
        if (!next) return;
        applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
          let applied = true;
          for (const range of ranges) {
            const ok = doc.setRangeFormat(sheetId, range, { numberFormat: next }, { label: "Number format" });
            if (ok === false) applied = false;
          }
          return applied;
        });
        return;
      }
      case "home.number.formatCells":
      case "home.number.moreFormats.formatCells":
      case "home.number.moreFormats.custom":
      case "home.cells.format.formatCells":
        openFormatCells();
        return;
      case "home.editing.autoSum":
      case "home.editing.autoSum.sum":
        executeBuiltinCommand("edit.autoSum");
        // `edit.autoSum` restores focus, but ensure the grid is focused even if the command is a no-op.
        app.focus();
        return;
      case "home.editing.fill.down":
        executeBuiltinCommand("edit.fillDown");
        app.focus();
        return;
      case "home.editing.fill.right":
        executeBuiltinCommand("edit.fillRight");
        app.focus();
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
        scheduleRibbonSelectionFormatStateUpdate();
        app.focus();
        return;
      case "view.appearance.theme.light":
        themeController.setThemePreference("light");
        scheduleRibbonSelectionFormatStateUpdate();
        app.focus();
        return;
      case "view.appearance.theme.dark":
        themeController.setThemePreference("dark");
        scheduleRibbonSelectionFormatStateUpdate();
        app.focus();
        return;
      case "view.appearance.theme.highContrast":
        themeController.setThemePreference("high-contrast");
        scheduleRibbonSelectionFormatStateUpdate();
        app.focus();
        return;

      // --- Debug / dev controls migrated from the legacy status bar ---------------
      // Keep these command ids stable because Playwright e2e depends on their `data-testid`s.
      case "open-ai-panel":
      case "open-panel-ai-chat":
        toggleDockPanel(PanelIds.AI_CHAT);
        return;
      case "open-inline-ai-edit":
        void commandRegistry.executeCommand("ai.inlineEdit").catch((err) => {
          showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
        });
        return;
      case "open-ai-audit-panel":
      case "open-panel-ai-audit":
        toggleDockPanel(PanelIds.AI_AUDIT);
        return;
      case "open-version-history-panel":
        toggleDockPanel(PanelIds.VERSION_HISTORY);
        return;
      case "open-branch-manager-panel":
        toggleDockPanel(PanelIds.BRANCH_MANAGER);
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
        // Opening the Marketplace panel should not eagerly start the extension host; users can
        // browse/search/install extensions without spinning up the extension runtime. Extensions
        // are loaded lazily when the user opens the Extensions panel or executes an extension
        // command (see `ensureExtensionsLoaded()`).
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
        // "Open Extensions" should be idempotent: if the panel is already open (for example,
        // restored from layout persistence after a reload), focus it instead of toggling
        // it closed. Closing remains available via the panel's close control.
        ribbonLayoutController?.openPanel(PanelIds.EXTENSIONS);
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
function isCollabSessionActive(): boolean {
  const anyApp = app as any;
  try {
    return typeof anyApp.getCollabSession === "function" && anyApp.getCollabSession() != null;
  } catch {
    // Ignore collab detection failures and fall back to normal dirty tracking.
    return false;
  }
}

function isDirtyForUnsavedChangesPrompts(): boolean {
  if (isCollabSessionActive()) return false;
  return app.getDocument().isDirty;
}

const collabAwareDirtyController = {
  get isDirty(): boolean {
    return isDirtyForUnsavedChangesPrompts();
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
  // If the active sheet is hidden (or otherwise missing from the visible list),
  // fall back to the first visible option so the <select> always has a value.
  const hasActive = sheets.some((sheet) => sheet.id === activeId);
  sheetSwitcherEl.value = hasActive ? activeId : sheets[0]?.id ?? "";
}

sheetSwitcherEl.addEventListener("change", () => {
  void (async () => {
    const next = sheetSwitcherEl.value;
    if (!next) return;

    // If a sheet tab rename is in progress, allow it to commit/cancel before switching sheets.
    //
    // IMPORTANT: the rename UI lives inside the React sheet tab strip, while the sheet switcher
    // is owned by main.ts. Detect rename mode via the presence/focus state of the inline rename
    // <input> and block sheet switching if the rename remains active (e.g. invalid name).
    const waitForRenameToResolve = async (): Promise<boolean> => {
      const start = typeof performance !== "undefined" ? performance.now() : Date.now();
      const timeoutMs = 5_000;
      while ((typeof performance !== "undefined" ? performance.now() : Date.now()) - start < timeoutMs) {
        const input = sheetTabsRootEl.querySelector<HTMLInputElement>("input.sheet-tab__input");
        if (!input) return true;
        // If the input has regained focus, the rename is still active (likely invalid); don't switch.
        if (document.activeElement === input) return false;
        await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      }
      return false;
    };

    const renameInput = sheetTabsRootEl.querySelector<HTMLInputElement>("input.sheet-tab__input");
    if (renameInput) {
      // Ensure any blur-commit handlers have a chance to run. (If the user interacted with the
      // sheet switcher via mouse, the rename input will already have blurred.)
      try {
        renameInput.blur();
      } catch {
        // ignore
      }
      const ok = await waitForRenameToResolve();
      if (!ok) {
        // Restore the dropdown value to the current visible sheet and keep the user in rename mode.
        const sheets = listSheetsForUi();
        renderSheetSwitcher(sheets, app.getCurrentSheetId());
        try {
          sheetTabsRootEl.querySelector<HTMLInputElement>("input.sheet-tab__input")?.focus();
        } catch {
          // ignore
        }
        return;
      }
    }

    // Defensive: only allow activating visible sheets via the dropdown.
    const sheets = listSheetsForUi();
    if (!sheets.some((sheet) => sheet.id === next)) {
      renderSheetSwitcher(sheets, app.getCurrentSheetId());
      return;
    }

    app.activateSheet(next);
    restoreFocusAfterSheetNavigation();
  })();
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
  const tauri = (globalThis as any).__TAURI__;
  const dialog = tauri?.dialog ?? tauri?.plugin?.dialog;
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

function normalizeSheetList(info: WorkbookInfo): SheetUiInfo[] {
  const sheets = Array.isArray(info.sheets) ? info.sheets : [];
  return sheets
    .map((sheet) => {
      const id = String((sheet as any)?.id ?? "").trim();
      const nameRaw = (sheet as any)?.name ?? (sheet as any)?.id ?? "";
      const name = String(nameRaw).trim() || id;

      const rawVisibility = (sheet as any)?.visibility;
      const visibility: SheetVisibility =
        rawVisibility === "visible" || rawVisibility === "hidden" || rawVisibility === "veryHidden"
          ? rawVisibility
          : "visible";

      const rawTabColor = (sheet as any)?.tabColor;
      const tabColor: TabColor | undefined = (() => {
        if (rawTabColor == null) return undefined;
        if (typeof rawTabColor === "string") {
          const rgb = rawTabColor.trim();
          return rgb ? { rgb } : undefined;
        }
        if (typeof rawTabColor !== "object") return undefined;
        const color = rawTabColor as any;
        const out: TabColor = {};
        if (typeof color.rgb === "string" && color.rgb.trim() !== "") out.rgb = color.rgb;
        if (typeof color.theme === "number") out.theme = color.theme;
        if (typeof color.indexed === "number") out.indexed = color.indexed;
        if (typeof color.tint === "number") out.tint = color.tint;
        if (typeof color.auto === "boolean") out.auto = color.auto;
        return Object.keys(out).length > 0 ? out : undefined;
      })();

      return { id, name, visibility, tabColor };
    })
    .filter((sheet) => sheet.id !== "");
}

function commitAllPendingEditsForCommand(): void {
  // Secondary pane has its own in-cell editor; ensure we commit it too so unsaved-change
  // prompts and file operations include the latest user input.
  secondaryGridView?.commitPendingEditsForCommand();
  app.commitPendingEditsForCommand();
}

async function confirmDiscardDirtyState(actionLabel: string): Promise<boolean> {
  // If the user triggers File commands while editing, the document may not yet be dirty
  // (the edit is still in the UI editor). Commit first so discard prompts are correct.
  commitAllPendingEditsForCommand();
  if (!isDirtyForUnsavedChangesPrompts()) return true;
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
      visibility: sheet.visibility ?? "visible",
      tabColor: sheet.tabColor,
    })),
  );
  installSheetStoreSubscription();
  // Keep the e2e harness up-to-date when we swap the store after opening a workbook.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).__workbookSheetStore = workbookSheetStore;

  const { maxRows: MAX_ROWS, maxCols: MAX_COLS, chunkRows: CHUNK_ROWS } = getWorkbookLoadLimits();
  // Backend-enforced Tauri limits for `get_range` requests. Keep in sync with:
  // `apps/desktop/src-tauri/src/resource_limits.rs`.
  const MAX_RANGE_DIM = 10_000;
  const MAX_RANGE_CELLS_PER_CALL = 1_000_000;

  const snapshotSheets: Array<{
    id: string;
    name: string;
    visibility: SheetVisibility;
    tabColor?: TabColor;
    cells: any[];
  }> = [];
  const truncations: WorkbookLoadTruncation[] = [];

  for (const sheet of sheets) {
    const cells: Array<{ row: number; col: number; value: unknown | null; formula: string | null; format: null }> = [];

    const usedRange = await tauriBackend.getSheetUsedRange(sheet.id);
    if (!usedRange) {
      snapshotSheets.push({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility,
        tabColor: sheet.tabColor,
        cells,
      });
      continue;
    }

    const { startRow, endRow, startCol, endCol, truncatedRows, truncatedCols } = clampUsedRange(usedRange, {
      maxRows: MAX_ROWS,
      maxCols: MAX_COLS,
    });
    if (truncatedRows || truncatedCols) {
      truncations.push({
        sheetId: sheet.id,
        sheetName: sheet.name,
        originalRange: usedRange,
        loadedRange: { startRow, endRow, startCol, endCol },
        truncatedRows,
        truncatedCols,
      });
    }

    if (startRow > endRow || startCol > endCol) {
      snapshotSheets.push({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility,
        tabColor: sheet.tabColor,
        cells,
      });
      continue;
    }

    const effectiveChunkRows = Math.max(1, Math.min(CHUNK_ROWS, MAX_RANGE_DIM));

    for (let chunkStartRow = startRow; chunkStartRow <= endRow; chunkStartRow += effectiveChunkRows) {
      const chunkEndRow = Math.min(endRow, chunkStartRow + effectiveChunkRows - 1);
      const rowCount = chunkEndRow - chunkStartRow + 1;
      // Tauri bounds each `get_range` call by total cell count and per-axis dimensions.
      // Chunk columns automatically based on the requested rowCount so large width
      // workbooks can still load without failing backend range-size checks.
      const maxColsByCellLimit = Math.max(1, Math.floor(MAX_RANGE_CELLS_PER_CALL / Math.max(1, rowCount)));
      const effectiveChunkCols = Math.max(1, Math.min(MAX_RANGE_DIM, maxColsByCellLimit));

      for (let chunkStartCol = startCol; chunkStartCol <= endCol; chunkStartCol += effectiveChunkCols) {
        const chunkEndCol = Math.min(endCol, chunkStartCol + effectiveChunkCols - 1);
        const range = await tauriBackend.getRange({
          sheetId: sheet.id,
          startRow: chunkStartRow,
          startCol: chunkStartCol,
          endRow: chunkEndRow,
          endCol: chunkEndCol,
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
              col: chunkStartCol + c,
              value: formula != null ? null : value,
              formula,
              format: null,
            });
          }
        }
      }
    }

    snapshotSheets.push({
      id: sheet.id,
      name: sheet.name,
      visibility: sheet.visibility,
      tabColor: sheet.tabColor,
      cells,
    });
  }

  const snapshot = encodeDocumentSnapshot({ schemaVersion: 1, sheets: snapshotSheets });
  const workbookSignature = await workbookSignaturePromise;
  // Reset Power Query table signatures before applying the snapshot so any
  // in-flight query executions cannot reuse cached table results from a
  // previously-opened workbook.
  refreshTableSignaturesFromBackend(doc, [], { workbookSignature });
  refreshDefinedNameSignaturesFromBackend(doc, [], { workbookSignature });
  await app.restoreDocumentState(snapshot);

  warnIfWorkbookLoadTruncated(truncations, { maxRows: MAX_ROWS, maxCols: MAX_COLS }, showToast);

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

  // Default to the first *visible* sheet so hidden sheets are never activated
  // (and therefore never shown as the selected value in the sheet switcher).
  const visibleSheets = listSheetsForUi();
  const firstSheetId = visibleSheets[0]?.id ?? sheets[0].id;
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
      const err = new Error("Open workbook cancelled");
      err.name = "AbortError";
      throw err;
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
        const openedPath = activeWorkbook.path ?? activeWorkbook.origin_path ?? path;
        emitWorkbookOpenedForExtensions(getWorkbookSnapshotForExtensions({ pathOverride: openedPath }));
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

    // Each workbook has its own persisted zoom. Restore it (or reset to the default zoom)
    // after switching the active workbook id.
    applyPersistedSharedGridZoom({ resetIfMissing: true });
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
  commitAllPendingEditsForCommand();
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (!workbookSync) return;

  if (!activeWorkbook.path) {
    await handleSaveAs(options);
    return;
  }

  if (options.notifyExtensions !== false) {
    try {
      emitBeforeSaveForExtensions(getWorkbookSnapshotForExtensions());
    } catch {
      // Ignore extension host errors; save should still succeed.
    }
  }
  await workbookSync.markSaved();
}

async function handleSaveAs(
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean; throwOnCancel?: boolean } = {},
): Promise<void> {
  commitAllPendingEditsForCommand();
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
      const err = new Error("Save cancelled");
      err.name = "AbortError";
      throw err;
    }
    return;
  }

  await handleSaveAsPath(path, { previousPanelWorkbookId, notifyExtensions: options.notifyExtensions });
}

async function handleSaveAsPath(
  path: string,
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean } = {},
): Promise<void> {
  commitAllPendingEditsForCommand();
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (typeof path !== "string" || path.trim() === "") return;

  const previousPanelWorkbookId = options.previousPanelWorkbookId ?? activePanelWorkbookId;

  // Ensure any pending microtask-batched workbook edits are flushed before saving.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await drainBackendSync();
  if (options.notifyExtensions !== false) {
    try {
      emitBeforeSaveForExtensions(getWorkbookSnapshotForExtensions({ pathOverride: path }));
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
  // Ensure zoom persistence follows the workbook when its id changes (e.g. Save As).
  persistCurrentSharedGridZoom();
  startPowerQueryService();
  rerenderLayout?.();
}

async function handleNewWorkbook(
  options: {
    notifyExtensions?: boolean;
    throwOnCancel?: boolean;
    /**
     * Human-readable action label used in the discard-unsaved-changes prompt.
     * Defaults to "create a new workbook".
     */
    actionLabel?: string;
    /**
     * Error message used when `throwOnCancel` is true and the user cancels the discard prompt.
     * Defaults to "Create workbook cancelled".
     */
    cancelMessage?: string;
  } = {},
): Promise<void> {
  if (!tauriBackend) return;
  const actionLabel = options.actionLabel ?? "create a new workbook";
  const cancelMessage = options.cancelMessage ?? "Create workbook cancelled";
  const ok = await confirmDiscardDirtyState(actionLabel);
  if (!ok) {
    if (options.throwOnCancel) {
      const err = new Error(cancelMessage);
      err.name = "AbortError";
      throw err;
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
        emitWorkbookOpenedForExtensions(getWorkbookSnapshotForExtensions());
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

    // New workbooks should not inherit zoom from the prior workbook; treat zoom as a per-workbook
    // view setting and restore (or reset) based on the new session id.
    applyPersistedSharedGridZoom({ resetIfMissing: true });
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
  //
  // Canonical desktop event names (keep in sync with the allowlist):
  // Rust -> JS (`listen`):
  // - close-prep, close-requested
  // - open-file, file-dropped
  // - tray-open, tray-new, tray-quit
  // - shortcut-quick-open, shortcut-command-palette
  // - menu-open, menu-new, menu-save, menu-save-as, menu-export-pdf, menu-close-window, menu-quit,
  //   menu-undo, menu-redo, menu-cut, menu-copy, menu-paste, menu-paste-special, menu-select-all,
  //   menu-zoom-in, menu-zoom-out, menu-zoom-reset, menu-about, menu-check-updates
  // - startup:window-visible, startup:webview-loaded, startup:tti, startup:metrics
  // - update-check-started, update-check-already-running, update-not-available, update-check-error, update-available
  // - oauth-redirect
  // JS -> Rust (`emit`):
  // - open-file-ready, oauth-redirect-ready
  // - close-prep-done, close-handled
  // - updater-ui-ready, coi-check-result
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
    isDirty: () => isDirtyForUnsavedChangesPrompts(),
    runWorkbookBeforeClose: async () => {
      commitAllPendingEditsForCommand();
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
      // If the user is mid-edit when a native window close is requested, ensure the edit
      // is committed before we flush workbook sync and before the backend runs Workbook_BeforeClose.
      commitAllPendingEditsForCommand();
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

  // Queue open-file requests until after the listener is registered, then signal readiness to
  // flush any pending paths from the Rust host.
  installOpenFileIpc({ listen, emit, onOpenPath: queueOpenWorkbook });

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

  // Native menu bar integration (desktop shell emits menu-open/menu-save/... events).
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
    try {
      const provider = await getClipboardProvider();
      await provider.write({ text: String(text ?? "") });
      return true;
    } catch {
      return false;
    }
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
  void listen("menu-paste-special", () => {
    // Paste Special is a spreadsheet command; do not invoke it while focus is inside
    // text editors (formula bar, dialogs, etc.).
    const target = getTextEditingTarget();
    if (target) return;
    void commandRegistry.executeCommand("clipboard.pasteSpecial");
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
    void (async () => {
      const tauri = (globalThis as any).__TAURI__;
      const appApi = tauri?.app;
      const getName = appApi?.getName as (() => Promise<unknown>) | undefined;
      const getVersion = appApi?.getVersion as (() => Promise<unknown>) | undefined;

      let name = "Formula";
      let version: string | null = null;

      if (typeof getName === "function") {
        try {
          const resolved = await getName();
          if (typeof resolved === "string" && resolved.trim()) name = resolved.trim();
        } catch {
          // ignore
        }
      }

      if (typeof getVersion === "function") {
        try {
          const resolved = await getVersion();
          if (typeof resolved === "string" && resolved.trim()) version = resolved.trim();
        } catch {
          // ignore
        }
      }

      const message = version ? `${name}\nVersion ${version}` : name;
      await nativeDialogs.alert(message, { title: `About ${name}` });
    })();
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

  void listen("menu-open-release-page", () => {
    void shellOpen(FORMULA_RELEASES_URL).catch((err) => {
      console.error("Failed to open release page:", err);
    });
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
      // Ensure in-progress cell/formula edits are committed so `doc.isDirty` is accurate
      // and Workbook_BeforeClose macros see the latest inputs.
      commitAllPendingEditsForCommand();

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

      if (isDirtyForUnsavedChangesPrompts()) {
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
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__workbookSheetStore = workbookSheetStore;

// Time-to-interactive instrumentation (best-effort, no-op for web builds).
void markStartupTimeToInteractive({ whenIdle: () => app.whenIdle() }).catch(() => {});
