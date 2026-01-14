// IMPORTANT: keep this as the very first import so startup timing listeners are installed
// as early as possible. The Rust host may emit `startup:*` events very early during app
// load; we ask it to re-emit cached timings once listeners are ready.
//
// NOTE: This uses the `.ts` extension (rather than the repo's usual `.js` specifier convention)
// so it matches the separate module script entry in `apps/desktop/index.html`, ensuring the
// bootstrap side effects run exactly once in the browser module graph.
import "./tauri/startupMetricsBootstrap.ts";

import { SpreadsheetApp } from "./app/spreadsheetApp";
import type { SheetNameResolver } from "./sheet/sheetNameResolver";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/command-palette.css";
import "./styles/dialogs.css";
import "./styles/print.css";
import "./styles/sort-filter.css";
import "./styles/extensions-ui.css";
import "./styles/extensions.css";
import "./styles/workspace.css";
import "./styles/ai-chat.css";
import "./styles/query-editor.css";
import "./styles/pivot-builder.css";
import "./styles/charts-overlay.css";
import "./styles/selection-pane.css";
import "./styles/scrollbars.css";
import "./styles/comments.css";
import "./styles/shell.css";
import "./styles/auditing.css";
import "./styles/format-cells-dialog.css";
import "./styles/context-menu.css";
import "./styles/conflicts.css";
import "./styles/presence-ui.css";
import "./styles/macros-runner.css";
import "./styles/script-editor.css";
import "./styles/python-panel.css";
import "./styles/data-queries.css";
import "./styles/what-if.css";
import "./styles/solver.css";

import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { createLazyImport } from "./startup/lazyImport.js";

import { SheetTabStrip } from "./sheets/SheetTabStrip";

import { ThemeController } from "./theme/themeController.js";

import { mountRibbon } from "./ribbon/index.js";
import { createRibbonActions } from "./ribbon/ribbonCommandRouter.js";

import { computeSelectionFormatState } from "./ribbon/selectionFormatState.js";
import { computeRibbonDisabledByIdFromCommandRegistry } from "./ribbon/ribbonCommandRegistryDisabling.js";
import { getRibbonUiStateSnapshot, setRibbonUiState } from "./ribbon/ribbonUiState.js";
import { deriveRibbonAriaKeyShortcutsById, deriveRibbonShortcutById } from "./ribbon/ribbonShortcuts.js";
import { MAX_AXIS_RESIZE_INDICES, promptAndApplyAxisSizing, selectedColIndices, selectedRowIndices } from "./ribbon/axisSizing.js";
import { RIBBON_DISABLED_BY_ID_WHILE_EDITING } from "./ribbon/ribbonEditingDisabledById.js";

import type { CellRange as GridCellRange } from "@formula/grid";

import { rewriteDocumentFormulasForSheetDelete, rewriteDocumentFormulasForSheetRename } from "./sheets/sheetFormulaRewrite";
import { pickAdjacentVisibleSheetId } from "./sheets/sheetNavigation";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import type { SecondaryGridView } from "./grid/splitView/secondaryGridView.js";
import { resolveDesktopGridMode } from "./grid/shared/desktopGridMode.js";
import { resolveEnableDrawingInteractions } from "./drawings/drawingInteractionsFlag.js";
import { FORMULA_AUDITING_RIBBON_COMMAND_IDS } from "./commands/formulaAuditingCommandIds.js";
import { getPanelTitle, panelRegistry, PanelIds } from "./panels/panelRegistry.js";
import type { PanelBodyRenderer, PanelBodyRendererOptions } from "./panels/panelBodyRenderer.js";
import { createWhatIfApi } from "./panels/what-if/api.js";
import type { MacroRecorder } from "./macro-recorder/index.js";
import { mountTitlebar } from "./titlebar/mountTitlebar.js";
import type { VbaEventMacrosHandle } from "./macros/event_macros";
import type { MacroRunRequest, MacroTrustDecision } from "./macros/types.js";
import { applyMacroCellUpdates } from "./macros/applyUpdates";
import { installUnsavedChangesPrompt } from "./document/index.js";
import type { DocumentController } from "./document/documentController.js";
import type { DocumentControllerWorkbookAdapter } from "./scripting/documentControllerWorkbookAdapter.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize } from "./formatting/selectionSizeGuard.js";
import { registerFindReplaceShortcuts, FindReplaceController } from "./panels/find-replace/index.js";
import { t, tWithVars } from "./i18n/index.js";
import { getOpenFileFilters, isOpenWorkbookPath } from "./file_dialog_filters.js";
import { formatRangeAddress, parseRangeAddress } from "@formula/scripting";
import { normalizeFormulaTextOpt } from "@formula/engine";
import type { CollabSession } from "@formula/collab-session";
import { exportDocumentRangeToCsv } from "./import-export/csv/export.js";
import { startWorkbookSync } from "./tauri/workbookSync";
import { TauriWorkbookBackend } from "./tauri/workbookBackend";
import {
  getTauriDialogOrThrow,
  getTauriEventApiOrThrow,
  getTauriAppGetNameOrNull,
  getTauriAppGetVersionOrNull,
  getTauriInvokeOrNull,
  getTauriWindowHandleOrThrow,
  hasTauri,
  hasTauriInvoke,
  hasTauriWindowApi,
  hasTauriWindowHandleApi,
} from "./tauri/api";
import * as nativeDialogs from "./tauri/nativeDialogs";
import { shellOpen } from "./tauri/shellOpen";
import { setTrayStatus } from "./tauri/trayStatus";
import { FORMULA_RELEASES_URL, installUpdaterUi } from "./tauri/updaterUi";
import { installOpenFileIpc } from "./tauri/openFileIpc";
import { notify } from "./tauri/notifications";
import { registerAppQuitHandlers, requestAppQuit } from "./tauri/appQuit";
import { flushCollabLocalPersistenceBestEffort } from "./tauri/quitFlush";
import { checkForUpdatesFromCommandPalette } from "./tauri/updater.js";
import type { WorkbookInfo } from "@formula/workbook-backend";
import { chartThemeFromWorkbookPalette } from "./charts/theme";
import { parseA1Range, splitSheetQualifier } from "../../../packages/search/index.js";
import type { DesktopPowerQueryService } from "./power-query/service.js";
import { createDesktopDlpContext } from "./dlp/desktopDlp.js";
import { enforceClipboardCopy } from "./dlp/enforceClipboardCopy.js";
import { parseImageCellValue } from "./shared/imageCellValue.js";
import { showInputBox, showQuickPick, showToast } from "./extensions/ui.js";
import { assertExtensionRangeWithinLimits } from "./extensions/rangeSizeGuard.js";
import { createOpenFormatCells } from "./formatting/openFormatCellsCommand.js";
import { formatValueWithNumberFormat } from "./formatting/numberFormat.ts";
import { getStyleNumberFormat } from "./formatting/styleFieldAccess.js";
import { handleCustomSortCommand } from "./sort-filter/openCustomSortDialog.js";
import { parseCollabShareLink, serializeCollabShareLink } from "./sharing/collabLink.js";
import { saveCollabConnectionForWorkbook, loadCollabConnectionForWorkbook } from "./sharing/collabConnectionStore.js";
import { loadCollabToken, preloadCollabTokenFromKeychain, storeCollabToken } from "./sharing/collabTokenStore.js";
import { getWorkbookMutationPermission, READ_ONLY_SHEET_MUTATION_MESSAGE } from "./collab/permissionGuards";
import { showCollabEditRejectedToast } from "./collab/editRejectionToast";
import { registerEncryptionUiCommands } from "./collab/encryption-ui/registerEncryptionUiCommands";
import type { DesktopExtensionHostManager } from "./extensions/extensionHostManager.js";
import { ExtensionPanelBridge } from "./extensions/extensionPanelBridge.js";
import { ContextKeyService } from "./extensions/contextKeys.js";
import { resolveMenuItems } from "./extensions/contextMenus.js";
import { CELL_CONTEXT_MENU_ID, COLUMN_CONTEXT_MENU_ID, CORNER_CONTEXT_MENU_ID, ROW_CONTEXT_MENU_ID } from "./extensions/menuIds.js";
import { buildContextMenuModel } from "./extensions/contextMenuModel.js";
import { getPrimaryCommandKeybindingDisplay, type ContributedKeybinding } from "./extensions/keybindings.js";
import { KeybindingService } from "./extensions/keybindingService.js";
import { isEventWithinKeybindingBarrier, KEYBINDING_BARRIER_ATTRIBUTE, markKeybindingBarrier } from "./keybindingBarrier.js";
import { deriveSelectionContextKeys } from "./extensions/selectionContextKeys.js";
import { installKeyboardContextKeys, KeyboardContextKeyIds } from "./keyboard/installKeyboardContextKeys.js";
import { CommandRegistry } from "./extensions/commandRegistry.js";
import { createCommandPalette, installCommandPaletteRecentsTracking } from "./command-palette/index.js";
import { registerDesktopCommands } from "./commands/registerDesktopCommands.js";
import { registerRibbonAutoFilterCommands } from "./commands/registerRibbonAutoFilterCommands.js";
import { HOME_STYLES_COMMAND_IDS } from "./commands/registerHomeStylesCommands.js";
import {
  FORMAT_FONT_NAME_PRESET_COMMAND_IDS,
  FORMAT_FONT_SIZE_PRESET_COMMAND_IDS,
} from "./commands/registerBuiltinFormatFontCommands.js";
import {
  isSpreadsheetEditingCommandBlockedError,
  SpreadsheetEditingCommandBlockedError,
} from "./commands/spreadsheetEditingCommandBlockedError.js";
import { PAGE_LAYOUT_COMMANDS } from "./commands/registerPageLayoutCommands.js";
import { WORKBENCH_FILE_COMMANDS } from "./commands/registerWorkbenchFileCommands.js";
import { FORMAT_PAINTER_COMMAND_ID } from "./commands/formatPainterCommand.js";
import { DEFAULT_GRID_LIMITS } from "./selection/selection.js";
import type { GridLimits, Range, SelectionState } from "./selection/types";
import { rangeToA1 } from "./selection/a1.js";
import { ContextMenu, type ContextMenuItem } from "./menus/contextMenu.js";
import { buildDrawingContextMenuItems, tryOpenDrawingContextMenuAtClientPoint } from "./mainContextMenuDrawing.js";
import { resolveIsMacPlatform, tagDrawingContextClickPointerDown } from "./drawings/tagDrawingContextClick.js";
import { getPasteSpecialMenuItems } from "./clipboard/pasteSpecial.js";
import {
  WorkbookSheetStore,
  validateSheetName,
  type SheetVisibility,
  type TabColor,
} from "./sheets/workbookSheetStore";
import {
  CollabWorkbookSheetStore,
  computeCollabSheetsKey,
  listSheetsFromCollabSession,
  type CollabSheetsKeyRef,
} from "./sheets/collabWorkbookSheetStore";
import { tryInsertCollabSheet } from "./sheets/collabSheetMutations";
import { startSheetStoreDocumentSync } from "./sheets/sheetStoreDocumentSync";
import { createAddSheetCommand, createDeleteActiveSheetCommand } from "./sheets/sheetCommands";
import {
  NUMBER_FORMATS,
  toggleBold,
  toggleItalic,
  toggleStrikethrough,
  toggleUnderline,
  toggleWrap,
  type CellRange,
} from "./formatting/toolbar.js";
import {
  applyFormatAsTablePreset,
  FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS,
  estimateFormatAsTableBandedRowOps,
} from "./formatting/formatAsTablePresets.js";
import { computeFilterHiddenRows, computeUniqueFilterValues, RibbonAutoFilterStore } from "./sort-filter/ribbonAutoFilter.js";
import type { CellRange as PrintCellRange, PageSetup } from "./print/index.js";
import { AutoFilterDropdown, type TableViewRow } from "./table/index.js";
import {
  getDefaultSeedStoreStorage,
  readContributedPanelsSeedStore,
  removeSeedPanelsForExtension,
  seedPanelRegistryFromContributedPanelsSeedStore,
  setSeedPanelsForExtension,
} from "./extensions/contributedPanelsSeedStore.js";
import { builtinKeybindings as builtinKeybindingsCatalog } from "./commands/builtinKeybindings.js";
import { DlpViolationError } from "../../../packages/security/dlp/src/errors.js";

import sampleHelloManifest from "../../../extensions/sample-hello/package.json";
import { purgeLegacyDesktopLLMSettings } from "./ai/llm/desktopLLMClient.js";
import { markStartupFirstRender, markStartupTimeToInteractive } from "./tauri/startupMetrics.js";
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
import { mergeEmbeddedCellImagesIntoSnapshot } from "./workbook/load/embeddedCellImages.js";
import { warnIfWorkbookLoadTruncated, type WorkbookLoadTruncation } from "./workbook/load/truncationWarning.js";
import { buildImportedDrawingLayerSnapshotAdditions } from "./workbook/load/hydrateImportedDrawings.js";
import { hydrateSheetBackgroundImagesFromBackend } from "./workbook/load/hydrateSheetBackgroundImages.js";
import { hiddenColsFromImportedColProperties, sheetColWidthsFromViewOrImportedColProperties } from "./workbook/load/importedColProperties.js";
import {
  mergeFormattingIntoSnapshot,
  type CellFormatClampBounds,
  type SheetFormattingSnapshot,
  type SnapshotCell,
} from "./workbook/mergeFormattingIntoSnapshot.js";
import { coerceSavePathToXlsx, getWorkbookFileMetadataFromWorkbookInfo } from "./workbook/workbookFileMetadata.js";

// ---------------------------------------------------------------------------
// Startup instrumentation + two-phase boot
// ---------------------------------------------------------------------------
const STARTUP_DEBUG =
  ((import.meta as any).env?.VITE_STARTUP_DEBUG ?? "") === "1" ||
  ((import.meta as any).env?.VITE_STARTUP_DEBUG ?? "") === "true" ||
  (((globalThis as any).process?.env as Record<string, unknown> | undefined)?.VITE_STARTUP_DEBUG ?? "") === "1";

function startupMark(name: string): void {
  try {
    globalThis.performance?.mark?.(name);
  } catch {
    // ignore
  }
  if (STARTUP_DEBUG) {
    // eslint-disable-next-line no-console
    console.log(`[formula][desktop][startup] ${name}`);
  }
}

async function startupImport<T>(name: string, importer: () => Promise<T>): Promise<T> {
  startupMark(`formula:desktop:import:${name}:start`);
  try {
    const mod = await importer();
    startupMark(`formula:desktop:import:${name}:end`);
    return mod;
  } catch (err) {
    startupMark(`formula:desktop:import:${name}:error`);
    throw err;
  }
}

startupMark("formula:desktop:boot:start");

// `requestAnimationFrame` fires before paint; the double-rAF pattern marks the first paint
// *after* our entry module yields back to the browser.
try {
  if (typeof window !== "undefined" && typeof window.requestAnimationFrame === "function") {
    window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        startupMark("formula:desktop:boot:first-render");
      });
    });
  }
} catch {
  // ignore
}

type YjsModule = typeof import("yjs");
let yjsModulePromise: Promise<YjsModule> | null = null;
function importYjs(): Promise<YjsModule> {
  if (!yjsModulePromise) {
    yjsModulePromise = startupImport("yjs", () => import("yjs"));
  }
  return yjsModulePromise;
}

type ExtensionHostManagerModule = typeof import("./extensions/extensionHostManager.js");
let extensionHostManagerModulePromise: Promise<ExtensionHostManagerModule> | null = null;
function importExtensionHostManagerModule(): Promise<ExtensionHostManagerModule> {
  if (!extensionHostManagerModulePromise) {
    extensionHostManagerModulePromise = startupImport("extension-host-manager", () => import("./extensions/extensionHostManager.js"));
  }
  return extensionHostManagerModulePromise;
}

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

/**
 * SharedArrayBuffer requires cross-origin isolation (COOP/COEP). When we ship a
 * packaged Tauri build without it, Pyodide cannot use its Worker backend.
 *
 * Vite dev/preview servers configure COOP/COEP headers; the packaged app relies
 * on Tauri's asset protocol configuration. Fail loudly (but non-fatally) in
 * production Tauri builds so this doesn't ship silently.
 */
function warnIfMissingCrossOriginIsolationInTauriProd(): void {
  const isTauri = hasTauriInvoke();
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

function reportLazyImportFailure(featureLabel: string, err: unknown): void {
  console.error(`[formula][desktop] Failed to load ${featureLabel}:`, err);
  try {
    showToast(`Failed to load ${featureLabel}. Please reload the app.`, "error");
  } catch {
    // ignore (toast root missing / non-DOM contexts)
  }
}

const loadPanelBodyRendererModule = createLazyImport(() => import("./panels/panelBodyRenderer.js"), {
  label: "Panels",
  onError: (err) => reportLazyImportFailure("Panels", err),
});

const loadScriptEditorPanelModule = createLazyImport(() => import("./panels/script-editor/index.js"), {
  label: "Script Editor",
  onError: (err) => reportLazyImportFailure("Script Editor", err),
});

const loadScriptingWorkbookAdapterModule = createLazyImport(() => import("./scripting/documentControllerWorkbookAdapter.js"), {
  label: "Scripting",
  onError: (err) => reportLazyImportFailure("Scripting", err),
});

const loadMacroRecorderModule = createLazyImport(() => import("./macro-recorder/index.js"), {
  label: "Macro recorder",
  onError: (err) => reportLazyImportFailure("Macro recorder", err),
});

const loadMacrosModule = createLazyImport(() => import("./macros/index.js"), {
  label: "Macros",
  onError: (err) => reportLazyImportFailure("Macros", err),
});

const loadVbaEventMacrosModule = createLazyImport(() => import("./macros/event_macros.js"), {
  label: "VBA event macros",
  onError: (err) => reportLazyImportFailure("VBA event macros", err),
});

const loadPowerQueryServiceModule = createLazyImport(() => import("./power-query/service.js"), {
  label: "Power Query",
  onError: (err) => reportLazyImportFailure("Power Query", err),
});

const loadPowerQueryRefreshStateStoreModule = createLazyImport(() => import("./power-query/refreshStateStore.js"), {
  label: "Power Query refresh state",
  onError: (err) => reportLazyImportFailure("Power Query refresh state", err),
});

const loadPowerQueryTableSignaturesModule = createLazyImport(() => import("./power-query/tableSignatures.js"), {
  label: "Power Query signatures",
  onError: (err) => reportLazyImportFailure("Power Query signatures", err),
});

const loadPowerQueryOauthBrokerModule = createLazyImport(() => import("./power-query/oauthBroker.js"), {
  label: "Power Query OAuth",
  onError: (err) => reportLazyImportFailure("Power Query OAuth", err),
});

const loadOrganizeSheetsDialogModule = createLazyImport(() => import("./sheets/OrganizeSheetsDialog"), {
  label: "Organize Sheets",
  onError: (err) => reportLazyImportFailure("Organize Sheets", err),
});

const loadSecondaryGridViewModule = createLazyImport(
  () => startupImport("secondary-grid-view", () => import("./grid/splitView/secondaryGridView.js")),
  {
    label: "Split View",
    onError: (err) => reportLazyImportFailure("Split View", err),
  },
);

const loadPageSetupDialogModule = createLazyImport(() => import("./print/PageSetupDialog.js"), {
  label: "Page Setup",
  onError: (err) => reportLazyImportFailure("Printing", err),
});

const loadPrintPreviewDialogModule = createLazyImport(() => import("./print/PrintPreviewDialog.js"), {
  label: "Print Preview",
  onError: (err) => reportLazyImportFailure("Printing", err),
});

const loadGoalSeekDialogModule = createLazyImport(() => import("./panels/what-if/GoalSeekDialog.js"), {
  label: "Goal Seek",
  onError: (err) => reportLazyImportFailure("Goal Seek", err),
});

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

function installExternalLinkInterceptor(): void {
  if (typeof document === "undefined") return;

  const handler = (event: MouseEvent) => {
      if (event.defaultPrevented) return;

      // `event.target` can be a `Text` node (e.g. clicking the text inside an <a>).
      // Normalize to an Element so we can use `.closest(...)`.
      const rawTarget = event.target as unknown;
      const target =
        rawTarget instanceof Element
          ? rawTarget
          : rawTarget && typeof (rawTarget as any).parentElement === "object"
            ? ((rawTarget as any).parentElement as Element | null)
            : null;
      if (!target || typeof (target as any).closest !== "function") return;
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
      const isTauri = hasTauri();
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
type SecondaryGridViewModule = typeof import("./grid/splitView/secondaryGridView.js");
let secondaryGridViewModule: SecondaryGridViewModule | null = null;
let secondaryGridViewInitToken = 0;
let secondaryGridViewInitPromise: Promise<void> | null = null;

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
  // Emit a DOM event so non-extension UI surfaces (e.g. dialogs) can react to sheet activation
  // without wiring directly into SpreadsheetApp internals.
  try {
    window.dispatchEvent(new CustomEvent("formula:sheet-activated", { detail: { sheetId: id } }));
  } catch {
    // ignore
  }
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
// File commands can be triggered from many entry points (ribbon, titlebar, native menu events).
// Some of those paths intentionally fire-and-forget command promises. Keep file operations
// serialized so "Save" cannot run while an "Open" is still loading and wiring up workbook sync.
let pendingFileCommand: Promise<void> = Promise.resolve();
let queuedInvoke: TauriInvoke | null = null;
let workbookSync: ReturnType<typeof startWorkbookSync> | null = null;
let rerenderLayout: (() => void) | null = null;
let vbaEventMacros: VbaEventMacrosHandle | null = null;
let ribbonLayoutController: LayoutController | null = null;
let activeMacroRecorder: MacroRecorder | null = null;
type MacrosPanelFocusTarget = "runner-select" | "runner-run" | "runner-trust-center" | "recorder-start" | "recorder-stop";
let pendingMacrosPanelFocus: MacrosPanelFocusTarget | null = null;
let ensureExtensionsLoadedRef: (() => Promise<void>) | null = null;
let ensureMacroRecorderRef: (() => Promise<MacroRecorder | null>) | null = null;
let syncContributedCommandsRef: (() => void) | null = null;
let syncContributedPanelsRef: (() => void) | null = null;
let updateKeybindingsRef: (() => void) | null = null;
let focusAfterSheetNavigationFromCommandRef: (() => void) | null = null;

function openRibbonPanel(panelId: string): void {
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
}

function focusScriptEditorPanel(): void {
  // Script editor is mounted via panel rendering; wait a frame (or two) so the
  // DOM nodes exist before focusing.
  if (typeof document === "undefined") return;
  if (typeof requestAnimationFrame !== "function") return;
  requestAnimationFrame(() =>
    requestAnimationFrame(() => {
      const el =
        document.querySelector<HTMLElement>('[data-testid="script-editor-code"]') ??
        document.querySelector<HTMLElement>('[data-testid="script-editor-run"]');
      try {
        el?.focus();
      } catch {
        // Best-effort: ignore focus errors (e.g. element not focusable in headless environments).
      }
    }),
  );
}

function focusVbaMigratePanel(): void {
  // The VBA migrate panel is a React mount; give it a couple of frames to render.
  if (typeof document === "undefined") return;
  if (typeof requestAnimationFrame !== "function") return;
  requestAnimationFrame(() =>
    requestAnimationFrame(() => {
      const el =
        document.querySelector<HTMLElement>('[data-testid="vba-entrypoint"]') ??
        document.querySelector<HTMLElement>('button[data-testid^="vba-module-"]') ??
        document.querySelector<HTMLElement>('[data-testid="vba-refresh"]');
      try {
        el?.focus();
      } catch {
        // Best-effort: ignore focus errors (e.g. element not focusable in headless environments).
      }
    }),
  );
}

// --- AutoSave ---------------------------------------------------------------
// Persisted globally (not per-workbook) since the ribbon toggle is global.
const AUTO_SAVE_STORAGE_KEY = "formula.desktop.autoSave.enabled";
// Debounce time after the most recent change before attempting an autosave.
// Keep this relatively short so background saves feel responsive, but long enough
// to avoid saving on every keystroke.
const AUTO_SAVE_DEBOUNCE_MS = 4_000;

function readAutoSaveEnabledFromStorage(): boolean {
  try {
    const raw = globalThis.localStorage?.getItem(AUTO_SAVE_STORAGE_KEY);
    return raw === "true" || raw === "1";
  } catch {
    return false;
  }
}

function writeAutoSaveEnabledToStorage(enabled: boolean): void {
  try {
    globalThis.localStorage?.setItem(AUTO_SAVE_STORAGE_KEY, enabled ? "true" : "false");
  } catch {
    // Ignore storage errors (disabled/quota/etc).
  }
}

let autoSaveEnabled = readAutoSaveEnabledFromStorage();

let autoSaveTimer: number | null = null;
let autoSaveLastChangeAt = 0;
let autoSaveSaving = false;
let autoSaveNeedsSaveAfterFlight = false;
let autoSaveNeedsSaveAfterEditing = false;
// When true, run an autosave even if the DocumentController is currently "clean".
// This protects against edge cases where a save marks the document saved while edits
// are still queued behind that save.
let autoSaveForceNextSave = false;

function isTauriInvokeAvailable(): boolean {
  return hasTauriInvoke();
}

function clearAutoSaveTimer(): void {
  if (autoSaveTimer == null) return;
  globalThis.clearTimeout(autoSaveTimer);
  autoSaveTimer = null;
}

function scheduleAutoSaveFromLastChange(): void {
  clearAutoSaveTimer();
  if (!autoSaveEnabled) return;
  // When not running under Tauri (e.g. web builds), the toggle is forced off and
  // background saves are never scheduled.
  if (!isTauriInvokeAvailable()) return;
  if (!activeWorkbook) return;

  const now = Date.now();
  const delay = Math.max(0, autoSaveLastChangeAt + AUTO_SAVE_DEBOUNCE_MS - now);
  autoSaveTimer = globalThis.setTimeout(() => {
    autoSaveTimer = null;
    void attemptAutoSave();
  }, delay) as unknown as number;
}

function noteAutoSaveChange(): void {
  if (!autoSaveEnabled) return;
  if (!isTauriInvokeAvailable()) return;
  autoSaveLastChangeAt = Date.now();
  // If a save is already in-flight, coalesce into one additional save after it completes.
  if (autoSaveSaving) autoSaveNeedsSaveAfterFlight = true;
  scheduleAutoSaveFromLastChange();
}

async function attemptAutoSave(): Promise<void> {
  if (!autoSaveEnabled) return;
  if (!isTauriInvokeAvailable()) return;
  if (!tauriBackend) return;
  if (!activeWorkbook) return;

  const doc = app.getDocument();

  // Only autosave when we actually have something to persist, unless we're forcing a save
  // due to a prior in-flight save that may have been superseded by later edits.
  if (!doc.isDirty && !autoSaveForceNextSave) return;

  // Never interrupt an in-progress edit: wait until editing ends.
  if (isSpreadsheetEditing()) {
    autoSaveNeedsSaveAfterEditing = true;
    return;
  }

  if (autoSaveSaving) {
    autoSaveNeedsSaveAfterFlight = true;
    autoSaveForceNextSave = true;
    return;
  }

  // If we don't have a path yet, autosave can't write to disk. Prompt for Save As.
  if (!activeWorkbook.path) {
    try {
      await handleSaveAs({ throwOnCancel: true });
      autoSaveForceNextSave = false;
      return;
    } catch (err) {
      const name = (err as any)?.name;
      if (name !== "AbortError") {
        console.error("AutoSave Save As failed:", err);
        showToast(`AutoSave failed: ${String(err)}`, "error");
      }
      // If the user cancels the Save As dialog, AutoSave must revert to OFF.
      autoSaveEnabled = false;
      writeAutoSaveEnabledToStorage(false);
      clearAutoSaveTimer();
      scheduleRibbonSelectionFormatStateUpdate();
      return;
    }
  }

  autoSaveSaving = true;
  try {
    // Ensure any pending microtask-batched workbook edits are flushed before saving.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await drainBackendSync();

    // Another save (manual or AutoSave) could have completed while we were waiting for the
    // backend sync queue to drain. Re-check whether we still have anything to persist
    // before issuing a new save_workbook.
    if (!autoSaveEnabled) return;
    if (!doc.isDirty && !autoSaveForceNextSave) return;

    if (workbookSync) {
      await workbookSync.markSaved();
    } else if (queuedInvoke) {
      await queuedInvoke("save_workbook", {});
      doc.markSaved();
    } else {
      await tauriBackend.saveWorkbook();
      doc.markSaved();
    }

    autoSaveForceNextSave = false;
  } catch (err) {
    console.error("AutoSave failed:", err);
    showToast(`AutoSave failed: ${String(err)}`, "error");
    // Keep AutoSave enabled on failure.
  } finally {
    autoSaveSaving = false;
  }

  if (!autoSaveEnabled) return;
  if (autoSaveNeedsSaveAfterEditing) {
    autoSaveNeedsSaveAfterEditing = false;
    autoSaveForceNextSave = true;
    scheduleAutoSaveFromLastChange();
    return;
  }
  if (autoSaveNeedsSaveAfterFlight) {
    autoSaveNeedsSaveAfterFlight = false;
    autoSaveForceNextSave = true;
    scheduleAutoSaveFromLastChange();
  }
}

async function setAutoSaveEnabledFromUi(nextEnabled: boolean): Promise<void> {
  const invokeAvailable = isTauriInvokeAvailable();
  if (!invokeAvailable) {
    showDesktopOnlyToast("AutoSave is available in the desktop app.");
    autoSaveEnabled = false;
    writeAutoSaveEnabledToStorage(false);
    clearAutoSaveTimer();
    scheduleRibbonSelectionFormatStateUpdate();
    return;
  }

  if (!nextEnabled) {
    autoSaveEnabled = false;
    writeAutoSaveEnabledToStorage(false);
    clearAutoSaveTimer();
    scheduleRibbonSelectionFormatStateUpdate();
    return;
  }

  if (!tauriBackend || !activeWorkbook) {
    autoSaveEnabled = false;
    writeAutoSaveEnabledToStorage(false);
    clearAutoSaveTimer();
    scheduleRibbonSelectionFormatStateUpdate();
    showToast("AutoSave is not available (workbook backend not ready).", "error");
    return;
  }

  if (!activeWorkbook.path) {
    try {
      await handleSaveAs({ throwOnCancel: true });
    } catch (err) {
      const name = (err as any)?.name;
      if (name !== "AbortError") {
        console.error("Failed to enable AutoSave (Save As failed):", err);
        showToast(`Failed to enable AutoSave: ${String(err)}`, "error");
      }
      autoSaveEnabled = false;
      writeAutoSaveEnabledToStorage(false);
      clearAutoSaveTimer();
      scheduleRibbonSelectionFormatStateUpdate();
      return;
    }

    if (!activeWorkbook.path) {
      autoSaveEnabled = false;
      writeAutoSaveEnabledToStorage(false);
      clearAutoSaveTimer();
      scheduleRibbonSelectionFormatStateUpdate();
      return;
    }
  }

  autoSaveEnabled = true;
  writeAutoSaveEnabledToStorage(true);
  scheduleRibbonSelectionFormatStateUpdate();

  // If the document already has unsaved edits, schedule an autosave now.
  if (app.getDocument().isDirty) {
    autoSaveForceNextSave = true;
    autoSaveLastChangeAt = Date.now();
    scheduleAutoSaveFromLastChange();
  }
}

function toggleDockPanel(panelId: string): void {
  const controller = ribbonLayoutController;
  if (!controller) return;
  const layout = controller.layout as any;
  const placement = getPanelPlacement(layout, panelId) as any;
  if (placement.kind === "closed") {
    controller.openPanel(panelId);
    return;
  }
  // Dock zones can be collapsed. Treat a collapsed docked panel as "closed" so status-bar toggles
  // restore the dock instead of removing the panel.
  if (placement.kind === "docked" && layout?.docks?.[placement.side]?.collapsed) {
    controller.openPanel(panelId);
    return;
  }
  // Floating panels can be minimized. Treat a minimized panel as "closed" so status-bar toggles
  // restore it instead of closing the panel.
  if (placement.kind === "floating" && layout?.floating?.[panelId]?.minimized) {
    controller.setFloatingPanelMinimized(panelId, false);
    return;
  }
  controller.closePanel(panelId);
}
let handleCloseRequestForRibbon: ((opts: { quit: boolean }) => Promise<void>) | null = null;

function installCollabStatusIndicator(app: SpreadsheetApp, element: HTMLElement): void {
  const abortController = new AbortController();

  const cleanup = (): void => {
    if (abortController.signal.aborted) return;
    abortController.abort();
  };

  window.addEventListener("unload", cleanup, { once: true });

  // Wrap destroy so collab listeners detach in tests / fast-refresh scenarios.
  if (typeof app.destroy === "function") {
    const originalDestroy = app.destroy.bind(app) as () => void;
    app.destroy = () => {
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

  const getSession = (): unknown | null => app.getCollabSession();

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

const ribbonRoot = document.getElementById("ribbon");
if (!ribbonRoot) {
  throw new Error("Missing #ribbon container");
}
const ribbonRootEl = ribbonRoot;
const ribbonReactRoot = document.getElementById("ribbon-react-root");
if (!ribbonReactRoot) {
  throw new Error("Missing #ribbon-react-root container");
}

const formulaBarRoot = document.getElementById("formula-bar");
if (!formulaBarRoot) {
  throw new Error("Missing #formula-bar container");
}
const formulaBarRootEl = formulaBarRoot;

const sheetTabsRoot = document.getElementById("sheet-tabs");
if (!sheetTabsRoot) {
  throw new Error("Missing #sheet-tabs container");
}
const sheetTabsRootEl = sheetTabsRoot;
// The shell uses `.sheet-bar` (with an inner `.sheet-tabs` strip) for styling.
// Normalize older HTML scaffolds that used `.sheet-tabs` on the container itself.
sheetTabsRootEl.classList.add("sheet-bar");
sheetTabsRootEl.classList.remove("sheet-tabs");

// Sheet tab mounting is triggered by `SpreadsheetApp.onEditStateChange`, which invokes listeners
// immediately with the current edit state. Declare the React root eagerly so early subscribers
// (including Playwright bootstrap code) don't trip the temporal dead zone for `let` bindings.
let sheetTabsReactRoot: ReturnType<typeof createRoot> | null = null;

const statusBarRoot = document.querySelector<HTMLElement>(".statusbar");
if (!statusBarRoot) {
  throw new Error("Missing .statusbar container");
}
const statusBarRootEl = statusBarRoot;

const statusMode = document.querySelector<HTMLElement>('[data-testid="status-mode"]');
const activeCell = document.querySelector<HTMLElement>('[data-testid="active-cell"]');
const selectionRange = document.querySelector<HTMLElement>('[data-testid="selection-range"]');
const activeValue = document.querySelector<HTMLElement>('[data-testid="active-value"]');
const collabStatus = document.querySelector<HTMLElement>('[data-testid="collab-status"]');
const readOnlyIndicator = document.querySelector<HTMLElement>('[data-testid="read-only-indicator"]');
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

// Best-effort: hydrate any persisted collaboration token from the desktop secure store
// into session-scoped storage before SpreadsheetApp resolves its collab connection options.
//
// We preload for:
// - collab URLs (`?collab=1&wsUrl=...&docId=...`) so reloading the same collab doc works after restart
// - stored connections (localStorage metadata) so "open workbook → auto-reconnect" works after restart
try {
  const url = new URL(window.location.href);
  const params = url.searchParams;
  const collabEnabled = params.get("collab");
  const fromUrl =
    collabEnabled === "1" || collabEnabled === "true"
      ? {
          wsUrl: String(params.get("collabWsUrl") ?? params.get("wsUrl") ?? "").trim(),
          docId: String(params.get("collabDocId") ?? params.get("docId") ?? "").trim(),
        }
      : null;

  const stored = loadCollabConnectionForWorkbook({ workbookKey: workbookId });
  const target = fromUrl?.wsUrl && fromUrl.docId ? fromUrl : stored ? { wsUrl: stored.wsUrl, docId: stored.docId } : null;
  if (target) await preloadCollabTokenFromKeychain(target);
} catch {
  // ignore (storage disabled / Tauri invoke unavailable / etc)
}

const legacyGridLimits = (() => {
  if (resolveDesktopGridMode() !== "legacy") return undefined;
  // `getWorkbookLoadLimits` also resolves snapshot chunking controls; strip it down to
  // `GridLimits` so SpreadsheetApp doesn't accidentally retain unrelated fields.
  const { maxRows, maxCols } = getWorkbookLoadLimits();
  return { maxRows, maxCols };
})();
const app = new SpreadsheetApp(
  gridRoot,
  { activeCell, selectionRange, activeValue, selectionSum, selectionAverage, selectionCount, readOnlyIndicator },
  {
    formulaBar: formulaBarRoot,
    workbookId,
    sheetNameResolver,
    enableDrawingInteractions: resolveEnableDrawingInteractions(),
    // Legacy renderer uses a smaller grid by default for performance reasons. When users
    // explicitly raise the workbook load limits (via query params/env), align the legacy
    // grid limits so the loaded data is reachable.
    limits: legacyGridLimits,
  },
);

// Startup performance instrumentation: "first meaningful render" / grid visible.
// Best-effort, no-op for web builds.
void markStartupFirstRender().catch(() => {});

// Expose a small API for Playwright assertions early so e2e can still attach even if
// optional desktop integrations (e.g. Tauri host wiring) fail during startup.
window.__formulaApp = app;
(app as unknown as { getWorkbookSheetStore: () => unknown }).getWorkbookSheetStore = () => workbookSheetStore;
window.__workbookSheetStore = workbookSheetStore;

// Ribbon AutoFilter MVP: keep filter state in-memory for the current session.
const ribbonAutoFilterStore = new RibbonAutoFilterStore();

// DocumentController creates sheets lazily whenever code reads/writes a new sheet id (via `getCell`).
// When undo/redo removes the currently active sheet id, downstream listeners can accidentally
// "recreate" the sheet by reading from it during the same `document.on("change")` dispatch.
//
// Keep the active sheet id valid *before* other `document.on("change")` listeners run so undo/redo
// of sheet add/delete behaves predictably (and stays in sync with the sheet metadata store).
app.getDocument().on("change", (payload: any) => {
  const source = typeof payload?.source === "string" ? payload.source : "";
  if (source !== "undo" && source !== "redo" && source !== "applyState") return;

  const activeId = app.getCurrentSheetId();
  if (!activeId) return;

  const doc = app.getDocument();
  const docSheetIds = doc.getSheetIds();
  if (docSheetIds.length === 0) return;
  if (docSheetIds.includes(activeId)) return;

  const docIdSet = new Set(docSheetIds);
  // Mirror Excel: when the active sheet disappears (undo/redo/applyState), prefer the next visible
  // sheet to the right, otherwise the previous visible sheet. Use the sheet-store ordering snapshot
  // (which is typically still pre-change at this point) so we choose a stable adjacent candidate.
  const sheetsSnapshot = workbookSheetStore.listAll();
  const preferred = pickAdjacentVisibleSheetId(sheetsSnapshot, activeId);
  const fallback =
    (preferred && docIdSet.has(preferred) ? preferred : null) ??
    workbookSheetStore.listVisible().map((s) => s.id).find((id) => docIdSet.has(id)) ??
    sheetsSnapshot.map((s) => s.id).find((id) => docIdSet.has(id)) ??
    docSheetIds[0] ??
    null;
  if (!fallback || fallback === activeId) return;

  app.activateSheet(fallback);
});

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
  const raw = app.getGridLimits();
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

function applyFormattingToSelection(
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
  options: { forceBatch?: boolean; allowReadOnlyBandSelection?: boolean } = {},
): void {
  // Match SpreadsheetApp guards: formatting commands should never mutate the sheet while the user
  // is actively editing (cell editor / formula bar / inline edit).
  if (isSpreadsheetEditing()) return;
  const isReadOnly = app.isReadOnly?.() === true;
  // By default, read-only collab roles (viewer/commenter) are allowed to apply formatting *only*
  // when the selection is a full row/column/sheet band (i.e. "formatting defaults"). These
  // mutations must remain local-only; collaboration binders are responsible for preventing
  // them from being persisted into shared Yjs state.
  //
  // Some commands routed through this helper (e.g. "Clear All") mutate cell values and must
  // remain blocked in read-only even for band selections. Callers can opt out via
  // `allowReadOnlyBandSelection: false`.
  const allowReadOnlyBandSelection = options.allowReadOnlyBandSelection !== false;

  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  const selection = app.getSelectionRanges();
  const limits = getGridLimitsForFormatting();
  const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });

  if (isReadOnly) {
    if (!allowReadOnlyBandSelection) {
      showToast("Read-only: you don't have permission to edit cell contents.", "warning");
      return;
    }
  }

  if (!decision.allowed) {
    showToast("Selection is too large to format. Try selecting fewer cells or an entire row/column.", "warning");
    return;
  }

  if (isReadOnly && !decision.allRangesBand) {
    showCollabEditRejectedToast([{ rejectionKind: "formatDefaults", rejectionReason: "permission" }]);
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

// --- Format Painter -----------------------------------------------------------
//
// Excel-style one-shot Format Painter:
// 1) Capture effective formatting from the active cell.
// 2) Arm the painter and wait for the next selection change.
// 3) Apply the captured style patch to the destination selection once.
// 4) Automatically disarm, or allow Escape/workbook/sheet changes to cancel.

type FormatPainterState = {
  doc: DocumentController;
  sourceWorkbookId: string;
  sourceSheetId: string;
  sourceSelectionKey: string;
  capturedFormat: Record<string, any>;
  applyTimeoutId: number | null;
  pendingSelectionKey: string | null;
};

const FORMAT_PAINTER_APPLY_DEBOUNCE_MS = 100;
let formatPainterState: FormatPainterState | null = null;

function cloneStylePatch(value: unknown): Record<string, any> {
  if (!value || typeof value !== "object") return {};
  try {
    if (typeof (globalThis as any).structuredClone === "function") {
      return (globalThis as any).structuredClone(value);
    }
  } catch {
    // Fall back below.
  }
  try {
    return JSON.parse(JSON.stringify(value)) as Record<string, any>;
  } catch {
    return {};
  }
}

function formatPainterSelectionKey(selection?: SelectionState): string {
  const sheetId = app.getCurrentSheetId();
  const active = selection?.active ?? app.getActiveCell();
  const ranges = selection?.ranges ?? app.getSelectionRanges();
  const parts: string[] = [`sheet=${sheetId}`, `active=${active.row},${active.col}`];
  for (const raw of ranges) {
    const r = normalizeSelectionRange(raw);
    parts.push(`range=${r.startRow},${r.startCol},${r.endRow},${r.endCol}`);
  }
  return parts.join("|");
}

function shouldRestoreFocusAfterArmingFormatPainter(): boolean {
  try {
    if (typeof document === "undefined") return true;
    const active = document.activeElement as HTMLElement | null;
    if (!active) return true;

    // Avoid stealing focus from modal overlays (command palette, dialogs, etc.).
    if (typeof active.closest === "function" && active.closest(`[${KEYBINDING_BARRIER_ATTRIBUTE}]`)) {
      return false;
    }

    const tag = active.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || active.isContentEditable) {
      return false;
    }
  } catch {
    // Best-effort: default to restoring focus.
  }

  return true;
}

function disarmFormatPainter(): void {
  const state = formatPainterState;
  if (!state) return;
  if (state.applyTimeoutId != null) {
    try {
      window.clearTimeout(state.applyTimeoutId);
    } catch {
      // ignore
    }
  }
  formatPainterState = null;
  scheduleRibbonSelectionFormatStateUpdate();
}

function armFormatPainter(): void {
  if (formatPainterState) {
    // Toggle off (Excel uses double-click to lock; we only support one-shot).
    disarmFormatPainter();
    try {
      showToast("Format Painter cancelled");
    } catch {
      // ignore (toast root missing in non-UI test environments)
    }
    return;
  }

  if (isSpreadsheetEditing()) {
    try {
      showToast("Finish editing to use Format Painter");
    } catch {
      // ignore
    }
    return;
  }

  const isReadOnly = app.isReadOnly?.() === true;

  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  const activeCell = app.getActiveCell();

  let captured: unknown = {};
  try {
    captured = doc.getCellFormat(sheetId, activeCell);
  } catch {
    captured = {};
  }

  formatPainterState = {
    doc,
    sourceWorkbookId: activePanelWorkbookId,
    sourceSheetId: sheetId,
    sourceSelectionKey: formatPainterSelectionKey(),
    capturedFormat: cloneStylePatch(captured),
    applyTimeoutId: null,
    pendingSelectionKey: null,
  };

  scheduleRibbonSelectionFormatStateUpdate();
  try {
    showToast(isReadOnly ? "Format Painter: select an entire row, column, or sheet" : "Format Painter: select destination cells");
  } catch {
    // ignore
  }

  // Restore grid focus so keyboard selection continues to work after clicking the ribbon.
  if (shouldRestoreFocusAfterArmingFormatPainter()) {
    queueMicrotask(() => app.focus());
  }
}

function handleFormatPainterSelectionChange(selection: SelectionState): void {
  const state = formatPainterState;
  if (!state) return;

  // Cancel if the active workbook changed (workbook open/version restore).
  if (activePanelWorkbookId !== state.sourceWorkbookId) {
    disarmFormatPainter();
    return;
  }

  const sheetId = app.getCurrentSheetId();
  if (sheetId !== state.sourceSheetId) {
    disarmFormatPainter();
    return;
  }

  // Defensive: SpreadsheetApp.restoreDocumentState can swap the underlying DocumentController.
  if (app.getDocument() !== state.doc) {
    disarmFormatPainter();
    return;
  }

  if (isSpreadsheetEditing()) {
    // Keep armed, but never apply while editing.
    return;
  }

  const key = formatPainterSelectionKey(selection);
  if (key === state.sourceSelectionKey) return;

  state.pendingSelectionKey = key;
  if (state.applyTimeoutId != null) {
    window.clearTimeout(state.applyTimeoutId);
  }

  state.applyTimeoutId = window.setTimeout(() => {
    const liveState = formatPainterState;
    if (!liveState) return;
    liveState.applyTimeoutId = null;

    if (isSpreadsheetEditing()) return;
    if (activePanelWorkbookId !== liveState.sourceWorkbookId) {
      disarmFormatPainter();
      return;
    }
    if (app.getCurrentSheetId() !== liveState.sourceSheetId) {
      disarmFormatPainter();
      return;
    }
    if (app.getDocument() !== liveState.doc) {
      disarmFormatPainter();
      return;
    }

    const currentKey = formatPainterSelectionKey();
    if (currentKey !== liveState.pendingSelectionKey) return;

    const format = liveState.capturedFormat;
    applyFormattingToSelection("Format Painter", (doc, sheetId, ranges) => {
      let applied = true;
      for (const range of ranges) {
        const ok = doc.setRangeFormat(sheetId, range, format, { label: "Format Painter" });
        if (ok === false) applied = false;
      }
      return applied;
    });

    disarmFormatPainter();
  }, FORMAT_PAINTER_APPLY_DEBOUNCE_MS);
}

function activeCellNumberFormat(): string | null {
  const sheetId = app.getCurrentSheetId();
  const cell = app.getActiveCell();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const docAny = app.getDocument() as any;
  const style = docAny.getCellFormat?.(sheetId, cell);
  return getStyleNumberFormat(style);
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
if (collabStatus) installCollabStatusIndicator(app, collabStatus);
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();

app.focus();

// Split-view secondary pane mounts its own CellEditorOverlay. SpreadsheetApp's `isEditing()`
// only reflects its *primary* editor/formula bar/inline edit state, so track the secondary
// editor separately for global UI state (status bar, shortcut gating, etc).
let splitViewSecondaryIsEditing = false;
let recomputeKeyboardContextKeys: (() => void) | null = null;

const isSpreadsheetEditing = (): boolean => app.isEditing() || splitViewSecondaryIsEditing;

// Some UI surfaces (panels, tab strip, etc) need to react to a unified "spreadsheet editing"
// signal that includes both:
// - SpreadsheetApp's primary edit state (cell editor / formula bar / inline edit), and
// - split-view secondary pane edit state (SecondaryGridView).
//
// SpreadsheetApp exposes `onEditStateChange`, but it only tracks the primary editor. Publish a
// desktop-shell-owned window event so panels can disable unsafe actions (Excel-style) without
// needing a direct reference to split-view internals.
const SPREADSHEET_EDITING_CHANGED_EVENT = "formula:spreadsheet-editing-changed";
let lastEmittedSpreadsheetEditing = isSpreadsheetEditing();
// Expose the current state for panels that mount while editing is already active.
(globalThis as any).__formulaSpreadsheetIsEditing = lastEmittedSpreadsheetEditing;

function emitSpreadsheetEditingChanged(): void {
  const next = isSpreadsheetEditing();
  if (next === lastEmittedSpreadsheetEditing) return;
  lastEmittedSpreadsheetEditing = next;
  (globalThis as any).__formulaSpreadsheetIsEditing = next;
  if (typeof window === "undefined") return;
  try {
    window.dispatchEvent(new CustomEvent(SPREADSHEET_EDITING_CHANGED_EVENT, { detail: { isEditing: next } }));
  } catch {
    // ignore
  }
}

// --- AutoSave wiring ---------------------------------------------------------
// Schedule autosave whenever the document becomes dirty or changes while dirty.
const unsubscribeAutoSaveDocChange = app.getDocument().on("change", () => {
  if (!autoSaveEnabled) return;
  // Some change events (e.g. authoritative backend deltas) don't mark the document dirty.
  // Only schedule autosave once we actually have local dirty state to persist.
  if (!app.getDocument().isDirty) return;
  noteAutoSaveChange();
});
const unsubscribeAutoSaveDocDirty = app.getDocument().on("dirty", (evt: any) => {
  if (!autoSaveEnabled) return;
  if (evt?.isDirty !== true) return;
  noteAutoSaveChange();
});
const unsubscribeAutoSaveEditState = app.onEditStateChange(() => {
  if (!autoSaveEnabled) return;
  if (isSpreadsheetEditing()) return;
  if (!autoSaveNeedsSaveAfterEditing) return;
  autoSaveNeedsSaveAfterEditing = false;
  autoSaveForceNextSave = true;
  scheduleAutoSaveFromLastChange();
});
window.addEventListener("unload", () => {
  unsubscribeAutoSaveDocChange();
  unsubscribeAutoSaveDocDirty();
  unsubscribeAutoSaveEditState();
  clearAutoSaveTimer();
});

const openFormatCells = createOpenFormatCells({
  isEditing: () => isSpreadsheetEditing(),
  isReadOnly: () => app.isReadOnly?.() === true,
  getDocument: () => app.getDocument(),
  getSheetId: () => app.getCurrentSheetId(),
  getActiveCell: () => app.getActiveCell(),
  getSelectionRanges: () => app.getSelectionRanges(),
  getGridLimits: () => getGridLimitsForFormatting(),
  focusGrid: () => app.focus(),
});

async function openOrganizeSheets(): Promise<void> {
  const getStore = () => createPermissionGuardedSheetStore(workbookSheetStore, () => app.getCollabSession?.() ?? null);
  // Reuse the existing add-sheet command logic, but avoid restoring focus to the grid while the
  // dialog is open (the dialog will restore focus on close instead).
  const addSheet = createAddSheetCommand({
    app,
    getWorkbookSheetStore: () => workbookSheetStore,
    restoreFocusAfterSheetNavigation: () => {},
    showToast,
  });
  const { openOrganizeSheetsDialog } = await loadOrganizeSheetsDialogModule();
  openOrganizeSheetsDialog({
    store: getStore(),
    getStore,
    addSheet,
    getActiveSheetId: () => app.getCurrentSheetId(),
    activateSheet: (sheetId) => {
      app.activateSheet(sheetId);
    },
    renameSheetById: (sheetId, newName) => renameSheetById(sheetId, newName),
    getDocument: () => app.getDocument(),
    isEditing: () => isSpreadsheetEditing(),
    readOnly: app.isReadOnly?.() === true,
    focusGrid: () => app.focus(),
    onError: (message) => showToast(message, "error"),
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
  if (!hasTauriWindowHandleApi()) return undefined;
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
  const isTauri = hasTauriInvoke();
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
    const canUseProvider = hasTauri() || Boolean(globalThis.navigator?.clipboard);
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
        void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.saveWorkbook).catch((err) => {
          showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
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
  // Titlebar undo/redo state is computed by `SpreadsheetApp.getUndoRedoState`, which relies on the
  // desktop-shell-owned `__formulaSpreadsheetIsEditing` flag (to include split-view secondary edits).
  // Ensure the flag is up-to-date before reading undo/redo state so we don't render stale disabled
  // controls when edit mode transitions (e.g. Escape to exit formula-bar editing).
  emitSpreadsheetEditingChanged();
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
let ribbonShortcutById: Record<string, string> = Object.create(null);
let ribbonAriaKeyShortcutsById: Record<string, string> = Object.create(null);
let ribbonCommandRegistryDisabledById: Record<string, boolean> = Object.create(null);
let ribbonCommandRegistryDisabledByIdReady = false;

const RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY: Record<string, true> = (() => {
  // Start from the "disabled while editing" set, which includes sheet structure edits and other
  // commands that should also remain disabled for read-only collab roles (viewer/commenter).
  const out: Record<string, true> = { ...RIBBON_DISABLED_BY_ID_WHILE_EDITING };

  // Pure view-only commands should remain enabled in read-only mode.
  delete out["view.toggleShowFormulas"];
  delete out["view.window.freezePanes"];
  delete out["view.freezePanes"];
  delete out["view.freezeTopRow"];
  delete out["view.freezeFirstColumn"];
  delete out["view.unfreezePanes"];
  delete out["view.toggleSplitView"];
  delete out["view.splitVertical"];
  delete out["view.splitHorizontal"];
  delete out["view.splitNone"];
  // Selection Pane is a view panel (object list). Allow it in read-only even though other Arrange
  // commands are disabled; sheet mutations are already guarded by permission checks.
  delete out["pageLayout.arrange.selectionPane"];
  delete out["audit.togglePrecedents"];
  delete out["audit.toggleDependents"];
  delete out["audit.toggleTransitive"];
  // Formula Auditing ribbon ids are registered commands (and are safe to keep enabled in read-only),
  // but we avoid embedding the full ids here so `node:test` drift-guards can ensure we aren't
  // routing them through main.ts fallback wiring.
  for (const id of FORMULA_AUDITING_RIBBON_COMMAND_IDS) {
    delete out[id];
  }

  // View-only mutations (axis sizing) should remain enabled in read-only.
  delete out["home.cells.format"];
  delete out["home.cells.format.rowHeight"];
  delete out["home.cells.format.columnWidth"];

  // Formatting defaults (row/col/sheet band selections) should be allowed in read-only, so
  // keep these formatting commands conditionally enabled based on the current selection.
  delete out["format.toggleSubscript"];
  delete out["format.toggleSuperscript"];
  delete out["home.number.moreFormats.custom"];
  delete out["format.clearFormats"];
  // Home → Styles commands are formatting surfaces; keep them conditionally enabled (band selections only)
  // so they match the same read-only "formatting defaults" policy as other formatting controls.
  delete out["home.styles.formatAsTable"];
  delete out[HOME_STYLES_COMMAND_IDS.formatAsTable.light];
  delete out[HOME_STYLES_COMMAND_IDS.formatAsTable.medium];
  delete out[HOME_STYLES_COMMAND_IDS.formatAsTable.dark];
  delete out["home.styles.formatAsTable.newStyle"];
  delete out["home.styles.cellStyles"];
  delete out[HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral];
  delete out["home.styles.cellStyles.dataModel"];
  delete out["home.styles.cellStyles.titlesHeadings"];
  delete out["home.styles.cellStyles.numberFormat"];
  delete out["home.styles.cellStyles.newStyle"];
  // Conditional formatting is not implemented yet (CommandRegistry disables it), but keep the ids aligned
  // with the same read-only formatting-defaults gating for future implementations.
  delete out["home.styles.conditionalFormatting"];
  delete out["home.styles.conditionalFormatting.highlightCellsRules"];
  delete out["home.styles.conditionalFormatting.topBottomRules"];
  delete out["home.styles.conditionalFormatting.dataBars"];
  delete out["home.styles.conditionalFormatting.colorScales"];
  delete out["home.styles.conditionalFormatting.iconSets"];
  delete out["home.styles.conditionalFormatting.manageRules"];
  delete out["home.styles.conditionalFormatting.clearRules"];

  // Clipboard + find/replace are non-mutating and should remain enabled in read-only mode.
  delete out["clipboard.copy"];
  delete out["home.editing.findSelect"];
  delete out["edit.find"];
  delete out["edit.replace"];
  delete out["navigation.goTo"];

  // Ribbon AutoFilter MVP is view-local (row hiding via outline.hidden.filter), so allow it in read-only.
  delete out["home.editing.sortFilter"];
  delete out["home.editing.sortFilter.filter"];
  delete out["home.editing.sortFilter.clear"];
  delete out["home.editing.sortFilter.reapply"];
  delete out["data.sortFilter.filter"];
  delete out["data.sortFilter.clear"];
  delete out["data.sortFilter.reapply"];
  delete out["data.sortFilter.advanced"];
  delete out["data.sortFilter.advanced.clearFilter"];

  return out;
})();

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
    const isReadOnly = app.isReadOnly?.() === true;
    const formattingLimits = getGridLimitsForFormatting();
    const formattingDecision = evaluateFormattingSelectionSize(ranges, formattingLimits, {
      maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT,
    });
    // In read-only collab roles (viewer/commenter), allow formatting only when the selection
    // represents "formatting defaults" (entire rows/columns/sheet). These should remain
    // local-only; collab binders prevent persisting them into shared Yjs state.
    const allowReadOnlyFormattingDefaults = isReadOnly && formattingDecision.allRangesBand;
    const perfStats = app.getGridPerfStats() as any;
    const perfStatsSupported = perfStats != null;
    const perfStatsEnabled = Boolean(perfStats?.enabled);
    const isPanelOpen = (panelId: string): boolean => {
      if (!ribbonLayoutController) return false;
      const layout = ribbonLayoutController.layout as any;
      const placement = getPanelPlacement(layout, panelId) as any;
      if (placement.kind === "closed") return false;
      if (placement.kind === "floating") return !Boolean(layout?.floating?.[panelId]?.minimized);
      if (placement.kind === "docked") return !Boolean(layout?.docks?.[placement.side]?.collapsed);
      return false;
    };

    const hasRibbonAutoFilter = ribbonAutoFilterStore.hasAny(sheetId);
    const hasAdvancedFilterCommand =
      ribbonCommandRegistryDisabledByIdReady && ribbonCommandRegistryDisabledById["data.sortFilter.advanced.advancedFilter"] !== true;

    const pressedById = {
      "format.toggleBold": formatState.bold,
      "format.toggleItalic": formatState.italic,
      "format.toggleUnderline": formatState.underline,
      "format.toggleStrikethrough": formatState.strikethrough,
      "format.toggleSubscript": formatState.fontVariantPosition === "subscript",
      "format.toggleSuperscript": formatState.fontVariantPosition === "superscript",
      "format.toggleWrapText": formatState.wrapText,
      "format.alignTop": formatState.verticalAlign === "top",
      "format.alignMiddle": formatState.verticalAlign === "center",
      "format.alignBottom": formatState.verticalAlign === "bottom",
      "format.alignLeft": formatState.align === "left",
      "format.alignCenter": formatState.align === "center",
      "format.alignRight": formatState.align === "right",
      // AutoSave is only supported in the desktop/Tauri runtime.
      "file.save.autoSave": autoSaveEnabled && isTauriInvokeAvailable(),
      "view.toggleShowFormulas": app.getShowFormulas(),
      "view.togglePerformanceStats": perfStatsEnabled,
      "view.toggleSplitView": ribbonLayoutController ? ribbonLayoutController.layout.splitView.direction !== "none" : false,
      "view.togglePanel.aiChat": isPanelOpen(PanelIds.AI_CHAT),
      "view.togglePanel.aiAudit": isPanelOpen(PanelIds.AI_AUDIT),
      "view.togglePanel.versionHistory": isPanelOpen(PanelIds.VERSION_HISTORY),
      "view.togglePanel.branchManager": isPanelOpen(PanelIds.BRANCH_MANAGER),
      "view.togglePanel.dataQueries": isPanelOpen(PanelIds.DATA_QUERIES),
      "view.togglePanel.macros": isPanelOpen(PanelIds.MACROS),
      "view.togglePanel.scriptEditor": isPanelOpen(PanelIds.SCRIPT_EDITOR),
      "view.togglePanel.vbaMigrate": isPanelOpen(PanelIds.VBA_MIGRATE),
      "view.togglePanel.python": isPanelOpen(PanelIds.PYTHON),
      "view.togglePanel.marketplace": isPanelOpen(PanelIds.MARKETPLACE),
      "view.togglePanel.extensions": isPanelOpen(PanelIds.EXTENSIONS),
      "data.queriesConnections.queriesConnections": isPanelOpen(PanelIds.DATA_QUERIES),
      "comments.togglePanel": app.isCommentsPanelVisible(),
      // MVP ribbon AutoFilter: treat "Filter" as pressed when any filter is active on the sheet.
      "data.sortFilter.filter": hasRibbonAutoFilter,
      [FORMAT_PAINTER_COMMAND_ID]: Boolean(formatPainterState),
    };

    const numberFormatLabel = (() => {
      const format = formatState.numberFormat;
      if (format === "mixed") return t("ribbon.label.mixed");

      const normalized = typeof format === "string" ? format.trim() : "";
      const normalizedLower = normalized.toLowerCase();
      const localizedGeneral = t("command.format.numberFormat.general").trim().toLowerCase();
      if (!normalized || normalizedLower === "general" || (localizedGeneral && normalizedLower === localizedGeneral)) {
        return t("command.format.numberFormat.general");
      }

      const compact = normalized.toLowerCase().replace(/\s+/g, "");
      // The ribbon needs to map Excel number format codes to one of the built-in labels (Currency,
      // Date, Percent, etc). Excel format strings often include extra characters for alignment/padding:
      // - quotes around currency symbols
      // - underscores for padding
      // - unmatched parentheses (e.g. `_)`) for aligning negative numbers
      // - separate positive/negative sections delimited by `;`
      //
      // Strip the most common “noise” so we can reliably detect standard formats even when they
      // were applied via the full Format Cells UI.
      const signature = (compact.split(";")[0] ?? compact).replace(/[_()"]/g, "").replace(/\*/g, "");

      const hasCurrencySymbol = /[$€£¥]/.test(signature);
      const hasNumericPlaceholders = /[#0]/.test(signature);
      if (/^[$€£¥](?:#,##0|0)(\.[0]+)?$/.test(signature) || (hasCurrencySymbol && hasNumericPlaceholders)) {
        return t("command.format.numberFormat.currency");
      }
      if (signature.includes("yyyy-mm-dd")) return t("command.format.numberFormat.longDate");
      if (signature.includes("m/d/yyyy")) return t("command.format.numberFormat.shortDate");
      if (signature === NUMBER_FORMATS.date.toLowerCase()) return t("command.format.numberFormat.shortDate");
      if (/^h{1,2}:m{1,2}(:s{1,2})?$/.test(signature)) return t("command.format.numberFormat.time");
      if (signature.includes("%")) return t("command.format.numberFormat.percent");
      if (/^#,##0(\.[0]+)?$/.test(signature)) return t("ribbon.label.comma");
      if (/^0(\.[0]+)?$/.test(signature)) return t("command.format.numberFormat.number");
      // Be careful: `signature` can include "e" in locale codes (e.g. `[$€-de-de]...`).
      // Only treat it as scientific when it looks like an exponent marker (e.g. `0.00e+00`).
      if (/e[+-]?\d/.test(signature)) return t("command.format.numberFormat.scientific");
      if (signature.includes("/")) return t("command.format.numberFormat.fraction");
      if (signature === "@") return t("command.format.numberFormat.text");

      return t("ribbon.label.custom");
    })();

    const themeLabel = (() => {
      const preference = themeController.getThemePreference();
      switch (preference) {
        case "system":
          return t("command.view.theme.system");
        case "light":
          return t("command.view.theme.light");
        case "dark":
          return t("command.view.theme.dark");
        case "high-contrast":
          return t("command.view.theme.highContrast");
        default:
          // Default to the UX baseline.
          return t("command.view.theme.light");
      }
    })();

    const stripMenuPrefix = (value: string): string => {
      const idx = value.indexOf(":");
      if (idx === -1) return value;
      return value.slice(idx + 1).trim() || value;
    };

    const stripEllipsis = (value: string): string => {
      const trimmed = String(value ?? "").trim();
      if (!trimmed) return trimmed;
      if (trimmed.endsWith("…")) return trimmed.slice(0, -1).trim();
      if (trimmed.endsWith("...")) return trimmed.slice(0, -3).trim();
      return trimmed;
    };

    const numberFormatAriaPrefix = t("quickPick.numberFormat.placeholder");
    const moreNumberFormatsLabel = t("ribbon.label.moreNumberFormats");
    const openFormatCellsLabel = t("command.format.openFormatCells");
    const customNumberFormatLabel = t("command.home.number.moreFormats.custom");
    const formatAriaLabelWithValue = (label: string, value: string): string =>
      tWithVars("ribbon.aria.labelWithValue", { label, value });

    const labelById: Record<string, string> = {
      "home.number.numberFormat": numberFormatLabel,
      // Include the current selection value in the accessible name so screen readers can
      // announce both the control purpose and the active format (e.g. "Number format: General").
      "home.number.numberFormat.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, numberFormatLabel),
      "home.number.moreFormats.ariaLabel": moreNumberFormatsLabel,
      "view.appearance.theme": themeLabel,
      // The visible label is already "Theme: X"; use the same value so screen readers can
      // announce the current preference.
      "view.appearance.theme.ariaLabel": themeLabel,
      // Panel toggles: localize schema-provided English labels/aria labels.
      "view.togglePanel.versionHistory": t("panels.versionHistory.title"),
      "view.togglePanel.versionHistory.ariaLabel": t("commandDescription.view.togglePanel.versionHistory"),
      "view.togglePanel.branchManager": t("branchManager.title"),
      "view.togglePanel.branchManager.ariaLabel": t("commandDescription.view.togglePanel.branchManager"),
      // File → Info → Manage Workbook submenu.
      "file.info.manageWorkbook.versions": t("panels.versionHistory.title"),
      "file.info.manageWorkbook.versions.ariaLabel": t("panels.versionHistory.title"),
      "file.info.manageWorkbook.branches": t("branchManager.title"),
      "file.info.manageWorkbook.branches.ariaLabel": t("branchManager.title"),
      // Common formatting toggles: use localized command titles for icon-only button tooltips/aria labels.
      "format.toggleBold": t("command.format.toggleBold"),
      "format.toggleItalic": t("command.format.toggleItalic"),
      "format.toggleUnderline": t("command.format.toggleUnderline"),
      "format.toggleStrikethrough": t("command.format.toggleStrikethrough"),
      // Localize dropdown menu item labels (schema labels are currently English-only).
      "format.numberFormat.general": t("command.format.numberFormat.general"),
      "format.numberFormat.number": t("command.format.numberFormat.number"),
      "format.numberFormat.currency": t("command.format.numberFormat.currency"),
      "format.numberFormat.accounting": t("command.format.numberFormat.accounting"),
      "format.numberFormat.shortDate": t("command.format.numberFormat.shortDate"),
      "format.numberFormat.longDate": t("command.format.numberFormat.longDate"),
      "format.numberFormat.time": t("command.format.numberFormat.time"),
      "format.numberFormat.percent": t("command.format.numberFormat.percent"),
      "format.numberFormat.commaStyle": t("command.format.numberFormat.commaStyle"),
      "format.numberFormat.increaseDecimal": t("command.format.numberFormat.increaseDecimal"),
      "format.numberFormat.decreaseDecimal": t("command.format.numberFormat.decreaseDecimal"),
      "format.numberFormat.fraction": t("command.format.numberFormat.fraction"),
      "format.numberFormat.scientific": t("command.format.numberFormat.scientific"),
      "format.numberFormat.text": t("command.format.numberFormat.text"),
      // Provide additional context for screen readers: the menu item label alone ("General") is
      // ambiguous without knowing it belongs to the Number Format menu.
      "format.numberFormat.general.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.general")),
      "format.numberFormat.number.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.number")),
      "format.numberFormat.currency.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.currency")),
      "format.numberFormat.accounting.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.accounting")),
      "format.numberFormat.shortDate.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.shortDate")),
      "format.numberFormat.longDate.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.longDate")),
      "format.numberFormat.time.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.time")),
      "format.numberFormat.percent.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.percent")),
      "format.numberFormat.commaStyle.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.commaStyle")),
      "format.numberFormat.increaseDecimal.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.increaseDecimal")),
      "format.numberFormat.decreaseDecimal.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.decreaseDecimal")),
      "format.numberFormat.fraction.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.fraction")),
      "format.numberFormat.scientific.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.scientific")),
      "format.numberFormat.text.ariaLabel": formatAriaLabelWithValue(numberFormatAriaPrefix, t("command.format.numberFormat.text")),
      // Number format dialog entrypoints.
      "format.openFormatCells": openFormatCellsLabel,
      "format.openFormatCells.ariaLabel": stripEllipsis(openFormatCellsLabel),
      "home.number.moreFormats.custom": customNumberFormatLabel,
      "home.number.moreFormats.custom.ariaLabel": formatAriaLabelWithValue(
        moreNumberFormatsLabel,
        stripEllipsis(customNumberFormatLabel),
      ),
      // Accounting symbol picker menu items (Home → Number → Accounting dropdown).
      "format.numberFormat.accounting.usd": stripMenuPrefix(t("command.format.numberFormat.accounting.usd")),
      "format.numberFormat.accounting.eur": stripMenuPrefix(t("command.format.numberFormat.accounting.eur")),
      "format.numberFormat.accounting.gbp": stripMenuPrefix(t("command.format.numberFormat.accounting.gbp")),
      "format.numberFormat.accounting.jpy": stripMenuPrefix(t("command.format.numberFormat.accounting.jpy")),
      "format.numberFormat.accounting.usd.ariaLabel": t("command.format.numberFormat.accounting.usd"),
      "format.numberFormat.accounting.eur.ariaLabel": t("command.format.numberFormat.accounting.eur"),
      "format.numberFormat.accounting.gbp.ariaLabel": t("command.format.numberFormat.accounting.gbp"),
      "format.numberFormat.accounting.jpy.ariaLabel": t("command.format.numberFormat.accounting.jpy"),
      "view.appearance.theme.system": t("ribbon.theme.system"),
      "view.appearance.theme.light": t("ribbon.theme.light"),
      "view.appearance.theme.dark": t("ribbon.theme.dark"),
      "view.appearance.theme.highContrast": t("ribbon.theme.highContrast"),
      "view.appearance.theme.system.ariaLabel": t("commandDescription.view.theme.system"),
      "view.appearance.theme.light.ariaLabel": t("commandDescription.view.theme.light"),
      "view.appearance.theme.dark.ariaLabel": t("commandDescription.view.theme.dark"),
      "view.appearance.theme.highContrast.ariaLabel": t("commandDescription.view.theme.highContrast"),
    };

    // Font dropdown labels should reflect the current selection (Excel-style).
    // Keep defaults ("Font"/"Size") when we don't have an explicit value, but show "Mixed"
    // when the selection contains multiple values.
    if (formatState.fontName === "mixed") {
      labelById["home.font.fontName"] = t("ribbon.label.mixed");
    } else if (typeof formatState.fontName === "string") {
      labelById["home.font.fontName"] = formatState.fontName;
    }

    if (formatState.fontSize === "mixed") {
      labelById["home.font.fontSize"] = t("ribbon.label.mixed");
    } else if (typeof formatState.fontSize === "number") {
      labelById["home.font.fontSize"] = String(formatState.fontSize);
    }

    const printExportAvailable = typeof queuedInvoke === "function" || hasTauriInvoke();

    const zoomDisabled = !app.supportsZoom();
    const outlineDisabled = app.getGridMode() === "shared";
    const canComment = (() => {
      // Fail closed for collab sessions: if permissions are unset/unknown or the
      // capability check throws, treat the user as unable to comment so we don't
      // surface comment write actions pre-permissions.
      let session: any = null;
      try {
        session = app.getCollabSession?.() ?? null;
      } catch {
        session = null;
      }
      if (!session) return true;
      try {
        return typeof session.canComment === "function" ? Boolean(session.canComment()) : false;
      } catch {
        return false;
      }
    })();
    const shouldDisableFormattingCommands = isEditing || (isReadOnly && !allowReadOnlyFormattingDefaults);

    const dynamicDisabledById = {
      ...(shouldDisableFormattingCommands
        ? {
            // Formatting commands are disabled while editing (Excel-style behavior).
            //
            // In read-only collab sessions (viewer/commenter), we only enable formatting commands
            // when the selection is an entire row/column/sheet (formatting defaults), since those
            // mutations remain local-only.
            "format.toggleBold": true,
            "format.toggleItalic": true,
            "format.toggleUnderline": true,
            "format.toggleStrikethrough": true,
            "format.toggleSubscript": true,
            "format.toggleSuperscript": true,
            "home.font.fontName": true,
            "home.font.fontSize": true,
            ...RIBBON_DISABLED_FONT_COMMANDS_WHILE_EDITING,
            "home.font.fontColor": true,
            "home.font.fillColor": true,
            "home.font.borders": true,
            "home.font.clearFormatting": true,
            [FORMAT_PAINTER_COMMAND_ID]: true,
            "format.toggleWrapText": true,
            "format.alignTop": true,
            "format.alignMiddle": true,
            "format.alignBottom": true,
            "format.alignLeft": true,
            "format.alignCenter": true,
            "format.alignRight": true,
            "home.alignment.orientation": true,
            "format.increaseIndent": true,
            "format.decreaseIndent": true,
            "home.number.numberFormat": true,
            "home.number.moreFormats": true,
            "format.numberFormat.percent": true,
            "format.numberFormat.accounting": true,
            "format.numberFormat.shortDate": true,
            "format.numberFormat.commaStyle": true,
            "format.numberFormat.increaseDecimal": true,
            "format.numberFormat.decreaseDecimal": true,
            "format.openFormatCells": true,
            "home.number.moreFormats.custom": true,
            "format.clearFormats": true,
             // Style commands should follow the same read-only formatting-defaults gating.
             "home.styles.formatAsTable": true,
             [HOME_STYLES_COMMAND_IDS.formatAsTable.light]: true,
             [HOME_STYLES_COMMAND_IDS.formatAsTable.medium]: true,
             [HOME_STYLES_COMMAND_IDS.formatAsTable.dark]: true,
             "home.styles.cellStyles": true,
             [HOME_STYLES_COMMAND_IDS.cellStyles.goodBadNeutral]: true,
             "home.styles.conditionalFormatting": true,
             "home.styles.conditionalFormatting.highlightCellsRules": true,
             "home.styles.conditionalFormatting.topBottomRules": true,
             "home.styles.conditionalFormatting.dataBars": true,
            "home.styles.conditionalFormatting.colorScales": true,
            "home.styles.conditionalFormatting.iconSets": true,
            "home.styles.conditionalFormatting.manageRules": true,
            "home.styles.conditionalFormatting.clearRules": true,
          }
        : null),
      ...(isEditing ? RIBBON_DISABLED_BY_ID_WHILE_EDITING : null),
      ...(isReadOnly
        ? {
            ...RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY,
            // Keep Format Painter disabled in read-only (it can apply local-only formatting to
            // arbitrary ranges; we only allow formatting defaults).
            [FORMAT_PAINTER_COMMAND_ID]: true,
            // Organize Sheets is a structural edit surface and must remain disabled in read-only.
            "home.cells.format.organizeSheets": true,
            // Workbook mutations should remain disabled in read-only mode even if formatting defaults are enabled.
            "format.clearAll": true,
          }
        : null),
      ...(isEditing || !canComment ? { "comments.addComment": true } : null),
      // Comment mutations should be disabled for viewers even if the UI surface is otherwise
      // visible (e.g. the Review tab includes delete actions). Use conditional spreads so we
      // don't accidentally override the CommandRegistry baseline (which disables unregistered
      // placeholder commands) with explicit `false` values.
      ...(!canComment
        ? {
            "review.comments.deleteComment": true,
            "review.comments.deleteComment.deleteThread": true,
            "review.comments.deleteComment.deleteAll": true,
          }
        : null),
      ...(isReadOnly
        ? {
            // Editing clipboard actions are disabled in read-only mode.
            "clipboard.cut": true,
            "clipboard.paste": true,
            "clipboard.pasteSpecial": true,
            "clipboard.pasteSpecial.values": true,
            "clipboard.pasteSpecial.formulas": true,
            "clipboard.pasteSpecial.formats": true,
            "clipboard.pasteSpecial.transpose": true,
            // Structural edits should not be available in read-only mode.
            "home.cells.insert.insertCells": true,
            "home.cells.delete.deleteCells": true,
            "home.cells.insert.insertSheetRows": true,
            "home.cells.insert.insertSheetColumns": true,
            "home.cells.delete.deleteSheetRows": true,
            "home.cells.delete.deleteSheetColumns": true,
          }
        : null),
      ...(printExportAvailable
        ? null
        : {
            // In web/demo builds we do not have access to the desktop print/export backend.
            "pageLayout.pageSetup.pageSetupDialog": true,
            "pageLayout.pageSetup.margins": true,
            "pageLayout.pageSetup.orientation": true,
            "pageLayout.pageSetup.size": true,
            "pageLayout.pageSetup.printArea": true,
            "pageLayout.printArea.setPrintArea": true,
            "pageLayout.printArea.clearPrintArea": true,
            "pageLayout.export.exportPdf": true,
          }),
      // View/zoom controls depend on the current runtime (e.g. shared-grid mode).
      ...(!perfStatsSupported ? { "view.togglePerformanceStats": true } : null),
      ...(zoomDisabled
        ? {
            "view.zoom.zoom": true,
            "view.zoom.zoom100": true,
            "view.zoom.zoomToSelection": true,
          }
        : null),
      ...(outlineDisabled
        ? {
            // Shared-grid mode does not yet support outline groups.
            "data.outline.group": true,
            "data.outline.ungroup": true,
            "data.outline.subtotal": true,
            "data.outline.showDetail": true,
            "data.outline.hideDetail": true,
            // MVP ribbon AutoFilter uses outline.hidden.filter, which isn't implemented in shared-grid mode yet.
            "data.sortFilter.filter": true,
             "data.sortFilter.clear": true,
             "data.sortFilter.reapply": true,
             "data.sortFilter.advanced": true,
             "data.sortFilter.advanced.clearFilter": true,
          }
        : null),
      ...(!hasRibbonAutoFilter
        ? {
            // Clear/Reapply should only be enabled when AutoFilter is enabled for the current sheet.
            // (Excel-style: these are no-ops when no filter exists.)
            "data.sortFilter.clear": true,
            "data.sortFilter.reapply": true,
            "data.sortFilter.advanced.clearFilter": true,
            // Keep the Advanced dropdown disabled when it would otherwise contain only disabled items.
            ...(hasAdvancedFilterCommand ? null : { "data.sortFilter.advanced": true }),
          }
        : null),
    };

    const disabledById = Object.assign(Object.create(ribbonCommandRegistryDisabledById), dynamicDisabledById) as Record<
      string,
      boolean
    >;

    setRibbonUiState({
      pressedById,
      labelById,
      disabledById,
      shortcutById: ribbonShortcutById,
      ariaKeyShortcutsById: ribbonAriaKeyShortcutsById,
    });
  });
}

app.subscribeSelection((selection) => {
  renderStatusMode();
  scheduleRibbonSelectionFormatStateUpdate();
  handleFormatPainterSelectionChange(selection);
});
app.getDocument().on("change", () => scheduleRibbonSelectionFormatStateUpdate());

// `SpreadsheetApp.onEditStateChange(...)` invokes the callback immediately to sync initial state.
// The sheet tabs UI (and its supporting variables like `sheetTabsReactRoot`, `handleAddSheet`, etc)
// is wired up later in this module. Defer sheet-tab rendering until wiring is ready so startup
// doesn't trip TDZ errors by calling `renderSheetTabs()` too early.
let sheetTabsRenderReady = false;
let sheetTabsRenderScheduled = false;
function scheduleSheetTabsRender(): void {
  if (sheetTabsRenderScheduled) return;
  sheetTabsRenderScheduled = true;
  const schedule = typeof queueMicrotask === "function" ? queueMicrotask : (cb: () => void) => Promise.resolve().then(cb);
  schedule(() => {
    sheetTabsRenderScheduled = false;
    if (!sheetTabsRenderReady) return;
    renderSheetTabs();
  });
}

app.onEditStateChange(() => {
  emitSpreadsheetEditingChanged();
  scheduleRibbonSelectionFormatStateUpdate();
  scheduleSheetTabsRender();
});
window.addEventListener("formula:view-changed", () => scheduleRibbonSelectionFormatStateUpdate());
window.addEventListener("formula:read-only-changed", () => {
  scheduleRibbonSelectionFormatStateUpdate();
  scheduleSheetTabsRender();
});
scheduleRibbonSelectionFormatStateUpdate();

window.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!formatPainterState) return;
  if (isSpreadsheetEditing()) return;
  disarmFormatPainter();
});

// Excel-style focus cycling (F6 / Shift+F6) is handled via the keybinding pipeline:
// KeybindingService -> CommandRegistry (`workbench.focusNextRegion` / `workbench.focusPrevRegion`).
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
let powerQueryServiceStartToken = 0;
type PowerQueryServiceModule = typeof import("./power-query/service.js");
let powerQueryServiceModule: PowerQueryServiceModule | null = null;
let powerQueryServiceInitPromise: Promise<void> | null = null;
let powerQueryServiceInitWorkbookId: string | null = null;

function updateTrayFromPowerQuery(): void {
  const status = powerQueryTrayJobs.size > 0 ? "syncing" : powerQueryTrayHadError ? "error" : "idle";
  void setTrayStatus(status);
}

function currentPowerQueryWorkbookId(): string {
  return activePanelWorkbookId;
}

function startPowerQueryService(): void {
  const serviceWorkbookId = currentPowerQueryWorkbookId();

  // Already running (or already starting) for this workbook id.
  if (powerQueryService && powerQueryServiceWorkbookId === serviceWorkbookId) return;
  if (powerQueryServiceInitPromise && powerQueryServiceInitWorkbookId === serviceWorkbookId) return;

  stopPowerQueryService();
  const startToken = ++powerQueryServiceStartToken;
  powerQueryServiceInitWorkbookId = serviceWorkbookId;

  let initPromise: Promise<void>;
  initPromise = (async () => {
    const mod = await loadPowerQueryServiceModule();
    if (!mod) return;
    powerQueryServiceModule = mod;

    // Bail if another start/stop happened while the chunk was loading.
    if (startToken !== powerQueryServiceStartToken) return;

    powerQueryServiceWorkbookId = serviceWorkbookId;
    const service = new mod.DesktopPowerQueryService({
      workbookId: serviceWorkbookId,
      document: app.getDocument(),
      concurrency: 1,
      batchSize: 1024,
    });
    powerQueryService = service;
    mod.setDesktopPowerQueryService(serviceWorkbookId, service);

    stopPowerQueryTrayListener?.();
    powerQueryTrayJobs.clear();
    powerQueryTrayHadError = false;
    updateTrayFromPowerQuery();

    stopPowerQueryTrayListener = service.onEvent((evt) => {
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
  })()
    .catch((err) => {
      console.error("Failed to start Power Query service:", err);
      showToast(`Failed to start Power Query service: ${String(err)}`, "error");
    })
    .finally(() => {
      if (powerQueryServiceInitPromise === initPromise) {
        powerQueryServiceInitPromise = null;
        powerQueryServiceInitWorkbookId = null;
      }
    });

  powerQueryServiceInitPromise = initPromise;
}

function stopPowerQueryService(): void {
  // Invalidate any in-flight lazy initialization.
  powerQueryServiceStartToken += 1;
  powerQueryServiceInitPromise = null;
  powerQueryServiceInitWorkbookId = null;

  const existingWorkbookId = powerQueryServiceWorkbookId;
  powerQueryServiceWorkbookId = null;
  if (existingWorkbookId && powerQueryServiceModule) {
    try {
      powerQueryServiceModule.setDesktopPowerQueryService(existingWorkbookId, null);
    } catch {
      // ignore
    }
  }
  stopPowerQueryTrayListener?.();
  stopPowerQueryTrayListener = null;
  powerQueryService?.dispose();
  powerQueryService = null;

  powerQueryTrayJobs.clear();
  powerQueryTrayHadError = false;
  updateTrayFromPowerQuery();
}

// Defer Power Query startup work so the initial grid render + time-to-interactive aren't
// blocked by parsing/executing Power Query (connectors, auth flows, etc).
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
// through UI affordances. Assign this as early as possible so e2e can still
// attach even if later desktop integrations fail during startup.
window.__formulaCommandRegistry = commandRegistry;

// Excel-style edit mode: some commands are disabled while a cell/formula edit is active.
// This prevents bypassing ribbon-disabled actions via non-ribbon surfaces (command palette,
// keybindings, etc) and avoids "click does nothing" UX when command implementations no-op.
const shouldBlockCommandWhileEditing = (commandId: string): boolean => {
  const id = String(commandId);
  // Comments can still be opened/read while editing, but add/mutate actions should not trigger.
  if (id === "comments.addComment") return true;
  // Format commands can open pickers/dialogs and should not run while editing.
  if (id.startsWith("format.")) return true;
  return RIBBON_DISABLED_BY_ID_WHILE_EDITING[id] === true;
};

commandRegistry.onWillExecuteCommand(({ commandId }) => {
  if (!isSpreadsheetEditing()) return;
  if (!shouldBlockCommandWhileEditing(commandId)) return;
  try {
    showToast("Finish editing to use this command.", "warning");
  } catch {
    // ignore (toast root missing in non-UI test environments)
  }
  throw new SpreadsheetEditingCommandBlockedError(commandId);
});

// Ribbon AutoFilter MVP command wiring.
//
// Register these early so keybindings (Ctrl+Shift+L / Cmd+Shift+L) can dispatch as soon as the
// KeybindingService is installed, without duplicating `registerBuiltinCommand` calls in main.ts
// (CommandRegistry guardrails require registrations live under `src/commands`).
const ribbonAutoFilterHandlers = {
  toggle: async (pressed?: boolean) => {
    try {
      const nextEnabled = typeof pressed === "boolean" ? pressed : !ribbonAutoFilterStore.hasAny(app.getCurrentSheetId());
      if (nextEnabled) {
        await applyRibbonAutoFilterFromSelection();
      } else {
        clearRibbonAutoFiltersForActiveSheet();
      }
    } catch (err) {
      console.error("Failed to toggle filter:", err);
      showToast(`Failed to toggle filter: ${String(err)}`, "error");
      scheduleRibbonSelectionFormatStateUpdate();
    } finally {
      app.focus();
    }
  },
  // Excel-style: "Clear" clears filter criteria (shows all rows) but keeps AutoFilter enabled.
  clear: () => clearRibbonAutoFilterCriteriaForActiveSheet(),
  reapply: () => reapplyRibbonAutoFiltersForActiveSheet(),
};
registerRibbonAutoFilterCommands({
  commandRegistry,
  getHandlers: () => ribbonAutoFilterHandlers,
  isEditing: isSpreadsheetEditing,
  category: t("commandCategory.data"),
});

const RIBBON_DISABLED_FONT_COMMANDS_WHILE_EDITING: Record<string, true> = Object.fromEntries(
  [
    ...FORMAT_FONT_NAME_PRESET_COMMAND_IDS,
    ...FORMAT_FONT_SIZE_PRESET_COMMAND_IDS,
    // Step commands are registered by `registerBuiltinCommands` (not `registerBuiltinFormatFontCommands`).
    "format.fontSize.increase",
    "format.fontSize.decrease",
  ].map((id) => [id, true] as const),
);

// --- Ribbon: auto-disable unimplemented CommandRegistry-backed controls ----------
//
// Many ribbon controls exist for Excel parity but are not implemented yet. When clicked, they
// fall back to a noisy `showToast("Ribbon: <id>")`. Use the CommandRegistry as the source of
// truth for which command ids exist, and disable missing ids by default.
//
// NOTE: `COMMAND_REGISTRY_EXEMPT_IDS` (in `ribbonCommandRegistryDisabling.ts`) is intentionally
// empty. Prefer registering UI-specific ribbon ids (e.g. `file.*`) as hidden CommandRegistry
// aliases so ribbon enable/disable, command palette, and keybindings can share a single source
// of truth.
let ribbonCommandRegistryRefreshScheduled = false;
const scheduleRibbonCommandRegistryDisabledRefresh = (): void => {
  if (ribbonCommandRegistryRefreshScheduled) return;
  ribbonCommandRegistryRefreshScheduled = true;

  const schedule = typeof queueMicrotask === "function" ? queueMicrotask : (cb: () => void) => Promise.resolve().then(cb);
  schedule(() => {
    ribbonCommandRegistryRefreshScheduled = false;
    ribbonCommandRegistryDisabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry);
    ribbonCommandRegistryDisabledByIdReady = true;
    scheduleRibbonSelectionFormatStateUpdate();
  });
};

scheduleRibbonCommandRegistryDisabledRefresh();
const unsubscribeRibbonCommandRegistry = commandRegistry.subscribe(() => scheduleRibbonCommandRegistryDisabledRefresh());
window.addEventListener("unload", () => {
  unsubscribeRibbonCommandRegistry();
});

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

// `updateContextKeys` is initialized once the extensions/context-key wiring is ready.
// Sheet UI helpers (like sheet rename) can call it safely; it is a no-op until wired.
let updateContextKeys: (selection?: SelectionState | null) => void = () => {};

// --- Sheet tabs (Excel-like multi-sheet UI) -----------------------------------
let stopSheetStoreListener: (() => void) | null = null;
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

function focusActiveSheetTab(): void {
  const activeSheetId = app.getCurrentSheetId();
  const buttons = sheetTabsRootEl.querySelectorAll<HTMLButtonElement>('.sheet-tabs button[role="tab"][data-sheet-id]');
  for (const btn of buttons) {
    if (btn.dataset.sheetId === activeSheetId) {
      btn.focus({ preventScroll: true });
      return;
    }
  }
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
  // In collab mode, the authoritative sheet list comes from the Yjs session (`session.sheets`).
  // Avoid reconciling against DocumentController's lazily-created sheet ids, which can drift
  // (and in read-only sessions would cause local-only UI mutations).
  const session = app.getCollabSession?.() ?? null;
  if (session) return;

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
      // allow the removal (e.g. last-visible-sheet guard).
      //
      // The DocumentController is the source of truth for sheet ids here. If the
      // sheet store refuses the removal due to Excel-style invariants, fall back
      // to replacing the store snapshot so we still drop the orphaned sheet id.
      try {
        workbookSheetStore.replaceAll(workbookSheetStore.listAll().filter((s) => s.id !== sheet.id));
      } catch {
        // ignore
      }
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
    workbookSheetStore = new CollabWorkbookSheetStore(session, sheets, collabSheetsKeyRef, {
      canEditWorkbook: () => getWorkbookMutationPermission(session).allowed,
    });
  } catch (err) {
    // If collab sheet names are invalid/duplicated (shouldn't happen, but can if a remote
    // client writes bad metadata), fall back to using the stable sheet id as the display name
    // so the UI remains functional.
    console.error("[formula][desktop] Failed to apply collab sheet metadata:", err);
    workbookSheetStore = new CollabWorkbookSheetStore(
      session,
      sheets.map((sheet) => ({ ...sheet, name: sheet.id })),
      collabSheetsKeyRef,
      { canEditWorkbook: () => getWorkbookMutationPermission(session).allowed },
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
  //
  // Defensive: if the workbook metadata is invalid (all sheets hidden/veryHidden), fall back
  // to exposing exactly one sheet so the UI remains functional and the user can unhide sheets
  // via the context menu. Prefer the current active sheet when possible.
  if (visible.length > 0) return visible.map((s) => ({ id: s.id, name: s.name }));

  const all = workbookSheetStore.listAll();
  if (all.length === 0) return [];

  const activeId = app.getCurrentSheetId();
  const active = all.find((s) => s.id === activeId) ?? null;

  // Avoid exposing `veryHidden` sheets when possible (Excel doesn't show them in the UI).
  const nonVeryHiddenActive = active && active.visibility !== "veryHidden" ? active : null;
  const firstNonVeryHidden = all.find((s) => s.visibility !== "veryHidden") ?? null;

  const fallback = nonVeryHiddenActive ?? firstNonVeryHidden ?? active ?? all[0]!;
  return [{ id: fallback.id, name: fallback.name }];
}

const handleAddSheet = createAddSheetCommand({
  app,
  getWorkbookSheetStore: () => workbookSheetStore,
  restoreFocusAfterSheetNavigation,
  showToast,
});

const handleDeleteActiveSheet = createDeleteActiveSheetCommand({
  app,
  getWorkbookSheetStore: () => workbookSheetStore,
  restoreFocusAfterSheetNavigation,
  showToast,
  confirm: (message) => nativeDialogs.confirm(message),
});

async function renameSheetById(
  sheetId: string,
  newName: string,
): Promise<{ id: string; name: string }> {
  const id = String(sheetId ?? "").trim();
  if (!id) throw new Error("Sheet id cannot be empty");

  // In case a sheet was created lazily in the DocumentController (or via an extension)
  // and the doc->store sync hasn't run yet, reconcile once before rejecting the id.
  let sheet = workbookSheetStore.getById(id);
  if (!sheet) {
    reconcileSheetStoreWithDocument(listDocumentSheetIds());
    sheet = workbookSheetStore.getById(id);
  }
  if (!sheet) throw new Error("Sheet not found");

  const oldDisplayName = sheet.name || id;
  const normalizedNewName = validateSheetName(String(newName ?? ""), {
    sheets: workbookSheetStore.listAll(),
    ignoreId: id,
  });

  // No-op rename; preserve the same return shape as other sheet APIs.
  if (oldDisplayName === normalizedNewName) {
    return { id, name: oldDisplayName };
  }

  const collabSession = app.getCollabSession?.() ?? null;
  if (collabSession) {
    const permission = getWorkbookMutationPermission(collabSession);
    if (!permission.allowed) {
      throw new Error(permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE);
    }
  }

  // Update UI metadata first so follow-up operations observe the new name.
  workbookSheetStore.rename(id, normalizedNewName);

  // Rewrite existing formulas that reference the old sheet name (Excel-style behavior).
  const doc = app.getDocument() as any;
  try {
    rewriteDocumentFormulasForSheetRename(doc, oldDisplayName, normalizedNewName);
  } catch (err) {
    showToast(`Failed to update formulas after rename: ${String((err as any)?.message ?? err)}`, "error");
  }

  syncSheetUi();
  try {
    updateContextKeys();
  } catch {
    // Best-effort; context keys should never block a rename.
  }

  return { id, name: workbookSheetStore.getName(id) ?? normalizedNewName };
}

const permissionGuardedSheetStoreCache = new WeakMap<WorkbookSheetStore, WorkbookSheetStore>();

function createPermissionGuardedSheetStore(
  store: WorkbookSheetStore,
  getSession: () => CollabSession | null,
): WorkbookSheetStore {
  const cached = permissionGuardedSheetStoreCache.get(store);
  if (cached) return cached;

  const guardedMutations = new Set([
    "addAfter",
    "rename",
    "move",
    "remove",
    "hide",
    "unhide",
    "setVisibility",
    "setTabColor",
    // Not used by sheet tabs directly, but include for completeness.
    "replaceAll",
  ]);

  const guarded = new Proxy(store, {
    get(target, prop) {
      const value = (target as any)[prop];
      if (typeof value !== "function") return value;

      // Bind non-mutating methods to the underlying store so `this` remains correct.
      const name = typeof prop === "string" ? prop : null;
      if (name && guardedMutations.has(name)) {
        return (...args: any[]) => {
          const session = getSession();
          const permission = getWorkbookMutationPermission(session);
          if (session && !permission.allowed) {
            throw new Error(permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE);
          }
          return value.apply(target, args);
        };
      }

      return value.bind(target);
    },
  });

  const out = guarded as unknown as WorkbookSheetStore;
  permissionGuardedSheetStoreCache.set(store, out);
  return out;
}

// Sheet tabs wiring is now ready (React root, sheet commands, permission guard helpers, etc).
// Allow earlier `onEditStateChange`/`read-only-changed` listeners to render, and render the
// initial tab strip immediately.
sheetTabsRenderReady = true;
renderSheetTabs();

function renderSheetTabs(): void {
  if (!sheetTabsReactRoot) {
    sheetTabsReactRoot = createRoot(sheetTabsRootEl);
  }

  const handleSheetDeletedFromTabs = (event: { sheetId: string; name: string; sheetOrder: string[] }) => {
    const { name, sheetOrder } = event;
    const doc = app.getDocument() as any;

    try {
      // `installSheetStoreSubscription()` routes the sheet removal into `DocumentController.deleteSheet(...)`
      // and keeps the batch open through the end of the current task. Rewrite formulas synchronously here
      // so the delete + rewrites become a single undo step (Excel-like).
      rewriteDocumentFormulasForSheetDelete(doc, name, sheetOrder);
    } catch (err) {
      showToast(`Failed to update formulas after delete: ${String((err as any)?.message ?? err)}`, "error");
    }
  };

  sheetTabsReactRoot.render(
    React.createElement(SheetTabStrip, {
      store: createPermissionGuardedSheetStore(workbookSheetStore, () => app.getCollabSession?.() ?? null),
      activeSheetId: app.getCurrentSheetId(),
      readOnly: app.isReadOnly?.() === true,
      disableMutations: isSpreadsheetEditing(),
      onActivateSheet: (sheetId: string) => {
        app.activateSheet(sheetId);
        restoreFocusAfterSheetNavigation();
      },
      onAddSheet: handleAddSheet,
      onRenameSheet: async (sheetId: string, newName: string) => {
        await renameSheetById(sheetId, newName);
      },
      onSheetsReordered: () => restoreFocusAfterSheetNavigation(),
      onSheetDeleted: handleSheetDeletedFromTabs,
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
    sheetPositionEl.textContent = tWithVars("statusBar.sheetPosition", { position: 0, total: 0 });
    sheetPositionEl.dataset.sheetPosition = "0";
    sheetPositionEl.dataset.sheetTotal = "0";
    return;
  }
  const index = sheets.findIndex((sheet) => sheet.id === activeId);
  const position = index >= 0 ? index + 1 : 1;
  sheetPositionEl.textContent = tWithVars("statusBar.sheetPosition", { position, total });
  sheetPositionEl.dataset.sheetPosition = String(position);
  sheetPositionEl.dataset.sheetTotal = String(total);
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
    // Capture the current sheet ordering before any potential collab metadata sync replaces the store.
    // This lets us choose an Excel-like adjacent visible sheet when the active sheet disappears due to
    // remote deletes (the new store no longer contains the deleted id, but the previous ordering does).
    const prevSheetsForAdjacentActivation = workbookSheetStore.listAll();

    ensureCollabSheetObserver();

    const session = app.getCollabSession?.() ?? null;
    if (session) {
      // Keep the UI store aligned with the authoritative sheet list in Yjs.
      syncSheetStoreFromCollabSession(session);
    }

    const sheets = listSheetsForUi();
    const activeId = app.getCurrentSheetId();
    if (!sheets.some((sheet) => sheet.id === activeId)) {
      // If the active sheet is missing from the visible list, prefer activating the adjacent
      // visible sheet (Excel-like). This can happen when the active sheet is hidden/veryHidden
      // or removed during state restores (version history, branch checkout, etc).
      const currentById = new Map(workbookSheetStore.listAll().map((sheet) => [sheet.id, sheet.visibility] as const));
      const orderedCurrentVisibility = prevSheetsForAdjacentActivation.map((sheet) => ({
        id: sheet.id,
        visibility: currentById.get(sheet.id) ?? ("hidden" satisfies SheetVisibility),
      }));
      const preferred = pickAdjacentVisibleSheetId(orderedCurrentVisibility, activeId);
      const fallback =
        (preferred && sheets.some((sheet) => sheet.id === preferred) ? preferred : null) ?? sheets[0]?.id ?? null;
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

  // Coalesce doc-driven sheet reorder persistence so we don't spam the backend when a single
  // DocumentController reorder is applied to the store as multiple `store.move(...)` calls.
  let pendingDocDrivenReorder:
    | { invoke: TauriInvoke; prevOrder: string[]; nextOrder: string[] }
    | null = null;
  let pendingDocDrivenReorderScheduled = false;

  const simulateMove = (order: string[], sheetId: string, toIndex: number): string[] => {
    const fromIndex = order.indexOf(sheetId);
    if (fromIndex === -1) return order.slice();
    if (fromIndex === toIndex) return order.slice();
    const next = order.slice();
    next.splice(fromIndex, 1);
    next.splice(toIndex, 0, sheetId);
    return next;
  };

  const detectSingleMove = (
    prevOrder: string[],
    nextOrder: string[],
  ): { sheetId: string; toIndex: number } | null => {
    if (prevOrder.length !== nextOrder.length) return null;
    if (prevOrder.length === 0) return null;

    // Ensure we're only dealing with pure reorders (no adds/removes).
    const prevSet = new Set(prevOrder);
    for (const id of nextOrder) {
      if (!prevSet.has(id)) return null;
    }

    const firstMismatch = prevOrder.findIndex((id, idx) => id !== nextOrder[idx]);
    if (firstMismatch === -1) return null;

    const tryMove = (sheetId: string, toIndex: number): { sheetId: string; toIndex: number } | null => {
      if (!sheetId) return null;
      if (!Number.isInteger(toIndex) || toIndex < 0 || toIndex >= prevOrder.length) return null;
      const simulated = simulateMove(prevOrder, sheetId, toIndex);
      const ok = simulated.length === nextOrder.length && simulated.every((id, idx) => id === nextOrder[idx]);
      return ok ? { sheetId, toIndex } : null;
    };

    // Prefer moving the first mismatched sheet from the *previous* order. This is deterministic,
    // and matches our current e2e expectations for undo/redo swap cases.
    const preferredId = prevOrder[firstMismatch] ?? "";
    const preferredToIndex = nextOrder.indexOf(preferredId);
    const preferred = tryMove(preferredId, preferredToIndex);
    if (preferred) return preferred;

    // Fall back to moving the sheet that appears at the mismatch location in the *next* order.
    const altId = nextOrder[firstMismatch] ?? "";
    const alt = tryMove(altId, firstMismatch);
    if (alt) return alt;

    return null;
  };

  const flushPersistDocDrivenReorder = (): void => {
    pendingDocDrivenReorderScheduled = false;
    const pending = pendingDocDrivenReorder;
    pendingDocDrivenReorder = null;
    if (!pending) return;

    const { invoke, prevOrder, nextOrder } = pending;
    if (prevOrder.join("|") === nextOrder.join("|")) return;

    const move = detectSingleMove(prevOrder, nextOrder);
    if (move) {
      void invoke("move_sheet", { sheet_id: move.sheetId, to_index: move.toIndex }).catch((err) => {
        console.error("[formula][desktop] Failed to persist sheet reorder to backend:", err);
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to sync sheet reorder to workbook: ${message}`, "error");
      });
      return;
    }

    // Fallback: apply a stable sequence of moves to transform the previous order into the next.
    // This is slower/noisier than a single move but should be rare (complex doc reorders).
    const current = prevOrder.slice();
    for (let targetIndex = 0; targetIndex < nextOrder.length; targetIndex += 1) {
      const sheetId = nextOrder[targetIndex]!;
      const currentIndex = current.indexOf(sheetId);
      if (currentIndex === -1) continue;
      if (currentIndex === targetIndex) continue;

      current.splice(currentIndex, 1);
      current.splice(targetIndex, 0, sheetId);

      void invoke("move_sheet", { sheet_id: sheetId, to_index: targetIndex }).catch((err) => {
        console.error("[formula][desktop] Failed to persist sheet reorder to backend:", err);
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to sync sheet reorder to workbook: ${message}`, "error");
      });
    }
  };

  const schedulePersistDocDrivenReorder = (invoke: TauriInvoke, prevOrder: string[], nextOrder: string[]): void => {
    if (!pendingDocDrivenReorder) {
      pendingDocDrivenReorder = { invoke, prevOrder: prevOrder.slice(), nextOrder: nextOrder.slice() };
    } else {
      // Preserve the earliest observed `prevOrder` so we can transform the backend state
      // (which should still match that ordering) into the final ordering after coalescing.
      pendingDocDrivenReorder.invoke = invoke;
      pendingDocDrivenReorder.nextOrder = nextOrder.slice();
    }

    if (pendingDocDrivenReorderScheduled) return;
    pendingDocDrivenReorderScheduled = true;
    queueMicrotask(flushPersistDocDrivenReorder);
  };

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

        // If the active sheet was removed (e.g. deleting the current tab), switch the app to a
        // remaining sheet *before* we apply DocumentController sheet deletion deltas.
        //
        // DocumentController emits synchronous `change` events during `deleteSheet(...)`. Several
        // SpreadsheetApp listeners read sheet view / cell state for the *current* sheet id, and
        // DocumentController lazily materializes sheets on read. If the app is still pointing at
        // the sheet being deleted, those listeners can unintentionally re-create the sheet in the
        // document model, causing the delete to appear to "not stick" (and later doc->store sync
        // can re-add the tab metadata).
        //
        // By moving to a valid sheet first, we ensure any change listeners observe a stable
        // active sheet id and do not resurrect the deleted sheet.
        const activeSheetId = app.getCurrentSheetId();
        if (removed.includes(activeSheetId)) {
          // Mirror Excel: deleting the active sheet activates the next visible sheet to the right,
          // otherwise the previous visible sheet. Use the *pre-delete* sheet ordering (prev snapshot)
          // to pick the adjacent sheet.
          const prevSheets = prevOrder.map((sheetId) => ({
            id: sheetId,
            visibility: prevById.get(sheetId)?.visibility ?? ("visible" satisfies SheetVisibility),
          }));
          const preferred = pickAdjacentVisibleSheetId(prevSheets, activeSheetId);
          const preferredMeta = preferred ? nextById.get(preferred) : null;
          const preferredIsValid = Boolean(preferred && preferredMeta && preferredMeta.visibility === "visible");

          const fallback = preferredIsValid
            ? preferred
            : (workbookSheetStore.listVisible().at(0)?.id ?? workbookSheetStore.listAll().at(0)?.id ?? null);
          if (fallback && fallback !== activeSheetId) {
            app.activateSheet(fallback);
          }
        }

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
            doc.deleteSheet(sheetId, { label: "Delete Sheet", source: "sheetTabs" });
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

    // When the sheet store is mutated as a *result* of DocumentController-driven state changes
    // (undo/redo/applyState/script-driven edits), reconcile the local workbook backend so future
    // saves reflect the restored sheet structure.
    if (syncingSheetUi && prevSnapshot) {
      const session = app.getCollabSession?.() ?? null;
      if (!session) {
        const baseInvoke = getTauriInvokeOrNull();
        // Prefer the queued invoke (it sequences behind pending `set_cell` / `set_range` sync work).
        const invoke =
          queuedInvoke ??
          (typeof baseInvoke === "function" ? ((cmd: string, args?: any) => queueBackendOp(() => baseInvoke(cmd, args))) : null);

        if (typeof invoke === "function") {
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

          const hasMetaChange = nextOrder.some((sheetId) => {
            if (added.includes(sheetId)) return false;
            if (removed.includes(sheetId)) return false;
            const before = prevById.get(sheetId);
            const after = nextById.get(sheetId);
            if (!before || !after) return false;
            return (
              before.name !== after.name ||
              before.visibility !== after.visibility ||
              !tabColorEqual(before.tabColor, after.tabColor)
            );
          });

          const prevKey = prevOrder.join("|");
          const nextKey = nextOrder.join("|");
          const hasOrderChange = prevKey !== nextKey;

          if (added.length > 0 || removed.length > 0 || hasMetaChange || hasOrderChange) {
            const reportError = (label: string, err: unknown): void => {
              console.error(`[formula][desktop] Failed to reconcile sheet ${label} to backend:`, err);
              const message = err instanceof Error ? err.message : String(err);
              showToast(`Failed to sync sheet changes to workbook: ${message}`, "error");
            };

            // Deletes.
            for (const sheetId of removed) {
              void invoke("delete_sheet", { sheet_id: sheetId }).catch((err) => reportError("delete", err));
            }

            // Adds.
            for (const sheetId of added) {
              const meta = nextById.get(sheetId);
              if (!meta) continue;
              const idx = nextOrder.indexOf(sheetId);
              const afterSheetId = idx > 0 ? nextOrder[idx - 1] : null;
              void invoke("add_sheet_with_id", {
                sheet_id: sheetId,
                name: meta.name,
                after_sheet_id: afterSheetId,
                index: idx,
              }).catch((err) => reportError("add", err));
            }

            // Metadata updates.
            for (const sheetId of nextOrder) {
              if (added.includes(sheetId)) continue;
              if (removed.includes(sheetId)) continue;
              const before = prevById.get(sheetId);
              const after = nextById.get(sheetId);
              if (!before || !after) continue;

              if (before.name !== after.name) {
                void invoke("rename_sheet", { sheet_id: sheetId, name: after.name }).catch((err) => reportError("rename", err));
              }

              if (before.visibility !== after.visibility) {
                void invoke("set_sheet_visibility", { sheet_id: sheetId, visibility: after.visibility }).catch((err) =>
                  reportError("visibility", err),
                );
              }

              if (!tabColorEqual(before.tabColor, after.tabColor)) {
                void invoke("set_sheet_tab_color", { sheet_id: sheetId, tab_color: after.tabColor ?? null }).catch((err) =>
                  reportError("tab color", err),
                );
              }
            }

            // Reorder (including hidden sheets).
            //
            // Reorder persistence is handled via a coalescer so we don't spam `move_sheet`
            // during doc->store sync (a single document reorder can be applied as multiple
            // `store.move(...)` calls).
            //
            // Note: when sheets are added/removed we already pass explicit insertion indices
            // to `add_sheet_with_id`, and deletions keep remaining relative order stable.
            if (hasOrderChange && added.length === 0 && removed.length === 0) {
              schedulePersistDocDrivenReorder(invoke, prevOrder, nextOrder);
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
      const preferred = pickAdjacentVisibleSheetId(workbookSheetStore.listAll(), activeId);
      const fallback =
        (preferred && sheets.some((sheet) => sheet.id === preferred) ? preferred : null) ?? sheets[0]?.id ?? null;
      if (fallback) {
        app.activateSheet(fallback);
        restoreFocusAfterSheetNavigation();
        if (syncingSheetUi) {
          // If this sheet activation is triggered while we're applying a DocumentController -> sheet-store sync
          // transaction (i.e. `syncingSheetUi` is true), the `app.activateSheet` hook's call to `syncSheetUi()`
          // will be gated by the same flag. Schedule a follow-up sync so tabs/switcher/status bar reflect the
          // new active sheet once the transaction completes.
          queueMicrotask(() => syncSheetUi());
        }
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

    // When a modal/overlay is open, do not allow global sheet-navigation shortcuts to
    // affect the workbook. Still prevent the browser default (tab switching).
    if (isEventWithinKeybindingBarrier(e)) {
      e.preventDefault();
      return;
    }

    // Ctrl/Cmd+PgUp/PgDn should generally not switch sheets while editing (cell editor,
    // inline AI edit, etc). Exception: when the formula bar is actively editing a *formula*
    // we still allow sheet navigation so users can build cross-sheet references (Excel behavior).
    const formulaBarFormulaEditing = app.isFormulaBarFormulaEditing();
    if (isSpreadsheetEditing() && !formulaBarFormulaEditing) {
      // Prevent browser tab switching / other defaults while editing spreadsheet content.
      e.preventDefault();
      e.stopPropagation();
      return;
    }

    const target = e.target as EventTarget | null;
    if (target instanceof HTMLElement) {
      const tabList = target.closest?.("#sheet-tabs .sheet-tabs");
      if (tabList) {
        // Let the sheet tab strip handle shortcuts when focus is on a tab.
        //
        // When inline rename is active the focused element is an <input>, and the tab strip
        // intentionally does not handle Ctrl/Cmd+PgUp/PgDn. In that case, prevent browser
        // defaults (e.g. tab switching) but keep focus in rename mode.
        const tag = target.tagName;
        if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) {
          e.preventDefault();
        }
        return;
      }

      // Never steal the shortcut from text inputs / contenteditable surfaces.
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) {
        // Exception: allow sheet navigation while the formula bar is editing a formula
        // (range selection / cross-sheet references).
        if (
          !formulaBarFormulaEditing ||
          !formulaBarRoot.contains(target) ||
          !target.classList.contains("formula-bar-input")
        ) {
          return;
        }
      }
    }
    // Use the sheet UI's visible list ordering so the shortcut matches the tab strip.
    // (In invalid workbooks where no sheets are visible, `listSheetsForUi()` falls back to
    // exposing a single sheet so keyboard navigation remains deterministic.)
    const ordered = listSheetsForUi().map((sheet) => sheet.id);
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
  // `restoreDocumentState` is used by workbook open + version restore. Never carry
  // Format Painter mode across workbook boundaries.
  disarmFormatPainter();
  // Ribbon AutoFilter MVP state is view-local and should not survive workbook/model replacement.
  ribbonAutoFilterStore.clearAll();
  // Clear any filter-hidden outline rows so they do not "leak" across workbooks (e.g. Sheet1 -> Sheet1).
  if (app.getGridMode() === "legacy") app.clearAllFilteredHiddenRows();
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
  if (nextSheet !== prevSheet) {
    disarmFormatPainter();
    emitSheetActivated(nextSheet);
  }
};

const originalActivateCell = app.activateCell.bind(app);
app.activateCell = (...args: Parameters<SpreadsheetApp["activateCell"]>): void => {
  const prevSheet = app.getCurrentSheetId();
  originalActivateCell(...args);
  const nextSheet = app.getCurrentSheetId();
  if (nextSheet !== prevSheet) {
    disarmFormatPainter();
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
    disarmFormatPainter();
    syncSheetUi();
    emitSheetActivated(nextSheet);
  }
};

// Keep the canvas renderer in sync with programmatic document mutations (e.g. AI tools)
// and re-render when edits create new sheets (DocumentController creates sheets lazily).
app.getDocument().on("change", (payload: any) => {
  // `DocumentController` creates sheets lazily on access (`getCell`, etc). When a sheet is removed
  // (via delete, undo, applyState restore, etc) and the UI is still "pointing at" that sheet id,
  // a synchronous refresh can immediately recreate the sheet by reading from it.
  //
  // Guard against this by switching away from any sheet ids that were deleted by this change event
  // *before* we refresh the grid.
  try {
    const sheetMetaDeltas = Array.isArray(payload?.sheetMetaDeltas) ? payload.sheetMetaDeltas : [];
    const deletedSheetIds = sheetMetaDeltas
      .filter((delta: any) => delta && typeof delta.sheetId === "string" && (delta.after ?? null) == null)
      .map((delta: any) => String(delta.sheetId));
    if (deletedSheetIds.length > 0) {
      // Ribbon AutoFilter MVP state is view-local. If a sheet is deleted (or removed via undo/redo),
      // drop any stored filter state for that sheet so we don't retain stale entries.
      for (const id of deletedSheetIds) {
        ribbonAutoFilterStore.clearSheet(id);
      }

      const currentSheetId = app.getCurrentSheetId();
      if (currentSheetId && deletedSheetIds.includes(currentSheetId)) {
        const doc = app.getDocument();
        const candidates = typeof (doc as any).getVisibleSheetIds === "function" ? doc.getVisibleSheetIds() : doc.getSheetIds();
        const fallback = candidates.find((id: string) => id && !deletedSheetIds.includes(id)) ?? null;
        if (fallback && fallback !== currentSheetId) {
          app.activateSheet(fallback);
          restoreFocusAfterSheetNavigation();
        }
      }
    }
  } catch {
    // Best-effort: sheet deletion should never crash the UI.
  }

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

  const workspaceManager = new LayoutWorkspaceManager({ storage: localStorage, panelRegistry });
  const layoutController = new LayoutController({
    workbookId,
    workspaceManager,
    primarySheetId: "Sheet1",
    workspaceId: "default",
  });
  ribbonLayoutController = layoutController;

  // Expose layout state for Playwright assertions (e.g. split view persistence).
  window.__layoutController = layoutController;

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
  // Scripting + macro tooling is large and not needed to show the initial grid. Load it lazily.
  let scriptingWorkbook: DocumentControllerWorkbookAdapter | null = null;
  let scriptingWorkbookInitPromise: Promise<DocumentControllerWorkbookAdapter | null> | null = null;
  let scriptEditorReady = false;
  let pendingScriptEditorCode: string | null = null;

  const dispatchScriptEditorSetCode = (code: string): void => {
    try {
      window.dispatchEvent(new CustomEvent("formula:script-editor:set-code", { detail: { code } }));
    } catch {
      // ignore
    }
  };

  const flushPendingScriptEditorCode = (): void => {
    if (!scriptEditorReady) return;
    const code = pendingScriptEditorCode;
    if (typeof code !== "string") return;
    pendingScriptEditorCode = null;
    dispatchScriptEditorSetCode(code);
  };

  async function ensureScriptingWorkbook(): Promise<DocumentControllerWorkbookAdapter | null> {
    if (scriptingWorkbook) return scriptingWorkbook;
    if (!scriptingWorkbookInitPromise) {
      scriptingWorkbookInitPromise = (async () => {
        const mod = await loadScriptingWorkbookAdapterModule();
        if (!mod) return null;
        const adapter = new mod.DocumentControllerWorkbookAdapter(app.getDocument(), {
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
            app.refresh();
          },
        });
        scriptingWorkbook = adapter;
        return adapter;
      })();
    }
    return scriptingWorkbookInitPromise;
  }

  let macroRecorder: MacroRecorder | null = null;
  let macroRecorderInitPromise: Promise<MacroRecorder | null> | null = null;
  let stopMacroRecorderSelectionSync: (() => void) | null = null;

  async function ensureMacroRecorder(): Promise<MacroRecorder | null> {
    if (macroRecorder) return macroRecorder;
    if (!macroRecorderInitPromise) {
      macroRecorderInitPromise = (async () => {
        const workbook = await ensureScriptingWorkbook();
        if (!workbook) return null;

        const mod = await loadMacroRecorderModule();
        if (!mod) return null;

        const recorder = new mod.MacroRecorder(workbook);
        macroRecorder = recorder;
        activeMacroRecorder = recorder;

        // SpreadsheetApp selection changes live outside the DocumentController mutation stream. Emit
        // selectionChanged events from the UI so the macro recorder can capture selection steps.
        let lastSelectionKey = "";
        try {
          (workbook as any).events?.on?.("selectionChanged", (evt: any) => {
            lastSelectionKey = `${evt.sheetName}:${evt.address}`;
          });
        } catch {
          // ignore
        }

        stopMacroRecorderSelectionSync?.();
        stopMacroRecorderSelectionSync = app.subscribeSelection((selection) => {
          const first = selection.ranges[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
          const address = formatRangeAddress(first);
          const sheetName = app.getCurrentSheetId();
          const key = `${sheetName}:${address}`;
          if (key === lastSelectionKey) return;
          try {
            (workbook as any).events?.emit?.("selectionChanged", { sheetName, address });
          } catch {
            // ignore
          }
        });

        window.addEventListener(
          "unload",
          () => {
            stopMacroRecorderSelectionSync?.();
            stopMacroRecorderSelectionSync = null;
          },
          { once: true },
        );

        return recorder;
      })();
    }
    return macroRecorderInitPromise;
  }

  ensureMacroRecorderRef = ensureMacroRecorder;

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

    const highlights = app.getFormulaReferenceHighlights();

    if (highlights.length === 0) {
      secondaryGridView.grid.renderer.setReferenceHighlights(null);
      return;
    }

    // Split-view secondary pane always uses a shared-grid renderer with a 1x1 frozen
    // header row/col (row/column labels), even when the primary pane is in legacy mode.
    const headerRows = 1;
    const headerCols = 1;

    const gridHighlights = highlights.map((h) => {
      const startRow = Math.min(h.start.row, h.end.row);
      const endRow = Math.max(h.start.row, h.end.row);
      const startCol = Math.min(h.start.col, h.end.col);
      const endCol = Math.max(h.start.col, h.end.col);

      const gridRange: GridCellRange = {
        startRow: startRow + headerRows,
        endRow: endRow + headerRows + 1,
        startCol: startCol + headerCols,
        endCol: endCol + headerCols + 1,
      };

      return { range: gridRange, color: h.color, active: h.active };
    });

    secondaryGridView.grid.renderer.setReferenceHighlights(gridHighlights.length > 0 ? gridHighlights : null);
  };

  const syncSecondaryGridInteractionMode = () => {
    if (!secondaryGridView) return;
    const mode = app.isFormulaBarFormulaEditing() ? "rangeSelection" : "default";
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
  const unsubscribeSplitViewFormulaBarOverlaySync = app.onFormulaBarOverlayChange(() => syncSecondaryGridInteractionMode());
  window.addEventListener("unload", () => unsubscribeSplitViewFormulaBarOverlaySync());

  // Programmatic formula bar commit/cancel (e.g. File → Save calling `commitPendingEditsForCommand()`)
  // may not produce textarea blur/input events. Subscribe to SpreadsheetApp edit-state changes so we
  // always leave split-view range-selection mode when formula editing ends.
  const unsubscribeSplitViewEditStateSync = app.onEditStateChange(() => syncSecondaryGridInteractionModeSoon());
  window.addEventListener("unload", () => unsubscribeSplitViewEditStateSync());

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

  // SpreadsheetApp draws drawing overlays (pictures/shapes/imported charts) on its own canvas layers.
  // Some drawing-related UI updates do not flow through DocumentController mutations (e.g. async
  // imported chart model hydration, drawing selection changes). In split view, ensure the secondary
  // pane redraws its drawings overlay in response so the two panes stay visually consistent.
  const repaintSecondaryDrawings = () => {
    if (!secondaryGridView) return;
    // `scrollBy(0,0)` is treated as a "sync + notify" call in DesktopSharedGrid; it triggers
    // the scroll callback without changing scroll offsets, which is where SecondaryGridView
    // repaints its drawings overlay.
    secondaryGridView.grid.scrollBy(0, 0);
  };
  window.addEventListener("formula:drawings-changed", repaintSecondaryDrawings);
  window.addEventListener("formula:drawing-selection-changed", repaintSecondaryDrawings);

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
    const sharedProvider = app.getSharedGridProvider();
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

  function syncPrimarySelectionFromSecondary(): void {
    if (!secondaryGridView) return;
    if (splitSelectionSyncInProgress) return;

    const gridSelection = secondaryGridView.grid.renderer.getSelection();
    const gridRanges = secondaryGridView.grid.renderer.getSelectionRanges();
    const activeIndex = secondaryGridView.grid.renderer.getActiveSelectionIndex();
    if (!gridSelection || gridRanges.length === 0) return;

    splitSelectionSyncInProgress = true;
    try {
      // Sync via SpreadsheetApp so we preserve shared-grid multi-range selection when available,
      // while keeping the legacy grid fallback (single-range) consistent.
      //
      // Never cross-scroll or steal focus: selection sync should not disturb the destination pane.
      app.setSharedGridSelectionRanges(gridRanges, {
        activeIndex,
        activeCell: gridSelection,
        scrollIntoView: false,
        focus: false,
      });
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
      secondaryGridViewInitToken += 1;
      secondaryGridViewInitPromise = null;
      secondaryGridView?.destroy();
      secondaryGridView = null;
      app.setSplitViewSecondaryGridView(null);
      if (splitViewSecondaryIsEditing) {
        splitViewSecondaryIsEditing = false;
        renderStatusMode();
        syncTitlebar();
        emitSpreadsheetEditingChanged();
        scheduleRibbonSelectionFormatStateUpdate();
        renderSheetTabs();
        recomputeKeyboardContextKeys?.();
      }
      stopPrimarySplitPanePersistence();
      primaryPaneViewportRestored = false;
      window.__formulaSecondaryGrid = null;
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
      const createSecondaryGridView = (mod: SecondaryGridViewModule) => {
        const pane = layoutController.layout.splitView.panes.secondary;
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
        const limits = app.getGridLimits();
        const rowCount = Number.isInteger(limits.maxRows) ? limits.maxRows + 1 : DEFAULT_DESKTOP_LOAD_MAX_ROWS + 1;
        const colCount = Number.isInteger(limits.maxCols) ? limits.maxCols + 1 : DEFAULT_DESKTOP_LOAD_MAX_COLS + 1;

        const view = new mod.SecondaryGridView({
          container: gridSecondaryEl,
          provider: app.getSharedGridProvider() ?? undefined,
          imageResolver: app.getSharedGridImageResolver() ?? undefined,
          document: app.getDocument(),
          getSheetId: getSecondarySheetId,
          rowCount,
          colCount,
          showFormulas: () => app.getShowFormulas(),
          getComputedValue: (cell) => app.getCellComputedValueForSheet(getSecondarySheetId(), cell),
          getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
          images: app.getDrawingImages(),
          chartRenderer: app.getDrawingChartRenderer(),
          getSelectedDrawingId: () => app.getSelectedDrawingId(),
          onRequestRefresh: () => app.refresh(),
          onSelectionChange: () => syncPrimarySelectionFromSecondary(),
          onSelectionRangeChange: () => syncPrimarySelectionFromSecondary(),
          callbacks: app.getSharedGridRangeSelectionCallbacks(),
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
            emitSpreadsheetEditingChanged();
            scheduleRibbonSelectionFormatStateUpdate();
            renderSheetTabs();
            recomputeKeyboardContextKeys?.();
          },
        });

        secondaryGridView = view;
        app.setSplitViewSecondaryGridView(view);

        // Ensure the secondary selection reflects the current primary selection without
        // affecting either pane's scroll positions.
        if (lastSplitSelection) {
          const selection = lastSplitSelection;
          splitSelectionSyncInProgress = true;
          try {
            const ranges = selection.ranges.map((r) => gridRangeFromDocRange(r));
            const activeCell = { row: selection.active.row + SPLIT_HEADER_ROWS, col: selection.active.col + SPLIT_HEADER_COLS };
            view.grid.setSelectionRanges(ranges, {
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
      };

      if (secondaryGridViewModule) {
        createSecondaryGridView(secondaryGridViewModule);
      } else if (!secondaryGridViewInitPromise) {
        const initToken = ++secondaryGridViewInitToken;
        const initPromise = (async () => {
          const mod = await loadSecondaryGridViewModule();
          if (!mod) return;
          secondaryGridViewModule = mod;
          if (initToken !== secondaryGridViewInitToken) return;
          if (layoutController.layout.splitView.direction === "none") return;
          if (secondaryGridView) return;
          createSecondaryGridView(mod);
          // Re-render so CSS state + Playwright hooks reflect the newly created pane.
          renderSplitView();
        })();
        secondaryGridViewInitPromise = initPromise;
        void initPromise.finally(() => {
          if (secondaryGridViewInitPromise === initPromise) secondaryGridViewInitPromise = null;
        });
      }
    }

    // Expose for Playwright (secondary pane autofill).
    window.__formulaSecondaryGrid = secondaryGridView ? secondaryGridView.grid : null;
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
    const image = parseImageCellValue(value);
    if (image) return image.altText ?? "[Image]";
    return null;
  }

  function findSheetIdByName(name: string): string | null {
    const query = String(name ?? "").trim();
    if (!query) return null;

    // In collab mode, the authoritative sheet metadata lives in the CollabSession Yjs schema.
    // The UI sheet store may lag behind (e.g. when tests stub out `observeDeep`), so resolve
    // by scanning the session sheet list before we fall back to local stores.
    const collabSession = app.getCollabSession?.() ?? null;
    if (collabSession) {
      // Match backend uniqueness semantics (best-effort): NFKC normalize + uppercase.
      const normalize = (value: string): string => {
        try {
          return value.normalize("NFKC").toUpperCase();
        } catch {
          return value.toUpperCase();
        }
      };
      const queryCi = normalize(query);
      for (const sheet of listSheetsFromCollabSession(collabSession)) {
        if (sheet.id === query) return sheet.id;
        if (normalize(sheet.name) === queryCi) return sheet.id;
      }
    }

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
  const disposeKeyboardContextKeys = installKeyboardContextKeys({
    contextKeys,
    app,
    formulaBarRoot,
    sheetTabsRoot: sheetTabsRootEl,
    gridRoot,
    gridSecondaryRoot: gridSecondaryEl,
    isCommandPaletteOpen: () => {
      // The palette is mounted lazily; treat "overlay exists and is visible" as open.
      const overlay = document.querySelector<HTMLElement>(".command-palette-overlay");
      return overlay != null && overlay.hidden === false;
    },
    isSplitViewSecondaryEditing: () => splitViewSecondaryIsEditing,
  });
  recomputeKeyboardContextKeys = disposeKeyboardContextKeys.recompute;
  window.addEventListener("unload", () => disposeKeyboardContextKeys());
  type GridArea = "cell" | "rowHeader" | "colHeader" | "corner";
  let currentGridArea: GridArea = "cell";
  let currentContextMenuTarget: "grid" | "drawing" = "grid";

  let lastSelection: SelectionState | null = null;

  const updateContextKeysInternal = (selection?: SelectionState | null) => {
    const resolvedSelection = selection ?? lastSelection;
    if (!resolvedSelection) return;
    const sheetId = app.getCurrentSheetId();
    // Avoid resurrecting deleted sheets while responding to DocumentController change events.
    //
    // DocumentController lazily creates sheets when `getCell` is called. During undo/redo of
    // sheet-structure operations (add/delete/hide), the active sheet id can temporarily point at
    // a sheet that was just removed. Calling `getCell` in that window would re-materialize the
    // sheet without recording a history entry, breaking undo semantics and dirty tracking.
    const doc: any = app.getDocument();
    const sheetExistsInDoc = (() => {
      // Prefer `getSheetMeta` when available because it can differentiate between
      // an actually-deleted sheet vs. an active sheet id that hasn't been materialized
      // yet (DocumentController creates sheets lazily on access).
      if (typeof doc.getSheetMeta === "function") {
        try {
          return Boolean(doc.getSheetMeta(sheetId));
        } catch {
          // Best-effort: if metadata access fails, assume the sheet exists so we don't
          // break context key updates in environments that don't expose this API.
          return true;
        }
      }

      // Fallback for older controllers without `getSheetMeta`: if we have a non-empty list
      // of sheet ids and the active id is missing, treat the sheet as deleted.
      try {
        const ids = typeof doc.getSheetIds === "function" ? doc.getSheetIds() : [];
        if (Array.isArray(ids) && ids.length > 0) {
          return ids.includes(sheetId);
        }
      } catch {
        // Best-effort: if `getSheetIds` fails, continue and let downstream logic handle it.
      }

      return true;
    })();
    if (!sheetExistsInDoc) return;
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;
    const active = resolvedSelection.active;
    const cell = doc.getCell(sheetId, { row: active.row, col: active.col }) as any;
    const value = normalizeExtensionCellValue(cell?.value ?? null);
    const formula = typeof cell?.formula === "string" ? cell.formula : null;
    const selectionKeys = deriveSelectionContextKeys(resolvedSelection);

    contextKeys.batch({
      sheetName,
      ...selectionKeys,
      cellHasValue: (value != null && String(value).trim().length > 0) || (formula != null && formula.trim().length > 0),
      commentsPanelVisible: app.isCommentsPanelVisible(),
      cellHasComment: app.activeCellHasComment(),
      // "legacy" vs "shared" grid mode. Used to gate keybindings/command-palette commands that
      // depend on legacy-only features (e.g. outline-based AutoFilter MVP).
      "spreadsheet.gridMode": app.getGridMode(),
      // Ribbon AutoFilter MVP is view-local (in-memory store + outline.hidden.filter). Expose
      // a context key so command-palette entries like "Clear" / "Reapply" can be hidden when
      // AutoFilter is not enabled for the active sheet.
      "spreadsheet.hasAutoFilter": ribbonAutoFilterStore.hasAny(sheetId),
      "spreadsheet.isReadOnly": app.isReadOnly?.() === true,
      // Some roles (e.g. `commenter`) are read-only for cell edits but can still create comments.
      // Expose a dedicated key so shortcuts (Shift+F2) can be gated without relying on
      // `spreadsheet.isReadOnly`.
      "spreadsheet.canComment": (() => {
        let session: any = null;
        try {
          session = app.getCollabSession?.() ?? null;
        } catch {
          session = null;
        }
        if (!session) return true;
        try {
          return typeof session.canComment === "function" ? Boolean(session.canComment()) : false;
        } catch {
          return false;
        }
      })(),
      gridArea: currentGridArea,
      isRowHeader: currentGridArea === "rowHeader",
      isColHeader: currentGridArea === "colHeader",
      isCorner: currentGridArea === "corner",
    });
  };

  // Expose `updateContextKeys` to sheet helpers outside this dock/layout block.
  updateContextKeys = updateContextKeysInternal;

  app.subscribeSelection((selection) => {
    lastSelection = selection;
    updateContextKeys(selection);
  });
  app.getDocument().on("change", () => updateContextKeys());
  window.addEventListener("formula:comments-panel-visibility-changed", () => updateContextKeys());
  window.addEventListener("formula:comments-changed", () => updateContextKeys());
  window.addEventListener("formula:view-changed", () => updateContextKeys());
  window.addEventListener("formula:read-only-changed", () => updateContextKeys());
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
      /**
       * Optional formulas matrix for the selection.
       *
       * Note: the extension API runtime will synthesize this when absent, but that synthesis
       * involves allocating a full 2D array. For very large selections we include an empty array
       * to avoid catastrophic allocations in the worker runtime.
       */
      formulas?: Array<Array<string | null>>;
      /**
       * Indicates that the selection range was too large to safely materialize `values`/`formulas`
       * in memory, so the payload may contain empty matrices.
       */
      truncated?: boolean;
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

      let values: Array<Array<string | number | boolean | null>> = [];
      let formulas: Array<Array<string | null>> | undefined;
      let truncated = false;

      try {
        // Keep selectionChanged payloads consistent with the hard cap used by
        // `cells.getSelection/getRange`: extensions should not be able to trigger multi-million-cell
        // allocations by selecting an entire sheet/row/column.
        assertExtensionRangeWithinLimits(range, { label: "Selection" });

        values = [];
        for (let r = range.startRow; r <= range.endRow; r++) {
          const row: Array<string | number | boolean | null> = [];
          for (let c = range.startCol; c <= range.endCol; c++) {
            const cell = app.getDocument().getCell(rect.sheetId, { row: r, col: c }) as any;
            row.push(normalizeExtensionCellValue(cell?.value ?? null));
          }
          values.push(row);
        }
      } catch {
        // Best-effort: still emit the event so extensions can observe selection movement, but do
        // not materialize per-cell matrices for huge selections (Excel-scale sheets).
        truncated = true;
        values = [];
        // Defensive: `packages/extension-api` runtime auto-fills `formulas` with a null matrix when
        // absent. For huge ranges, that would allocate millions of elements, so provide an empty
        // matrix here to skip that path.
        formulas = [];
      }

      const payload: ExtensionSelectionChangedEvent = {
        sheetId: rect.sheetId,
        selection: { ...range, address, values, ...(formulas ? { formulas } : {}), ...(truncated ? { truncated } : {}) },
      };
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
  let extensionHostManager: DesktopExtensionHostManager | null = null;
  let extensionHostManagerInitPromise: Promise<DesktopExtensionHostManager> | null = null;

  // Extensions can access spreadsheet data via `formula.cells.*` / `formula.events.*` and then write
  // arbitrary text to the system clipboard via `formula.clipboard.writeText()`. SpreadsheetApp's
  // copy/cut handlers already enforce clipboard-copy DLP, but extensions would otherwise bypass it.
  //
  // BrowserExtensionHost tracks read taint (API reads + event payloads) and passes those ranges to this
  // optional `clipboardWriteGuard`, which runs *before* any clipboard write. Here we enforce clipboard-copy
  // DLP against any workbook ranges the extension has observed (taintedRanges).
  //
  // Selection-based clipboard-copy DLP enforcement (active-cell fallback) is handled separately by the
  // desktop adapter's `clipboardApi.writeText` implementation.
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

  const clipboardWriteGuard = async (params: { extensionId: string; taintedRanges: any[] }) => {
    try {
      const taintedRanges = Array.isArray(params.taintedRanges) ? params.taintedRanges : [];
      if (taintedRanges.length === 0) return;
      for (const raw of taintedRanges) {
        if (!raw || typeof raw !== "object") continue;
        const sheetId = typeof (raw as any).sheetId === "string" ? String((raw as any).sheetId) : "";
        if (!sheetId.trim()) continue;
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
        const message =
          typeof (err as any)?.message === "string" && String((err as any).message).trim().length > 0
            ? String((err as any).message)
            : "Clipboard copy is blocked by data loss prevention policy.";
        try {
          showToast(message, "error");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        throw err;
      }

      // Best-effort: guard failures should never take down clipboard writes.
      console.error(`[formula][desktop] clipboardWriteGuard error for ${String(params.extensionId)}:`, err);
    }
  };

  // The desktop UI is used both inside the Tauri shell and as a pure-web fallback (e2e, local dev).
  // Only expose real workbook lifecycle operations to extensions when the Tauri bridge is present;
  // otherwise let BrowserExtensionHost fall back to its in-memory stub workbook implementation.
  const hasTauriWorkbookBridge = hasTauriInvoke();

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

        const existingSheets = listSheetsFromCollabSession(collabSession);
        const existingIds = new Set(existingSheets.map((sheet) => sheet.id));

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

        const permission = getWorkbookMutationPermission(collabSession);
        if (!permission.allowed) {
          throw new Error(permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE);
        }

        const sheetsArray: any = (collabSession as any)?.sheets ?? null;
        const isYjsSheetsArray = Boolean(sheetsArray && typeof sheetsArray === "object" && sheetsArray.doc);
        const Y = isYjsSheetsArray ? await importYjs() : null;

        const activeIdx = existingSheets.findIndex((sheet) => sheet.id === activeId);
        const insertIndex = activeIdx >= 0 ? activeIdx + 1 : collabSession.sheets.length;

        collabSession.transactLocal(() => {
          // Prefer inserting a real Y.Map when `session.sheets` is backed by Yjs.
          // Some environments provide a lightweight `session.sheets` implementation that stores
          // plain JS objects and does not attach Yjs types to a Y.Doc; in that case insert a
          // plain `{ id, name, visibility }` entry instead.
          if (isYjsSheetsArray && Y) {
            const sheet = new Y.Map<unknown>();
            // Attach first, then populate fields (Yjs types warn when accessed before attachment).
            collabSession.sheets.insert(insertIndex, [sheet as any]);
            sheet.set("id", id);
            sheet.set("name", normalizedName);
            sheet.set("visibility", "visible");
          } else {
            collabSession.sheets.insert(insertIndex, [{ id, name: normalizedName, visibility: "visible" } as any]);
          }
        });

        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
        app.activateSheet(id);
        restoreFocusAfterSheetNavigation();
        // In collab mode the authoritative sheet list lives in the Yjs session (`session.sheets`).
        // `createSheet` mutates that list directly; ensure the desktop sheet store/UI are refreshed
        // immediately so follow-up extension calls (e.g. deleteSheet, getSheet) can resolve the new
        // sheet by name even when collab observers are stubbed in tests.
        syncSheetUi();
        updateContextKeys();
        return { id, name: normalizedName };
      }

      // Local (non-collab) behavior: update the UI sheet store first so the corresponding
      // DocumentController sheet op becomes undoable (Ctrl+Z/Ctrl+Y). The workbook sync bridge
      // will persist the structural change to the native backend.
      const existingIdCi = new Set(workbookSheetStore.listAll().map((s) => s.id.trim().toLowerCase()));
      const baseId = validatedName;
      let id = baseId;
      let counter = 1;
      while (existingIdCi.has(id.toLowerCase())) {
        counter += 1;
        id = `${baseId}-${counter}`;
      }

      workbookSheetStore.addAfter(activeId, { id, name: validatedName });

      // Best-effort: ensure the sheet exists for follow-up APIs.
      try {
        doc.getCell(id, { row: 0, col: 0 });
      } catch {
        // ignore
      }

      app.activateSheet(id);
      restoreFocusAfterSheetNavigation();
      const storedName = workbookSheetStore.getName(id);
      return { id, name: storedName ?? validatedName };
    },
    async renameSheet(_oldName: string, _newName: string) {
      const oldName = String(_oldName ?? "");
      const sheetId = findSheetIdByName(oldName);
      if (!sheetId) {
        throw new Error(`Unknown sheet: ${oldName}`);
      }
      return await renameSheetById(sheetId, String(_newName ?? ""));
    },
    async deleteSheet(_name: string) {
      const name = String(_name ?? "");
      const sheetId = findSheetIdByName(name);
      if (!sheetId) {
        throw new Error(`Unknown sheet: ${name}`);
      }

      const collabSession = app.getCollabSession?.() ?? null;
      if (collabSession) {
        const permission = getWorkbookMutationPermission(collabSession);
        if (!permission.allowed) {
          throw new Error(permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE);
        }
      }
 
      const doc = app.getDocument();
      const wasActive = app.getCurrentSheetId() === sheetId;
      const deletedName = workbookSheetStore.getName(sheetId) ?? sheetId;
      const allSheets = workbookSheetStore.listAll();
      const sheetOrder = allSheets.map((s) => s.name);
      const nextActiveId = wasActive ? pickAdjacentVisibleSheetId(allSheets, sheetId) : null;
 
      // Update sheet metadata to enforce workbook invariants (e.g. last-sheet guard) and drive UI
      // reconciliation. The workbook sync bridge will persist the structural change to the native backend.
      workbookSheetStore.remove(sheetId);
      if (wasActive) {
        const fallback = workbookSheetStore.listVisible().at(0)?.id ?? workbookSheetStore.listAll().at(0)?.id ?? null;
        const next = nextActiveId ?? fallback;
        if (next && next !== sheetId && app.getCurrentSheetId() !== next) {
          app.activateSheet(next);
          restoreFocusAfterSheetNavigation();
        }
      }

      try {
        rewriteDocumentFormulasForSheetDelete(doc as any, deletedName, sheetOrder);
      } catch {
        // ignore
      }
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
      if (app.isReadOnly?.() === true) {
        throw new Error("Read-only: you don't have permission to edit cell contents.");
      }
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
      if (app.isReadOnly?.() === true) {
        throw new Error("Read-only: you don't have permission to edit cell contents.");
      }
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
      return getWorkbookSnapshotForExtensions();
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

  const ensureExtensionHostManager = async (): Promise<DesktopExtensionHostManager> => {
    if (extensionHostManager) return extensionHostManager;
    if (!extensionHostManagerInitPromise) {
      extensionHostManagerInitPromise = (async () => {
        const { DesktopExtensionHostManager } = await importExtensionHostManagerModule();
        const manager = new DesktopExtensionHostManager({
          engineVersion: "1.0.0",
          spreadsheetApi: extensionSpreadsheetApi,
          clipboardWriteGuard,
          clipboardApi: {
            readText: async () => {
              const provider = await getClipboardProvider();
              const { text } = await provider.read();
              return text ?? "";
            },
            writeText: async (text: string) => {
              // Extensions can write arbitrary text to the system clipboard via `formula.clipboard.writeText()`.
              //
              // Enforce clipboard-copy DLP against the current UI selection (active-cell fallback) so
              // extensions cannot bypass SpreadsheetApp's copy/cut policy enforcement by directly writing
              // to the system clipboard.
              //
              // Note: DLP enforcement based on per-extension taint tracking (ranges read via `cells.*` and
              // `events.*`) is handled separately by `clipboardWriteGuard`.
              try {
                const sheetId = app.getCurrentSheetId();
                const active = app.getActiveCell();
                const selectionRanges = app.getSelectionRanges();
                const rangesToCheck =
                  selectionRanges.length > 0
                    ? selectionRanges
                    : [{ startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col }];

                for (const range of rangesToCheck) {
                  enforceExtensionClipboardDlpForRange({ sheetId, range: normalizeSelectionRange(range) });
                }
              } catch (err) {
                const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
                if (isDlpViolation) {
                  const message =
                    typeof (err as any)?.message === "string" && String((err as any).message).trim().length > 0
                      ? String((err as any).message)
                      : "Clipboard copy is blocked by data loss prevention policy.";
                  try {
                    showToast(message, "error");
                  } catch {
                    // `showToast` requires a #toast-root; unit tests don't always include it.
                  }
                  // Surface DLP violations to the extension worker as a normal Error (serialized across
                  // the worker boundary), preserving the `name` so extensions can detect policy blocks.
                  if (err instanceof Error) {
                    throw err;
                  }
                  const normalized = new Error(message);
                  normalized.name = "DlpViolationError";
                  throw normalized;
                }
                throw err;
              }
              const provider = await getClipboardProvider();
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

        extensionHostManager = manager;
        extensionHostManagerForE2e = manager;
        try {
          // Surface to e2e/debug tooling as soon as the host is created.
          (window as any).__formulaExtensionHostManager = manager;
          (window as any).__formulaExtensionHost = manager.host ?? null;
        } catch {
          // ignore
        }

        // Bridge extension webviews into the dock layout once the host exists.
        extensionPanelBridge = new ExtensionPanelBridge({
          host: manager.host as any,
          panelRegistry,
          layoutController: layoutController as any,
        });

        return manager;
      })();
    }
    return extensionHostManagerInitPromise;
  };

  // Pre-initialize the extension host manager (without loading extensions) so e2e + UI surfaces
  // can safely reference `window.__formulaExtensionHostManager` / `window.__formulaExtensionHost`
  // without having to first open the Extensions panel.
  //
  // This does *not* load built-in or installed extensions; worker startup remains lazy and is
  // still gated behind `ensureExtensionsLoaded()`.
  void ensureExtensionHostManager().catch((err) => {
    console.error("Failed to initialize extension host manager:", err);
  });

  const focusAfterSheetNavigationFromCommand = (): void => {
    if (!shouldRestoreFocusAfterSheetNavigation()) return;
    if (contextKeys.get(KeyboardContextKeyIds.focusInSheetTabs) === true) {
      focusActiveSheetTab();
      return;
    }
    app.focusAfterSheetNavigation();
  };
  focusAfterSheetNavigationFromCommandRef = focusAfterSheetNavigationFromCommand;
  registerEncryptionUiCommands({ commandRegistry, app });

  // Loading extensions spins up an additional Worker for each workbook tab. That can be
  // expensive in e2e + low-spec environments, so defer actually loading extensions until
  // the user opens the Extensions panel or invokes an extension command.
  let extensionsLoadPromise: Promise<void> | null = null;

  const ensureExtensionsLoaded = async () => {
    const manager = await ensureExtensionHostManager();
    if (manager.ready) return;
    if (!extensionsLoadPromise) {
      extensionsLoadPromise = manager.loadBuiltInExtensions();
    }
    await extensionsLoadPromise;

    // Once extensions are loaded, sync their contributions into the UI surfaces that
    // were initialized earlier during boot (commands, panels, keybindings).
    try {
      updateKeybindingsRef?.();
      syncContributedCommandsRef?.();
      syncContributedPanelsRef?.();
    } catch {
      // ignore
    }
    try {
      activateOpenExtensionPanels();
    } catch {
      // ignore
    }
  };

  const executeExtensionCommand = async (commandId: string, ...args: any[]) => {
    try {
      await ensureExtensionsLoaded();
      syncContributedCommands();
      await commandRegistry.executeCommand(commandId, ...args);
    } catch (err) {
      // DLP policy violations are already surfaced via a dedicated toast (e.g. clipboard copy blocked).
      // Avoid double-toasting "Command failed" for expected policy restrictions.
      if ((err as any)?.name === "DlpViolationError") return;
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    }
  };

  const executeCommand = (commandId: string) => {
    const cmd = commandRegistry.getCommand(commandId);
    if (cmd?.source.kind === "builtin") {
      executeBuiltinCommand(commandId);
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
    const manager = extensionHostManager;
    if (!manager?.ready || manager.error) return;
    try {
      commandRegistry.setExtensionCommands(
        manager.getContributedCommands(),
        (commandId, ...args) => manager.executeCommand(commandId, ...args),
      );
    } catch (err) {
      showToast(`Failed to register extension commands: ${String((err as any)?.message ?? err)}`, "error");
    }
  };

  const syncContributedPanels = () => {
    const manager = extensionHostManager;
    if (!manager?.ready || manager.error) return;
    const contributed = manager.getContributedPanels() as Array<{ extensionId: string; id: string; title: string; icon?: string | null }>;
    const contributedIds = new Set(contributed.map((p) => p.id));

    // Update the synchronous seed store from loaded extension manifests.
    if (contributedPanelsSeedStorage) {
      try {
        const extensions = manager.host.listExtensions?.() ?? [];
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
    // Built-in keybindings use `when` clauses + focus/editing context keys to decide
    // whether a shortcut should run. Allow those built-ins to dispatch even when the
    // keydown target is an input/textarea (e.g. command palette).
    //
    // Extension keybindings remain blocked in text inputs for now as a conservative default.
    ignoreInputTargets: "extensions",
    onBeforeExecuteCommand: async (_commandId, source) => {
      if (source.kind !== "extension") return;
      await ensureExtensionsLoaded();
      syncContributedCommands();
    },
    onCommandError: (commandId, err) => {
      if (isSpreadsheetEditingCommandBlockedError(err)) return;
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    },
  });
  keybindingService.setBuiltinKeybindings(builtinKeybindingHints);
  const commandKeybindingDisplayIndex = keybindingService.getCommandKeybindingDisplayIndex();
  const commandKeybindingAriaIndex = keybindingService.getCommandKeybindingAriaIndex();

  const updateRibbonShortcuts = () => {
    ribbonShortcutById = deriveRibbonShortcutById(commandKeybindingDisplayIndex);
    ribbonAriaKeyShortcutsById = deriveRibbonAriaKeyShortcutsById(commandKeybindingAriaIndex);
    // Preserve current pressed/label/disabled state while updating shortcut hints.
    setRibbonUiState({
      ...getRibbonUiStateSnapshot(),
      shortcutById: ribbonShortcutById,
      ariaKeyShortcutsById: ribbonAriaKeyShortcutsById,
    });
  };
  updateRibbonShortcuts();
  // Split dispatch across phases:
  // - Capture: built-in keybindings only (needed for some global shortcuts).
  // - Bubble: extension keybindings only, so SpreadsheetApp can `preventDefault()` first for
  //   grid-local typing/navigation and prevent extensions from stealing those keys.
  window.addEventListener(
    "keydown",
    (e) => {
      void keybindingService.dispatchKeydown(e, { allowBuiltins: true, allowExtensions: false });
    },
    { capture: true },
  );
  window.addEventListener("keydown", (e) => {
    void keybindingService.dispatchKeydown(e, { allowBuiltins: false, allowExtensions: true });
  });

  const updateKeybindings = () => {
    const manager = extensionHostManager;
    const contributed =
      manager && manager.ready && !manager.error ? (manager.getContributedKeybindings() as ContributedKeybinding[]) : [];
    keybindingService.setExtensionKeybindings(contributed);
    updateRibbonShortcuts();
  };

  const activateOpenExtensionPanels = () => {
    const manager = extensionHostManager;
    if (!manager?.ready || manager.error) return;
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
    if (!extensionHostManager || (!extensionHostManager.ready && !extensionsLoadPromise)) return;

    void ensureExtensionsLoaded()
      .then(async () => {
        try {
          // Pick up any installs that happened after the initial `loadAllInstalled()` pass.
          await extensionHostManager?.getMarketplaceExtensionManager().loadAllInstalled();
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
      currentContextMenuTarget = "grid";
      updateContextKeys();

      // Best-effort: restore focus to the appropriate editing surface after closing.
      // (Typically the grid, but preserve formula-bar focus during range-selection workflows.)
      app.focusAfterSheetNavigation();
    },
  });
  sharedContextMenu = contextMenu;

  function executeBuiltinCommand(commandId: string, ...args: any[]): void {
    void commandRegistry.executeCommand(commandId, ...args).catch((err) => {
      // DLP policy violations are already surfaced via a dedicated toast (e.g. clipboard copy blocked).
      // Avoid double-toasting "Command failed" for expected policy restrictions.
      if ((err as any)?.name === "DlpViolationError") return;
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    });
  }

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
    return app.hitTestGridAreaAtClientPoint(clientX, clientY);
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

  const applyRowHeight = () => promptAndApplyAxisSizing(app, "rowHeight", { isEditing: () => isSpreadsheetEditing() });
  const applyColWidth = () => promptAndApplyAxisSizing(app, "colWidth", { isEditing: () => isSpreadsheetEditing() });

  const buildGridContextMenuItems = (): ContextMenuItem[] => {
    const allowEditCommands = !isSpreadsheetEditing();
    // View-only mutations (row/col sizing) are allowed in read-only collab roles (viewer/commenter),
    // but must remain local-only (collab binders prevent persisting these into shared Yjs state).
    const allowViewMutations = allowEditCommands;
    const isReadOnly = app.isReadOnly?.() === true;
    // Workbook/cell mutations should remain disabled in read-only roles.
    const allowSheetMutations = allowEditCommands && !isReadOnly;
    // In read-only collab sessions (viewer/commenter), allow formatting only when the selection is an
    // entire row/column/sheet band (formatting defaults).
    const allowReadOnlyFormattingDefaults = (() => {
      if (!isReadOnly) return false;
      const selection = app.getSelectionRanges();
      const limits = getGridLimitsForFormatting();
      const decision = evaluateFormattingSelectionSize(selection, limits, { maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT });
      return decision.allRangesBand;
    })();
    const allowFormattingCommands = allowEditCommands && (!isReadOnly || allowReadOnlyFormattingDefaults);
    const canComment = (() => {
      let session: any = null;
      try {
        session = app.getCollabSession?.() ?? null;
      } catch {
        session = null;
      }
      if (!session) return true;
      try {
        return typeof session.canComment === "function" ? Boolean(session.canComment()) : false;
      } catch {
        return false;
      }
    })();

    let menuItems: ContextMenuItem[] = [];
    if (currentGridArea === "rowHeader") {
      menuItems = [
        { type: "item", label: "Row Height…", enabled: allowViewMutations, onSelect: applyRowHeight },
        {
          type: "item",
          label: "Hide",
          enabled: allowSheetMutations,
          onSelect: () => {
            // `selectedRowIndices()` enumerates every row in every selection range into a Set.
            // Keep this bounded so Excel-scale select-all (1M rows) can't cause huge allocations.
            const selection = app.getSelectionRanges();
            let rowUpperBound = 0;
            for (const range of selection) {
              const r = normalizeSelectionRange(range);
              rowUpperBound += Math.max(0, r.endRow - r.startRow + 1);
              if (rowUpperBound > MAX_AXIS_RESIZE_INDICES) break;
            }
            if (rowUpperBound > MAX_AXIS_RESIZE_INDICES) {
              showToast("Selection too large to hide rows. Select fewer rows and try again.", "warning");
              return;
            }
            app.hideRows(selectedRowIndices(selection));
          },
        },
        {
          type: "item",
          label: "Unhide",
          enabled: allowSheetMutations,
          onSelect: () => {
            const selection = app.getSelectionRanges();
            let rowUpperBound = 0;
            for (const range of selection) {
              const r = normalizeSelectionRange(range);
              rowUpperBound += Math.max(0, r.endRow - r.startRow + 1);
              if (rowUpperBound > MAX_AXIS_RESIZE_INDICES) break;
            }
            if (rowUpperBound > MAX_AXIS_RESIZE_INDICES) {
              showToast("Selection too large to unhide rows. Select fewer rows and try again.", "warning");
              return;
            }
            app.unhideRows(selectedRowIndices(selection));
          },
        },
      ];
    } else if (currentGridArea === "colHeader") {
      menuItems = [
        { type: "item", label: "Column Width…", enabled: allowViewMutations, onSelect: applyColWidth },
        {
          type: "item",
          label: "Hide",
          enabled: allowSheetMutations,
          onSelect: () => {
            // `selectedColIndices()` enumerates every column in every selection range into a Set.
            // Keep this bounded so Excel-scale select-all (16k cols) can't cause huge allocations.
            const selection = app.getSelectionRanges();
            let colUpperBound = 0;
            for (const range of selection) {
              const r = normalizeSelectionRange(range);
              colUpperBound += Math.max(0, r.endCol - r.startCol + 1);
              if (colUpperBound > MAX_AXIS_RESIZE_INDICES) break;
            }
            if (colUpperBound > MAX_AXIS_RESIZE_INDICES) {
              showToast("Selection too large to hide columns. Select fewer columns and try again.", "warning");
              return;
            }
            app.hideCols(selectedColIndices(selection));
          },
        },
        {
          type: "item",
          label: "Unhide",
          enabled: allowSheetMutations,
          onSelect: () => {
            const selection = app.getSelectionRanges();
            let colUpperBound = 0;
            for (const range of selection) {
              const r = normalizeSelectionRange(range);
              colUpperBound += Math.max(0, r.endCol - r.startCol + 1);
              if (colUpperBound > MAX_AXIS_RESIZE_INDICES) break;
            }
            if (colUpperBound > MAX_AXIS_RESIZE_INDICES) {
              showToast("Selection too large to unhide columns. Select fewer columns and try again.", "warning");
              return;
            }
            app.unhideCols(selectedColIndices(selection));
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
          enabled: allowSheetMutations,
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
          enabled: allowSheetMutations,
          shortcut: getPrimaryCommandKeybindingDisplay("clipboard.paste", commandKeybindingDisplayIndex) ?? primaryShortcut("V"),
          onSelect: () => executeBuiltinCommand("clipboard.paste"),
        },
        { type: "separator" },
        {
          type: "submenu",
          label: t("clipboard.pasteSpecial.title"),
          enabled: allowSheetMutations,
          // Retain the shortcut hint for the Paste Special command (still available via
          // keybinding/command palette), even though the context menu now exposes direct
          // paste-special modes via a submenu.
          shortcut:
            getPrimaryCommandKeybindingDisplay("clipboard.pasteSpecial", commandKeybindingDisplayIndex) ??
            (isMac ? "⇧⌘V" : "Ctrl+Shift+V"),
          items: getPasteSpecialMenuItems().map((item) => ({
            type: "item",
            label: item.label,
            onSelect: () => executeBuiltinCommand(`clipboard.pasteSpecial.${item.mode}`),
          })),
        },
        { type: "separator" },
        {
          type: "item",
          label: t("menu.clearContents"),
          enabled: (() => {
            if (!allowSheetMutations) return false;

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
          enabled: allowSheetMutations,
          shortcut: getPrimaryCommandKeybindingDisplay("ai.inlineEdit", commandKeybindingDisplayIndex) ?? primaryShortcut("K"),
          onSelect: () => executeBuiltinCommand("ai.inlineEdit"),
        },
        { type: "separator" },
        {
          type: "submenu",
          label: t("menu.format"),
          enabled: allowFormattingCommands,
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
                (isMac ? "⌃I" : "Ctrl+I"),
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
          enabled: allowEditCommands && canComment,
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

    const manager = extensionHostManager;
    if (!manager || !manager.ready) {
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

    if (manager.error) {
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
    const contributed = resolveMenuItems(manager.getContributedMenu(menuId), contextKeys.asLookup());
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
    if (currentContextMenuTarget === "drawing") {
      contextMenu.update(buildDrawingContextMenuItems({ app, isEditing: isSpreadsheetEditing() }));
    } else {
      contextMenu.update(buildGridContextMenuItems());
    }
  });

  let contextMenuSession = 0;

  const openGridContextMenuAtPoint = (x: number, y: number) => {
    currentContextMenuTarget = "grid";
    const session = (contextMenuSession += 1);
    currentContextMenuTarget = "grid";
    contextMenu.open({ x, y, items: buildGridContextMenuItems() });

    // Extensions are lazy-loaded to keep startup light. Right-clicking should still
    // surface extension-contributed context menu items, so load them on-demand and
    // update the menu if it is still open.
    if (!extensionHostManager?.ready) {
      void ensureExtensionsLoaded()
        .then(() => {
          if (session !== contextMenuSession) return;
          if (!contextMenu.isOpen()) return;
          if (!extensionHostManager?.ready) return;
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

    if (
      tryOpenDrawingContextMenuAtClientPoint({
        app,
        contextMenu,
        clientX: anchorX,
        clientY: anchorY,
        isEditing: isSpreadsheetEditing(),
        onWillOpen: () => {
          // Invalidate any pending extension-driven updates from a previously-open grid context menu.
          // (Those updates should never overwrite a drawing-specific menu.)
          contextMenuSession += 1;
          currentContextMenuTarget = "drawing";
        },
      })
    ) {
      return;
    }

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

  // In split view, the secondary pane is a CanvasGrid instance mounted outside of SpreadsheetApp's
  // `.grid-root`. SpreadsheetApp's capture-phase drawing hit testing cannot tag pointer events for
  // the secondary pane, so right-clicking on a drawing would otherwise move the active cell
  // underneath it (DesktopSharedGrid's default context-click behavior).
  //
  // Tag secondary-pane pointer events that hit drawings so DesktopSharedGrid can ignore them
  // (Excel-like behavior).
  const isMacPlatform = resolveIsMacPlatform();
  gridSecondaryEl.addEventListener(
    "pointerdown",
    (e) => {
      tagDrawingContextClickPointerDown(e, app, { isMacPlatform, requireCanvasTarget: true });
    },
    { capture: true, passive: true },
  );

  gridSecondaryEl.addEventListener("contextmenu", (e) => {
    // Always prevent the native context menu; we render our own.
    e.preventDefault();

    const anchorX = e.clientX;
    const anchorY = e.clientY;

    if (
      tryOpenDrawingContextMenuAtClientPoint({
        app,
        contextMenu,
        clientX: anchorX,
        clientY: anchorY,
        isEditing: isSpreadsheetEditing(),
        onWillOpen: () => {
          // Invalidate any pending extension-driven updates from a previously-open grid context menu.
          // (Those updates should never overwrite a drawing-specific menu.)
          contextMenuSession += 1;
          currentContextMenuTarget = "drawing";
        },
      })
    ) {
      return;
    }

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

  // Central command entrypoint so Shift+F10 / ContextMenu can migrate to KeybindingService
  // cleanly (e.g. if keybinding dispatch moves earlier in the capture phase).
  commandRegistry.registerBuiltinCommand(
    "ui.openContextMenu",
    t("command.ui.openContextMenu"),
    () => {
      // Fail-closed: only allow opening the grid context menu when we can
      // confidently say focus is not inside a text input.
      if (contextKeys.get("focus.inTextInput") !== false) return;

      // If *any* context menu is already open, let it manage focus/keyboard handling.
      // This avoids opening the grid context menu while another menu (e.g. sheet tabs)
      // is active.
      const openContextMenu = document.querySelector<HTMLElement>(".context-menu-overlay:not([hidden])");
      if (openContextMenu) return;

      // When focus is on a sheet tab, Shift+F10 / ContextMenu should behave like Excel and
      // open the sheet-tab context menu instead of the active-cell context menu.
      // (The tab strip is responsible for handling these keys.)
      const target = document.activeElement as HTMLElement | null;
      if (target?.closest?.('#sheet-tabs button[role="tab"]')) return;

      openGridContextMenuAtActiveCell();
    },
    {
      category: t("commandCategory.view"),
      icon: null,
      description: t("commandDescription.ui.openContextMenu"),
      keywords: ["context menu", "menu"],
    },
  );

  window.addEventListener(
    "keydown",
    (e) => {
      if (e.defaultPrevented) return;
      // Fail-closed: only allow opening the grid context menu when we can
      // confidently say focus is not inside a text input.
      if (contextKeys.get(KeyboardContextKeyIds.focusInTextInput) !== false) return;

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

  let panelBodyRenderer: PanelBodyRenderer | null = null;
  let panelBodyRendererInitPromise: Promise<PanelBodyRenderer | null> | null = null;
  let macrosPanelRenderToken = 0;

  const panelBodyRendererOptions: PanelBodyRendererOptions = {
    getDocumentController: () => app.getDocument(),
    getSpreadsheetApp: () => app,
    getActiveSheetId: () => app.getCurrentSheetId(),
    getSearchWorkbook: () => app.getSearchWorkbook(),
    getCharts: () => app.listCharts(),
    getSelectedChartId: () => app.getSelectedChartId(),
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
        activeRow: selection.activeRow,
        activeCol: selection.activeCol,
      };
    },
    workbookId,
    getWorkbookId: () => activePanelWorkbookId,
    getCollabSession: () => app.getCollabSession(),
    invoke:
      hasTauriInvoke()
        ? (cmd, args) => {
            const baseInvoke = getTauriInvokeOrNull();
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
    getExtensionPanelBridge: () => extensionPanelBridge,
    getExtensionHostManager: () => extensionHostManager,
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
      const renderToken = ++macrosPanelRenderToken;
      (body as any).__formulaMacrosRenderToken = renderToken;
      body.textContent = "Loading macros…";

      const schedule =
        typeof queueMicrotask === "function" ? queueMicrotask : (cb: () => void) => window.setTimeout(cb, 0);
      schedule(() => {
        void (async () => {
          const isStale = () => (body as any).__formulaMacrosRenderToken !== renderToken;

          const macrosMod = await loadMacrosModule();
          if (!macrosMod) {
            if (isStale()) return;
            body.textContent = "Failed to load macros.";
            return;
          }

          const recorderMod = await loadMacroRecorderModule();

          if (isStale()) return;
          body.replaceChildren();
          body.classList.add("macros-panel");

          const recorderPanel = document.createElement("div");
          recorderPanel.className = "macros-panel__recorder";

          const runnerPanel = document.createElement("div");
          runnerPanel.className = "macros-panel__runner";

          body.appendChild(recorderPanel);
          body.appendChild(runnerPanel);

          const scriptStorageId = activePanelWorkbookId;

          const scriptBackend = new macrosMod.WebMacroBackend({
            getDocumentController: () => app.getDocument(),
            getActiveSheetId: () => app.getCurrentSheetId(),
            sheetNameResolver,
          });

          const scriptBackendForWorkbook = {
            listMacros: async (_workbookId: string) => await scriptBackend.listMacros(scriptStorageId),
            getMacroSecurityStatus: async (_workbookId: string) => await scriptBackend.getMacroSecurityStatus(scriptStorageId),
            setMacroTrust: async (_workbookId: string, decision: MacroTrustDecision) =>
              await scriptBackend.setMacroTrust(scriptStorageId, decision),
            runMacro: async (request: MacroRunRequest) => await scriptBackend.runMacro({ ...request, workbookId: scriptStorageId }),
          };

          // Prefer a composite backend in desktop builds: run VBA via Tauri (when available),
          // and run modern scripts (TypeScript/Python) via the web backend + local storage.
          const backend = (() => {
            if (!isTauriInvokeAvailable()) return scriptBackendForWorkbook;
            try {
              const baseBackend = new macrosMod.TauriMacroBackend({ invoke: queuedInvoke ?? undefined });
              const tauriBackend = macrosMod.wrapTauriMacroBackendWithUiContext(
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
                },
              );

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
                    return await scriptBackend.runMacro({ ...request, workbookId: scriptStorageId });
                  }

                  return await tauriBackend.runMacro(request);
                },
              };
            } catch (err) {
              console.warn("[formula][desktop] Failed to create Tauri macros backend; falling back to web backend:", err);
              return scriptBackendForWorkbook;
            }
          })();

          // Match ribbon edit-mode/read-only disabling: macros should not run while a cell/formula edit
          // is active, and should be blocked for read-only collab roles. Keep this guard close to the
          // panel wiring so any macro entrypoint (Tauri VBA or web scripts) follows the same policy.
          const guardedBackend = {
            ...backend,
            runMacro: async (request: MacroRunRequest) => {
              if (isSpreadsheetEditing()) {
                throw new Error("Finish editing before running a macro.");
              }
              if (app.isReadOnly?.() === true) {
                throw new Error(READ_ONLY_SHEET_MUTATION_MESSAGE);
              }
              return await backend.runMacro(request);
            },
          };

          const focusMacrosPanelElement = (el: HTMLElement | null) => {
            if (!el) return;
            try {
              el.focus();
            } catch {
              // Best-effort: ignore focus errors (e.g. element not focusable in a headless environment).
            }
          };

           const refreshRunner = async () => {
             await macrosMod.renderMacroRunner(runnerPanel, guardedBackend as any, workbookId, {
               isEditing: () => isSpreadsheetEditing(),
               isReadOnly: () => app.isReadOnly?.() === true,
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
              const provider = await getClipboardProvider();
              await provider.write({ text });
            } catch {
              // Fall back to execCommand in case Clipboard API permissions are unavailable.
              // This is best-effort; ignore failures.
            }

            // Clipboard provider writes are best-effort, and in web contexts clipboard access
            // can still be permission-gated. Keep an execCommand fallback when not running
            // under Tauri to preserve "Copy" behavior in dev/preview browsers.
            const isTauri = hasTauri();
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
            pendingScriptEditorCode = code;
            const placement = getPanelPlacement(layoutController.layout, PanelIds.SCRIPT_EDITOR);
            if (placement.kind === "closed") layoutController.openPanel(PanelIds.SCRIPT_EDITOR);
            // Allow the layout to mount the panel before we attempt to populate the editor.
            // If the panel is still being lazy-loaded, the mount path will flush the pending code once ready.
            requestAnimationFrame(() => requestAnimationFrame(flushPendingScriptEditorCode));
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

          let recorder: MacroRecorder | null = null;
          let recorderLoaded = false;

          const currentActions = () => recorder?.getOptimizedActions() ?? [];

          const updateRecorderUi = () => {
            if (!recorderLoaded) {
              status.textContent = "Loading…";
              meta.textContent = "";
              preview.textContent = "";
              return;
            }
            if (!recorder) {
              status.textContent = "Macro recorder not available.";
              meta.textContent = "";
              preview.textContent = "";
              return;
            }

            status.textContent = recorder.recording ? "Recording…" : "Not recording";
            const raw = recorder.getRawActions();
            const optimized = recorder.getOptimizedActions();
            meta.textContent = `Recorded: ${raw.length} steps · Optimized: ${optimized.length} steps`;
            preview.textContent = optimized.length ? JSON.stringify(optimized, null, 2) : "(no recorded actions)";
          };

          startButton.onclick = () => {
            if (!recorder) return;
            recorder.start();
            updateRecorderUi();
          };

          stopButton.onclick = () => {
            recorder?.stop();
            updateRecorderUi();
          };

          copyTsButton.onclick = async () => {
            if (!recorderMod) return;
            const actions = currentActions();
            if (actions.length === 0) return;
            await copyText(recorderMod.generateTypeScriptMacro(actions));
          };

          copyPyButton.onclick = async () => {
            if (!recorderMod) return;
            const actions = currentActions();
            if (actions.length === 0) return;
            await copyText(recorderMod.generatePythonMacro(actions));
          };

          openScriptEditorButton.onclick = () => {
            if (!recorderMod) return;
            const actions = currentActions();
            if (actions.length === 0) return;
            openScriptEditor(recorderMod.generateTypeScriptMacro(actions));
          };

          saveButton.onclick = async () => {
            if (!recorderMod) return;
            const actions = currentActions();
            if (actions.length === 0) return;
            const name = await showInputBox({ prompt: "Macro name:", value: "Recorded Macro" });
            if (!name) return;
            saveScriptsToStorage(name, { ts: recorderMod.generateTypeScriptMacro(actions), py: recorderMod.generatePythonMacro(actions) });
            await refreshRunner();
          };

          updateRecorderUi();

          void ensureMacroRecorder()
            .then((value) => {
              recorderLoaded = true;
              recorder = value;
              updateRecorderUi();
            })
            .catch(() => {
              recorderLoaded = true;
              recorder = null;
              updateRecorderUi();
            });

          void refreshRunner().catch((err) => {
            runnerPanel.textContent = `Failed to load macros: ${String(err)}`;
          });
        })().catch((err) => {
          if ((body as any).__formulaMacrosRenderToken !== renderToken) return;
          body.textContent = `Failed to load macros: ${String(err)}`;
        });
      });
    },
  };

  async function ensurePanelBodyRenderer(): Promise<PanelBodyRenderer | null> {
    if (panelBodyRenderer) return panelBodyRenderer;
    if (!panelBodyRendererInitPromise) {
      panelBodyRendererInitPromise = (async () => {
        const mod = await loadPanelBodyRendererModule();
        if (!mod) return null;
        try {
          const renderer = mod.createPanelBodyRenderer(panelBodyRendererOptions);
          panelBodyRenderer = renderer;
          return renderer;
        } catch (err) {
          panelBodyRenderer = null;
          panelBodyRendererInitPromise = null;
          reportLazyImportFailure("Panels", err);
          return null;
        }
      })();
    }
    const renderer = await panelBodyRendererInitPromise;
    if (!renderer) {
      // Allow retry (e.g. transient chunk load failures).
      panelBodyRendererInitPromise = null;
    }
    return renderer;
  }

  function renderPanelBody(panelId: string, body: HTMLDivElement) {
    if (panelId === PanelIds.SCRIPT_EDITOR) {
      let mount = panelMounts.get(panelId);
      if (!mount) {
        const container = document.createElement("div");
        container.className = "panel-body__container";
        container.textContent = "Loading Script Editor…";

        let disposed = false;
        mount = {
          container,
          dispose: () => {
            disposed = true;
            scriptEditorReady = false;
            container.innerHTML = "";
          },
        };
        panelMounts.set(panelId, mount);

        void (async () => {
          const [mod, workbook] = await Promise.all([loadScriptEditorPanelModule(), ensureScriptingWorkbook()]);
          if (disposed) return;
          if (!mod || !workbook) {
            container.textContent = "Script Editor unavailable.";
            return;
          }

          const mounted = mod.mountScriptEditorPanel({
            workbook,
            container,
            isEditing: () => isSpreadsheetEditing(),
            isReadOnly: () => app.isReadOnly?.() === true,
          });
          scriptEditorReady = true;
          flushPendingScriptEditorCode();
          mount!.dispose = () => {
            disposed = true;
            scriptEditorReady = false;
            try {
              mounted.dispose();
            } finally {
              container.innerHTML = "";
            }
          };
        })().catch((err) => {
          if (disposed) return;
          scriptEditorReady = false;
          container.textContent = `Failed to load Script Editor: ${String(err)}`;
          reportLazyImportFailure("Script Editor", err);
        });
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
              isEditing: () => isSpreadsheetEditing(),
              isReadOnly: () => app.isReadOnly?.() === true,
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
    if (panelDef?.source?.kind === "extension" && !extensionHostManager?.ready) {
      // If the user persisted an extension panel in their layout, ensure the extension host is
      // loaded so the view can activate and populate its webview.
      void ensureExtensionsLoaded();
    }

    if (panelId === PanelIds.DATA_QUERIES || panelId === PanelIds.QUERY_EDITOR) {
      startPowerQueryService();
    }

    if (!panelBodyRenderer) {
      body.textContent = "Loading panel…";
      void ensurePanelBodyRenderer().then((renderer) => {
        if (!body.isConnected) return;
        if (!renderer) {
          body.textContent = "Failed to load panel.";
          return;
        }
        rerenderLayout?.();
      });
      return;
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
    panelBodyRenderer?.cleanup(openPanels);
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
  app.setFocusTargetProvider(() => {
    const split = layoutController.layout.splitView;
    if (split.direction !== "none" && (split.activePane ?? "primary") === "secondary") {
      return gridSecondaryEl;
    }
    return null;
  });

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
    try {
      gridSplitterEl.setPointerCapture(pointerId);
    } catch {
      // Best-effort; some environments (tests/jsdom) may not implement pointer capture.
    }

    // Preserve the pointer's initial offset within the splitter handle so dragging
    // doesn't "snap" the pointer to the edge of the splitter (keeps the grab point stable).
    const pointerOffsetInSplitter = direction === "vertical" ? event.offsetX : event.offsetY;

    // `getBoundingClientRect()` can force layout; defer until the drag actually moves so
    // clicking the splitter without dragging stays cheap.
    let rect: DOMRect | null = null;
    const initialRatio = layoutController.layout.splitView.ratio;
    let latestClientX = event.clientX;
    let latestClientY = event.clientY;
    let didMove = false;
    let lastRatio: number | null = null;
    let rafHandle: number | null = null;

    const applyRatio = (emit: boolean) => {
      if (!rect) rect = gridSplitEl.getBoundingClientRect();
      const splitRect = rect;
      const size = direction === "vertical" ? splitRect.width : splitRect.height;
      if (size <= 0) return;
      const offset =
        direction === "vertical"
          ? latestClientX - splitRect.left - pointerOffsetInSplitter
          : latestClientY - splitRect.top - pointerOffsetInSplitter;
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

    const onMove = (move: PointerEvent) => {
      if (move.pointerId !== pointerId) return;
      didMove = true;
      latestClientX = move.clientX;
      latestClientY = move.clientY;
      scheduleDragApply();
    };

    const onUp = (up: PointerEvent) => {
      if (up.pointerId !== pointerId) return;
      gridSplitterEl.removeEventListener("lostpointercapture", onUp);
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

      // pointercancel events may not have meaningful coordinates; in that case, fall back
      // to the last known pointermove position.
      if (up.type !== "pointercancel" && up.type !== "lostpointercapture") {
        latestClientX = up.clientX;
        latestClientY = up.clientY;
      }
      if (!didMove) return;

      // Flush any pending silent updates so `lastRatio` reflects the final pointer position.
      applyRatio(false);

      // Emit a final layout change (one-time) and persist it so split ratio restores after reload.
      // If the ratio never changed, avoid unnecessary layout normalization/persistence churn.
      if (lastRatio != null && Math.abs(lastRatio - initialRatio) > 1e-6) {
        // Only emit once at drag-end; the rAF-throttled drag updates above use `{ emit:false }`
        // to avoid triggering full `renderLayout()` churn on every pointermove.
        layoutController.setSplitRatio(lastRatio, { persist: false, emit: true });
        persistLayoutNow();
      }
    };

    // Ensure we always clean up the drag if capture is lost unexpectedly (e.g. OS gesture,
    // element teardown). Note: `releasePointerCapture()` also triggers `lostpointercapture`,
    // so the handler removes itself before calling `releasePointerCapture` to avoid reentrancy.
    gridSplitterEl.addEventListener("lostpointercapture", onUp, { passive: true });

    // `touch-action: none` on `#grid-splitter` prevents native touch panning during drags,
    // so these listeners can remain passive (avoids scroll-blocking overhead).
    window.addEventListener("pointermove", onMove, { passive: true });
    window.addEventListener("pointerup", onUp, { passive: true });
    window.addEventListener("pointercancel", onUp, { passive: true });
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
  endBatch: () => app.getDocument().endBatch(),
  // Replace operations mutate cell contents; block them in read-only collab roles (viewer/commenter)
  // and while the spreadsheet is actively editing to avoid silent no-ops (DocumentController filters
  // disallowed cell deltas) and confusing UI behavior.
  canReplace: () => {
    if (isSpreadsheetEditing()) {
      return { allowed: false, reason: "Finish editing before replacing cells." };
    }
    if (app.isReadOnly?.() === true) {
      return { allowed: false, reason: "Read-only: you don't have permission to edit cell contents." };
    }
    return { allowed: true, reason: null };
  },
  showToast,
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

function showDialogModalBestEffort(dialog: HTMLDialogElement): void {
  const showModal = (dialog as any).showModal as (() => void) | undefined;
  if (typeof showModal === "function") {
    try {
      showModal.call(dialog);
      return;
    } catch {
      // Fall through to non-modal open attribute.
    }
  }

  // jsdom doesn't implement showModal(). Best-effort fallback so tests/headless contexts can
  // still exercise dialog contents without crashing.
  dialog.setAttribute("open", "");
}

function showDialogAndFocus(dialog: HTMLDialogElement): void {
  if (!dialog.open) {
    showDialogModalBestEffort(dialog);
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

function showDesktopOnlyToast(message: string): void {
  showToast(`Desktop-only: ${message}`);
}

registerDesktopCommands({
  commandRegistry,
  app,
  layoutController: ribbonLayoutController,
  isEditing: isSpreadsheetEditing,
  commitPendingEditsForCommand: commitAllPendingEditsForCommand,
  focusAfterSheetNavigation: focusAfterSheetNavigationFromCommandRef,
  getVisibleSheetIds: () => listSheetsForUi().map((sheet) => sheet.id),
  sheetStructureHandlers: {
    insertSheet: handleAddSheet,
    deleteActiveSheet: handleDeleteActiveSheet,
    openOrganizeSheets,
  },
  autoFilterHandlers: ribbonAutoFilterHandlers,
  ensureExtensionsLoaded: () => ensureExtensionsLoadedRef?.() ?? Promise.resolve(),
  onExtensionsLoaded: () => {
    updateKeybindingsRef?.();
    syncContributedCommandsRef?.();
    syncContributedPanelsRef?.();
  },
  themeController,
  refreshRibbonUiState: scheduleRibbonSelectionFormatStateUpdate,
  applyFormattingToSelection,
  getActiveCellNumberFormat: activeCellNumberFormat,
  getActiveCellIndentLevel: activeCellIndentLevel,
  formatPainter: {
    isArmed: () => Boolean(formatPainterState),
    arm: () => armFormatPainter(),
    disarm: () => disarmFormatPainter(),
    onCancel: () => {
      try {
        showToast("Format Painter cancelled");
      } catch {
        // ignore (toast root missing in non-UI test environments)
      }
    },
  },
  openFormatCells,
  showQuickPick,
  findReplace: {
    openFind: () => showExclusiveFindReplaceDialog(findDialog as any),
    openReplace: () => {
      if (app.isReadOnly?.() === true) {
        try {
          showToast("Read-only: you don't have permission to edit cell contents.", "warning");
        } catch {
          // `showToast` requires a #toast-root; ignore in headless contexts/tests.
        }
        return;
      }
      showExclusiveFindReplaceDialog(replaceDialog as any);
    },
    openGoTo: () => showExclusiveFindReplaceDialog(goToDialog as any),
  },
  ribbonMacroHandlers: {
    openPanel: openRibbonPanel,
    focusScriptEditorPanel,
    focusVbaMigratePanel,
    setPendingMacrosPanelFocus: (target) => {
      pendingMacrosPanelFocus = target;
    },
    startMacroRecorder: () => {
      const ensure = ensureMacroRecorderRef;
      if (!ensure) {
        activeMacroRecorder?.start();
        return;
      }
      void ensure()
        .then((recorder) => recorder?.start())
        .catch((err) => {
          console.error("Failed to start macro recorder:", err);
          showToast(`Failed to start macro recorder: ${String(err)}`, "error");
        });
    },
    stopMacroRecorder: () => activeMacroRecorder?.stop(),
    isTauri: () => isTauriInvokeAvailable(),
  },
  dataQueriesHandlers: {
    getPowerQueryService: () => {
      if (powerQueryService) return powerQueryService as any;
      startPowerQueryService();

      const waitForService = async () => {
        // `startPowerQueryService` is fire-and-forget; wait for the current init promise
        // (if any) and then verify we have a service.
        const init = powerQueryServiceInitPromise;
        if (init) await init;
        const service = powerQueryService;
        if (!service) throw new Error("Queries service not available");
        await service.ready;
        return service;
      };

      return {
        ready: waitForService().then(() => undefined),
        getQueries: () => powerQueryService?.getQueries() ?? [],
        refreshAll: () => {
          startPowerQueryService();
          return {
            promise: waitForService().then(async (service) => {
              const handle = service.refreshAll();
              return await handle.promise;
            }),
          };
        },
      };
    },
    showToast,
    notify,
    focusAfterExecute: () => app.focus(),
  },
  pageLayoutHandlers: {
    openPageSetupDialog: () => handleRibbonPageSetup(),
    updatePageSetup: (patch) => handleRibbonUpdatePageSetup(patch),
    setPrintArea: () => handleRibbonSetPrintArea(),
    clearPrintArea: () => handleRibbonClearPrintArea(),
    addToPrintArea: () => handleRibbonAddToPrintArea(),
    exportPdf: () => handleRibbonExportPdf(),
  },
  workbenchFileHandlers: {
    newWorkbook: async () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Creating new workbooks is available in the desktop app.");
        return;
      }
      return queueFileCommand(async () => {
        try {
          await handleNewWorkbook();
        } catch (err) {
          console.error("Failed to create workbook:", err);
          showToast(`Failed to create workbook: ${String(err)}`, "error");
        }
      });
    },
    openWorkbook: async () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Opening workbooks is available in the desktop app.");
        return;
      }
      return queueFileCommand(async () => {
        try {
          await promptOpenWorkbook();
        } catch (err) {
          console.error("Failed to open workbook:", err);
          showToast(`Failed to open workbook: ${String(err)}`, "error");
        }
      });
    },
    saveWorkbook: async () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Saving workbooks is available in the desktop app.");
        return;
      }
      return queueFileCommand(async () => {
        try {
          await handleSave();
        } catch (err) {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        }
      });
    },
    saveWorkbookAs: async () => {
      if (!tauriBackend) {
        showDesktopOnlyToast("Save As is available in the desktop app.");
        return;
      }
      return queueFileCommand(async () => {
        try {
          await handleSaveAs();
        } catch (err) {
          console.error("Failed to save workbook:", err);
          showToast(`Failed to save workbook: ${String(err)}`, "error");
        }
      });
    },
    setAutoSaveEnabled: (enabled?: boolean) => {
      const nextEnabled = typeof enabled === "boolean" ? enabled : !autoSaveEnabled;
      void setAutoSaveEnabledFromUi(nextEnabled);
    },
    print: () => {
      const invokeAvailable = hasTauriInvoke();
      if (!invokeAvailable) {
        showDesktopOnlyToast("Print is available in the desktop app.");
        return;
      }
      void handleRibbonPrintPreview({ autoPrint: true }).catch((err) => {
        console.error("Failed to print:", err);
        showToast(`Failed to print: ${String(err)}`, "error");
      });
    },
    printPreview: () => {
      const invokeAvailable = hasTauriInvoke();
      if (!invokeAvailable) {
        showDesktopOnlyToast("Print Preview is available in the desktop app.");
        return;
      }
      void handleRibbonPrintPreview({ autoPrint: false }).catch((err) => {
        console.error("Failed to open print preview:", err);
        showToast(`Failed to open print preview: ${String(err)}`, "error");
      });
    },
    closeWorkbook: () => {
      if (handleCloseRequestForRibbon) {
        void handleCloseRequestForRibbon({ quit: false }).catch((err) => {
          console.error("Failed to close window:", err);
          showToast(`Failed to close window: ${String(err)}`, "error");
        });
        return;
      }

      // When running under Tauri, the close-request handler is normally installed by the desktop
      // host integration. If it isn't available (e.g. permission/config mismatch), fall back to
      // the window API so Close Window still works.
      if (!hasTauriWindowApi()) {
        showDesktopOnlyToast("Closing windows is available in the desktop app.");
        return;
      }

      void (async () => {
        try {
          // Best-effort prompt to avoid data loss if the host close-request wiring isn't installed.
          if (isDirtyForUnsavedChangesPrompts()) {
            const discard = await nativeDialogs.confirm(t("prompt.unsavedChangesDiscardConfirm"));
            if (!discard) return;
          }
          await hideTauriWindow();
        } catch (err) {
          console.error("Failed to close window:", err);
          showToast(`Failed to close window: ${String(err)}`, "error");
        }
      })();
    },
    quit: () => {
      if (!handleCloseRequestForRibbon) {
        const invokeAvailable = hasTauriInvoke();
        if (!invokeAvailable) {
          showDesktopOnlyToast("Quitting is available in the desktop app.");
          return;
        }

        void requestAppQuit().catch((err) => {
          console.error("Failed to quit app:", err);
          showToast(`Failed to quit app: ${String(err)}`, "error");
        });
        return;
      }
      void handleCloseRequestForRibbon({ quit: true }).catch((err) => {
        console.error("Failed to quit app:", err);
        showToast(`Failed to quit app: ${String(err)}`, "error");
      });
    },
  },
  encryptWithPassword: () => {
    if (!tauriBackend) {
      showDesktopOnlyToast("Encrypt with Password is available in the desktop app.");
      return;
    }
    void handleSaveWithPassword().catch((err) => {
      console.error("Failed to save workbook with password:", err);
      showToast(`Failed to save workbook with password: ${String(err)}`, "error");
    });
  },
  openCommandPalette: () => openCommandPalette?.(),
  openGoalSeekDialog: () => showGoalSeekDialogModal(),
});

function getTauriInvokeForPrint(): TauriInvoke | null {
  const invoke = (queuedInvoke ?? getTauriInvokeOrNull()) as TauriInvoke | null;
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
  // `Blob` expects ArrayBuffer-backed views. TypeScript models `Uint8Array` as possibly backed by a
  // `SharedArrayBuffer` (`ArrayBufferLike`), so normalize for type safety.
  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  const blob = new Blob([normalized], { type: mime });
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
    // Export should reflect the latest user input (including in-progress edits that haven't been
    // committed into the DocumentController yet).
    commitAllPendingEditsForCommand();
    if (isSpreadsheetEditing()) return;

    if (typeof document === "undefined" || typeof Blob === "undefined") return;
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return;

    const sheetId = app.getCurrentSheetId();
    const sheetName =
      typeof (app as any)?.getCurrentSheetDisplayName === "function"
        ? (app as any).getCurrentSheetDisplayName()
        : workbookSheetStore.getName(sheetId) ?? sheetId;
    const doc = app.getDocument();
    const limits = getGridLimitsForFormatting();
    const active = app.getActiveCell();

    const clipBandSelectionToUsedRange = (range0: CellRange): CellRange => {
      const normalized: CellRange = {
        start: { row: Math.min(range0.start.row, range0.end.row), col: Math.min(range0.start.col, range0.end.col) },
        end: { row: Math.max(range0.start.row, range0.end.row), col: Math.max(range0.start.col, range0.end.col) },
      };

      const activeCellFallback0: CellRange = {
        start: { row: active.row, col: active.col },
        end: { row: active.row, col: active.col },
      };

      const isFullHeight = normalized.start.row === 0 && normalized.end.row === limits.maxRows - 1;
      const isFullWidth = normalized.start.col === 0 && normalized.end.col === limits.maxCols - 1;
      if (!isFullHeight && !isFullWidth) return normalized;

      const used = doc.getUsedRange(sheetId);
      if (!used) return activeCellFallback0;

      const startRow = Math.max(normalized.start.row, used.startRow);
      const endRow = Math.min(normalized.end.row, used.endRow);
      const startCol = Math.max(normalized.start.col, used.startCol);
      const endCol = Math.min(normalized.end.col, used.endCol);
      const clipped =
        startRow <= endRow && startCol <= endCol
          ? { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } }
          : null;
      return clipped ?? activeCellFallback0;
    };

    const exportRange0 = clipBandSelectionToUsedRange(selectionBoundingBox0Based());
    const dlp = createDesktopDlpContext({ documentId: workbookId });

    const text = exportDocumentRangeToCsv(doc, sheetId, exportRange0, {
      delimiter: args.delimiter,
      dlp: { documentId: dlp.documentId, classificationStore: dlp.classificationStore, policy: dlp.policy },
    });

    const blob = new Blob([text], { type: args.mime });
    const url = URL.createObjectURL(blob);
    try {
      const a = document.createElement("a");
      a.href = url;
      a.download = `${sanitizeFilename(sheetName)}.${args.extension}`;
      a.rel = "noopener";
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } finally {
      URL.revokeObjectURL(url);
    }
  } catch (err) {
    console.error(`Failed to export ${args.label}:`, err);
    showToast(`Failed to export ${args.label}: ${String(err)}`, "error");
  } finally {
    app.focus();
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
  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) {
    if (openModal.classList.contains("page-setup-dialog")) return;
    return;
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog page-setup-dialog";
  markKeybindingBarrier(dialog);
  dialog.setAttribute("aria-label", t("print.pageSetup.title"));

  const container = document.createElement("div");
  dialog.appendChild(container);
  document.body.appendChild(dialog);

  const root = createRoot(container);

  const close = () => dialog.close();

  function Wrapper() {
    const [value, setValue] = React.useState<PageSetup>(args.initialValue);
    const [printModule, setPrintModule] = React.useState<Awaited<ReturnType<typeof loadPageSetupDialogModule>>>(null);
    const [loadFailed, setLoadFailed] = React.useState(false);

    React.useEffect(() => {
      let cancelled = false;
      void loadPageSetupDialogModule().then((mod) => {
        if (cancelled) return;
        if (!mod) {
          setLoadFailed(true);
          return;
        }
        setPrintModule(mod);
      });
      return () => {
        cancelled = true;
      };
    }, []);

    const handleChange = React.useCallback(
      (next: PageSetup) => {
        setValue(next);
        args.onChange(next);
      },
      [args],
    );

    if (loadFailed) {
      return React.createElement(
        "div",
        { className: "dialog__loading" },
        React.createElement("div", null, "Failed to load Page Setup dialog."),
        React.createElement("div", { className: "dialog__controls" }, React.createElement("button", { type: "button", onClick: close }, "Close")),
      );
    }

    if (!printModule) {
      return React.createElement("div", { className: "dialog__loading" }, "Loading…");
    }

    return React.createElement(printModule.PageSetupDialog, { value, onChange: handleChange, onClose: close });
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

  showDialogModalBestEffort(dialog);
}

async function handleRibbonPageSetup(): Promise<void> {
  if (isSpreadsheetEditing()) {
    app.focus();
    return;
  }
  if (app.isReadOnly?.() === true) {
    showToast("Read-only: you don't have permission to edit page setup.", "warning");
    app.focus();
    return;
  }

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

async function handleRibbonUpdatePageSetup(patch: (current: PageSetup) => PageSetup): Promise<void> {
  if (isSpreadsheetEditing()) {
    app.focus();
    return;
  }
  if (app.isReadOnly?.() === true) {
    showToast("Read-only: you don't have permission to edit page setup.", "warning");
    app.focus();
    return;
  }

  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const sheetId = app.getCurrentSheetId();
    const settings = await invoke("get_sheet_print_settings", { sheet_id: sheetId });
    const current = pageSetupFromTauri((settings as any)?.page_setup);
    const next = patch(current);
    await invoke("set_sheet_page_setup", { sheet_id: sheetId, page_setup: pageSetupToTauri(next) });
    app.focus();
  } catch (err) {
    console.error("Failed to update page setup:", err);
    showToast(`Failed to update page setup: ${String(err)}`, "error");
  }
}

async function handleRibbonSetPrintArea(): Promise<void> {
  if (isSpreadsheetEditing()) {
    app.focus();
    return;
  }
  if (app.isReadOnly?.() === true) {
    showToast("Read-only: you don't have permission to set a print area.", "warning");
    app.focus();
    return;
  }

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
  if (isSpreadsheetEditing()) {
    app.focus();
    return;
  }
  if (app.isReadOnly?.() === true) {
    showToast("Read-only: you don't have permission to clear the print area.", "warning");
    app.focus();
    return;
  }

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

async function handleRibbonAddToPrintArea(): Promise<void> {
  if (isSpreadsheetEditing()) {
    app.focus();
    return;
  }
  if (app.isReadOnly?.() === true) {
    showToast("Read-only: you don't have permission to edit the print area.", "warning");
    app.focus();
    return;
  }

  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const sheetId = app.getCurrentSheetId();
    const selection = selectionBoundingBox1Based();

    const settings = await invoke("get_sheet_print_settings", { sheet_id: sheetId });
    const existing = (settings as any)?.print_area;
    const ranges = Array.isArray(existing) ? existing : [];

    const next = ranges
      .map((r: any) => ({
        start_row: Number(r?.start_row),
        end_row: Number(r?.end_row),
        start_col: Number(r?.start_col),
        end_col: Number(r?.end_col),
      }))
      .filter(
        (r: any) =>
          Number.isFinite(r.start_row) &&
          Number.isFinite(r.end_row) &&
          Number.isFinite(r.start_col) &&
          Number.isFinite(r.end_col),
      );

    const selectionTauri = {
      start_row: selection.startRow,
      end_row: selection.endRow,
      start_col: selection.startCol,
      end_col: selection.endCol,
    };

    const alreadyIncluded = next.some(
      (r: any) =>
        r.start_row === selectionTauri.start_row &&
        r.end_row === selectionTauri.end_row &&
        r.start_col === selectionTauri.start_col &&
        r.end_col === selectionTauri.end_col,
    );
    if (!alreadyIncluded) next.push(selectionTauri);

    await invoke("set_sheet_print_area", { sheet_id: sheetId, print_area: next });
    app.focus();
  } catch (err) {
    console.error("Failed to add to print area:", err);
    showToast(`Failed to add to print area: ${String(err)}`, "error");
  }
}

async function generateSheetPdfBytes(invoke: TauriInvoke): Promise<{ sheetId: string; bytes: Uint8Array }> {
  // Export should reflect the latest user input. If a cell edit is in progress, it may not yet
  // have been committed into the DocumentController / backend sync pipeline.
  commitAllPendingEditsForCommand();

  // Best-effort: ensure any pending workbook sync changes are flushed before exporting.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await drainBackendSync();

  const sheetId = app.getCurrentSheetId();
  const doc = app.getDocument();
  const limits = getGridLimitsForFormatting();
  const active = app.getActiveCell();

  const clipBandSelectionToUsedRange = (range0: CellRange): CellRange => {
    const normalized: CellRange = {
      start: { row: Math.min(range0.start.row, range0.end.row), col: Math.min(range0.start.col, range0.end.col) },
      end: { row: Math.max(range0.start.row, range0.end.row), col: Math.max(range0.start.col, range0.end.col) },
    };

    const activeCellFallback0: CellRange = {
      start: { row: active.row, col: active.col },
      end: { row: active.row, col: active.col },
    };

    const isFullHeight = normalized.start.row === 0 && normalized.end.row === limits.maxRows - 1;
    const isFullWidth = normalized.start.col === 0 && normalized.end.col === limits.maxCols - 1;
    if (!isFullHeight && !isFullWidth) return normalized;

    const used = doc.getUsedRange(sheetId);
    if (!used) return activeCellFallback0;

    const startRow = Math.max(normalized.start.row, used.startRow);
    const endRow = Math.min(normalized.end.row, used.endRow);
    const startCol = Math.max(normalized.start.col, used.startCol);
    const endCol = Math.min(normalized.end.col, used.endCol);
    const clipped =
      startRow <= endRow && startCol <= endCol
        ? { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } }
        : null;
    return clipped ?? activeCellFallback0;
  };

  // Use the selection by default, but clip full-row/full-column/full-sheet selections to the
  // used range to avoid generating PDFs that span millions of empty cells.
  let exportRange0 = clipBandSelectionToUsedRange(selectionBoundingBox0Based());

  try {
    const settings = await invoke("get_sheet_print_settings", { sheet_id: sheetId });
    const printArea = (settings as any)?.print_area;
    const first = Array.isArray(printArea) ? printArea[0] : null;
    if (first) {
      const startRow = Number(first.start_row);
      const endRow = Number(first.end_row);
      const startCol = Number(first.start_col);
      const endCol = Number(first.end_col);
      if ([startRow, endRow, startCol, endCol].every((v) => Number.isFinite(v) && v > 0)) {
        exportRange0 = clipBandSelectionToUsedRange({
          start: { row: Math.min(startRow, endRow) - 1, col: Math.min(startCol, endCol) - 1 },
          end: { row: Math.max(startRow, endRow) - 1, col: Math.max(startCol, endCol) - 1 },
        });
      }
    }
  } catch (err) {
    console.warn("Failed to fetch print area settings; exporting selection instead:", err);
  }

  const range: PrintCellRange = {
    startRow: exportRange0.start.row + 1,
    endRow: exportRange0.end.row + 1,
    startCol: exportRange0.start.col + 1,
    endCol: exportRange0.end.col + 1,
  };

  const b64 = await invoke("export_sheet_range_pdf", {
    sheet_id: sheetId,
    range: { start_row: range.startRow, end_row: range.endRow, start_col: range.startCol, end_col: range.endCol },
    col_widths_points: undefined,
    row_heights_points: undefined,
  });

  return { sheetId, bytes: decodeBase64ToBytes(String(b64)) };
}

function showPrintPreviewDialogModal(args: { bytes: Uint8Array; filename: string; autoPrint?: boolean }): void {
  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) {
    if (openModal.classList.contains("print-preview-dialog")) return;
    return;
  }

  const dialog = document.createElement("dialog");
  // `data-keybinding-barrier` prevents global spreadsheet keybindings from firing while the
  // user is interacting with the modal (e.g. Cmd+P should not re-trigger Print).
  dialog.className = "dialog print-preview-dialog";
  markKeybindingBarrier(dialog);
  dialog.setAttribute("aria-label", "Print Preview");

  const container = document.createElement("div");
  dialog.appendChild(container);
  document.body.appendChild(dialog);

  const root = createRoot(container);

  const close = () => dialog.close();
  function Wrapper() {
    const [printModule, setPrintModule] = React.useState<Awaited<ReturnType<typeof loadPrintPreviewDialogModule>>>(null);
    const [loadFailed, setLoadFailed] = React.useState(false);

    React.useEffect(() => {
      let cancelled = false;
      void loadPrintPreviewDialogModule().then((mod) => {
        if (cancelled) return;
        if (!mod) {
          setLoadFailed(true);
          return;
        }
        setPrintModule(mod);
      });
      return () => {
        cancelled = true;
      };
    }, []);

    if (loadFailed) {
      return React.createElement(
        "div",
        { className: "dialog__loading" },
        React.createElement("div", null, "Failed to load Print Preview."),
        React.createElement("div", { className: "dialog__controls" }, React.createElement("button", { type: "button", onClick: close }, "Close")),
      );
    }

    if (!printModule) {
      return React.createElement("div", { className: "dialog__loading" }, "Loading…");
    }

    return React.createElement(printModule.PrintPreviewDialog, {
      pdfBytes: args.bytes,
      filename: args.filename,
      autoPrint: Boolean(args.autoPrint),
      onDownload: () => downloadBytes(args.bytes, args.filename, "application/pdf"),
      onClose: close,
    });
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

  // Trap Tab navigation within the modal so focus doesn't escape back to the grid/ribbon.
  dialog.addEventListener("keydown", (event) => {
    if (event.key !== "Tab") return;
    const focusables = Array.from(
      dialog.querySelectorAll<HTMLElement>(
        'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
      ),
    ).filter((el) => el.getAttribute("aria-hidden") !== "true");
    if (focusables.length === 0) return;
    const first = focusables[0]!;
    const last = focusables[focusables.length - 1]!;
    const active = document.activeElement as HTMLElement | null;
    if (!active) return;

    if (event.shiftKey) {
      if (active === first) {
        event.preventDefault();
        last.focus();
      }
      return;
    }

    if (active === last) {
      event.preventDefault();
      first.focus();
    }
  });

  showDialogModalBestEffort(dialog);
}

const whatIfApi = createWhatIfApi();
let goalSeekDialogMount: { root: Root; container: HTMLDivElement } | null = null;

function showGoalSeekDialogModal(): void {
  if (typeof document === "undefined") return;
  if (goalSeekDialogMount) return;

  const container = document.createElement("div");
  document.body.appendChild(container);

  const root = createRoot(container);
  goalSeekDialogMount = { root, container };

  function Wrapper() {
    const [open, setOpen] = React.useState(true);
    const [goalSeekModule, setGoalSeekModule] = React.useState<Awaited<ReturnType<typeof loadGoalSeekDialogModule>>>(null);
    const [loadFailed, setLoadFailed] = React.useState(false);

    React.useEffect(() => {
      let cancelled = false;
      void loadGoalSeekDialogModule().then((mod) => {
        if (cancelled) return;
        if (!mod) {
          setLoadFailed(true);
          return;
        }
        setGoalSeekModule(mod);
      });
      return () => {
        cancelled = true;
      };
    }, []);

    React.useEffect(() => {
      if (open) return;
      // Defer cleanup so `GoalSeekDialog` can run its focus-restoration effect on close.
      const timeout = window.setTimeout(() => {
        try {
          root.unmount();
        } finally {
          container.remove();
          goalSeekDialogMount = null;
          app.focus();
        }
      }, 0);

      return () => {
        window.clearTimeout(timeout);
      };
    }, [open]);

    if (loadFailed) {
      return React.createElement(
        "div",
        { className: "dialog__loading" },
        React.createElement("div", null, "Failed to load Goal Seek dialog."),
        React.createElement("button", { type: "button", onClick: () => setOpen(false) }, "Close"),
      );
    }

    if (!goalSeekModule) {
      return React.createElement("div", { className: "dialog__loading" }, "Loading…");
    }

    return React.createElement(goalSeekModule.GoalSeekDialog, { api: whatIfApi, open, onClose: () => setOpen(false) });
  }

  root.render(React.createElement(Wrapper));
}

async function handleRibbonPrintPreview(args: { autoPrint: boolean }): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const { bytes, sheetId } = await generateSheetPdfBytes(invoke);
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;
    const filename = `${sanitizeFilename(sheetName)}.pdf`;
    showPrintPreviewDialogModal({ bytes, filename, autoPrint: args.autoPrint });
  } catch (err) {
    console.error("Failed to open print preview:", err);
    showToast(`Failed to open print preview: ${String(err)}`, "error");
  }
}

async function handleRibbonExportPdf(): Promise<void> {
  const invoke = getTauriInvokeForPrint();
  if (!invoke) return;

  try {
    const { bytes, sheetId } = await generateSheetPdfBytes(invoke);
    const sheetName = workbookSheetStore.getName(sheetId) ?? sheetId;
    downloadBytes(bytes, `${sanitizeFilename(sheetName)}.pdf`, "application/pdf");
    app.focus();
  } catch (err) {
    console.error("Failed to export PDF:", err);
    showToast(`Failed to export PDF: ${String(err)}`, "error");
  }
}
// --- Ribbon: AutoFilter MVP ----------------------------------------------------
//
// Excel's AutoFilter is normally driven by header dropdowns inside the grid. For an MVP we wire the
// ribbon commands to a simple checklist UI that filters the current selection range by hiding rows
// via the legacy outline layer (`outline.hidden.filter`).
//
// Note: Shared-grid mode does not currently support outline-based hidden rows; we disable the ribbon
// controls in `scheduleRibbonSelectionFormatStateUpdate()` to avoid noisy toasts.
//
// Performance note: This MVP implementation scans the filter range to compute distinct values and
// then scans again to apply the row-hiding filter. Keep the scan bounded to avoid freezing the UI
// on very large selections.

const DEFAULT_RIBBON_AUTOFILTER_ROW_LIMIT = 50_000;
const DEFAULT_RIBBON_AUTOFILTER_UNIQUE_VALUE_LIMIT = 5_000;

function getRibbonAutoFilterCellText(sheetId: string, cell: { row: number; col: number }): string {
  const doc = app.getDocument();
  const state = doc.getCell(sheetId, cell) as { value?: unknown; formula?: string | null } | null;

  const formatValue = (raw: unknown): string => {
    if (raw == null) return "";
    if (typeof raw === "object" && raw && "text" in raw && typeof (raw as any).text === "string") {
      return String((raw as any).text);
    }

    const image = parseImageCellValue(raw);
    if (image) return image.altText ?? "[Image]";

    if (typeof raw === "boolean") return raw ? "TRUE" : "FALSE";

    if (typeof raw === "number" && Number.isFinite(raw)) {
      const fmt = doc.getCellFormat(sheetId, cell) as any;
      const numberFormat = getStyleNumberFormat(fmt);
      if (numberFormat) return formatValueWithNumberFormat(raw, numberFormat);
    }

    return String(raw);
  };

  // Prefer computed values for formula cells so filtering matches what users see in the grid.
  const hasFormula = typeof state?.formula === "string" && state.formula.trim() !== "";
  if (hasFormula) {
    // Respect the View → Show Formulas toggle so the filter dialog matches the grid display.
    if (app.getShowFormulas()) {
      return state?.formula ?? "";
    }
    const computed = app.getCellComputedValueForSheet(sheetId, cell);
    if ((computed == null || computed === "") && state?.value != null) {
      const image = parseImageCellValue(state.value);
      if (image) return image.altText ?? "[Image]";
    }
    return formatValue(computed);
  }

  const value = state?.value ?? null;
  return formatValue(value);
}

async function showRibbonAutoFilterDialog(args: {
  rows: TableViewRow[];
  initialSelected: string[] | null;
}): Promise<string[] | null> {
  // Avoid throwing when another modal dialog is already open. Treat this as a cancel so callers
  // don't hang awaiting a dialog that can never become modal.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) return null;

  return await new Promise<string[] | null>((resolve) => {
    const dialog = document.createElement("dialog");
    dialog.className = "dialog ribbon-auto-filter-dialog";
    markKeybindingBarrier(dialog);
    dialog.setAttribute("aria-label", "Filter");

    const container = document.createElement("div");
    dialog.appendChild(container);
    document.body.appendChild(dialog);
    const root = createRoot(container);

    let result: string[] | null = null;

    const close = () => dialog.close();
    root.render(
      React.createElement(AutoFilterDropdown, {
        rows: args.rows,
        colId: 0,
        initialSelected: args.initialSelected,
        onApply: (selected: string[]) => {
          result = selected;
        },
        onClose: close,
      }),
    );

    dialog.addEventListener(
      "close",
      () => {
        try {
          root.unmount();
        } finally {
          dialog.remove();
          resolve(result);
          app.focus();
        }
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      dialog.close();
    });

    showDialogModalBestEffort(dialog);
  });
}

async function applyRibbonAutoFilterFromSelection(): Promise<boolean> {
  if (app.getGridMode() !== "legacy") {
    // Defensive: shared-grid mode should have these commands disabled in the ribbon, but keep the
    // toggle state consistent if invoked via another surface.
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
    return false;
  }

  if (isSpreadsheetEditing()) {
    scheduleRibbonSelectionFormatStateUpdate();
    return false;
  }

  const selection = currentSelectionRect();
  const sheetId = selection.sheetId;
  const selectionRange = {
    startRow: selection.startRow,
    endRow: selection.endRow,
    startCol: selection.startCol,
    endCol: selection.endCol,
  };

  // If the active cell is already inside an existing filter range, treat the ribbon command as
  // "edit that filter" (so users don't need to re-select the original range).
  const existingForCell = ribbonAutoFilterStore.findByCell(sheetId, { row: selection.activeRow, col: selection.activeCol });
  const range = (() => {
    if (existingForCell) return parseA1Range(existingForCell.rangeA1);
    // Excel's "Filter" button commonly starts from a single active cell and expands to the used range.
    // For this MVP, use the DocumentController's used range as a simple heuristic (only when the
    // active cell is within the used range so we don't unexpectedly filter unrelated data).
    const isSingleCellSelection =
      selectionRange.startRow === selectionRange.endRow && selectionRange.startCol === selectionRange.endCol;
    if (!isSingleCellSelection) return selectionRange;
    const used = app.getDocument().getUsedRange(sheetId);
    if (!used) return selectionRange;
    const inUsed =
      selection.activeRow >= used.startRow &&
      selection.activeRow <= used.endRow &&
      selection.activeCol >= used.startCol &&
      selection.activeCol <= used.endCol;
    if (!inUsed) return selectionRange;
    return used;
  })();
  const rangeA1 = existingForCell ? existingForCell.rangeA1 : rangeToA1(range);

  const headerRows = existingForCell?.headerRows ?? 1;
  const dataStartRow = range.startRow + headerRows;
  if (dataStartRow > range.endRow) {
    showToast("Select a range with a header row and at least one data row to filter.", "warning");
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
    return false;
  }
  const dataRowCount = range.endRow - dataStartRow + 1;
  if (dataRowCount > DEFAULT_RIBBON_AUTOFILTER_ROW_LIMIT) {
    showToast(
      `Selection too large to filter (${dataRowCount.toLocaleString()} rows). ` +
        `Reduce the selection to under ${DEFAULT_RIBBON_AUTOFILTER_ROW_LIMIT.toLocaleString()} rows and try again.`,
      "warning",
    );
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
    return false;
  }

  const colId = selection.activeCol - range.startCol;
  if (colId < 0 || range.startCol + colId > range.endCol) {
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
    return false;
  }

  const getValue = (row: number, col: number) => getRibbonAutoFilterCellText(sheetId, { row, col });
  // Build the dialog list from distinct values only. Avoid allocating a row entry per data row,
  // which can be expensive for large selections.
  const uniqueValues = computeUniqueFilterValues({
    range,
    headerRows,
    colId,
    getValue,
    maxValues: DEFAULT_RIBBON_AUTOFILTER_UNIQUE_VALUE_LIMIT,
  });
  if (uniqueValues.length > DEFAULT_RIBBON_AUTOFILTER_UNIQUE_VALUE_LIMIT) {
    showToast(
      `Too many unique values to filter (more than ${DEFAULT_RIBBON_AUTOFILTER_UNIQUE_VALUE_LIMIT.toLocaleString()}). ` +
        "Select a smaller range and try again.",
      "warning",
    );
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
    return false;
  }
  const dialogRows: TableViewRow[] = uniqueValues.map((value, idx) => ({ row: idx, values: [value] }));

  const existing = ribbonAutoFilterStore.get(sheetId, rangeA1);
  const initialSelected = existing?.filterColumns.find((c) => c.colId === colId)?.values ?? null;

  const selected = await showRibbonAutoFilterDialog({ rows: dialogRows, initialSelected });
  if (selected == null) {
    scheduleRibbonSelectionFormatStateUpdate();
    return false;
  }

  // If the user selected all values, treat it as "no filter" for that column (Excel-like),
  // so future newly-entered values remain visible by default.
  const uniqueSet = new Set(uniqueValues);
  const selectedSet = new Set(selected);
  const isSelectAll = selectedSet.size === uniqueSet.size && Array.from(selectedSet).every((v) => uniqueSet.has(v));

  const currentColumns = existing?.filterColumns ?? [];
  const nextColumns = currentColumns.filter((c) => c.colId !== colId);
  if (!isSelectAll) {
    nextColumns.push({ colId, values: selected });
  }
  ribbonAutoFilterStore.set(sheetId, { rangeA1, headerRows, filterColumns: nextColumns });
  updateContextKeys();

  // Reapply all stored filters so row hidden state stays consistent even if multiple filter ranges overlap.
  reapplyRibbonAutoFiltersForActiveSheet();
  return true;
}

function clearRibbonAutoFilterCriteriaForActiveSheet(): void {
  if (app.getGridMode() !== "legacy") {
    app.focus();
    return;
  }
  if (isSpreadsheetEditing()) return;

  const sheetId = app.getCurrentSheetId();
  const filters = ribbonAutoFilterStore.list(sheetId);
  if (filters.length === 0) {
    app.focus();
    return;
  }

  // Keep the filter ranges (so the ribbon toggle remains pressed), but clear any column criteria.
  for (const filter of filters) {
    ribbonAutoFilterStore.set(sheetId, { rangeA1: filter.rangeA1, headerRows: filter.headerRows, filterColumns: [] });
  }
  updateContextKeys();

  // Apply the updated (empty) criteria by clearing any existing filter-hidden rows.
  reapplyRibbonAutoFiltersForActiveSheet();
}

function clearRibbonAutoFiltersForActiveSheet(): void {
  if (app.getGridMode() !== "legacy") {
    app.focus();
    return;
  }
  if (isSpreadsheetEditing()) return;

  const sheetId = app.getCurrentSheetId();
  ribbonAutoFilterStore.clearSheet(sheetId);
  updateContextKeys();

  const limits = app.getGridLimits();
  app.clearFilteredHiddenRowsInRange(0, limits.maxRows - 1);

  scheduleRibbonSelectionFormatStateUpdate();
  app.focus();
}

function reapplyRibbonAutoFiltersForActiveSheet(): void {
  if (app.getGridMode() !== "legacy") {
    app.focus();
    return;
  }
  if (isSpreadsheetEditing()) return;

  const sheetId = app.getCurrentSheetId();
  const filters = ribbonAutoFilterStore.list(sheetId);
  if (filters.length === 0) {
    app.focus();
    return;
  }

  const limits = app.getGridLimits();
  app.clearFilteredHiddenRowsInRange(0, limits.maxRows - 1);

  const getValue = (row: number, col: number) => getRibbonAutoFilterCellText(sheetId, { row, col });
  for (const filter of filters) {
    const range = parseA1Range(filter.rangeA1);
    const hidden = computeFilterHiddenRows({
      range,
      headerRows: filter.headerRows,
      filterColumns: filter.filterColumns,
      getValue,
    });
    app.setRowsFilteredHidden(hidden, true);
  }

  scheduleRibbonSelectionFormatStateUpdate();
  app.focus();
}

function toggleRibbonAutoFilter(): void {
  // In Excel, the ribbon "Filter" button is a toggle. For this MVP we treat it as:
  // - if any filter is active on the sheet, clear it
  // - otherwise, prompt for values and apply a simple row-hiding filter
  if (ribbonAutoFilterStore.hasAny(app.getCurrentSheetId())) {
    clearRibbonAutoFiltersForActiveSheet();
    return;
  }
  void applyRibbonAutoFilterFromSelection().catch((err) => {
    console.error("Failed to apply filter:", err);
    showToast(`Failed to apply filter: ${String(err)}`, "error");
    scheduleRibbonSelectionFormatStateUpdate();
    app.focus();
  });
}

const ribbonActions = createRibbonActions({
  app,
  commandRegistry,
  isSpreadsheetEditing,
  showToast,
  showQuickPick,
  showInputBox,
  openOrganizeSheets,
  handleAddSheet,
  handleDeleteActiveSheet,
  openCustomSortDialog: (commandId) => {
    handleCustomSortCommand(commandId, {
      isEditing: isSpreadsheetEditing,
      getDocument: () => app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      getSelectionRanges: () => app.getSelectionRanges(),
      getCellValue: (sheetId, cell) => app.getCellComputedValueForSheet(sheetId, cell),
      focusGrid: () => app.focus(),
    });
  },
  toggleAutoFilter: toggleRibbonAutoFilter,
  // Excel-style: "Clear" clears filter criteria (show all rows) but keeps AutoFilter enabled.
  clearAutoFilter: clearRibbonAutoFilterCriteriaForActiveSheet,
  reapplyAutoFilter: reapplyRibbonAutoFiltersForActiveSheet,
  applyFormattingToSelection,
  getActiveCellNumberFormat: activeCellNumberFormat,
  getEnsureExtensionsLoadedRef: () => ensureExtensionsLoadedRef,
  getSyncContributedCommandsRef: () => syncContributedCommandsRef,
});

mountRibbon(ribbonReactRoot, ribbonActions);
// In Yjs-backed collaboration mode the workbook is continuously persisted, but
// DocumentController's `isDirty` flips to true on essentially every local/remote
// change (including `applyExternalDeltas`). That makes the browser/Tauri
// beforeunload "unsaved changes" prompt effectively permanent and incorrect.
//
// SpreadsheetApp may attach collaboration support asynchronously, so we check
// `getCollabSession()` at prompt time instead of only once at startup.
function isCollabSessionActive(): boolean {
  try {
    return app.getCollabSession() != null;
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

async function hideTauriWindow(): Promise<void> {
  try {
    const win = getTauriWindowHandleOrThrow();
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

async function showTauriWindowBestEffort(): Promise<void> {
  try {
    const win = getTauriWindowHandleOrThrow();
    // Prefer restoring from a minimized state first so `show()` makes the window visible.
    try {
      if (typeof win.unminimize === "function") {
        await win.unminimize();
      }
    } catch {
      // Best-effort.
    }
    try {
      // Some Tauri variants expose minimize state via a setter API.
      if (typeof win.setMinimized === "function") {
        await win.setMinimized(false);
      }
    } catch {
      // Best-effort.
    }
    if (typeof win.show === "function") {
      await win.show();
    }
    if (typeof win.setFocus === "function") {
      await win.setFocus();
    }
  } catch {
    // Best-effort.
  }
}

async function minimizeTauriWindow(): Promise<void> {
  try {
    const win = getTauriWindowHandleOrThrow();
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
    const win = getTauriWindowHandleOrThrow();
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

function queueFileCommand<T>(op: () => Promise<T>): Promise<T> {
  const result = pendingFileCommand.then(op, op);
  // Include shell-owned file operations (open/save/close) in the app's idle tracker so
  // Playwright helpers that await `app.whenIdle()` don't race work that runs outside
  // SpreadsheetApp internals.
  try {
    app.trackIdleTask(result);
  } catch {
    // ignore (defensive: non-standard harnesses may not expose the idle tracker hook)
  }
  pendingFileCommand = result.then(
    () => undefined,
    (err) => {
      console.error("Failed to run file command:", err);
    },
  );
  return result;
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

  const invoke = (queuedInvoke ?? getTauriInvokeOrNull()) as TauriInvoke | null;
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
  window.__workbookSheetStore = workbookSheetStore;

  // Start chart extraction early so it runs in parallel with the cell snapshot load below.
  // This is best-effort; failures should never block opening the workbook.
  const importedChartObjectsPromise: Promise<unknown[]> = tauriBackend
    .listImportedChartObjects()
    .catch((err) => {
      console.warn("[formula][desktop] Failed to extract imported chart objects:", err);
      return [];
    });

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
    cells: SnapshotCell[];
    defaultFormat?: unknown | null;
    rowFormats?: unknown;
    colFormats?: unknown;
    formatRunsByCol?: unknown;
    colWidths?: unknown;
    rowHeights?: unknown;
  }> = [];
  const truncations: WorkbookLoadTruncation[] = [];

  const formattingBySheetIdPromise: Promise<Array<{ sheetId: string; formatting: SheetFormattingSnapshot | null }>> = Promise.all(
    sheets.map(async (sheet) => {
      try {
        const formatting = (await tauriBackend.getSheetFormatting(sheet.id)) as SheetFormattingSnapshot | null;
        return { sheetId: sheet.id, formatting };
      } catch (err) {
        // Best-effort: treat formatting load failures as "no persisted formatting".
        console.warn(`[formula][desktop] Failed to load formatting for sheet ${sheet.id}:`, err);
        return { sheetId: sheet.id, formatting: null };
      }
    }),
  );

  const viewStateBySheetIdPromise: Promise<Array<{ sheetId: string; view: unknown | null }>> = Promise.all(
    sheets.map(async (sheet) => {
      try {
        const view = (await tauriBackend.getSheetViewState(sheet.id)) as unknown | null;
        return { sheetId: sheet.id, view };
      } catch (err) {
        // Best-effort: treat view state load failures as "no persisted view state".
        console.warn(`[formula][desktop] Failed to load view state for sheet ${sheet.id}:`, err);
        return { sheetId: sheet.id, view: null };
      }
    }),
  );

  const importedColPropertiesBySheetIdPromise: Promise<Array<{ sheetId: string; colProperties: unknown | null }>> = Promise.all(
    sheets.map(async (sheet) => {
      try {
        const colProperties = (await (tauriBackend as any).getSheetImportedColProperties?.(sheet.id)) as unknown | null;
        return { sheetId: sheet.id, colProperties };
      } catch (err) {
        // Best-effort: treat import metadata failures as "no imported column metadata".
        console.warn(`[formula][desktop] Failed to load imported column properties for sheet ${sheet.id}:`, err);
        return { sheetId: sheet.id, colProperties: null };
      }
    }),
  );

  const embeddedCellImagesPromise = tauriBackend
    .listImportedEmbeddedCellImages()
    .catch((err) => {
      // Best-effort: embedded image extraction should never prevent workbook loading.
      console.warn("[formula][desktop] Failed to list embedded cell images:", err);
      return [];
    });

  const clampCellFormatBoundsBySheetId = new Map<string, CellFormatClampBounds | null>();

  for (const sheet of sheets) {
    const cells: SnapshotCell[] = [];

    const usedRange = await tauriBackend.getSheetUsedRange(sheet.id);
    if (!usedRange) {
      clampCellFormatBoundsBySheetId.set(sheet.id, null);
      snapshotSheets.push({
        id: sheet.id,
        name: sheet.name,
        visibility: sheet.visibility ?? "visible",
        tabColor: sheet.tabColor,
        cells,
      });
      continue;
    }

    const { startRow, endRow, startCol, endCol, truncatedRows, truncatedCols } = clampUsedRange(usedRange, {
      maxRows: MAX_ROWS,
      maxCols: MAX_COLS,
    });
    clampCellFormatBoundsBySheetId.set(sheet.id, truncatedRows || truncatedCols ? { startRow, endRow, startCol, endCol } : null);
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
        visibility: sheet.visibility ?? "visible",
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
      visibility: sheet.visibility ?? "visible",
      tabColor: sheet.tabColor,
      cells,
    });
  }

  const formattingBySheetId = new Map<string, SheetFormattingSnapshot | null>();
  for (const { sheetId, formatting } of await formattingBySheetIdPromise) {
    formattingBySheetId.set(sheetId, formatting);
  }

  const viewStateBySheetId = new Map<string, unknown | null>();
  for (const { sheetId, view } of await viewStateBySheetIdPromise) {
    viewStateBySheetId.set(sheetId, view);
  }

  const importedColPropertiesBySheetId = new Map<string, unknown | null>();
  for (const { sheetId, colProperties } of await importedColPropertiesBySheetIdPromise) {
    importedColPropertiesBySheetId.set(sheetId, colProperties);
  }

  // Track user-hidden columns imported from the source workbook so we can seed the UI's outline
  // hidden state ahead of the first render.
  const importedHiddenColsBySheetId = new Map<string, number[]>();

  for (const sheet of snapshotSheets) {
    const formatting = formattingBySheetId.get(sheet.id) ?? null;
    const clampCellFormatsTo = clampCellFormatBoundsBySheetId.get(sheet.id) ?? null;
    const merged = mergeFormattingIntoSnapshot({ cells: sheet.cells, formatting, clampCellFormatsTo });
    Object.assign(sheet, merged);

    const view = viewStateBySheetId.get(sheet.id) ?? null;
    const importedColProperties = importedColPropertiesBySheetId.get(sheet.id) ?? null;
    const colWidths = sheetColWidthsFromViewOrImportedColProperties(view, importedColProperties);
    if (colWidths) sheet.colWidths = colWidths;
    const rowHeights = (view as any)?.rowHeights;
    if (rowHeights && typeof rowHeights === "object") {
      sheet.rowHeights = rowHeights;
    }
    const hiddenCols = hiddenColsFromImportedColProperties(importedColProperties);
    if (hiddenCols.length > 0) importedHiddenColsBySheetId.set(sheet.id, hiddenCols);
  }

  const importedChartObjectsRaw = await importedChartObjectsPromise;
  const importedChartObjects = Array.isArray(importedChartObjectsRaw) ? importedChartObjectsRaw : [];
  const importedChartModels: unknown[] = [];

  // Hydrate DrawingML chart placeholders into the DocumentController snapshot so the drawing overlay
  // can render them. We use a lightweight JSON shape (`{ id, zOrder, anchor, kind }`) that
  // SpreadsheetApp's drawing adapter can consume.
  const drawingsBySheet: Record<string, any[]> = {};
  const knownSheetIds = new Set(sheets.map((s) => s.id));
  const nextZOrderBySheetId = new Map<string, number>();

  for (const entry of importedChartObjects) {
    if (!entry || typeof entry !== "object") continue;
    const e = entry as any;
    const sheetName: string | null =
      typeof e.sheet_name === "string" ? e.sheet_name : typeof e.sheetName === "string" ? e.sheetName : null;
    if (!sheetName) continue;

    const sheetId = workbookSheetStore.resolveIdByName(sheetName) ?? sheetName;
    if (!knownSheetIds.has(sheetId)) continue;

    const rawChartId: string | null =
      typeof e.chart_id === "string"
        ? e.chart_id
        : typeof e.chartId === "string"
          ? e.chartId
          : null;
    const relId: string | null =
      typeof e.rel_id === "string" ? e.rel_id : typeof e.relId === "string" ? e.relId : null;
    const drawingObjectId: number | null =
      typeof e.drawing_object_id === "number"
        ? e.drawing_object_id
        : typeof e.drawingObjectId === "number"
          ? e.drawingObjectId
          : null;

    // Prefer a stable chart id keyed by the internal sheet id rather than the visible sheet name
    // (sheet names can be renamed without changing `sheetId`).
    const stableChartId: string | null = drawingObjectId != null ? `${sheetId}:${String(drawingObjectId)}` : null;
    const chartId: string | null = stableChartId ?? rawChartId ?? relId;

    const id: string | null = drawingObjectId != null ? String(drawingObjectId) : chartId ?? relId;
    if (!id) continue;

    const anchor = e.anchor;
    if (!anchor || typeof anchor !== "object") continue;

    const drawingFrameXml: string =
      typeof e.drawing_frame_xml === "string"
        ? e.drawing_frame_xml
        : typeof e.drawingFrameXml === "string"
          ? e.drawingFrameXml
          : "";
    const drawingObjectName: string | null =
      typeof e.drawing_object_name === "string"
        ? e.drawing_object_name
        : typeof e.drawingObjectName === "string"
          ? e.drawingObjectName
          : null;

    const zOrder = nextZOrderBySheetId.get(sheetId) ?? 0;
    nextZOrderBySheetId.set(sheetId, zOrder + 1);

    const drawing = {
      id,
      zOrder,
      anchor,
      kind: {
        type: "chart",
        ...(chartId ? { chart_id: chartId } : {}),
        ...(relId ? { rel_id: relId } : {}),
        ...(drawingFrameXml ? { raw_xml: drawingFrameXml } : {}),
        ...(drawingObjectName ? { label: drawingObjectName } : {}),
      },
    };

    (drawingsBySheet[sheetId] ??= []).push(drawing);

    // Feed the chart model store using the same stable chart ids used by the drawing layer.
    importedChartModels.push({
      chart_id: chartId,
      rel_id: relId,
      sheet_name: sheetName,
      drawing_object_id: drawingObjectId,
      model: e.model,
    });
  }

  const embeddedCellImages = await embeddedCellImagesPromise;
  const { images: snapshotImages } = mergeEmbeddedCellImagesIntoSnapshot({
    sheets: snapshotSheets,
    embeddedCellImages,
    resolveSheetIdByName: (name) => workbookSheetStore.resolveIdByName(name) ?? null,
    sheetIdsInOrder: sheets.map((s) => s.id),
    maxRows: MAX_ROWS,
    maxCols: MAX_COLS,
  });

  // Hydrate imported DrawingML floating objects + image store entries from the workbook origin bytes
  // so the drawings overlay renders real workbook drawings on open (best-effort).
  let importedDrawingImages: Array<{ id: string; bytesBase64: string; mimeType?: string | null }> = [];
  let importedDrawingsBySheet: Record<string, any[]> = {};
  try {
    const imported = await tauriBackend.getImportedDrawingLayer();
    const additions = buildImportedDrawingLayerSnapshotAdditions(imported, workbookSheetStore);
    if (additions) {
      importedDrawingImages = additions.images;
      importedDrawingsBySheet = additions.drawingsBySheet;
    }
  } catch (err) {
    console.warn("[formula][desktop] Failed to load imported drawing objects/images:", err);
  }

  // Merge imported drawing-layer images into the snapshot image store, deduping by `id`.
  const imagesById = new Map<string, any>();
  for (const img of snapshotImages) {
    const id = typeof (img as any)?.id === "string" ? String((img as any).id) : "";
    const bytesBase64 = typeof (img as any)?.bytesBase64 === "string" ? (img as any).bytesBase64 : "";
    if (!id || !bytesBase64) continue;
    imagesById.set(id, img);
  }
  for (const img of importedDrawingImages) {
    const id = typeof (img as any)?.id === "string" ? String((img as any).id) : "";
    const bytesBase64 = typeof (img as any)?.bytesBase64 === "string" ? (img as any).bytesBase64 : "";
    if (!id || !bytesBase64) continue;
    if (imagesById.has(id)) continue;
    imagesById.set(id, img);
  }
  const mergedImages = Array.from(imagesById.values()).sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  // Merge imported drawing objects into `drawingsBySheet`. Prefer chart placeholder objects already
  // populated from `listImportedChartObjects()` when there are id collisions.
  for (const [sheetId, importedList] of Object.entries(importedDrawingsBySheet)) {
    if (!knownSheetIds.has(sheetId)) continue;
    if (!Array.isArray(importedList) || importedList.length === 0) continue;

    const existing = drawingsBySheet[sheetId] ? [...drawingsBySheet[sheetId]!] : [];
    const existingIds = new Set<string>();
    for (const d of existing) {
      const id = (d as any)?.id != null ? String((d as any).id) : "";
      if (id) existingIds.add(id);
    }

    for (const d of importedList) {
      if (!d || typeof d !== "object") continue;
      const id = (d as any).id != null ? String((d as any).id) : "";
      if (!id || existingIds.has(id)) continue;
      existingIds.add(id);
      existing.push(d);
    }

    if (existing.length > 0) drawingsBySheet[sheetId] = existing;
  }

  const snapshotPayload: any = { schemaVersion: 1, sheets: snapshotSheets };
  if (Object.keys(drawingsBySheet).length > 0) snapshotPayload.drawingsBySheet = drawingsBySheet;
  if (mergedImages.length > 0) snapshotPayload.images = mergedImages;
  const snapshot = encodeDocumentSnapshot(snapshotPayload);
  const workbookSignature = await workbookSignaturePromise;
  // Reset Power Query table signatures before applying the snapshot so any
  // in-flight query executions cannot reuse cached table results from a
  // previously-opened workbook.
  const signaturesMod = await loadPowerQueryTableSignaturesModule();
  signaturesMod?.refreshTableSignaturesFromBackend(doc, [], { workbookSignature });
  signaturesMod?.refreshDefinedNameSignaturesFromBackend(doc, [], { workbookSignature });

  // Keep the WASM formula engine's workbook file metadata (directory/filename) in sync with the
  // desktop workbook identity so `CELL("filename")` / `INFO("directory")` update immediately.
  //
  // Set the metadata *before* `restoreDocumentState` hydrates the worker so the initial recalc
  // observes the correct path without needing an extra recalculation pass.
  const workbookFileMetadata = getWorkbookFileMetadataFromWorkbookInfo(info);
  await app.setWorkbookFileMetadata(workbookFileMetadata.directory, workbookFileMetadata.filename, { syncEngine: false });

  // Seed per-sheet hidden columns from the imported workbook model before restoring the
  // DocumentController snapshot. SpreadsheetApp's row/col visibility is tracked in a local
  // per-sheet outline map; if we don't seed it here, hidden columns can momentarily render
  // as visible until the user interacts with the grid.
  try {
    // Clear any outline state from a previously-opened workbook. Sheet ids are often derived from
    // sheet names (e.g. "Sheet1"), so they can collide across workbook boundaries.
    //
    // Even when the newly-opened workbook has no imported hidden columns, we still need to clear
    // this map so user-hidden rows/cols (and outline groups) don't leak into the next workbook.
    const outlinesBySheet = (app as any).outlinesBySheet as Map<string, unknown> | undefined;
    if (outlinesBySheet && typeof outlinesBySheet.clear === "function") outlinesBySheet.clear();

    if (importedHiddenColsBySheetId.size > 0) {
      const getOutlineForSheet = (app as any).getOutlineForSheet as ((sheetId: string) => any) | undefined;
      if (typeof getOutlineForSheet === "function") {
        for (const [sheetId, cols] of importedHiddenColsBySheetId) {
          const outline = getOutlineForSheet.call(app, sheetId);
          if (!outline) continue;
          const axis = outline?.cols;
          const entryMut = axis?.entryMut;
          if (typeof entryMut !== "function") continue;
          for (const col of cols) {
            const idx = Number(col);
            if (!Number.isInteger(idx) || idx < 0) continue;
            const entry = entryMut.call(axis, idx + 1);
            if (entry?.hidden) entry.hidden.user = true;
          }
        }
      }
    }

    // SpreadsheetApp's legacy renderer caches row/col visibility to avoid consulting outline state
    // for every paint. When the active sheet id stays the same across workbook opens (e.g. two
    // unrelated workbooks both have `Sheet1`), `restoreDocumentState` may not rebuild those caches.
    //
    // Rebuild the caches after clearing/seeding outline state so the first post-open render matches
    // Excel (both for imported hidden columns and for the "no hidden columns" case).
    const rebuildAxisVisibilityCache = (app as any).rebuildAxisVisibilityCache as (() => void) | undefined;
    if (typeof rebuildAxisVisibilityCache === "function") rebuildAxisVisibilityCache.call(app);
  } catch (err) {
    console.warn("[formula][desktop] Failed to reset outline state on workbook open:", err);
  }
  await app.restoreDocumentState(snapshot);

  // Best-effort: propagate imported hidden-column flags into the WASM engine metadata so
  // Excel-compatible worksheet functions (e.g. `CELL("width")`) can observe them once the
  // engine-side semantics are implemented.
  if (importedHiddenColsBySheetId.size > 0) {
    try {
      const wasmEngine = (app as any).wasmEngine as any;
      if (wasmEngine && typeof wasmEngine.setColHidden === "function") {
        for (const [sheetId, cols] of importedHiddenColsBySheetId) {
          for (const col of cols) {
            const idx = Number(col);
            if (!Number.isInteger(idx) || idx < 0) continue;
            // The WASM engine workbook is keyed by stable sheet ids (not display/tab names). Use
            // the DocumentController sheet id directly so we don't accidentally create phantom
            // sheets after a sheet rename.
            await wasmEngine.setColHidden(idx, true, sheetId);
          }
        }

        // If the engine produces value-change deltas (e.g. volatile `CELL("width")`), apply them
        // to keep computed-value caches coherent.
        if (typeof wasmEngine.recalculate === "function" && typeof (app as any).applyComputedChanges === "function") {
          const changes = await wasmEngine.recalculate();
          await (app as any).applyComputedChanges(changes);
        }
      }
    } catch (err) {
      console.warn("[formula][desktop] Failed to sync imported hidden columns into WASM engine:", err);
    }
  }

  // Hydrate worksheet background images (`<picture r:id="...">`) from the opened XLSX package.
  // Best-effort: failures should never prevent the workbook from loading.
  try {
    await hydrateSheetBackgroundImagesFromBackend({ app, workbookSheetStore, backend: tauriBackend });
  } catch (err) {
    console.warn("[formula][desktop] Failed to hydrate worksheet background images:", err);
  }

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
  signaturesMod?.refreshTableSignaturesFromBackend(doc, normalizedTables as any, { workbookSignature });
  const normalizedDefinedNames = definedNames.map((entry) => {
    const refers_to = typeof (entry as any)?.refers_to === "string" ? String((entry as any).refers_to) : "";
    const { sheetName: explicitSheetName } = splitSheetQualifier(refers_to);
    const sheetIdFromRef = explicitSheetName ? workbookSheetStore.resolveIdByName(explicitSheetName) ?? explicitSheetName : null;
    const rawScopeSheet = typeof (entry as any)?.sheet_id === "string" ? String((entry as any).sheet_id) : null;
    const sheetIdFromScope = rawScopeSheet ? workbookSheetStore.resolveIdByName(rawScopeSheet) ?? rawScopeSheet : null;
    return { ...(entry as any), sheet_id: sheetIdFromScope ?? sheetIdFromRef };
  });
  signaturesMod?.refreshDefinedNameSignaturesFromBackend(doc, normalizedDefinedNames as any, { workbookSignature });

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

  // Populate chart models extracted from imported XLSX chart parts so DrawingML chart placeholders
  // can be rendered via the canvas chart renderer.
  app.setImportedChartModels(importedChartModels);

  doc.markSaved();

  // Default to the first *visible* sheet so hidden sheets are never activated
  // (and therefore never shown as the selected value in the sheet switcher).
  const visibleSheets = listSheetsForUi();
  const firstSheetId = visibleSheets[0]?.id ?? sheets[0].id;
  app.activateSheet(firstSheetId);
  app.activateCell({ sheetId: firstSheetId, row: 0, col: 0 });
  app.refresh();
}

const OPEN_WORKBOOK_PASSWORD_REQUIRED_PREFIX = "PASSWORD_REQUIRED:";
const OPEN_WORKBOOK_INVALID_PASSWORD_PREFIX = "INVALID_PASSWORD:";

function extractInvokeErrorMessage(err: unknown): string {
  if (typeof err === "string") return err;
  if (err && typeof err === "object") {
    const msg = (err as any).message;
    if (typeof msg === "string") return msg;
  }
  return String(err);
}

function isAbortError(err: unknown): boolean {
  return Boolean(err && typeof err === "object" && (err as any).name === "AbortError");
}

async function openWorkbookWithPasswordPrompt(
  path: string,
): Promise<WorkbookInfo> {
  if (!tauriBackend) throw new Error("Desktop backend is not available");

  const abort = () => {
    const err = new Error("Open workbook cancelled");
    err.name = "AbortError";
    throw err;
  };

  const promptPassword = async (prompt: string): Promise<string> => {
    const password = await showInputBox({
      prompt,
      placeHolder: "Password",
      type: "password",
      okLabel: "Open",
      cancelLabel: "Cancel",
    });
    if (password == null) abort();
    return password;
  };

  let password: string | null = null;
  for (;;) {
    try {
      if (password == null) return await tauriBackend.openWorkbook(path);
      return await tauriBackend.openWorkbook(path, { password });
    } catch (err) {
      const msg = extractInvokeErrorMessage(err).trim();
      if (msg.includes(OPEN_WORKBOOK_PASSWORD_REQUIRED_PREFIX)) {
        password = await promptPassword("This workbook is password-protected. Enter password to open.");
        continue;
      }
      if (msg.includes(OPEN_WORKBOOK_INVALID_PASSWORD_PREFIX)) {
        password = await promptPassword("Invalid password. Try again.");
        continue;
      }
      throw err;
    }
  }
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
        const mod = await loadVbaEventMacrosModule();
        if (mod) {
          await mod.fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        }
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

    activeWorkbook = await openWorkbookWithPasswordPrompt(path);
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
    void startPowerQueryService();
    rerenderLayout?.();

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });

    if (queuedInvoke) {
      void (async () => {
        const mod = await loadVbaEventMacrosModule();
        if (!mod) return;
        vbaEventMacros = mod.installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
      })().catch((err) => {
        console.error("Failed to install VBA event macros:", err);
      });
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
        void (async () => {
          const mod = await loadVbaEventMacrosModule();
          if (!mod) return;
          vbaEventMacros = mod.installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        })().catch((err) => {
          console.error("Failed to install VBA event macros:", err);
        });
      }
    }
    void startPowerQueryService();
    if (isAbortError(err) && !options.throwOnCancel) return;
    throw err;
  }
}

async function promptOpenWorkbook(): Promise<void> {
  if (!tauriBackend) return;
  const { open } = getTauriDialogOrThrow();
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
    const mod = await loadPowerQueryServiceModule();
    if (mod) {
      const queries = mod.loadQueriesFromStorage(fromWorkbookId);
      if (queries.length > 0) mod.saveQueriesToStorage(toWorkbookId, queries);
    }
  } catch {
    // Ignore storage failures (disabled storage, quota, etc).
  }

  // Scheduled refresh metadata is persisted via the RefreshStateStore abstraction.
  try {
    const mod = await loadPowerQueryRefreshStateStoreModule();
    if (mod) {
      const fromStore = mod.createPowerQueryRefreshStateStore({ workbookId: fromWorkbookId });
      const toStore = mod.createPowerQueryRefreshStateStore({ workbookId: toWorkbookId });
      const state = await fromStore.load();
      if (state && Object.keys(state).length > 0) {
        await toStore.save(state);
      }
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

  const savePath = coerceSavePathToXlsx(activeWorkbook.path);

  if (options.notifyExtensions !== false) {
    try {
      emitBeforeSaveForExtensions(getWorkbookSnapshotForExtensions({ pathOverride: savePath }));
    } catch {
      // Ignore extension host errors; save should still succeed.
    }
  }
  const previousPanelWorkbookId = activePanelWorkbookId;
  await workbookSync.markSaved();

  // Some desktop backends (CSV/XLS imports, etc) coerce "Save" targets to `.xlsx` by updating the
  // underlying workbook path without an explicit Save As flow. Keep the UI + WASM workbook metadata
  // in sync with that new identity so Excel-compatible functions update immediately.
  if (savePath !== activeWorkbook.path) {
    activeWorkbook = { ...activeWorkbook, path: savePath };
    syncTitlebar();

    const fileMetadata = getWorkbookFileMetadataFromWorkbookInfo({ path: savePath, origin_path: null });
    void app.setWorkbookFileMetadata(fileMetadata.directory, fileMetadata.filename);

    await copyPowerQueryPersistence(previousPanelWorkbookId, savePath);
    activePanelWorkbookId = savePath;
    // Ensure zoom persistence follows the workbook when its id changes due to save-path coercion.
    persistCurrentSharedGridZoom();
    startPowerQueryService();
    rerenderLayout?.();
  }
}

async function handleSaveWithPassword(): Promise<void> {
  commitAllPendingEditsForCommand();
  if (!tauriBackend) return;
  if (!activeWorkbook) return;

  const password = await showInputBox({
    prompt: "Encrypt workbook with password",
    placeHolder: "Password",
    type: "password",
    okLabel: "Encrypt",
    cancelLabel: "Cancel",
  });
  if (password == null) {
    // If the user cancels the password prompt, fall back to the normal save flow.
    await handleSave();
    return;
  }
  if (password.trim() === "") {
    showToast("Password cannot be empty.", "error");
    return;
  }

  if (!activeWorkbook.path) {
    const { save } = getTauriDialogOrThrow();
    const path = await save({
      filters: [
        { name: t("fileDialog.filters.excelWorkbook"), extensions: ["xlsx"] },
        { name: "Excel Macro-Enabled Workbook", extensions: ["xlsm"] },
      ],
    });
    if (!path) return;
    await handleSaveAsPath(path, { password });
    return;
  }

  const savePath = coerceSavePathToXlsx(activeWorkbook.path);
  await handleSaveAsPath(savePath, { password });
}

async function handleSaveAs(
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean; throwOnCancel?: boolean } = {},
): Promise<void> {
  commitAllPendingEditsForCommand();
  if (!tauriBackend) return;
  if (!activeWorkbook) return;

  const previousPanelWorkbookId = options.previousPanelWorkbookId ?? activePanelWorkbookId;
  const { save } = getTauriDialogOrThrow();
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
  options: { previousPanelWorkbookId?: string; notifyExtensions?: boolean; password?: string } = {},
): Promise<void> {
  commitAllPendingEditsForCommand();
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (typeof path !== "string" || path.trim() === "") return;

  const previousPanelWorkbookId = options.previousPanelWorkbookId ?? activePanelWorkbookId;
  const savePath = coerceSavePathToXlsx(path);

  // Ensure any pending microtask-batched workbook edits are flushed before saving.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await drainBackendSync();
  if (options.notifyExtensions !== false) {
    try {
      emitBeforeSaveForExtensions(getWorkbookSnapshotForExtensions({ pathOverride: savePath }));
    } catch {
      // Ignore extension host errors; save should still succeed.
    }
  }
  if (queuedInvoke) {
    const args: Record<string, unknown> = { path: savePath };
    if (options.password != null) args.password = options.password;
    await queuedInvoke("save_workbook", args);
  } else {
    await tauriBackend.saveWorkbook(savePath, options.password != null ? { password: options.password } : undefined);
  }
  activeWorkbook = { ...activeWorkbook, path: savePath };
  app.getDocument().markSaved();
  syncTitlebar();

  // Save As establishes a new file identity; keep the WASM engine workbook metadata in sync so
  // `CELL("filename")` / `INFO("directory")` update immediately (Excel semantics).
  const fileMetadata = getWorkbookFileMetadataFromWorkbookInfo({ path: savePath, origin_path: null });
  void app.setWorkbookFileMetadata(fileMetadata.directory, fileMetadata.filename);

  await copyPowerQueryPersistence(previousPanelWorkbookId, savePath);
  activePanelWorkbookId = savePath;
  // Ensure zoom persistence follows the workbook when its id changes (e.g. Save As).
  persistCurrentSharedGridZoom();
  void startPowerQueryService();
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
        const mod = await loadVbaEventMacrosModule();
        if (mod) {
          await mod.fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        }
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
    void startPowerQueryService();
    rerenderLayout?.();

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });

    if (queuedInvoke) {
      void (async () => {
        const mod = await loadVbaEventMacrosModule();
        if (!mod) return;
        vbaEventMacros = mod.installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
      })().catch((err) => {
        console.error("Failed to install VBA event macros:", err);
      });
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
        void (async () => {
          const mod = await loadVbaEventMacrosModule();
          if (!mod) return;
          vbaEventMacros = mod.installVbaEventMacros({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
        })().catch((err) => {
          console.error("Failed to install VBA event macros:", err);
        });
      }
    }
    void startPowerQueryService();
    throw err;
  }
}

try {
  tauriBackend = new TauriWorkbookBackend();
  const invoke = getTauriInvokeOrNull() as TauriInvoke | null;
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
  // - menu-open, menu-new, menu-save, menu-save-as, menu-print, menu-print-preview, menu-export-pdf, menu-close-window, menu-quit,
  //   menu-undo, menu-redo, menu-cut, menu-copy, menu-paste, menu-paste-special, menu-select-all,
  //   menu-zoom-in, menu-zoom-out, menu-zoom-reset, menu-about, menu-check-updates, menu-open-release-page
  // - startup:window-visible, startup:webview-loaded, startup:tti, startup:metrics
  // - update-check-started, update-check-already-running, update-not-available, update-check-error, update-available
  // - update-download-started, update-download-progress, update-downloaded, update-download-error
  // - oauth-redirect
  // JS -> Rust (`emit`):
  // - open-file-ready, oauth-redirect-ready
  // - close-prep-done, close-handled
  // - updater-ui-ready, coi-check-result
  const { listen, emit } = getTauriEventApiOrThrow();
  // File/open/save operations can trigger expensive async work (backend sync, workbook loading).
  // Serialize them so overlapping menu/IPC triggers can't interleave and cause inconsistent UI
  // state or missed extension lifecycle events.
  let pendingWorkbookFileCommands: Promise<void> = Promise.resolve();

  const queueWorkbookFileCommand = (task: () => Promise<void>) => {
    pendingWorkbookFileCommands = pendingWorkbookFileCommands
      .then(task)
      .catch((err) => console.error("Workbook file command failed:", err));
  };

  const queueOpenWorkbook = (path: string) => {
    queueWorkbookFileCommand(async () => {
      const task = (async () => {
        // Ensure the window is visible so any UI feedback (toasts, prompts) is visible even
        // when the app is running in the background (e.g. opened via file association while
        // minimized to tray).
        await showTauriWindowBestEffort();
        await openWorkbookFromPath(path);
      })();
      // Include workbook open in `__formulaApp.whenIdle()` so e2e tests (and any UI flows
      // that rely on `whenIdle`) don't race sheet-store/UI updates while the workbook is
      // still loading.
      app.trackIdleTask(task, { swallowErrors: true });
      try {
        await task;
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
      const mod = await loadVbaEventMacrosModule();
      if (!mod) return;
      await mod.fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
    },
    drainBackendSync,
    quitApp: async () => {
      await flushCollabLocalPersistenceBestEffort({
        session: app.getCollabSession?.() ?? null,
        whenIdle: async () => {
          await app.whenCollabBinderIdle();
          await app.whenIdle();
        },
        // Allow slightly more time for large paste/edit batches to propagate through
        // the DocumentController → Yjs binder (encryption can make this async).
        idleTimeoutMs: 1_000,
      });
      if (!invoke) {
        window.close();
        return;
      }
      // Exit the desktop shell. The backend command hard-exits the process so this promise
      // will never resolve in the success path.
      await invoke("quit_app");
    },
    restartApp: async () => {
      await flushCollabLocalPersistenceBestEffort({
        session: app.getCollabSession?.() ?? null,
        whenIdle: async () => {
          await app.whenCollabBinderIdle();
          await app.whenIdle();
        },
        idleTimeoutMs: 1_000,
      });
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
    void (async () => {
      const mod = await loadPowerQueryOauthBrokerModule();
      if (!mod) return;
      mod.oauthBroker.observeRedirect(redirectUrl);
    })().catch((err) => {
      console.error("Failed to handle oauth redirect:", err);
    });
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
  installOpenFileIpc({
    listen,
    emit,
    onOpenPath: (path) => {
      // Ignore non-spreadsheet open requests (defensive): the desktop app may receive file-open IPC
      // events from OS integrations and we should avoid attempting to open images/other assets.
      if (!isOpenWorkbookPath(path)) return;
      queueOpenWorkbook(path);
    },
  });

  void listen("file-dropped", async (event) => {
    const paths = (event as any)?.payload;
    const firstWorkbookPath = Array.isArray(paths)
      ? paths.find((p) => typeof p === "string" && p.trim() !== "" && isOpenWorkbookPath(p))
      : null;
    if (!firstWorkbookPath) return;
    // Ignore non-spreadsheet drops (e.g. images) so drag/drop picture insertion doesn't trigger
    // a spurious "open workbook" flow.
    queueOpenWorkbook(firstWorkbookPath);
  });

  void listen("tray-open", () => {
    queueWorkbookFileCommand(async () => {
      try {
        // Ensure the window is visible so any toast-based error handling is observable.
        await showTauriWindowBestEffort();
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.openWorkbook);
      } catch (err) {
        console.error("Failed to open workbook:", err);
        void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
      }
    });
  });

  void listen("tray-new", () => {
    queueWorkbookFileCommand(async () => {
      try {
        // Ensure the window is visible so any toast-based error handling is observable.
        await showTauriWindowBestEffort();
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.newWorkbook);
      } catch (err) {
        console.error("Failed to create workbook:", err);
        void nativeDialogs.alert(`Failed to create workbook: ${String(err)}`);
      }
    });
  });

  void listen("tray-quit", () => {
    void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.quit).catch((err) => {
      console.error("Failed to quit app:", err);
    });
  });

  // Native menu bar integration (desktop shell emits menu-open/menu-save/... events).
  void listen("menu-open", () => {
    queueWorkbookFileCommand(async () => {
      try {
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.openWorkbook);
      } catch (err) {
        console.error("Failed to open workbook:", err);
        void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
      }
    });
  });

  void listen("menu-new", () => {
    queueWorkbookFileCommand(async () => {
      try {
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.newWorkbook);
      } catch (err) {
        console.error("Failed to create workbook:", err);
        void nativeDialogs.alert(`Failed to create workbook: ${String(err)}`);
      }
    });
  });

  void listen("menu-save", () => {
    queueWorkbookFileCommand(async () => {
      try {
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.saveWorkbook);
      } catch (err) {
        console.error("Failed to save workbook:", err);
        void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
      }
    });
  });

  void listen("menu-save-as", () => {
    queueWorkbookFileCommand(async () => {
      try {
        await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.saveWorkbookAs);
      } catch (err) {
        console.error("Failed to save workbook:", err);
        void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
      }
    });
  });

  void listen("menu-print", () => {
    void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.print).catch((err) => {
      console.error("Failed to print:", err);
      showToast(`Failed to print: ${String(err)}`, "error");
    });
  });

  void listen("menu-print-preview", () => {
    void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.printPreview).catch((err) => {
      console.error("Failed to open print preview:", err);
      showToast(`Failed to open print preview: ${String(err)}`, "error");
    });
  });

  void listen("menu-export-pdf", () => {
    void commandRegistry.executeCommand(PAGE_LAYOUT_COMMANDS.exportPdf).catch((err) => {
      console.error("Failed to export PDF:", err);
      showToast(`Failed to export PDF: ${String(err)}`, "error");
    });
  });

  void listen("menu-close-window", () => {
    void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.closeWorkbook).catch((err) => {
      console.error("Failed to close window:", err);
    });
  });

  void listen("menu-quit", () => {
    void commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.quit).catch((err) => {
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
    try {
      const provider = await getClipboardProvider();
      const { text } = await provider.read();
      return typeof text === "string" ? text : null;
    } catch {
      return null;
    }
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
      const getName = getTauriAppGetNameOrNull();
      const getVersion = getTauriAppGetVersionOrNull();

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
    void (async () => {
      // Ensure the window is visible so any toast-based error handling is observable.
      await showTauriWindowBestEffort();
      await commandRegistry.executeCommand(WORKBENCH_FILE_COMMANDS.openWorkbook);
    })().catch((err) => {
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
          const mod = await loadVbaEventMacrosModule();
          if (mod) {
            await mod.fireWorkbookBeforeCloseBestEffort({ app, workbookId, invoke: queuedInvoke, drainBackendSync });
          }
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
      await flushCollabLocalPersistenceBestEffort({
        session: app.getCollabSession?.() ?? null,
        whenIdle: async () => {
          await app.whenCollabBinderIdle();
          await app.whenIdle();
        },
        idleTimeoutMs: 1_000,
      });
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
} catch {
  // Not running under Tauri; desktop host integration is unavailable.
}

// Expose a small API for Playwright assertions.
window.__formulaApp = app;
window.__formulaExtensionHostManager = extensionHostManagerForE2e;
window.__formulaExtensionHost = extensionHostManagerForE2e?.host ?? null;
window.__workbookSheetStore = workbookSheetStore;

// Time-to-interactive instrumentation (best-effort, no-op for web builds).
void markStartupTimeToInteractive({ whenIdle: () => app.whenIdle() })
  .catch(() => {})
  .finally(() => {
    // Phase 2 boot: kick off heavy/optional subsystems after TTI so the first paint + initial
    // spreadsheet interactions don't pay the parse/execute cost of rarely-used features.
    startupMark("formula:desktop:boot:phase2-start");

    const schedule = (task: () => void) => {
      try {
        const ric = (window as any).requestIdleCallback as ((cb: () => void, opts?: { timeout: number }) => void) | undefined;
        if (typeof ric === "function") {
          ric(task, { timeout: 5_000 });
          return;
        }
      } catch {
        // ignore
      }
      setTimeout(task, 0);
    };

    schedule(() => {
      // Power Query spins up a WASM/worker-backed engine and can be expensive to initialize.
      // Start it in the background after TTI so query panels can appear instantly when opened.
      if (tauriBackend) {
        startPowerQueryService();
      }
    });

    schedule(() => {
      // Split view (secondary grid) is large and not required for the initial shell render.
      // Preload it after TTI so the toggle command can create the secondary pane instantly.
      void loadSecondaryGridViewModule().then((mod) => {
        if (mod) secondaryGridViewModule = mod;
      });
    });
  });
