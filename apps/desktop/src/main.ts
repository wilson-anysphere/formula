import { SpreadsheetApp } from "./app/spreadsheetApp";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/command-palette.css";
import "./styles/workspace.css";
import "./styles/comments.css";
import "./styles/shell.css";
import "./styles/auditing.css";

import { ThemeController } from "./theme/themeController.js";

import { mountRibbon } from "./ribbon/index.js";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
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
import { DocumentControllerWorkbookAdapter } from "./scripting/documentControllerWorkbookAdapter.js";
import { registerFindReplaceShortcuts, FindReplaceController } from "./panels/find-replace/index.js";
import { t } from "./i18n/index.js";
import { formatRangeAddress, parseRangeAddress } from "@formula/scripting";
import { normalizeFormulaTextOpt } from "@formula/engine";
import { startWorkbookSync } from "./tauri/workbookSync";
import { TauriWorkbookBackend } from "./tauri/workbookBackend";
import * as nativeDialogs from "./tauri/nativeDialogs";
import { setTrayStatus } from "./tauri/trayStatus";
import { installUpdaterUi } from "./tauri/updaterUi";
import type { WorkbookInfo } from "@formula/workbook-backend";
import { chartThemeFromWorkbookPalette } from "./charts/theme";
import { parseA1Range, splitSheetQualifier } from "../../../packages/search/index.js";
import { refreshDefinedNameSignaturesFromBackend, refreshTableSignaturesFromBackend } from "./power-query/tableSignatures";
import {
  DesktopPowerQueryService,
  loadQueriesFromStorage,
  saveQueriesToStorage,
  setDesktopPowerQueryService,
} from "./power-query/service.js";
import { createPowerQueryRefreshStateStore } from "./power-query/refreshStateStore.js";
import { showInputBox, showQuickPick, showToast } from "./extensions/ui.js";
import { DesktopExtensionHostManager } from "./extensions/extensionHostManager.js";
import { ExtensionPanelBridge } from "./extensions/extensionPanelBridge.js";
import { ContextKeyService } from "./extensions/contextKeys.js";
import { resolveMenuItems } from "./extensions/contextMenus.js";
import { buildContextMenuModel } from "./extensions/contextMenuModel.js";
import { matchesKeybinding, parseKeybinding, platformKeybinding, type ContributedKeybinding } from "./extensions/keybindings.js";
import { deriveSelectionContextKeys } from "./extensions/selectionContextKeys.js";
import { evaluateWhenClause } from "./extensions/whenClause.js";
import { CommandRegistry } from "./extensions/commandRegistry.js";
import type { Range, SelectionState } from "./selection/types";
import { ContextMenu, type ContextMenuItem } from "./menus/contextMenu.js";
import { getPasteSpecialMenuItems } from "./clipboard/pasteSpecial.js";
import {
  applyAllBorders,
  applyNumberFormatPreset,
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
import {
  getDefaultSeedStoreStorage,
  readContributedPanelsSeedStore,
  removeSeedPanelsForExtension,
  seedPanelRegistryFromContributedPanelsSeedStore,
  setSeedPanelsForExtension,
} from "./extensions/contributedPanelsSeedStore.js";

import sampleHelloManifest from "../../../extensions/sample-hello/package.json";

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

const workbookSheetNames = new Map<string, string>();

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

const gridRoot = document.getElementById("grid");
if (!gridRoot) {
  throw new Error("Missing #grid container");
}

const titlebarRoot = document.getElementById("titlebar");
if (!titlebarRoot) {
  throw new Error("Missing #titlebar container");
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

const activeCell = document.querySelector<HTMLElement>('[data-testid="active-cell"]');
const selectionRange = document.querySelector<HTMLElement>('[data-testid="selection-range"]');
const activeValue = document.querySelector<HTMLElement>('[data-testid="active-value"]');
const statusMode = document.querySelector<HTMLElement>('[data-testid="status-mode"]');
const sheetSwitcher = document.querySelector<HTMLSelectElement>('[data-testid="sheet-switcher"]');
const openComments = document.querySelector<HTMLButtonElement>('[data-testid="open-comments-panel"]');
const auditPrecedents = document.querySelector<HTMLButtonElement>('[data-testid="audit-precedents"]');
const auditDependents = document.querySelector<HTMLButtonElement>('[data-testid="audit-dependents"]');
const auditTransitive = document.querySelector<HTMLButtonElement>('[data-testid="audit-transitive"]');
const openVbaMigratePanel = document.querySelector<HTMLButtonElement>('[data-testid="open-vba-migrate-panel"]');
if (!activeCell || !selectionRange || !activeValue || !statusMode || !sheetSwitcher) {
  throw new Error("Missing status bar elements");
}
const sheetSwitcherEl = sheetSwitcher;
if (!openComments) {
  throw new Error("Missing comments panel toggle button");
}
if (!auditPrecedents || !auditDependents || !auditTransitive) {
  throw new Error("Missing auditing toolbar buttons");
}

const workbookId = "local-workbook";
const app = new SpreadsheetApp(
  gridRoot,
  { activeCell, selectionRange, activeValue },
  { formulaBar: formulaBarRoot, workbookId },
);

function normalizeSelectionRange(range: Range): CellRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

function selectionRangesForFormatting(): CellRange[] {
  const ranges = app.getSelectionRanges();
  if (ranges.length === 0) {
    const cell = app.getActiveCell();
    return [{ start: { row: cell.row, col: cell.col }, end: { row: cell.row, col: cell.col } }];
  }
  return ranges.map(normalizeSelectionRange);
}

function rgbHexToArgb(rgb: string): string | null {
  if (!/^#[0-9A-Fa-f]{6}$/.test(rgb)) return null;
  // DocumentController formatting expects #AARRGGBB.
  return ["#", "FF", rgb.slice(1)].join("");
}

function applyToSelection(label: string, fn: (sheetId: string, ranges: CellRange[]) => void): void {
  const doc = app.getDocument();
  const sheetId = app.getCurrentSheetId();
  const ranges = selectionRangesForFormatting();
  const shouldBatch = ranges.length > 1;

  if (shouldBatch) doc.beginBatch({ label });
  try {
    fn(sheetId, ranges);
  } finally {
    if (shouldBatch) doc.endBatch();
  }
  app.focus();
}

function createHiddenColorInput(): HTMLInputElement {
  const input = document.createElement("input");
  input.type = "color";
  input.tabIndex = -1;
  input.style.position = "fixed";
  input.style.left = "-1000px";
  input.style.top = "-1000px";
  input.style.opacity = "0";
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
      applyToSelection(label, (sheetId, ranges) => apply(sheetId, ranges, argb));
    },
    { once: true },
  );
  input.click();
}

// Panels persist state keyed by a workbook/document identifier. For file-backed workbooks we use
// their on-disk path; for unsaved sessions we generate a random session id so distinct new
// workbooks don't collide.
let activePanelWorkbookId = workbookId;
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();

app.onEditStateChange((isEditing) => {
  statusMode.textContent = isEditing ? "Edit" : "Ready";
});

app.focus();

const titlebar = mountTitlebar(titlebarRootEl, {
  actions: [],
  undoRedo: {
    ...app.getUndoRedoState(),
    onUndo: () => {
      app.undo();
      app.focus();
    },
    onRedo: () => {
      app.redo();
      app.focus();
    },
  },
});

const syncTitlebarUndoRedo = () => {
  titlebar.update({
    actions: [],
    undoRedo: {
      ...app.getUndoRedoState(),
      onUndo: () => {
        app.undo();
        app.focus();
      },
      onRedo: () => {
        app.redo();
        app.focus();
      },
    },
  });
};

syncTitlebarUndoRedo();
const unsubscribeTitlebarHistory = app.getDocument().on("history", () => syncTitlebarUndoRedo());
window.addEventListener("unload", () => {
  unsubscribeTitlebarHistory();
  titlebar.dispose();
});

openComments.addEventListener("click", () => app.toggleCommentsPanel());
auditPrecedents.addEventListener("click", () => {
  app.toggleAuditingPrecedents();
  app.focus();
});
auditDependents.addEventListener("click", () => {
  app.toggleAuditingDependents();
  app.focus();
});
auditTransitive.addEventListener("click", () => {
  app.toggleAuditingTransitive();
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

// --- Sheet tabs (minimal multi-sheet support) ---------------------------------

const sheetTabsRoot = document.getElementById("sheet-tabs");
if (!sheetTabsRoot) {
  throw new Error("Missing #sheet-tabs container");
}
const sheetTabsRootEl = sheetTabsRoot;
// The HTML scaffold historically used `.sheet-tabs` for the container itself. The
// modern sheet bar uses `.sheet-bar` with an inner `.sheet-tabs` strip (matching
// `mockups/spreadsheet-main.html`), so update the class list once on boot.
sheetTabsRootEl.classList.add("sheet-bar");
sheetTabsRootEl.classList.remove("sheet-tabs");

let lastSheetIds: string[] = [];

type SheetUiInfo = { id: string; name: string };

function listSheetsForUi(): SheetUiInfo[] {
  const sheetIds = app.getDocument().getSheetIds();
  const ids = sheetIds.length > 0 ? sheetIds : ["Sheet1"];
  return ids.map((id) => ({ id, name: workbookSheetNames.get(id) ?? id }));
}

function renderSheetTabs(sheets: SheetUiInfo[] = listSheetsForUi()) {
  lastSheetIds = sheets.map((sheet) => sheet.id);
  sheetTabsRootEl.replaceChildren();

  const active = app.getCurrentSheetId();

  const nav = document.createElement("div");
  nav.className = "sheet-nav";

  const navLeft = document.createElement("button");
  navLeft.type = "button";
  navLeft.className = "sheet-nav-btn";
  navLeft.textContent = "◀";
  navLeft.setAttribute("aria-label", "Scroll sheets left");

  const navRight = document.createElement("button");
  navRight.type = "button";
  navRight.className = "sheet-nav-btn";
  navRight.textContent = "▶";
  navRight.setAttribute("aria-label", "Scroll sheets right");

  nav.append(navLeft, navRight);

  const tabStrip = document.createElement("div");
  tabStrip.className = "sheet-tabs";
  tabStrip.setAttribute("role", "tablist");

  let activeTabEl: HTMLElement | null = null;

  for (const sheet of sheets) {
    const sheetId = sheet.id;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sheet-tab";
    button.dataset.sheetId = sheetId;
    button.dataset.testid = `sheet-tab-${sheetId}`;
    button.dataset.active = sheetId === active ? "true" : "false";
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", sheetId === active ? "true" : "false");
    button.textContent = sheet.name;
    button.addEventListener("click", () => {
      app.activateSheet(sheetId);
      app.focus();
    });
    tabStrip.appendChild(button);
    if (sheetId === active) activeTabEl = button;
  }

  const addSheetBtn = document.createElement("button");
  addSheetBtn.type = "button";
  addSheetBtn.className = "sheet-add";
  addSheetBtn.dataset.testid = "sheet-add";
  addSheetBtn.textContent = "+";
  addSheetBtn.setAttribute("aria-label", "Add sheet");
  addSheetBtn.addEventListener("click", () => {
    const existingIds = new Set(listSheetsForUi().map((sheet) => sheet.id));
    let n = 1;
    while (existingIds.has(`Sheet${n}`)) n += 1;
    const newSheetId = `Sheet${n}`;

    const doc = app.getDocument();
    // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
    doc.getCell(newSheetId, { row: 0, col: 0 });
    app.activateSheet(newSheetId);
    app.focus();
  });

  sheetTabsRootEl.append(nav, tabStrip, addSheetBtn);

  const scrollStep = 120;
  navLeft.addEventListener("click", () => {
    tabStrip.scrollBy({ left: -scrollStep, behavior: "smooth" });
  });
  navRight.addEventListener("click", () => {
    tabStrip.scrollBy({ left: scrollStep, behavior: "smooth" });
  });

  function updateNavDisabledState() {
    const maxScrollLeft = tabStrip.scrollWidth - tabStrip.clientWidth;
    navLeft.disabled = tabStrip.scrollLeft <= 0;
    navRight.disabled = tabStrip.scrollLeft >= maxScrollLeft - 1;
  }
  tabStrip.addEventListener("scroll", updateNavDisabledState, { passive: true });
  updateNavDisabledState();

  // Best-effort: keep the active tab visible after re-rendering.
  activeTabEl?.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function syncSheetUi(): void {
  const sheets = listSheetsForUi();
  renderSheetTabs(sheets);
  renderSheetSwitcher(sheets, app.getCurrentSheetId());
}

syncSheetUi();

const originalActivateSheet = app.activateSheet.bind(app);
app.activateSheet = (sheetId: string): void => {
  originalActivateSheet(sheetId);
  syncSheetUi();
};

const originalActivateCell = app.activateCell.bind(app);
app.activateCell = (target: Parameters<SpreadsheetApp["activateCell"]>[0]): void => {
  const prevSheet = app.getCurrentSheetId();
  originalActivateCell(target);
  if (target.sheetId && target.sheetId !== prevSheet) syncSheetUi();
};

const originalSelectRange = app.selectRange.bind(app);
app.selectRange = (target: Parameters<SpreadsheetApp["selectRange"]>[0]): void => {
  const prevSheet = app.getCurrentSheetId();
  originalSelectRange(target);
  if (target.sheetId && target.sheetId !== prevSheet) syncSheetUi();
};

// Keep the canvas renderer in sync with programmatic document mutations (e.g. AI tools)
// and re-render when edits create new sheets (DocumentController creates sheets lazily).
app.getDocument().on("change", () => {
  app.refresh();
  const sheetIds = app.getDocument().getSheetIds();
  const nextSheetIds = sheetIds.length > 0 ? sheetIds : ["Sheet1"];
  if (nextSheetIds.length !== lastSheetIds.length || nextSheetIds.some((id, idx) => id !== lastSheetIds[idx])) {
    syncSheetUi();
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
const openAiPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-ai-panel"]');
const openAiAuditPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-ai-audit-panel"]');
const openDataQueriesPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-data-queries-panel"]');
const openMacrosPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-macros-panel"]');
const openScriptEditorPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-script-editor-panel"]');
const openPythonPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-python-panel"]');
const openExtensionsPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-extensions-panel"]');
const splitVertical = document.querySelector<HTMLButtonElement>('[data-testid="split-vertical"]');
const splitHorizontal = document.querySelector<HTMLButtonElement>('[data-testid="split-horizontal"]');
const splitNone = document.querySelector<HTMLButtonElement>('[data-testid="split-none"]');
const freezePanes = document.querySelector<HTMLButtonElement>('[data-testid="freeze-panes"]');
const freezeTopRow = document.querySelector<HTMLButtonElement>('[data-testid="freeze-top-row"]');
const freezeFirstColumn = document.querySelector<HTMLButtonElement>('[data-testid="freeze-first-column"]');
const unfreezePanes = document.querySelector<HTMLButtonElement>('[data-testid="unfreeze-panes"]');

if (
  dockLeft &&
  dockRight &&
  dockBottom &&
  floatingRoot &&
  workspaceRoot &&
  gridSplit &&
  gridSecondary &&
  gridSplitter &&
  openAiPanel &&
  openAiAuditPanel &&
  openDataQueriesPanel &&
  openMacrosPanel &&
  openScriptEditorPanel &&
  splitVertical &&
  splitHorizontal &&
  splitNone
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

  function renderSplitView() {
    const split = layoutController.layout.splitView;
    const ratio = typeof split.ratio === "number" ? split.ratio : 0.5;
    const clamped = Math.max(0.1, Math.min(0.9, ratio));
    const primaryPct = Math.round(clamped * 1000) / 10;
    const secondaryPct = Math.round((100 - primaryPct) * 10) / 10;

    if (split.direction === "none") {
      gridSplitEl.style.gridTemplateColumns = "1fr 0px 0px";
      gridSplitEl.style.gridTemplateRows = "1fr";
      gridSecondaryEl.style.display = "none";
      gridSplitterEl.style.display = "none";
      return;
    }

    gridSecondaryEl.style.display = "block";
    gridSplitterEl.style.display = "block";

    if (split.direction === "vertical") {
      gridSplitEl.style.gridTemplateColumns = `${primaryPct}% 4px ${secondaryPct}%`;
      gridSplitEl.style.gridTemplateRows = "1fr";
      gridSplitterEl.style.cursor = "col-resize";
    } else {
      gridSplitEl.style.gridTemplateColumns = "1fr";
      gridSplitEl.style.gridTemplateRows = `${primaryPct}% 4px ${secondaryPct}%`;
      gridSplitterEl.style.cursor = "row-resize";
    }

    const sheetLabel = split.panes.secondary.sheetId ?? "Sheet";
    gridSecondaryEl.textContent = `Secondary view (${sheetLabel})`;
    gridSecondaryEl.style.display = "flex";
    gridSecondaryEl.style.alignItems = "center";
    gridSecondaryEl.style.justifyContent = "center";
    gridSecondaryEl.style.color = "var(--text-secondary)";
    gridSecondaryEl.style.fontSize = "12px";
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

  const contextKeys = new ContextKeyService();

  let lastSelection: SelectionState | null = null;

  const updateContextKeys = (selection: SelectionState | null = lastSelection) => {
    if (!selection) return;
    const sheetId = app.getCurrentSheetId();
    const sheetName = workbookSheetNames.get(sheetId) ?? sheetId;
    const active = selection.active;
    const cell = app.getDocument().getCell(sheetId, { row: active.row, col: active.col }) as any;
    const value = normalizeExtensionCellValue(cell?.value ?? null);
    const formula = typeof cell?.formula === "string" ? cell.formula : null;
    const selectionKeys = deriveSelectionContextKeys(selection);

    contextKeys.batch({
      sheetName,
      ...selectionKeys,
      cellHasValue: (value != null && String(value).trim().length > 0) || (formula != null && formula.trim().length > 0),
    });
  };

  app.subscribeSelection((selection) => {
    lastSelection = selection;
    updateContextKeys(selection);
  });
  app.getDocument().on("change", () => updateContextKeys());

  let extensionPanelBridge: ExtensionPanelBridge | null = null;

  const extensionHostManager = new DesktopExtensionHostManager({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getActiveSheet() {
        const sheetId = app.getCurrentSheetId();
        return { id: sheetId, name: workbookSheetNames.get(sheetId) ?? sheetId };
      },
      listSheets() {
        const ids = app.getDocument().getSheetIds();
        const list = ids.length > 0 ? ids : ["Sheet1"];
        return list.map((id) => ({ id, name: workbookSheetNames.get(id) ?? id }));
      },
      async getSelection() {
        const sheetId = app.getCurrentSheetId();
        const range = app.getSelectionRanges()[0] ?? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
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

  const commandRegistry = new CommandRegistry();
  // Built-in spreadsheet commands. These must be registered in the CommandRegistry so
  // context menus (and eventually the command palette) can execute them uniformly.
  commandRegistry.registerBuiltinCommand("clipboard.cut", t("clipboard.cut"), () => app.clipboardCut());
  commandRegistry.registerBuiltinCommand("clipboard.copy", t("clipboard.copy"), () => app.clipboardCopy());
  commandRegistry.registerBuiltinCommand("clipboard.paste", t("clipboard.paste"), () => app.clipboardPaste());
  commandRegistry.registerBuiltinCommand("clipboard.pasteSpecial", "Paste Special…", (mode?: unknown) =>
    app.clipboardPasteSpecial((mode as any) ?? "all"),
  );
  commandRegistry.registerBuiltinCommand("edit.clearContents", "Clear Contents", () => app.clearContents());

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

  const executeExtensionCommand = (commandId: string) => {
    void (async () => {
      await ensureExtensionsLoaded();
      syncContributedCommands();
      await commandRegistry.executeCommand(commandId);
    })().catch((err) => {
      showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
    });
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

  // Keybindings (foundation): execute contributed commands.
  const parsedKeybindings: Array<ReturnType<typeof parseKeybinding>> = [];
  let lastLoadedExtensionIds = new Set<string>();

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

  const updateKeybindings = () => {
    parsedKeybindings.length = 0;
    if (!extensionHostManager.ready) return;
    const platform = /Mac|iPhone|iPad|iPod/.test(navigator.platform) ? "mac" : "other";
    const contributed = extensionHostManager.getContributedKeybindings() as ContributedKeybinding[];
    for (const kb of contributed) {
      const binding = platformKeybinding(kb, platform);
      const parsed = parseKeybinding(kb.command, binding, kb.when ?? null);
      if (parsed) parsedKeybindings.push(parsed);
    }
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
    syncContributedCommands();
    syncContributedPanels();
    updateKeybindings();
    activateOpenExtensionPanels();
  });
  syncContributedCommands();
  syncContributedPanels();
  updateKeybindings();

  window.addEventListener(
    "keydown",
    (e) => {
      if (!extensionHostManager.ready) return;
      if (e.defaultPrevented) return;
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) {
        return;
      }

      for (const binding of parsedKeybindings) {
        if (!binding) continue;
        if (!matchesKeybinding(binding, e)) continue;
        if (!evaluateWhenClause(binding.when, contextKeys.asLookup())) continue;
        e.preventDefault();
        executeExtensionCommand(binding.command);
        return;
      }
    },
    true,
  );

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

  gridRoot.addEventListener("contextmenu", (e) => {
    // Always prevent the native context menu; we render our own.
    e.preventDefault();

    const picked = app.pickCellAtClientPoint(e.clientX, e.clientY);
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

    const menuItems: ContextMenuItem[] = [
      {
        type: "item",
        label: t("clipboard.cut"),
        shortcut: primaryShortcut("X"),
        onSelect: () => executeBuiltinCommand("clipboard.cut"),
      },
      {
        type: "item",
        label: t("clipboard.copy"),
        shortcut: primaryShortcut("C"),
        onSelect: () => executeBuiltinCommand("clipboard.copy"),
      },
      {
        type: "item",
        label: t("clipboard.paste"),
        shortcut: primaryShortcut("V"),
        onSelect: () => executeBuiltinCommand("clipboard.paste"),
      },
      { type: "separator" },
      {
        type: "item",
        label: "Paste Special…",
        shortcut: isMac ? "⌘⌥V" : "Ctrl+Alt+V",
        onSelect: async () => {
          const options = getPasteSpecialMenuItems().map((item) => ({ label: item.label, value: item.mode }));
          const selected = await showQuickPick(options, { placeHolder: "Paste Special" });
          if (!selected) return;
          executeBuiltinCommand("clipboard.pasteSpecial", selected);
        },
      },
      { type: "separator" },
      {
        type: "item",
        label: "Clear Contents",
        shortcut: isMac ? "⌫" : "Del",
        onSelect: () => executeBuiltinCommand("edit.clearContents"),
      },
    ];

    if (extensionHostManager.ready) {
      // Ensure command labels are available.
      syncContributedCommands();

      const contributed = resolveMenuItems(extensionHostManager.getContributedMenu("cell/context"), contextKeys.asLookup());
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
            onSelect: () => executeExtensionCommand(entry.commandId),
          });
        }
      }
    }

    contextMenu.open({ x: e.clientX, y: e.clientY, items: menuItems });
  });

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
          body.style.display = "flex";
          body.style.flexDirection = "column";
          body.style.gap = "12px";
          body.style.padding = "8px";
          body.style.overflow = "auto";

          const recorderPanel = document.createElement("div");
          recorderPanel.style.display = "flex";
          recorderPanel.style.flexDirection = "column";
          recorderPanel.style.gap = "8px";
          recorderPanel.style.paddingBottom = "12px";
          recorderPanel.style.borderBottom = "1px solid var(--panel-border)";

          const runnerPanel = document.createElement("div");
          runnerPanel.style.flex = "1";
          runnerPanel.style.minHeight = "0";

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
          title.style.fontWeight = "600";
          recorderPanel.appendChild(title);

          const status = document.createElement("div");
          status.style.fontSize = "12px";
          status.style.color = "var(--text-secondary)";
          recorderPanel.appendChild(status);

          const buttons = document.createElement("div");
          buttons.style.display = "flex";
          buttons.style.flexWrap = "wrap";
          buttons.style.gap = "8px";
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
          meta.style.fontSize = "12px";
          meta.style.color = "var(--text-secondary)";
          recorderPanel.appendChild(meta);

          const preview = document.createElement("pre");
          preview.style.whiteSpace = "pre-wrap";
          preview.style.margin = "0";
          preview.style.padding = "8px";
          preview.style.border = "1px solid var(--panel-border)";
          preview.style.borderRadius = "6px";
          preview.style.maxHeight = "240px";
          preview.style.overflow = "auto";
          recorderPanel.appendChild(preview);

          const copyText = async (text: string) => {
            try {
              await navigator.clipboard.writeText(text);
            } catch {
              const textarea = document.createElement("textarea");
              textarea.value = text;
              textarea.style.position = "fixed";
              textarea.style.left = "-9999px";
              document.body.appendChild(textarea);
              textarea.select();
              document.execCommand("copy");
              textarea.remove();
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
        container.style.height = "100%";
        container.style.display = "flex";
        container.style.flexDirection = "column";
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
        container.style.height = "100%";
        container.style.display = "flex";
        container.style.flexDirection = "column";
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

      for (const panelId of zone.panels) {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "dock-panel__tab";
        tab.textContent = panelTitle(panelId);
        tab.setAttribute("role", "tab");
        tab.setAttribute("aria-selected", panelId === active ? "true" : "false");
        tab.addEventListener("click", (e) => {
          e.preventDefault();
          if (panelId === active) return;
          layoutController.activateDockedPanel(panelId, currentSide);
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

    const body = document.createElement("div");
    body.className = "dock-panel__body";
    renderPanelBody(active, body);

    panel.appendChild(header);
    panel.appendChild(body);

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

  splitVertical.addEventListener("click", () => layoutController.setSplitDirection("vertical", 0.5));
  splitHorizontal.addEventListener("click", () => layoutController.setSplitDirection("horizontal", 0.5));
  splitNone.addEventListener("click", () => layoutController.setSplitDirection("none", 0.5));

  openAiPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.AI_CHAT);
    else layoutController.closePanel(PanelIds.AI_CHAT);
  });

  openAiAuditPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.AI_AUDIT);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.AI_AUDIT);
    else layoutController.closePanel(PanelIds.AI_AUDIT);
  });

  openDataQueriesPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.DATA_QUERIES);
    else layoutController.closePanel(PanelIds.DATA_QUERIES);
  });

  openMacrosPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.MACROS);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.MACROS);
    else layoutController.closePanel(PanelIds.MACROS);
  });

  openScriptEditorPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.SCRIPT_EDITOR);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.SCRIPT_EDITOR);
    else layoutController.closePanel(PanelIds.SCRIPT_EDITOR);
  });

  openPythonPanel?.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.PYTHON);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.PYTHON);
    else layoutController.closePanel(PanelIds.PYTHON);
  });

  openExtensionsPanel?.addEventListener("click", () => {
    void ensureExtensionsLoaded().then(() => {
      // Ensure registries are up-to-date once the host finishes loading.
      syncContributedCommands();
      syncContributedPanels();
      updateKeybindings();
    });
    const placement = getPanelPlacement(layoutController.layout, PanelIds.EXTENSIONS);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.EXTENSIONS);
    else layoutController.closePanel(PanelIds.EXTENSIONS);
  });

  openVbaMigratePanel?.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.VBA_MIGRATE);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.VBA_MIGRATE);
    else layoutController.closePanel(PanelIds.VBA_MIGRATE);
  });

  // --- Command palette (minimal) ------------------------------------------------

  const paletteOverlay = document.createElement("div");
  paletteOverlay.className = "command-palette-overlay";
  paletteOverlay.hidden = true;
  paletteOverlay.setAttribute("role", "dialog");
  paletteOverlay.setAttribute("aria-modal", "true");

  const palette = document.createElement("div");
  palette.className = "command-palette";
  palette.dataset.testid = "command-palette";

  const paletteInput = document.createElement("input");
  paletteInput.className = "command-palette__input";
  paletteInput.dataset.testid = "command-palette-input";
  paletteInput.placeholder = t("commandPalette.placeholder");

  const paletteList = document.createElement("ul");
  paletteList.className = "command-palette__list";
  paletteList.dataset.testid = "command-palette-list";
  // The command palette is a modal dialog, so keep at least two tabbable targets
  // (input + list) to support a basic focus trap.
  paletteList.tabIndex = 0;

  palette.appendChild(paletteInput);
  palette.appendChild(paletteList);
  paletteOverlay.appendChild(palette);
  document.body.appendChild(paletteOverlay);

  type PaletteCommand = { id: string; label: string; run: () => void };

  const paletteCommands: PaletteCommand[] = [
    {
      id: "insertPivotTable",
      label: t("commandPalette.command.insertPivotTable"),
      run: () => {
        layoutController.openPanel(PanelIds.PIVOT_BUILDER);
        // If the panel is already open, we still want to refresh its source range from
        // the latest selection.
        window.dispatchEvent(new CustomEvent("pivot-builder:use-selection"));
      },
    },
    {
      id: "tracePrecedents",
      label: "Trace precedents",
      run: () => {
        app.clearAuditing();
        app.toggleAuditingPrecedents();
        app.focus();
      },
    },
    {
      id: "traceDependents",
      label: "Trace dependents",
      run: () => {
        app.clearAuditing();
        app.toggleAuditingDependents();
        app.focus();
      },
    },
    {
      id: "traceBoth",
      label: "Trace precedents + dependents",
      run: () => {
        app.clearAuditing();
        app.toggleAuditingPrecedents();
        app.toggleAuditingDependents();
        app.focus();
      },
    },
    {
      id: "clearAuditing",
      label: "Clear auditing",
      run: () => {
        app.clearAuditing();
        app.focus();
      },
    },
    {
      id: "toggleTransitiveAuditing",
      label: "Toggle transitive auditing",
      run: () => {
        app.toggleAuditingTransitive();
        app.focus();
      },
    },
    {
      id: "freezePanes",
      label: "Freeze Panes",
      run: () => {
        app.freezePanes();
        app.focus();
      },
    },
    {
      id: "freezeTopRow",
      label: "Freeze Top Row",
      run: () => {
        app.freezeTopRow();
        app.focus();
      },
    },
    {
      id: "freezeFirstColumn",
      label: "Freeze First Column",
      run: () => {
        app.freezeFirstColumn();
        app.focus();
      },
    },
    {
      id: "unfreezePanes",
      label: "Unfreeze Panes",
      run: () => {
        app.unfreezePanes();
        app.focus();
      },
    },
  ];

  let paletteQuery = "";
  let paletteSelected = 0;

  function filteredCommands(): PaletteCommand[] {
    const q = paletteQuery.trim().toLowerCase();
    if (!q) return paletteCommands;
    return paletteCommands.filter((cmd) => cmd.label.toLowerCase().includes(q));
  }

  function renderPalette(): void {
    const list = filteredCommands();
    if (paletteSelected >= list.length) paletteSelected = Math.max(0, list.length - 1);
    paletteList.replaceChildren();

    for (let i = 0; i < list.length; i += 1) {
      const cmd = list[i]!;
      const li = document.createElement("li");
      li.className = "command-palette__item";
      li.textContent = cmd.label;
      li.setAttribute("aria-selected", i === paletteSelected ? "true" : "false");
      li.addEventListener("mousedown", (e) => {
        // Prevent focus leaving the input before we run the command.
        e.preventDefault();
      });
      li.addEventListener("click", () => {
        closePalette();
        cmd.run();
      });
      paletteList.appendChild(li);
    }
  }

  function openPalette(): void {
    paletteQuery = "";
    paletteSelected = 0;
    paletteInput.value = "";
    paletteOverlay.hidden = false;
    document.addEventListener("focusin", handlePaletteDocumentFocusIn);
    renderPalette();
    paletteInput.focus();
    paletteInput.select();
  }

  function closePalette(): void {
    paletteOverlay.hidden = true;
    document.removeEventListener("focusin", handlePaletteDocumentFocusIn);
    // Best-effort: return focus to the grid.
    app.focus();
  }

  paletteOverlay.addEventListener("click", (e) => {
    if (e.target === paletteOverlay) closePalette();
  });

  function handlePaletteDocumentFocusIn(e: FocusEvent): void {
    if (paletteOverlay.hidden) return;
    const target = e.target as Node | null;
    if (!target) return;
    if (paletteOverlay.contains(target)) return;
    // If focus escapes the modal dialog, bring it back to the input.
    paletteInput.focus();
  }

  paletteOverlay.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    if (paletteOverlay.hidden) return;

    const focusable = [paletteInput, paletteList].filter((el) => !el.hasAttribute("disabled"));
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
  });

  paletteInput.addEventListener("input", () => {
    paletteQuery = paletteInput.value;
    paletteSelected = 0;
    renderPalette();
  });

  function handlePaletteKeydown(e: KeyboardEvent): void {
    if (e.key === "Escape") {
      e.preventDefault();
      closePalette();
      return;
    }

    const list = filteredCommands();
    if (e.key === "ArrowDown") {
      e.preventDefault();
      paletteSelected = list.length === 0 ? 0 : Math.min(list.length - 1, paletteSelected + 1);
      renderPalette();
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      paletteSelected = list.length === 0 ? 0 : Math.max(0, paletteSelected - 1);
      renderPalette();
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      const cmd = list[paletteSelected];
      if (!cmd) return;
      closePalette();
      cmd.run();
    }
  }

  paletteInput.addEventListener("keydown", handlePaletteKeydown);
  paletteList.addEventListener("keydown", handlePaletteKeydown);

  openCommandPalette = openPalette;

  window.addEventListener("keydown", (e) => {
    if (e.defaultPrevented) return;
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || !e.shiftKey) return;
    if (e.key !== "P" && e.key !== "p") return;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return;
    }

    e.preventDefault();
    openPalette();
  });

  layoutController.on("change", () => renderLayout());
  rerenderLayout = renderLayout;
  renderLayout();

  // Allow panel content to request opening another panel.
  window.addEventListener("formula:open-panel", (evt) => {
    const panelId = (evt as any)?.detail?.panelId;
    if (typeof panelId !== "string") return;
    const placement = getPanelPlacement(layoutController.layout, panelId);
    if (placement.kind === "closed") layoutController.openPanel(panelId);
    else if (placement.kind === "docked") layoutController.activateDockedPanel(panelId, placement.side);
  });
}

freezePanes?.addEventListener("click", () => {
  app.freezePanes();
  app.focus();
});
freezeTopRow?.addEventListener("click", () => {
  app.freezeTopRow();
  app.focus();
});
freezeFirstColumn?.addEventListener("click", () => {
  app.freezeFirstColumn();
  app.focus();
});
unfreezePanes?.addEventListener("click", () => {
  app.unfreezePanes();
  app.focus();
});

const workbook = app.getSearchWorkbook();

const findReplaceController = new FindReplaceController({
  workbook,
  getCurrentSheetName: () => app.getCurrentSheetId(),
  getActiveCell: () => {
    const cell = app.getActiveCell();
    return { sheetName: app.getCurrentSheetId(), row: cell.row, col: cell.col };
  },
  setActiveCell: ({ sheetName, row, col }) => app.activateCell({ sheetId: sheetName, row, col }),
  getSelectionRanges: () => app.getSelectionRanges(),
  beginBatch: (opts) => app.getDocument().beginBatch(opts),
  endBatch: () => app.getDocument().endBatch()
});

const { findDialog, replaceDialog, goToDialog } = registerFindReplaceShortcuts({
  controller: findReplaceController,
  workbook,
  getCurrentSheetName: () => app.getCurrentSheetId(),
  setActiveCell: ({ sheetName, row, col }) => app.activateCell({ sheetId: sheetName, row, col }),
  selectRange: ({ sheetName, range }) => app.selectRange({ sheetId: sheetName, range })
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

mountRibbon(ribbonRoot, {
  onToggle: (commandId, pressed) => {
    switch (commandId) {
      case "home.font.bold":
        applyToSelection("Bold", (sheetId, ranges) => toggleBold(app.getDocument(), sheetId, ranges, { next: pressed }));
        return;
      case "home.font.italic":
        applyToSelection("Italic", (sheetId, ranges) =>
          toggleItalic(app.getDocument(), sheetId, ranges, { next: pressed }),
        );
        return;
      case "home.font.underline":
        applyToSelection("Underline", (sheetId, ranges) =>
          toggleUnderline(app.getDocument(), sheetId, ranges, { next: pressed }),
        );
        return;
      case "home.alignment.wrapText":
        applyToSelection("Wrap", (sheetId, ranges) => toggleWrap(app.getDocument(), sheetId, ranges, { next: pressed }));
        return;
      default:
        return;
    }
  },
  onCommand: (commandId) => {
    switch (commandId) {
      case "home.font.borders":
        applyToSelection("Borders", (sheetId, ranges) => applyAllBorders(app.getDocument(), sheetId, ranges));
        return;
      case "home.font.fontColor":
        openColorPicker(fontColorPicker, "Font color", (sheetId, ranges, argb) =>
          setFontColor(app.getDocument(), sheetId, ranges, argb),
        );
        return;
      case "home.font.fillColor":
        openColorPicker(fillColorPicker, "Fill color", (sheetId, ranges, argb) =>
          setFillColor(app.getDocument(), sheetId, ranges, argb),
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
          applyToSelection("Font size", (sheetId, ranges) => setFontSize(app.getDocument(), sheetId, ranges, picked));
        })();
        return;

      case "home.alignment.alignLeft":
        applyToSelection("Align left", (sheetId, ranges) => setHorizontalAlign(app.getDocument(), sheetId, ranges, "left"));
        return;
      case "home.alignment.center":
        applyToSelection("Align center", (sheetId, ranges) =>
          setHorizontalAlign(app.getDocument(), sheetId, ranges, "center"),
        );
        return;
      case "home.alignment.alignRight":
        applyToSelection("Align right", (sheetId, ranges) =>
          setHorizontalAlign(app.getDocument(), sheetId, ranges, "right"),
        );
        return;

      case "home.number.percent":
        applyToSelection("Number format", (sheetId, ranges) =>
          applyNumberFormatPreset(app.getDocument(), sheetId, ranges, "percent"),
        );
        return;
      case "home.number.accounting":
        applyToSelection("Number format", (sheetId, ranges) =>
          applyNumberFormatPreset(app.getDocument(), sheetId, ranges, "currency"),
        );
        return;
      case "home.number.date":
        applyToSelection("Number format", (sheetId, ranges) => applyNumberFormatPreset(app.getDocument(), sheetId, ranges, "date"));
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
      default:
        return;
    }
  },
});

installUnsavedChangesPrompt(window, app.getDocument());

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
  app.focus();
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

function encodeDocumentSnapshot(snapshot: unknown): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(snapshot));
}

function normalizeSheetList(info: WorkbookInfo): { id: string; name: string }[] {
  const sheets = Array.isArray(info.sheets) ? info.sheets : [];
  return sheets
    .map((s) => ({ id: String((s as any).id ?? ""), name: String((s as any).name ?? (s as any).id ?? "") }))
    .filter((s) => s.id.trim() !== "");
}

async function confirmDiscardDirtyState(actionLabel: string): Promise<boolean> {
  const doc = app.getDocument();
  if (!doc.isDirty) return true;
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

  workbookSheetNames.clear();
  for (const sheet of sheets) {
    workbookSheetNames.set(sheet.id, sheet.name);
  }

  const MAX_COLS = 200;
  const CHUNK_ROWS = 200;
  const MAX_ROWS = 10_000;

  const snapshotSheets: Array<{ id: string; cells: any[] }> = [];

  for (const sheet of sheets) {
    const cells: Array<{ row: number; col: number; value: unknown | null; formula: string | null; format: null }> = [];

    const usedRange = await tauriBackend.getSheetUsedRange(sheet.id);
    if (!usedRange) {
      snapshotSheets.push({ id: sheet.id, cells });
      continue;
    }

    const startCol = Math.max(0, Math.min(usedRange.start_col, MAX_COLS - 1));
    const endCol = Math.max(0, Math.min(usedRange.end_col, MAX_COLS - 1));
    const startRow = Math.max(0, Math.min(usedRange.start_row, MAX_ROWS - 1));
    const endRow = Math.max(0, Math.min(usedRange.end_row, MAX_ROWS - 1));

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
  const sheetIdByName = new Map<string, string>();
  for (const [id, name] of workbookSheetNames.entries()) {
    sheetIdByName.set(name, id);
  }

  const [definedNames, tables] = await Promise.all([
    tauriBackend.listDefinedNames().catch(() => []),
    tauriBackend.listTables().catch(() => []),
  ]);

  const normalizedTables = tables.map((table) => {
    const rawSheetId = typeof (table as any)?.sheet_id === "string" ? String((table as any).sheet_id) : "";
    const sheet_id = rawSheetId ? sheetIdByName.get(rawSheetId) ?? rawSheetId : rawSheetId;
    return { ...(table as any), sheet_id };
  });
  refreshTableSignaturesFromBackend(doc, normalizedTables as any, { workbookSignature });
  const normalizedDefinedNames = definedNames.map((entry) => {
    const refers_to = typeof (entry as any)?.refers_to === "string" ? String((entry as any).refers_to) : "";
    const { sheetName: explicitSheetName } = splitSheetQualifier(refers_to);
    const sheetIdFromRef = explicitSheetName ? sheetIdByName.get(explicitSheetName) ?? explicitSheetName : null;
    const rawScopeSheet = typeof (entry as any)?.sheet_id === "string" ? String((entry as any).sheet_id) : null;
    const sheetIdFromScope = rawScopeSheet ? sheetIdByName.get(rawScopeSheet) ?? rawScopeSheet : null;
    return { ...(entry as any), sheet_id: sheetIdFromScope ?? sheetIdFromRef };
  });
  refreshDefinedNameSignaturesFromBackend(doc, normalizedDefinedNames as any, { workbookSignature });

  for (const entry of definedNames) {
    const name = typeof (entry as any)?.name === "string" ? String((entry as any).name) : "";
    const refersTo =
      typeof (entry as any)?.refers_to === "string" ? String((entry as any).refers_to) : "";
    if (!name || !refersTo) continue;

    const { sheetName: explicitSheetName, ref } = splitSheetQualifier(refersTo);
    const sheetIdFromRef = explicitSheetName ? sheetIdByName.get(explicitSheetName) ?? explicitSheetName : null;
    const rawScopeSheet = typeof (entry as any)?.sheet_id === "string" ? String((entry as any).sheet_id) : null;
    const sheetIdFromScope = rawScopeSheet ? sheetIdByName.get(rawScopeSheet) ?? rawScopeSheet : null;
    const sheetName = sheetIdFromRef ?? sheetIdFromScope;
    if (!sheetName) continue;

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
    const rawSheetName = typeof (table as any)?.sheet_id === "string" ? String((table as any).sheet_id) : "";
    const sheetName = rawSheetName ? sheetIdByName.get(rawSheetName) ?? rawSheetName : "";
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

async function openWorkbookFromPath(path: string): Promise<void> {
  if (!tauriBackend) return;
  if (typeof path !== "string" || path.trim() === "") return;
  const ok = await confirmDiscardDirtyState("open another workbook");
  if (!ok) return;

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
    activePanelWorkbookId = activeWorkbook.path ?? activeWorkbook.origin_path ?? path;
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
    filters: [
      { name: "Spreadsheets", extensions: ["xlsx", "xlsm", "xls", "xlsb", "csv"] },
      { name: "Excel", extensions: ["xlsx", "xlsm", "xls", "xlsb"] },
      { name: "CSV", extensions: ["csv"] },
    ],
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

async function handleSave(): Promise<void> {
  if (!tauriBackend) return;
  if (!activeWorkbook) return;
  if (!workbookSync) return;

  if (!activeWorkbook.path) {
    await handleSaveAs();
    return;
  }

  await workbookSync.markSaved();
}

async function handleSaveAs(): Promise<void> {
  if (!tauriBackend) return;
  if (!activeWorkbook) return;

  const previousPanelWorkbookId = activePanelWorkbookId;
  const { save } = getTauriDialog();
  const path = await save({
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (!path) return;

  // Ensure any pending microtask-batched workbook edits are flushed before saving.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await drainBackendSync();
  if (queuedInvoke) {
    await queuedInvoke("save_workbook", { path });
  } else {
    await tauriBackend.saveWorkbook(path);
  }
  activeWorkbook = { ...activeWorkbook, path };
  app.getDocument().markSaved();

  await copyPowerQueryPersistence(previousPanelWorkbookId, path);
  activePanelWorkbookId = path;
  startPowerQueryService();
  rerenderLayout?.();
}

async function handleNewWorkbook(): Promise<void> {
  if (!tauriBackend) return;
  const ok = await confirmDiscardDirtyState("create a new workbook");
  if (!ok) return;

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
    activePanelWorkbookId = nextPanelWorkbookId;
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

  installUpdaterUi(listen);

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
    void handleCloseRequest({ quit: true }).catch((err) => {
      console.error("Failed to quit app:", err);
    });
  });

  void listen("shortcut-quick-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      void nativeDialogs.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("shortcut-command-palette", () => {
    openCommandPalette?.();
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
      if (quit && queuedInvoke) {
        try {
          // Best-effort Workbook_BeforeClose when quitting via the tray menu. (Window close
          // requests fire this event from the Rust host already.)
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
      if (doc.isDirty) {
        const discard = await nativeDialogs.confirm("You have unsaved changes. Discard them?");
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

  void listen("close-requested", async (event) => {
    const payload = (event as any)?.payload;
    const token = typeof payload?.token === "string" ? String(payload.token) : undefined;
    await handleCloseRequest({ quit: false, beforeCloseUpdates: payload?.updates, closeToken: token });
  });

  window.addEventListener("keydown", (e) => {
    const isSaveCombo = (e.ctrlKey || e.metaKey) && (e.key === "s" || e.key === "S");
    const isSaveAsCombo = (e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === "S" || e.key === "s");
    if (isSaveAsCombo) {
      e.preventDefault();
      void handleSaveAs().catch((err) => {
        console.error("Failed to save workbook:", err);
        void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
      });
      return;
    }
    if (!isSaveCombo) return;
    e.preventDefault();
    void handleSave().catch((err) => {
      console.error("Failed to save workbook:", err);
      void nativeDialogs.alert(`Failed to save workbook: ${String(err)}`);
    });
  });
} catch {
  // Not running under Tauri; desktop host integration is unavailable.
}

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
