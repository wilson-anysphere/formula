import { SpreadsheetApp } from "./app/spreadsheetApp";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/workspace.css";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import { getPanelTitle, panelRegistry, PanelIds } from "./panels/panelRegistry.js";
import { createPanelBodyRenderer } from "./panels/panelBodyRenderer.js";
import { MacroRecorder, generatePythonMacro, generateTypeScriptMacro } from "./macro-recorder/index.js";
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
import { matchesKeybinding, parseKeybinding, platformKeybinding, type ContributedKeybinding } from "./extensions/keybindings.js";
import { evaluateWhenClause } from "./extensions/whenClause.js";
import { CommandRegistry } from "./extensions/commandRegistry.js";

import sampleHelloManifest from "../../../extensions/sample-hello/package.json";

const workbookSheetNames = new Map<string, string>();

// Seed contributed panels early so layout persistence doesn't drop their ids before the
// extension host finishes loading installed extensions.
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

const formulaBarRoot = document.getElementById("formula-bar");
if (!formulaBarRoot) {
  throw new Error("Missing #formula-bar container");
}

const activeCell = document.querySelector<HTMLElement>('[data-testid="active-cell"]');
const selectionRange = document.querySelector<HTMLElement>('[data-testid="selection-range"]');
const activeValue = document.querySelector<HTMLElement>('[data-testid="active-value"]');
const sheetSwitcher = document.querySelector<HTMLSelectElement>('[data-testid="sheet-switcher"]');
const openComments = document.querySelector<HTMLButtonElement>('[data-testid="open-comments-panel"]');
const auditPrecedents = document.querySelector<HTMLButtonElement>('[data-testid="audit-precedents"]');
const auditDependents = document.querySelector<HTMLButtonElement>('[data-testid="audit-dependents"]');
const auditTransitive = document.querySelector<HTMLButtonElement>('[data-testid="audit-transitive"]');
const openVbaMigratePanel = document.querySelector<HTMLButtonElement>('[data-testid="open-vba-migrate-panel"]');
if (!activeCell || !selectionRange || !activeValue || !sheetSwitcher) {
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
const app = new SpreadsheetApp(gridRoot, { activeCell, selectionRange, activeValue }, { formulaBar: formulaBarRoot, workbookId });
// Panels persist state keyed by a workbook/document identifier. For file-backed workbooks we use
// their on-disk path; for unsaved sessions we generate a random session id so distinct new
// workbooks don't collide.
let activePanelWorkbookId = workbookId;
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();
app.focus();
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
}

function stopPowerQueryService(): void {
  const existingWorkbookId = powerQueryServiceWorkbookId;
  powerQueryServiceWorkbookId = null;
  if (existingWorkbookId) setDesktopPowerQueryService(existingWorkbookId, null);
  powerQueryService?.dispose();
  powerQueryService = null;
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

  for (const sheet of sheets) {
    const sheetId = sheet.id;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sheet-tab";
    button.dataset.sheetId = sheetId;
    button.dataset.testid = `sheet-tab-${sheetId}`;
    button.dataset.active = sheetId === active ? "true" : "false";
    button.textContent = sheet.name;
    button.addEventListener("click", () => {
      app.activateSheet(sheetId);
      app.focus();
    });
    sheetTabsRootEl.appendChild(button);
  }
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

  const updateContextKeys = () => {
    const sheetId = app.getCurrentSheetId();
    const sheetName = workbookSheetNames.get(sheetId) ?? sheetId;
    const active = app.getActiveCell();
    const cell = app.getDocument().getCell(sheetId, { row: active.row, col: active.col }) as any;
    const value = normalizeExtensionCellValue(cell?.value ?? null);
    const formula = typeof cell?.formula === "string" ? cell.formula : null;
    const hasSelection = (() => {
      const range = app.getSelectionRanges()[0];
      if (!range) return false;
      return range.startRow !== range.endRow || range.startCol !== range.endCol;
    })();

    contextKeys.batch({
      sheetName,
      hasSelection,
      cellHasValue: (value != null && String(value).trim().length > 0) || (formula != null && formula.trim().length > 0),
    });
  };

  app.subscribeSelection(() => updateContextKeys());
  app.getDocument().on("change", () => updateContextKeys());
  updateContextKeys();

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

    for (const panel of contributed) {
      panelRegistry.registerPanel(
        panel.id,
        {
          title: panel.title,
          icon: (panel as any).icon ?? null,
          defaultDock: "right",
          defaultFloatingRect: { x: 140, y: 140, width: 520, height: 640 },
          source: { kind: "extension", extensionId: panel.extensionId, contributed: true },
        },
        { owner: panel.extensionId, overwrite: true },
      );
    }

    for (const id of panelRegistry.listPanelIds()) {
      const def = panelRegistry.get(id) as any;
      const source = def?.source;
      if (source?.kind !== "extension" || source.contributed !== true) continue;
      if (contributedIds.has(id)) continue;
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

  // Context menus (foundation): show cell/context contributions.
  const contextMenu = document.createElement("div");
  contextMenu.dataset.testid = "context-menu";
  contextMenu.style.position = "fixed";
  contextMenu.style.display = "none";
  contextMenu.style.flexDirection = "column";
  contextMenu.style.minWidth = "220px";
  contextMenu.style.padding = "6px";
  contextMenu.style.borderRadius = "10px";
  contextMenu.style.border = "1px solid var(--border)";
  contextMenu.style.background = "var(--dialog-bg)";
  contextMenu.style.boxShadow = "var(--dialog-shadow)";
  contextMenu.style.zIndex = "200";
  document.body.appendChild(contextMenu);

  const hideContextMenu = () => {
    contextMenu.style.display = "none";
    contextMenu.replaceChildren();
  };

  gridRoot.addEventListener("contextmenu", (e) => {
    if (!extensionHostManager.ready) return;
    const items = resolveMenuItems(extensionHostManager.getContributedMenu("cell/context"), contextKeys.asLookup());
    if (items.length === 0) return;

    e.preventDefault();
    contextMenu.replaceChildren();

    for (const item of items) {
      const btn = document.createElement("button");
      btn.type = "button";
      const command = commandRegistry.getCommand(item.command);
      btn.textContent = command ? (command.category ? `${command.category}: ${command.title}` : command.title) : item.command;
      btn.disabled = !item.enabled;
      btn.style.display = "block";
      btn.style.width = "100%";
      btn.style.textAlign = "left";
      btn.style.padding = "8px 10px";
      btn.style.borderRadius = "8px";
      btn.style.border = "1px solid transparent";
      btn.style.background = "transparent";
      btn.style.color = item.enabled ? "var(--text-primary)" : "var(--text-secondary)";
      btn.style.cursor = item.enabled ? "pointer" : "default";
      btn.addEventListener("click", () => {
        if (!item.enabled) return;
        hideContextMenu();
        executeExtensionCommand(item.command);
      });
      contextMenu.appendChild(btn);
    }

    contextMenu.style.left = `${e.clientX}px`;
    contextMenu.style.top = `${e.clientY}px`;
    contextMenu.style.display = "flex";
  });

  window.addEventListener("click", () => hideContextMenu(), true);
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") hideContextMenu();
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

    const title = document.createElement("div");
    title.className = "dock-panel__title";
    title.textContent = panelTitle(active);

    const controls = document.createElement("div");
    controls.className = "dock-panel__controls";

    function button(label: string, testId: string, onClick: () => void) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
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
        button("Dock left", active === PanelIds.AI_CHAT ? "dock-ai-panel-left" : "dock-panel-left", () => {
          layoutController.dockPanel(active, "left");
        }),
      );
    }

    if (currentSide !== "right") {
      controls.appendChild(
        button("Dock right", "dock-panel-right", () => {
          layoutController.dockPanel(active, "right");
        }),
      );
    }

    if (currentSide !== "bottom") {
      controls.appendChild(
        button("Dock bottom", "dock-panel-bottom", () => {
          layoutController.dockPanel(active, "bottom");
        }),
      );
    }

    controls.appendChild(
      button("Float", "float-panel", () => {
        const rect = (panelRegistry.get(active) as any)?.defaultFloatingRect ?? { x: 80, y: 80, width: 420, height: 560 };
        layoutController.floatPanel(active, rect);
      }),
    );

    controls.appendChild(
      button("Close", active === PanelIds.AI_CHAT ? "close-ai-panel" : "close-panel", () => {
        layoutController.closePanel(active);
      }),
    );

    header.appendChild(title);
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
      dockLeftBtn.textContent = "Dock left";
      dockLeftBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "left"));

      const dockRightBtn = document.createElement("button");
      dockRightBtn.type = "button";
      dockRightBtn.textContent = "Dock right";
      dockRightBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "right"));

      const dockBottomBtn = document.createElement("button");
      dockBottomBtn.type = "button";
      dockBottomBtn.textContent = "Dock bottom";
      dockBottomBtn.addEventListener("click", () => layoutController.dockPanel(panelId, "bottom"));

      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.textContent = "Close";
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
  paletteOverlay.style.position = "fixed";
  paletteOverlay.style.inset = "0";
  paletteOverlay.style.display = "none";
  paletteOverlay.style.alignItems = "flex-start";
  paletteOverlay.style.justifyContent = "center";
  paletteOverlay.style.paddingTop = "80px";
  paletteOverlay.style.background = "var(--dialog-backdrop)";
  paletteOverlay.style.zIndex = "1000";

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
    paletteOverlay.style.display = "flex";
    renderPalette();
    paletteInput.focus();
    paletteInput.select();
  }

  function closePalette(): void {
    paletteOverlay.style.display = "none";
    // Best-effort: return focus to the grid.
    app.focus();
  }

  paletteOverlay.addEventListener("click", (e) => {
    if (e.target === paletteOverlay) closePalette();
  });

  paletteInput.addEventListener("input", () => {
    paletteQuery = paletteInput.value;
    paletteSelected = 0;
    renderPalette();
  });

  paletteInput.addEventListener("keydown", (e) => {
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
  });

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

type CommandPaletteCommand = { id: string; title: string; run: () => void; keywords?: string[] };

function createCommandPalette(commands: readonly CommandPaletteCommand[]): { open: () => void } {
  const dialog = document.createElement("dialog");
  dialog.className = "command-palette";
  dialog.dataset.testid = "command-palette";
  dialog.addEventListener("click", (e) => {
    if (e.target === dialog) dialog.close();
  });

  const input = document.createElement("input");
  input.type = "text";
  input.className = "command-palette__input";
  input.placeholder = "Type a command…";

  const list = document.createElement("ul");
  list.className = "command-palette__list";

  dialog.appendChild(input);
  dialog.appendChild(list);
  document.body.appendChild(dialog);

  let filtered: CommandPaletteCommand[] = [];
  let selectedIndex = 0;

  const updateSelection = (next: number) => {
    selectedIndex = Math.max(0, Math.min(filtered.length - 1, next));
    const children = Array.from(list.children) as HTMLElement[];
    for (let i = 0; i < children.length; i += 1) {
      children[i]?.setAttribute("aria-selected", i === selectedIndex ? "true" : "false");
    }
  };

  const runSelected = () => {
    const cmd = filtered[selectedIndex];
    if (!cmd) return;
    dialog.close();
    cmd.run();
  };

  const render = () => {
    const query = input.value.trim().toLowerCase();
    filtered = commands.filter((cmd) => {
      if (query === "") return true;
      const haystack = `${cmd.title} ${cmd.id} ${(cmd.keywords ?? []).join(" ")}`.toLowerCase();
      return haystack.includes(query);
    });

    list.replaceChildren();
    selectedIndex = 0;

    for (const [idx, cmd] of filtered.entries()) {
      const item = document.createElement("li");
      item.className = "command-palette__item";
      item.textContent = cmd.title;
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", idx === selectedIndex ? "true" : "false");
      item.addEventListener("pointermove", () => updateSelection(idx));
      item.addEventListener("click", () => {
        dialog.close();
        cmd.run();
      });
      list.appendChild(item);
    }
  };

  input.addEventListener("input", render);
  input.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      updateSelection(selectedIndex + 1);
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      updateSelection(selectedIndex - 1);
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      runSelected();
    }
  });

  dialog.addEventListener("close", () => {
    input.value = "";
    list.replaceChildren();
    filtered = [];
    selectedIndex = 0;
  });

  return {
    open: () => {
      if (dialog.open) {
        input.focus();
        return;
      }
      if (typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute("open", "");
      render();
      input.focus();
    },
  };
}

const commandPalette = createCommandPalette([
  {
    id: "view.freezePanes",
    title: "Freeze Panes",
    run: () => {
      app.freezePanes();
      app.focus();
    },
    keywords: ["frozen", "pane"]
  },
  {
    id: "view.freezeTopRow",
    title: "Freeze Top Row",
    run: () => {
      app.freezeTopRow();
      app.focus();
    },
    keywords: ["frozen", "row"]
  },
  {
    id: "view.freezeFirstColumn",
    title: "Freeze First Column",
    run: () => {
      app.freezeFirstColumn();
      app.focus();
    },
    keywords: ["frozen", "column"]
  },
  {
    id: "view.unfreezePanes",
    title: "Unfreeze Panes",
    run: () => {
      app.unfreezePanes();
      app.focus();
    },
    keywords: ["frozen", "pane"]
  }
]);

const openCommandPalette = () => commandPalette.open();

window.addEventListener("keydown", (e) => {
  const primary = e.ctrlKey || e.metaKey;
  if (!primary || !e.shiftKey) return;
  if (e.key !== "P" && e.key !== "p") return;
  e.preventDefault();
  openCommandPalette();
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

registerFindReplaceShortcuts({
  controller: findReplaceController,
  workbook,
  getCurrentSheetName: () => app.getCurrentSheetId(),
  setActiveCell: ({ sheetName, row, col }) => app.activateCell({ sheetId: sheetName, row, col }),
  selectRange: ({ sheetName, range }) => app.selectRange({ sheetId: sheetName, range })
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
  return window.confirm(`You have unsaved changes. Discard them and ${actionLabel}?`);
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

  // Query definitions are currently persisted in LocalStorage for the desktop shell.
  // Copy them when the workbook id changes (e.g. Save As from an unsaved session).
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
  window.addEventListener("unload", () => {
    vbaEventMacros?.dispose();
    workbookSync?.stop();
  });

  const listen = getTauriListen();
  const emit = getTauriEmit();

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

  void listen("file-dropped", async (event) => {
    const paths = (event as any)?.payload;
    const first = Array.isArray(paths) ? paths[0] : null;
    if (typeof first !== "string" || first.trim() === "") return;
    try {
      await openWorkbookFromPath(first);
    } catch (err) {
      console.error("Failed to open workbook:", err);
      window.alert(`Failed to open workbook: ${String(err)}`);
    }
  });

  void listen("tray-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      window.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("tray-new", () => {
    void handleNewWorkbook().catch((err) => {
      console.error("Failed to create workbook:", err);
      window.alert(`Failed to create workbook: ${String(err)}`);
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
      window.alert(`Failed to open workbook: ${String(err)}`);
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
        const discard = window.confirm("You have unsaved changes. Discard them?");
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
        window.alert(`Failed to save workbook: ${String(err)}`);
      });
      return;
    }
    if (!isSaveCombo) return;
    e.preventDefault();
    void handleSave().catch((err) => {
      console.error("Failed to save workbook:", err);
      window.alert(`Failed to save workbook: ${String(err)}`);
    });
  });
} catch {
  // Not running under Tauri; desktop host integration is unavailable.
}

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
