import { SpreadsheetApp } from "./app/spreadsheetApp";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/workspace.css";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import { getPanelTitle, PANEL_REGISTRY, PanelIds } from "./panels/panelRegistry.js";
import { createPanelBodyRenderer } from "./panels/panelBodyRenderer.js";
import { renderMacroRunner, TauriMacroBackend } from "./macros";
import { mountScriptEditorPanel } from "./panels/script-editor/index.js";
import { installUnsavedChangesPrompt } from "./document/index.js";
import { DocumentWorkbookAdapter } from "./search/documentWorkbookAdapter.js";
import { DocumentControllerWorkbookAdapter } from "./scripting/documentControllerWorkbookAdapter.js";
import { registerFindReplaceShortcuts, FindReplaceController } from "./panels/find-replace/index.js";
import { formatRangeAddress, parseRangeAddress } from "@formula/scripting";
import { startWorkbookSync } from "./tauri/workbookSync";
import { TauriWorkbookBackend, type WorkbookInfo } from "./tauri/workbookBackend";

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
const openComments = document.querySelector<HTMLButtonElement>('[data-testid="open-comments-panel"]');
const openVbaMigratePanel = document.querySelector<HTMLButtonElement>('[data-testid="open-vba-migrate-panel"]');
if (!activeCell || !selectionRange || !activeValue) {
  throw new Error("Missing status bar elements");
}
if (!openComments) {
  throw new Error("Missing comments panel toggle button");
}

const app = new SpreadsheetApp(gridRoot, { activeCell, selectionRange, activeValue }, { formulaBar: formulaBarRoot });
// Treat the seeded demo workbook as an initial "saved" baseline so web reloads
// and Playwright tests aren't blocked by unsaved-changes prompts.
app.getDocument().markSaved();
app.focus();
openComments.addEventListener("click", () => app.toggleCommentsPanel());

// Keep the canvas renderer in sync with programmatic document mutations (e.g. AI tools).
app.getDocument().on("change", () => app.refresh());

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
const openMacrosPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-macros-panel"]');
const openScriptEditorPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-script-editor-panel"]');
const splitVertical = document.querySelector<HTMLButtonElement>('[data-testid="split-vertical"]');
const splitHorizontal = document.querySelector<HTMLButtonElement>('[data-testid="split-horizontal"]');
const splitNone = document.querySelector<HTMLButtonElement>('[data-testid="split-none"]');

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
  openMacrosPanel &&
  openScriptEditorPanel &&
  splitVertical &&
  splitHorizontal &&
  splitNone
) {
  const workbookId = "local-workbook";
  const workspaceManager = new LayoutWorkspaceManager({ storage: localStorage, panelRegistry: PANEL_REGISTRY });
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

  function zoneVisible(zone: { panels: string[]; collapsed: boolean }) {
    return zone.panels.length > 0 && !zone.collapsed;
  }

  function applyDockSizes() {
    const layout = layoutController.layout;

    const leftSize = zoneVisible(layout.docks.left) ? layout.docks.left.size : 0;
    const rightSize = zoneVisible(layout.docks.right) ? layout.docks.right.size : 0;
    const bottomSize = zoneVisible(layout.docks.bottom) ? layout.docks.bottom.size : 0;

    workspaceRoot.style.setProperty("--dock-left-size", `${leftSize}px`);
    workspaceRoot.style.setProperty("--dock-right-size", `${rightSize}px`);
    workspaceRoot.style.setProperty("--dock-bottom-size", `${bottomSize}px`);

    dockLeft.dataset.hidden = zoneVisible(layout.docks.left) ? "false" : "true";
    dockRight.dataset.hidden = zoneVisible(layout.docks.right) ? "false" : "true";
    dockBottom.dataset.hidden = zoneVisible(layout.docks.bottom) ? "false" : "true";
  }

  function renderSplitView() {
    const split = layoutController.layout.splitView;
    const ratio = typeof split.ratio === "number" ? split.ratio : 0.5;
    const clamped = Math.max(0.1, Math.min(0.9, ratio));
    const primaryPct = Math.round(clamped * 1000) / 10;
    const secondaryPct = Math.round((100 - primaryPct) * 10) / 10;

    if (split.direction === "none") {
      gridSplit.style.gridTemplateColumns = "1fr 0px 0px";
      gridSplit.style.gridTemplateRows = "1fr";
      gridSecondary.style.display = "none";
      gridSplitter.style.display = "none";
      return;
    }

    gridSecondary.style.display = "block";
    gridSplitter.style.display = "block";

    if (split.direction === "vertical") {
      gridSplit.style.gridTemplateColumns = `${primaryPct}% 4px ${secondaryPct}%`;
      gridSplit.style.gridTemplateRows = "1fr";
      gridSplitter.style.cursor = "col-resize";
    } else {
      gridSplit.style.gridTemplateColumns = "1fr";
      gridSplit.style.gridTemplateRows = `${primaryPct}% 4px ${secondaryPct}%`;
      gridSplitter.style.cursor = "row-resize";
    }

    const sheetLabel = split.panes.secondary.sheetId ?? "Sheet";
    gridSecondary.textContent = `Secondary view (${sheetLabel})`;
    gridSecondary.style.display = "flex";
    gridSecondary.style.alignItems = "center";
    gridSecondary.style.justifyContent = "center";
    gridSecondary.style.color = "var(--text-secondary)";
    gridSecondary.style.fontSize = "12px";
  }

  function panelTitle(panelId: string) {
    return getPanelTitle(panelId);
  }

  const panelBodyRenderer = createPanelBodyRenderer({
    getDocumentController: () => app.getDocument(),
    getActiveSheetId: () => app.getCurrentSheetId(),
    workbookId,
    createChart: (spec) => app.addChart(spec),
    renderMacrosPanel: (body) => {
      body.textContent = "Loading macros…";
      queueMicrotask(() => {
        try {
          const backend = new TauriMacroBackend();
          void renderMacroRunner(body, backend, workbookId, {
            onApplyUpdates: async (updates) => {
              const document = app.getDocument();
              document.beginBatch({ label: "Run macro" });
              try {
                for (const update of updates) {
                  const sheetId = update.sheetId;
                  const cell = { row: update.row, col: update.col };
                  if (update.formula != null) {
                    document.setCellFormula(sheetId, cell, update.formula);
                  } else {
                    document.setCellValue(sheetId, cell, update.value);
                  }
                }
                document.endBatch();
              } catch (err) {
                document.cancelBatch();
                throw err;
              }
              app.repaint();
            },
          }).catch((err) => {
            body.textContent = `Failed to load macros: ${String(err)}`;
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
        const rect = (PANEL_REGISTRY as any)?.[active]?.defaultFloatingRect ?? { x: 80, y: 80, width: 420, height: 560 };
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
    floatingRoot.replaceChildren();
    const layout = layoutController.layout;

    for (const [panelId, rect] of Object.entries(layout.floating)) {
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
      floatingRoot.appendChild(panel);
    }
  }

  function renderLayout() {
    applyDockSizes();
    renderSplitView();
    renderDock(dockLeft, layoutController.layout.docks.left, "left");
    renderDock(dockRight, layoutController.layout.docks.right, "right");
    renderDock(dockBottom, layoutController.layout.docks.bottom, "bottom");
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

  openVbaMigratePanel?.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.VBA_MIGRATE);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.VBA_MIGRATE);
    else layoutController.closePanel(PanelIds.VBA_MIGRATE);
  });

  layoutController.on("change", () => renderLayout());
  renderLayout();
}

const workbook = new DocumentWorkbookAdapter({ document: app.getDocument() });

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

const sheetSwitcher = document.querySelector<HTMLSelectElement>('[data-testid="sheet-switcher"]');
if (!sheetSwitcher) {
  throw new Error("Missing sheet switcher element");
}

function renderSheetSwitcher(sheets: { id: string; name: string }[], activeId: string) {
  sheetSwitcher.replaceChildren();
  for (const sheet of sheets) {
    const option = document.createElement("option");
    option.value = sheet.id;
    option.textContent = sheet.name;
    sheetSwitcher.appendChild(option);
  }
  sheetSwitcher.value = activeId;
}

renderSheetSwitcher([{ id: app.getCurrentSheetId(), name: app.getCurrentSheetId() }], app.getCurrentSheetId());
sheetSwitcher.addEventListener("change", () => {
  app.activateSheet(sheetSwitcher.value);
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

let tauriBackend: TauriWorkbookBackend | null = null;
let activeWorkbook: WorkbookInfo | null = null;
let pendingBackendSync: Promise<void> = Promise.resolve();
type TauriInvoke = (cmd: string, args?: any) => Promise<any>;
let queuedInvoke: TauriInvoke | null = null;
let workbookSync: ReturnType<typeof startWorkbookSync> | null = null;

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

async function loadWorkbookIntoDocument(info: WorkbookInfo): Promise<void> {
  if (!tauriBackend) {
    throw new Error("Workbook backend not available");
  }

  const doc = app.getDocument();
  const sheets = normalizeSheetList(info);
  if (sheets.length === 0) {
    throw new Error("Workbook contains no sheets");
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
  await app.restoreDocumentState(snapshot);

  // Ensure sheets exist even if they were empty (DocumentController lazily creates models).
  for (const sheet of sheets) {
    doc.getCell(sheet.id, { row: 0, col: 0 });
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

  // Drain any queued workbook sync IPC before swapping the backend workbook.
  // Otherwise stale `set_cell` / `set_range` calls could land in the newly-opened workbook.
  await new Promise<void>((resolve) => queueMicrotask(resolve));

  const hadActiveWorkbook = activeWorkbook != null;

  workbookSync?.stop();
  workbookSync = null;

  try {
    await pendingBackendSync;
    pendingBackendSync = Promise.resolve();

    activeWorkbook = await tauriBackend.openWorkbook(path);
    await loadWorkbookIntoDocument(activeWorkbook);

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });
  } catch (err) {
    // If we were unable to swap workbooks, restore syncing for the previously-active
    // workbook so edits remain persistable.
    if (hadActiveWorkbook) {
      workbookSync = startWorkbookSync({
        document: app.getDocument(),
        engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
      });
    }
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

  const { save } = getTauriDialog();
  const path = await save({
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (!path) return;

  // Ensure any pending microtask-batched workbook edits are flushed before saving.
  await new Promise<void>((resolve) => queueMicrotask(resolve));
  await pendingBackendSync;
  await tauriBackend.saveWorkbook(path);
  activeWorkbook = { ...activeWorkbook, path };
  app.getDocument().markSaved();
}

async function handleNewWorkbook(): Promise<void> {
  if (!tauriBackend) return;
  const ok = await confirmDiscardDirtyState("create a new workbook");
  if (!ok) return;

  // Ensure any scheduled workbook IPC flush runs before we replace the backend workbook.
  await new Promise<void>((resolve) => queueMicrotask(resolve));

  const hadActiveWorkbook = activeWorkbook != null;

  workbookSync?.stop();
  workbookSync = null;

  try {
    await pendingBackendSync;
    pendingBackendSync = Promise.resolve();

    activeWorkbook = await tauriBackend.newWorkbook();
    await loadWorkbookIntoDocument(activeWorkbook);

    workbookSync = startWorkbookSync({
      document: app.getDocument(),
      engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
    });
  } catch (err) {
    if (hadActiveWorkbook) {
      workbookSync = startWorkbookSync({
        document: app.getDocument(),
        engineBridge: queuedInvoke ? { invoke: queuedInvoke } : undefined,
      });
    }
    throw err;
  }
}

try {
  tauriBackend = new TauriWorkbookBackend();
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (invoke) {
    queuedInvoke = (cmd, args) => queueBackendOp(() => invoke(cmd, args));
  }
  window.addEventListener("unload", () => workbookSync?.stop());

  const listen = getTauriListen();
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

  void listen("shortcut-quick-open", () => {
    void promptOpenWorkbook().catch((err) => {
      console.error("Failed to open workbook:", err);
      window.alert(`Failed to open workbook: ${String(err)}`);
    });
  });

  void listen("unsaved-changes", async () => {
    const discard = window.confirm("You have unsaved changes. Discard them?");
    if (discard) {
      await hideTauriWindow();
    }
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
