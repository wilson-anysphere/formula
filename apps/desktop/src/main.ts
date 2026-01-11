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
import { DocumentWorkbookAdapter } from "./search/documentWorkbookAdapter.js";
import { registerFindReplaceShortcuts, FindReplaceController } from "./panels/find-replace/index.js";

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
app.focus();
openComments.addEventListener("click", () => app.toggleCommentsPanel());

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
    renderMacrosPanel: (body) => {
      body.textContent = "Loading macros…";
      queueMicrotask(() => {
        try {
          const backend = new TauriMacroBackend();
          void renderMacroRunner(body, backend, workbookId).catch((err) => {
            body.textContent = `Failed to load macros: ${String(err)}`;
          });
        } catch (err) {
          body.textContent = `Macros backend not available: ${String(err)}`;
        }
      });
    },
  });

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
    panelBodyRenderer.renderPanelBody(active, body);

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
      panelBodyRenderer.renderPanelBody(panelId, body);

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
    panelBodyRenderer.cleanup(openPanelIds());
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

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
