import { SpreadsheetApp } from "./app/spreadsheetApp";
import "./styles/tokens.css";
import "./styles/ui.css";
import "./styles/workspace.css";

import { LayoutController } from "./layout/layoutController.js";
import { LayoutWorkspaceManager } from "./layout/layoutPersistence.js";
import { getPanelPlacement } from "./layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "./panels/panelRegistry.js";

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
const openAiPanel = document.querySelector<HTMLButtonElement>('[data-testid="open-ai-panel"]');

if (dockLeft && dockRight && dockBottom && floatingRoot && workspaceRoot && openAiPanel) {
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

  function panelTitle(panelId: string) {
    return (PANEL_REGISTRY as any)?.[panelId]?.title ?? panelId;
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
    body.textContent =
      active === PanelIds.AI_CHAT
        ? "Ask a question about your data…"
        : active === PanelIds.VERSION_HISTORY
          ? "Version history will appear here."
          : `Panel: ${active}`;

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
      body.textContent = `Floating panel: ${panelId}`;

      inner.appendChild(header);
      inner.appendChild(body);

      panel.appendChild(inner);
      floatingRoot.appendChild(panel);
    }
  }

  function renderLayout() {
    applyDockSizes();
    renderDock(dockLeft, layoutController.layout.docks.left, "left");
    renderDock(dockRight, layoutController.layout.docks.right, "right");
    renderDock(dockBottom, layoutController.layout.docks.bottom, "bottom");
    renderFloating();
  }

  openAiPanel.addEventListener("click", () => {
    const placement = getPanelPlacement(layoutController.layout, PanelIds.AI_CHAT);
    if (placement.kind === "closed") layoutController.openPanel(PanelIds.AI_CHAT);
    else layoutController.closePanel(PanelIds.AI_CHAT);
  });

  layoutController.on("change", () => renderLayout());
  renderLayout();
}

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
