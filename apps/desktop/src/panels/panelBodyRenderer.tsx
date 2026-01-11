import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { PanelIds } from "./panelRegistry.js";
import { AIChatPanelContainer } from "./ai-chat/AIChatPanelContainer.js";
import { QueryEditorPanelContainer } from "./query-editor/QueryEditorPanelContainer.js";
import { createAIAuditPanel } from "./ai-audit/index.js";
import { mountPythonPanel } from "./python/index.js";
import { VbaMigratePanel } from "./vba-migrate/index.js";
import type { SpreadsheetApi } from "../../../../packages/ai-tools/src/spreadsheet/api.js";

export interface PanelBodyRendererOptions {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  workbookId?: string;
  renderMacrosPanel?: (body: HTMLDivElement) => void;
  createChart?: SpreadsheetApi["createChart"];
}

export interface PanelBodyRenderer {
  renderPanelBody: (panelId: string, body: HTMLDivElement) => void;
  cleanup: (openPanelIds: Iterable<string>) => void;
}

interface ReactPanelInstance {
  root: Root;
  container: HTMLDivElement;
}

interface DomPanelInstance {
  container: HTMLDivElement;
  dispose: () => void;
  refresh?: () => Promise<void> | void;
}

export function createPanelBodyRenderer(options: PanelBodyRendererOptions): PanelBodyRenderer {
  const reactPanels = new Map<string, ReactPanelInstance>();
  const domPanels = new Map<string, DomPanelInstance>();

  function renderReactPanel(panelId: string, body: HTMLDivElement, element: React.ReactElement) {
    let instance = reactPanels.get(panelId);
    if (!instance) {
      const container = document.createElement("div");
      container.style.height = "100%";
      container.style.display = "flex";
      container.style.flexDirection = "column";
      instance = { root: createRoot(container), container };
      reactPanels.set(panelId, instance);
    }

    body.appendChild(instance.container);
    instance.root.render(element);
  }

  function renderDomPanel(panelId: string, body: HTMLDivElement, mount: (container: HTMLDivElement) => DomPanelInstance) {
    let instance = domPanels.get(panelId);
    if (!instance) {
      const container = document.createElement("div");
      container.style.height = "100%";
      container.style.display = "flex";
      container.style.flexDirection = "column";
      instance = mount(container);
      domPanels.set(panelId, instance);
    }

    body.appendChild(instance.container);
    void instance.refresh?.();
  }

  function makeBodyFillAvailableHeight(body: HTMLDivElement) {
    body.style.flex = "1";
    body.style.minHeight = "0";
    body.style.padding = "0";
    body.style.display = "flex";
    body.style.flexDirection = "column";
  }

  function renderPanelBody(panelId: string, body: HTMLDivElement) {
    if (panelId === PanelIds.AI_CHAT) {
      // Ensure the chat UI can own the full panel height (dock panels are flex columns).
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <AIChatPanelContainer
          getDocumentController={options.getDocumentController}
          getActiveSheetId={options.getActiveSheetId}
          workbookId={options.workbookId}
          createChart={options.createChart}
        />,
      );
      return;
    }

    if (panelId === PanelIds.QUERY_EDITOR) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <QueryEditorPanelContainer
          getDocumentController={options.getDocumentController}
          getActiveSheetId={options.getActiveSheetId}
          workbookId={options.workbookId}
        />,
      );
      return;
    }

    if (panelId === PanelIds.PYTHON) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        const dispose = mountPythonPanel({
          // `DocumentControllerBridge` expects the desktop `DocumentController` shape.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          documentController: options.getDocumentController() as any,
          container,
          getActiveSheetId: options.getActiveSheetId,
        });
        return { container, dispose };
      });
      return;
    }

    if (panelId === PanelIds.AI_AUDIT) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        const panel = createAIAuditPanel({
          container,
          initialWorkbookId: options.workbookId,
          autoRefreshMs: 1_000,
        });
        return { container, dispose: panel.dispose, refresh: panel.refresh };
      });
      return;
    }

    if (panelId === PanelIds.VERSION_HISTORY) {
      body.textContent = "Version history will appear here.";
      return;
    }

    if (panelId === PanelIds.MACROS) {
      if (options.renderMacrosPanel) return options.renderMacrosPanel(body);
      body.textContent = "Macros will appear here.";
      return;
    }

    if (panelId === PanelIds.VBA_MIGRATE) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(panelId, body, <VbaMigratePanel workbookId={options.workbookId} />);
      return;
    }

    body.textContent = `Panel: ${panelId}`;
  }

  function cleanup(openPanelIds: Iterable<string>) {
    const open = new Set(openPanelIds);
    for (const [panelId, instance] of reactPanels) {
      if (open.has(panelId)) continue;
      instance.root.unmount();
      instance.container.remove();
      reactPanels.delete(panelId);
    }

    for (const [panelId, instance] of domPanels) {
      if (open.has(panelId)) continue;
      instance.dispose();
      instance.container.remove();
      domPanels.delete(panelId);
    }
  }

  return { renderPanelBody, cleanup };
}
