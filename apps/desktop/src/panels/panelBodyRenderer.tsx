import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { PanelIds } from "./panelRegistry.js";
import { AIChatPanelContainer } from "./ai-chat/AIChatPanelContainer.js";
import { createAIAuditPanel } from "./ai-audit/index.js";
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

export function createPanelBodyRenderer(options: PanelBodyRendererOptions): PanelBodyRenderer {
  const reactPanels = new Map<string, ReactPanelInstance>();

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

  function renderPanelBody(panelId: string, body: HTMLDivElement) {
    if (panelId === PanelIds.AI_CHAT) {
      // Ensure the React chat UI can own the full panel height (dock panels are flex columns).
      body.style.flex = "1";
      body.style.minHeight = "0";
      body.style.padding = "0";
      body.style.display = "flex";
      body.style.flexDirection = "column";
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

    if (panelId === PanelIds.AI_AUDIT) {
      createAIAuditPanel({ container: body });
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
      body.textContent = "VBA migration tools will appear here.";
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
  }

  return { renderPanelBody, cleanup };
}
