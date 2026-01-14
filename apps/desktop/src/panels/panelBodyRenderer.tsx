import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { PanelIds } from "./panelRegistry.js";
import { AIChatPanelContainer } from "./ai-chat/AIChatPanelContainer.js";
import { DataQueriesPanelContainer } from "./data-queries/DataQueriesPanelContainer.js";
import { QueryEditorPanelContainer } from "./query-editor/QueryEditorPanelContainer.js";
import { createAIAuditPanel } from "./ai-audit/index.js";
import { mountPythonPanel } from "./python/index.js";
import { VbaMigratePanel } from "./vba-migrate/index.js";
import { createMarketplacePanel } from "./marketplace/index.js";
import { ExtensionPanelBody } from "../extensions/ExtensionPanelBody.js";
import { ExtensionsPanel } from "../extensions/ExtensionsPanel.js";
import type { ExtensionPanelBridge } from "../extensions/extensionPanelBridge.js";
import type { DesktopExtensionHostManager } from "../extensions/extensionHostManager.js";
import type { PanelRegistry } from "./panelRegistry.js";
import { PivotBuilderPanelContainer } from "./pivot-builder/PivotBuilderPanelContainer.js";
import type { SpreadsheetApi } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import { MarketplaceClient, WebExtensionManager } from "@formula/extension-marketplace";
import type { SheetNameResolver } from "../sheet/sheetNameResolver.js";
import { CollabVersionHistoryPanel } from "./version-history/index.js";
import { CollabBranchManagerPanel } from "./branch-manager/CollabBranchManagerPanel.js";
import { getMarketplaceBaseUrl } from "./marketplace/getMarketplaceBaseUrl.ts";
import { verifyExtensionPackageV2Desktop } from "./marketplace/verifyExtensionPackageV2Desktop.ts";
import { t } from "../i18n/index.js";
export interface PanelBodyRendererOptions {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  getSelection?: () => unknown;
  getSearchWorkbook?: () => unknown;
  getCharts?: () => unknown;
  getSelectedChartId?: () => string | null;
  /**
   * Optional stable-id <-> display-name resolver for sheet-qualified strings.
   *
   * Panels that surface sheet locations to users should prefer the display name
   * (e.g. `Budget!A1`) over stable ids (e.g. `sheet_<uuid>!A1`).
   */
  sheetNameResolver?: SheetNameResolver | null;
  workbookId?: string;
  /**
   * Optional invoke wrapper (typically a queued/serialized Tauri invoke).
   *
   * When provided, panels that call into the backend (e.g. VBA migration
   * validation) should prefer it so commands run after any pending workbook
   * sync operations.
   */
  invoke?: (cmd: string, args?: any) => Promise<any>;
  /**
   * Optional hook to drain pending backend sync operations (e.g. microtask-batched
   * `set_cell` / `set_range` calls) before running a long-running command like
   * macro validation.
   */
  drainBackendSync?: () => Promise<void>;
  /**
   * Optional callback returning the current macro UI context (active sheet/cell +
   * selection). Used by panels that need to run VBA in a UI-consistent context.
   */
  getMacroUiContext?: () => {
    sheetId: string;
    activeRow: number;
    activeCol: number;
    selection?: { startRow: number; startCol: number; endRow: number; endCol: number } | null;
  };
  /**
   * Optional dynamic workbook id accessor.
   *
   * Some desktop flows (open workbook / new workbook) swap the active document
   * without recreating the panel renderer. Using a getter allows panels that
   * persist per-workbook state (Power Query refresh schedules, chat context, etc)
   * to re-render with the latest workbook id.
   */
  getWorkbookId?: () => string | undefined;
  /**
   * Optional accessor for the active collaboration session.
   *
   * When provided and returns a session, panels like Version History / Branch
   * Manager can use the shared Y.Doc as their backend.
   */
  getCollabSession?: () => import("@formula/collab-session").CollabSession | null;
  /**
   * Optional factory for the Version History panel's VersionStore.
   *
   * Reserved-root-guard deployments (SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED)
   * reject Yjs updates touching reserved roots like `versions*`. The default
   * in-doc store used by `@formula/collab-versioning` (YjsVersionStore) writes to
   * those roots, which can cause the sync server to close the websocket (1008).
   *
   * Provide an out-of-doc store implementation (e.g. ApiVersionStore, SQLite) to
   * keep version history out of the collaborative document.
   */
  createVersionStore?: (
    session: import("@formula/collab-session").CollabSession,
  ) =>
    | import("../../../../packages/collab/versioning/src/index.ts").VersionStore
    | Promise<import("../../../../packages/collab/versioning/src/index.ts").VersionStore>;
  /**
   * Optional factory for the Branch Manager panel's BranchStore.
   *
   * Reserved-root-guard deployments reject Yjs updates touching reserved roots
   * like `branching:*`. The default in-doc store (YjsBranchStore) writes to those
   * roots.
   *
   * Provide an out-of-doc store implementation to avoid reserved root
   * mutations.
   */
  createBranchStore?: (
    session: import("@formula/collab-session").CollabSession,
  ) =>
    | import("./branch-manager/branchStoreTypes.js").BranchStore
    | Promise<import("./branch-manager/branchStoreTypes.js").BranchStore>;
  renderMacrosPanel?: (body: HTMLDivElement) => void;
  createChart?: SpreadsheetApi["createChart"];
  panelRegistry?: PanelRegistry;
  extensionPanelBridge?: ExtensionPanelBridge;
  extensionHostManager?: DesktopExtensionHostManager;
  /**
   * Load the lazy extension host (used by built-in UI panels that want to show extension
   * contributions even when opened via layout restoration rather than an explicit command).
   */
  ensureExtensionsLoaded?: () => Promise<void>;
  onExecuteExtensionCommand?: (commandId: string, ...args: any[]) => Promise<unknown> | void;
  onOpenExtensionPanel?: (panelId: string) => void;
  onSyncExtensions?: () => void;
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

  // Marketplace wiring (lazy so desktop builds without a marketplace service don't
  // eagerly touch IndexedDB / crypto).
  let marketplaceClient: MarketplaceClient | null = null;
  let marketplaceExtensionManager: WebExtensionManager | null = null;
  let marketplaceExtensionHostManager:
    | {
        syncInstalledExtensions: () => Promise<void>;
        reloadExtension: (id: string) => Promise<void>;
        unloadExtension: (id: string) => Promise<void>;
        resetExtensionState: (id: string) => Promise<void>;
      }
    | null = null;

  function getMarketplaceServices() {
    // Prefer using the DesktopExtensionHostManager's shared marketplace services so that:
    // - installed extensions loaded at boot are tracked in a single WebExtensionManager instance
    //   (so unload/update can revoke blob URLs correctly)
    // - panel/UI code doesn't accidentally create a second manager with a separate loaded-state map
    if (options.extensionHostManager) {
      marketplaceClient = options.extensionHostManager.getMarketplaceClient();
      marketplaceExtensionManager = options.extensionHostManager.getMarketplaceExtensionManager();
    }

    if (!marketplaceClient) {
      marketplaceClient = marketplaceExtensionManager?.marketplaceClient ?? new MarketplaceClient({ baseUrl: getMarketplaceBaseUrl() });
    }

    if (!marketplaceExtensionManager) {
      marketplaceExtensionManager = new WebExtensionManager({
        marketplaceClient,
        host: (options.extensionHostManager?.host as any) ?? null,
        engineVersion: options.extensionHostManager?.engineVersion ?? "1.0.0",
        verifyPackage: verifyExtensionPackageV2Desktop,
      });
    }

    if (!marketplaceExtensionHostManager) {
      marketplaceExtensionHostManager = {
        syncInstalledExtensions: async () => {
          // Preserve the desktop "lazy-load extensions" behavior: allow users to install
          // extensions without immediately starting the extension host.
          if (options.extensionHostManager && !options.extensionHostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          await manager.loadAllInstalled().catch(() => {});
          try {
            options.extensionHostManager?.notifyDidChange();
          } catch {
            // ignore
          }
        },
        reloadExtension: async (id: string) => {
          if (options.extensionHostManager && !options.extensionHostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          try {
            if (manager.isLoaded(id)) {
              await manager.unload(id);
            }
            await manager.loadInstalled(id);
          } finally {
            try {
              options.extensionHostManager?.notifyDidChange();
            } catch {
              // ignore
            }
          }
        },
        unloadExtension: async (id: string) => {
          if (options.extensionHostManager && !options.extensionHostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          try {
            await manager.unload(id);
          } finally {
            try {
              options.extensionHostManager?.notifyDidChange();
            } catch {
              // ignore
            }
          }
        },
        resetExtensionState: async (id: string) => {
          // Clearing persisted state (permissions + storage) should be possible even when the
          // extension host has not been started yet. This keeps Marketplace uninstall behavior
          // consistent with a clean slate reinstall.
          try {
            await (options.extensionHostManager?.host as any)?.resetExtensionState?.(id);
          } catch {
            // Best-effort: state cleanup should not block uninstall.
          }
        },
      };
    }

    return {
      marketplaceClient,
      extensionManager: marketplaceExtensionManager,
      extensionHostManager: marketplaceExtensionHostManager,
    };
  }

  function renderReactPanel(panelId: string, body: HTMLDivElement, element: React.ReactElement) {
    let instance = reactPanels.get(panelId);
    if (!instance) {
      const container = document.createElement("div");
      container.className = "panel-body__container";
      instance = { root: createRoot(container), container };
      reactPanels.set(panelId, instance);
    }

    // Re-assert the sizing/flex class in case callers (or devtools) mutate it.
    instance.container.classList.add("panel-body__container");
    body.appendChild(instance.container);
    instance.root.render(element);
  }

  function renderDomPanel(panelId: string, body: HTMLDivElement, mount: (container: HTMLDivElement) => DomPanelInstance) {
    let instance = domPanels.get(panelId);
    if (!instance) {
      const container = document.createElement("div");
      container.className = "panel-body__container";
      instance = mount(container);
      instance.container.classList.add("panel-body__container");
      domPanels.set(panelId, instance);
    }

    // Re-assert the sizing/flex class in case the mount implementation overwrote it.
    instance.container.classList.add("panel-body__container");
    body.appendChild(instance.container);
    void instance.refresh?.();
  }

  function makeBodyFillAvailableHeight(body: HTMLDivElement) {
    body.classList.add("panel-body--fill");
  }

  function renderPanelBody(panelId: string, body: HTMLDivElement) {
    // Reset any previous renderer-specific body modifiers before applying this panel's layout.
    body.classList.remove("panel-body--fill");
    const workbookId = options.getWorkbookId?.() ?? options.workbookId;

    if (panelId === PanelIds.AI_CHAT) {
      // Ensure the chat UI can own the full panel height (dock panels are flex columns).
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <AIChatPanelContainer
          key={workbookId ?? "default"}
          getDocumentController={options.getDocumentController}
          getActiveSheetId={options.getActiveSheetId}
          getSelection={options.getSelection as any}
          getSearchWorkbook={options.getSearchWorkbook}
          getCharts={options.getCharts as any}
          getSelectedChartId={options.getSelectedChartId}
          sheetNameResolver={options.sheetNameResolver}
          workbookId={workbookId}
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
          key={workbookId ?? "default"}
          getDocumentController={options.getDocumentController}
          getActiveSheetId={options.getActiveSheetId}
          workbookId={workbookId}
        />,
      );
      return;
    }

    if (panelId === PanelIds.EXTENSIONS && options.extensionHostManager && options.onExecuteExtensionCommand && options.onOpenExtensionPanel) {
      // Extensions are lazy-loaded. If the Extensions panel is restored from persisted layout
      // (rather than opened via the ribbon command), we still want it to trigger loading.
      void options.ensureExtensionsLoaded?.().catch(() => {
        // Best-effort: keep the panel UI responsive even if extensions fail to load.
      });
      makeBodyFillAvailableHeight(body);
      const marketplaceManager = (() => {
        try {
          return getMarketplaceServices().extensionManager;
        } catch {
          return null;
        }
      })();
      renderReactPanel(
        panelId,
        body,
        <ExtensionsPanel
          manager={options.extensionHostManager}
          webExtensionManager={marketplaceManager}
          onSyncExtensions={options.onSyncExtensions}
          onExecuteCommand={options.onExecuteExtensionCommand}
          onOpenPanel={options.onOpenExtensionPanel}
        />,
      );
      return;
    }

    if (panelId === PanelIds.PIVOT_BUILDER) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <PivotBuilderPanelContainer
          key={workbookId ?? "default"}
          getDocumentController={options.getDocumentController}
          getActiveSheetId={options.getActiveSheetId}
          getSelection={options.getSelection as any}
          sheetNameResolver={options.sheetNameResolver ?? null}
          invoke={options.invoke as any}
          drainBackendSync={options.drainBackendSync}
        />,
      );
      return;
    }

    if (panelId === PanelIds.DATA_QUERIES) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <DataQueriesPanelContainer
          key={workbookId ?? "default"}
          getDocumentController={options.getDocumentController}
          workbookId={workbookId}
          sheetNameResolver={options.sheetNameResolver ?? null}
        />,
      );
      return;
    }

    if (panelId === PanelIds.MARKETPLACE) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        const services = getMarketplaceServices();
        const panel = createMarketplacePanel({
          container,
          marketplaceClient: services.marketplaceClient,
          extensionManager: services.extensionManager,
          extensionHostManager: services.extensionHostManager,
        });
        return { container, dispose: panel.dispose };
      });
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

    const panelDef = options.panelRegistry?.get(panelId) as any;
    if (panelDef?.source?.kind === "extension" && options.extensionPanelBridge) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(panelId, body, <ExtensionPanelBody panelId={panelId} bridge={options.extensionPanelBridge} />);
      return;
    }

    if (panelId === PanelIds.AI_AUDIT) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        const panel = createAIAuditPanel({
          container,
          initialWorkbookId: workbookId,
          autoRefreshMs: 1_000,
        });
        return { container, dispose: panel.dispose, refresh: panel.refresh };
      });
      return;
    }

    if (panelId === PanelIds.VERSION_HISTORY) {
      makeBodyFillAvailableHeight(body);
      const session = options.getCollabSession?.() ?? null;
      if (!session) {
        renderReactPanel(
          panelId,
          body,
          <div className="collab-panel__message collab-panel__message--error">{t("versionHistory.panel.noSession")}</div>,
        );
        return;
      }
      renderReactPanel(
        panelId,
        body,
        <CollabVersionHistoryPanel
          session={session}
          sheetNameResolver={options.sheetNameResolver ?? null}
          createVersionStore={options.createVersionStore}
        />,
      );
      return;
    }

    if (panelId === PanelIds.BRANCH_MANAGER) {
      makeBodyFillAvailableHeight(body);
      const session = options.getCollabSession?.() ?? null;
      if (!session) {
        renderReactPanel(
          panelId,
          body,
          <div className="collab-panel__message collab-panel__message--error">Branch manager requires collaboration mode.</div>,
        );
        return;
      }
      renderReactPanel(
        panelId,
        body,
        <CollabBranchManagerPanel
          session={session}
          sheetNameResolver={options.sheetNameResolver ?? null}
          createBranchStore={options.createBranchStore}
        />,
      );
      return;
    }

    if (panelId === PanelIds.MACROS) {
      if (options.renderMacrosPanel) return options.renderMacrosPanel(body);
      body.textContent = "Macros will appear here.";
      return;
    }

    if (panelId === PanelIds.VBA_MIGRATE) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <VbaMigratePanel
          key={workbookId ?? "default"}
          workbookId={workbookId}
          invoke={options.invoke}
          drainBackendSync={options.drainBackendSync}
          getMacroUiContext={options.getMacroUiContext}
        />,
      );
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
