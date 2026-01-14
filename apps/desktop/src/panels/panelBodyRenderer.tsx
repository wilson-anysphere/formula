import React, { useEffect, useState } from "react";
import { createRoot, type Root } from "react-dom/client";

import { PanelIds } from "./panelRegistry.js";
import { SolverPanel } from "./solver/SolverPanel.js";
import { ScenarioManagerPanel } from "./what-if/ScenarioManagerPanel.js";
import { MonteCarloWizard } from "./what-if/MonteCarloWizard.js";
import { createWhatIfApi } from "./what-if/api.js";
import { SelectionPanePanel } from "./selection-pane/SelectionPanePanel.js";
import type { ExtensionPanelBridge } from "../extensions/extensionPanelBridge.js";
import type { DesktopExtensionHostManager } from "../extensions/extensionHostManager.js";
import type { PanelRegistry } from "./panelRegistry.js";
import type { SpreadsheetApi } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import type { SheetNameResolver } from "../sheet/sheetNameResolver.js";
import { t } from "../i18n/index.js";

type MarketplaceClient = import("@formula/extension-marketplace").MarketplaceClient;
type WebExtensionManager = import("@formula/extension-marketplace").WebExtensionManager;

export interface PanelBodyRendererOptions {
  getDocumentController: () => unknown;
  /**
   * Optional SpreadsheetApp accessor for panels that need to interact with UI state beyond the
   * DocumentController (e.g. selection pane for drawings).
   */
  getSpreadsheetApp?: () => unknown;
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
   * Alias for {@link createVersionStore}.
   *
   * This name matches older wiring code and helps keep call sites explicit about the
   * collab/version-history context.
   */
  createCollabVersioningStore?: (
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
  /**
   * Alias for {@link createBranchStore}.
   */
  createCollabBranchStore?: (
    session: import("@formula/collab-session").CollabSession,
  ) =>
    | import("./branch-manager/branchStoreTypes.js").BranchStore
    | Promise<import("./branch-manager/branchStoreTypes.js").BranchStore>;
  renderMacrosPanel?: (body: HTMLDivElement) => void;
  createChart?: SpreadsheetApi["createChart"];
  panelRegistry?: PanelRegistry;
  extensionPanelBridge?: ExtensionPanelBridge;
  getExtensionPanelBridge?: () => ExtensionPanelBridge | null | undefined;
  extensionHostManager?: DesktopExtensionHostManager;
  getExtensionHostManager?: () => DesktopExtensionHostManager | null | undefined;
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
  const whatIfApi = createWhatIfApi();

  const getExtensionHostManager = (): DesktopExtensionHostManager | null =>
    options.getExtensionHostManager?.() ?? options.extensionHostManager ?? null;
  const getExtensionPanelBridge = (): ExtensionPanelBridge | null =>
    options.getExtensionPanelBridge?.() ?? options.extensionPanelBridge ?? null;

  // Marketplace wiring (lazy so desktop builds without a marketplace service don't
  // eagerly touch IndexedDB / crypto).
  type MarketplaceServices = {
    marketplaceClient: MarketplaceClient;
    extensionManager: WebExtensionManager;
    extensionHostManager: {
      syncInstalledExtensions: () => Promise<void>;
      reloadExtension: (id: string) => Promise<void>;
      unloadExtension: (id: string) => Promise<void>;
      resetExtensionState: (id: string) => Promise<void>;
    };
  };

  let marketplaceServicesPromise: Promise<MarketplaceServices> | null = null;

  async function getMarketplaceServices(): Promise<MarketplaceServices> {
    if (marketplaceServicesPromise) return marketplaceServicesPromise;

    marketplaceServicesPromise = (async () => {
      // Prefer using the DesktopExtensionHostManager's shared marketplace services so that:
      // - installed extensions loaded at boot are tracked in a single WebExtensionManager instance
      //   (so unload/update can revoke blob URLs correctly)
      // - panel/UI code doesn't accidentally create a second manager with a separate loaded-state map
      let marketplaceClient: MarketplaceClient | null = null;
      let marketplaceExtensionManager: WebExtensionManager | null = null;

      const extensionHostManager = getExtensionHostManager();
      if (extensionHostManager) {
        try {
          marketplaceClient = extensionHostManager.getMarketplaceClient();
          marketplaceExtensionManager = extensionHostManager.getMarketplaceExtensionManager();
        } catch {
          // ignore
        }
      }

      const [{ MarketplaceClient, WebExtensionManager }, { getMarketplaceBaseUrl }, { verifyExtensionPackageV2Desktop }] =
        await Promise.all([
          import("@formula/extension-marketplace"),
          import("./marketplace/getMarketplaceBaseUrl.ts"),
          import("./marketplace/verifyExtensionPackageV2Desktop.ts"),
        ]);

      if (!marketplaceClient) {
        marketplaceClient =
          (marketplaceExtensionManager as any)?.marketplaceClient ?? new MarketplaceClient({ baseUrl: getMarketplaceBaseUrl() });
      }

      if (!marketplaceExtensionManager) {
        marketplaceExtensionManager = new WebExtensionManager({
          marketplaceClient,
          host: (extensionHostManager?.host as any) ?? null,
          engineVersion: extensionHostManager?.engineVersion ?? "1.0.0",
          verifyPackage: verifyExtensionPackageV2Desktop,
        });
      }

      const marketplaceExtensionHostManager: MarketplaceServices["extensionHostManager"] = {
        syncInstalledExtensions: async () => {
          // Preserve the desktop "lazy-load extensions" behavior: allow users to install
          // extensions without immediately starting the extension host.
          const hostManager = getExtensionHostManager();
          if (hostManager && !hostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          await manager.loadAllInstalled().catch(() => {});
          try {
            getExtensionHostManager()?.notifyDidChange();
          } catch {
            // ignore
          }
        },
        reloadExtension: async (id: string) => {
          const hostManager = getExtensionHostManager();
          if (hostManager && !hostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          try {
            if (manager.isLoaded(id)) {
              await manager.unload(id);
            }
            await manager.loadInstalled(id);
          } finally {
            try {
              hostManager?.notifyDidChange();
            } catch {
              // ignore
            }
          }
        },
        unloadExtension: async (id: string) => {
          const hostManager = getExtensionHostManager();
          if (hostManager && !hostManager.ready) return;
          const manager = marketplaceExtensionManager!;
          try {
            await manager.unload(id);
          } finally {
            try {
              hostManager?.notifyDidChange();
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
            await (getExtensionHostManager()?.host as any)?.resetExtensionState?.(id);
          } catch {
            // Best-effort: state cleanup should not block uninstall.
          }
        },
      };

      return {
        marketplaceClient,
        extensionManager: marketplaceExtensionManager,
        extensionHostManager: marketplaceExtensionHostManager,
      };
    })();

    return marketplaceServicesPromise;
  }

  const LazyAIChatPanelContainer = React.lazy(() =>
    import("./ai-chat/AIChatPanelContainer.js").then((mod) => ({ default: (mod as any).AIChatPanelContainer })),
  );
  const LazyQueryEditorPanelContainer = React.lazy(() =>
    import("./query-editor/QueryEditorPanelContainer.js").then((mod) => ({ default: (mod as any).QueryEditorPanelContainer })),
  );
  const LazyPivotBuilderPanelContainer = React.lazy(() =>
    import("./pivot-builder/PivotBuilderPanelContainer.js").then((mod) => ({ default: (mod as any).PivotBuilderPanelContainer })),
  );
  const LazyDataQueriesPanelContainer = React.lazy(() =>
    import("./data-queries/DataQueriesPanelContainer.js").then((mod) => ({ default: (mod as any).DataQueriesPanelContainer })),
  );
  const LazyVbaMigratePanel = React.lazy(() =>
    import("./vba-migrate/index.js").then((mod) => ({ default: (mod as any).VbaMigratePanel })),
  );
  const LazyExtensionsPanel = React.lazy(() =>
    import("../extensions/ExtensionsPanel.js").then((mod) => ({ default: (mod as any).ExtensionsPanel })),
  );
  const LazyExtensionPanelBody = React.lazy(() =>
    import("../extensions/ExtensionPanelBody.js").then((mod) => ({ default: (mod as any).ExtensionPanelBody })),
  );
  const LazyCollabVersionHistoryPanel = React.lazy(() =>
    import("./version-history/CollabVersionHistoryPanel.js").then((mod) => ({
      default: (mod as any).CollabVersionHistoryPanel,
    })),
  );
  const LazyCollabBranchManagerPanel = React.lazy(() =>
    import("./branch-manager/CollabBranchManagerPanel.js").then((mod) => ({
      default: (mod as any).CollabBranchManagerPanel,
    })),
  );

  function ExtensionsPanelLoader(props: {
    onSyncExtensions?: (() => void) | undefined;
    onExecuteCommand: (commandId: string, ...args: any[]) => Promise<unknown> | void;
    onOpenPanel: (panelId: string) => void;
  }) {
    const [manager, setManager] = useState<DesktopExtensionHostManager | null>(() => getExtensionHostManager());
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
      let cancelled = false;
      void (async () => {
        try {
          setError(null);
          // Opening the Extensions panel is the primary trigger for extension host startup.
          await options.ensureExtensionsLoaded?.();
        } catch (err) {
          setError(String((err as any)?.message ?? err));
        } finally {
          if (cancelled) return;
          setManager(getExtensionHostManager());
        }
      })().catch(() => {
        // Best-effort: avoid unhandled rejections from effect cleanup/bookkeeping.
      });
      return () => {
        cancelled = true;
      };
    }, []);

    if (!manager) {
      return (
        <div className={error ? "collab-panel__message collab-panel__message--error" : "collab-panel__message"}>
          {error ? `Extensions unavailable: ${error}` : "Loading extensions…"}
        </div>
      );
    }

    return (
      <ExtensionsPanelWrapper
        manager={manager}
        onSyncExtensions={props.onSyncExtensions}
        onExecuteCommand={props.onExecuteCommand}
        onOpenPanel={props.onOpenPanel}
      />
    );
  }

  function ExtensionsPanelWrapper(props: {
    manager: DesktopExtensionHostManager;
    onSyncExtensions?: (() => void) | undefined;
    onExecuteCommand: (commandId: string, ...args: any[]) => Promise<unknown> | void;
    onOpenPanel: (panelId: string) => void;
  }) {
    const [marketplaceManager, setMarketplaceManager] = useState<any | null>(null);

    useEffect(() => {
      let disposed = false;
      void getMarketplaceServices()
        .then((services) => {
          if (disposed) return;
          setMarketplaceManager(services.extensionManager as any);
        })
        .catch(() => {
          // ignore
        });
      return () => {
        disposed = true;
      };
    }, []);

    return (
      <LazyExtensionsPanel
        manager={props.manager}
        webExtensionManager={marketplaceManager}
        onSyncExtensions={props.onSyncExtensions}
        onExecuteCommand={props.onExecuteCommand}
        onOpenPanel={props.onOpenPanel}
      />
    );
  }

  function ExtensionPanelBodyLoader(props: { panelId: string }) {
    const [bridge, setBridge] = useState<ExtensionPanelBridge | null>(() => getExtensionPanelBridge());
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
      let cancelled = false;
      if (bridge) return;
      void (async () => {
        try {
          setError(null);
          await options.ensureExtensionsLoaded?.();
        } catch (err) {
          setError(String((err as any)?.message ?? err));
        } finally {
          if (cancelled) return;
          setBridge(getExtensionPanelBridge());
        }
      })().catch(() => {
        // Best-effort: avoid unhandled rejections from effect cleanup/bookkeeping.
      });
      return () => {
        cancelled = true;
      };
    }, [bridge]);

    if (!bridge) {
      return (
        <div className={error ? "collab-panel__message collab-panel__message--error" : "collab-panel__message"}>
          {error ? `Extension panel unavailable: ${error}` : "Loading extension panel…"}
        </div>
      );
    }

    return (
      <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
        <LazyExtensionPanelBody panelId={props.panelId} bridge={bridge} />
      </React.Suspense>
    );
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
    try {
      const result = instance.refresh?.();
      if (typeof (result as any)?.then === "function") {
        void Promise.resolve(result).catch(() => {
          // Best-effort: panel refresh should not surface as an unhandled rejection.
        });
      }
    } catch {
      // Best-effort: ignore refresh failures.
    }
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
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyAIChatPanelContainer
            key={workbookId ?? "default"}
            getDocumentController={options.getDocumentController}
            getSpreadsheetApp={options.getSpreadsheetApp}
            getActiveSheetId={options.getActiveSheetId}
            getSelection={options.getSelection as any}
            getSearchWorkbook={options.getSearchWorkbook}
            getCharts={options.getCharts as any}
            getSelectedChartId={options.getSelectedChartId}
            sheetNameResolver={options.sheetNameResolver}
            workbookId={workbookId}
            createChart={options.createChart}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.SELECTION_PANE) {
      makeBodyFillAvailableHeight(body);
      const app = options.getSpreadsheetApp?.();
      if (!app) {
        body.textContent = "Selection Pane is unavailable.";
        return;
      }
      renderReactPanel(panelId, body, <SelectionPanePanel app={app as any} />);
      return;
    }

    if (panelId === PanelIds.QUERY_EDITOR) {
      makeBodyFillAvailableHeight(body);
      const app = options.getSpreadsheetApp?.() ?? null;
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyQueryEditorPanelContainer
            key={workbookId ?? "default"}
            getDocumentController={options.getDocumentController}
            getActiveSheetId={options.getActiveSheetId}
            app={app as any}
            workbookId={workbookId}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.EXTENSIONS && options.onExecuteExtensionCommand && options.onOpenExtensionPanel) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <ExtensionsPanelLoader
            onSyncExtensions={options.onSyncExtensions}
            onExecuteCommand={options.onExecuteExtensionCommand}
            onOpenPanel={options.onOpenExtensionPanel}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.PIVOT_BUILDER) {
      makeBodyFillAvailableHeight(body);
      const app = options.getSpreadsheetApp?.() ?? null;
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyPivotBuilderPanelContainer
            key={workbookId ?? "default"}
            getDocumentController={options.getDocumentController}
            getActiveSheetId={options.getActiveSheetId}
            getSelection={options.getSelection as any}
            sheetNameResolver={options.sheetNameResolver ?? null}
            app={app as any}
            invoke={options.invoke as any}
            drainBackendSync={options.drainBackendSync}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.SOLVER) {
      renderReactPanel(panelId, body, <SolverPanel />);
      return;
    }

    if (panelId === PanelIds.SCENARIO_MANAGER) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(panelId, body, <ScenarioManagerPanel api={whatIfApi} />);
      return;
    }

    if (panelId === PanelIds.MONTE_CARLO) {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(panelId, body, <MonteCarloWizard api={whatIfApi} />);
      return;
    }

    if (panelId === PanelIds.DATA_QUERIES) {
      makeBodyFillAvailableHeight(body);
      const app = options.getSpreadsheetApp?.() ?? null;
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyDataQueriesPanelContainer
            key={workbookId ?? "default"}
            getDocumentController={options.getDocumentController}
            app={app as any}
            workbookId={workbookId}
            sheetNameResolver={options.sheetNameResolver ?? null}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.MARKETPLACE) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        container.textContent = "Loading marketplace…";
        let disposed = false;
        let disposeFn: (() => void) | null = null;

        void (async () => {
          try {
            const [{ createMarketplacePanel }, services] = await Promise.all([
              import("./marketplace/index.js"),
              getMarketplaceServices(),
            ]);
            if (disposed) return;
            const panel = createMarketplacePanel({
              container,
              marketplaceClient: services.marketplaceClient,
              extensionManager: services.extensionManager,
              extensionHostManager: services.extensionHostManager,
            });
            disposeFn = panel.dispose;
          } catch (err) {
            if (disposed) return;
            // eslint-disable-next-line no-console
            console.error("[formula][desktop] Failed to load marketplace panel:", err);
            container.textContent = `Failed to load marketplace: ${String((err as any)?.message ?? err)}`;
          }
        })().catch(() => {
          // Best-effort: avoid unhandled rejections from panel bootstrapping.
        });

        return {
          container,
          dispose: () => {
            disposed = true;
            disposeFn?.();
          },
        };
      });
      return;
    }

    if (panelId === PanelIds.PYTHON) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        container.textContent = "Loading Python…";

        let disposed = false;
        let disposeFn: (() => void) | null = null;

        void (async () => {
          try {
            const { mountPythonPanel } = await import("./python/index.js");
            if (disposed) return;
            disposeFn = mountPythonPanel({
              // `DocumentControllerBridge` expects the desktop `DocumentController` shape.
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              documentController: options.getDocumentController() as any,
              container,
              getActiveSheetId: options.getActiveSheetId,
            });
          } catch (err) {
            if (disposed) return;
            // eslint-disable-next-line no-console
            console.error("[formula][desktop] Failed to load python panel:", err);
            container.textContent = `Failed to load Python: ${String((err as any)?.message ?? err)}`;
          }
        })().catch(() => {
          // Best-effort: avoid unhandled rejections from panel bootstrapping.
        });

        return {
          container,
          dispose: () => {
            disposed = true;
            try {
              disposeFn?.();
            } catch {
              // ignore
            }
            if (!disposeFn) container.innerHTML = "";
          },
        };
      });
      return;
    }

    const panelDef = options.panelRegistry?.get(panelId) as any;
    if (panelDef?.source?.kind === "extension") {
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <ExtensionPanelBodyLoader panelId={panelId} />,
      );
      return;
    }

    if (panelId === PanelIds.AI_AUDIT) {
      makeBodyFillAvailableHeight(body);
      renderDomPanel(panelId, body, (container) => {
        container.textContent = "Loading audit log…";
        let disposed = false;
        let disposeFn: (() => void) | null = null;
        let refreshFn: (() => Promise<void> | void) | undefined;

        void (async () => {
          try {
            const { createAIAuditPanel } = await import("./ai-audit/index.js");
            if (disposed) return;
            const panel = createAIAuditPanel({
              container,
              initialWorkbookId: workbookId,
              autoRefreshMs: 1_000,
            });
            disposeFn = panel.dispose;
            refreshFn = panel.refresh;
          } catch (err) {
            if (disposed) return;
            // eslint-disable-next-line no-console
            console.error("[formula][desktop] Failed to load audit panel:", err);
            container.textContent = `Failed to load audit panel: ${String((err as any)?.message ?? err)}`;
          }
        })().catch(() => {
          // Best-effort: avoid unhandled rejections from panel bootstrapping.
        });

        return {
          container,
          dispose: () => {
            disposed = true;
            disposeFn?.();
          },
          refresh: () => refreshFn?.(),
        };
      });
      return;
    }

    if (panelId === PanelIds.VERSION_HISTORY) {
      const session = options.getCollabSession?.() ?? null;
      if (!session) {
        body.textContent = t("versionHistory.panel.noSession");
        return;
      }
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">{t("versionHistory.panel.loading")}</div>}>
          <LazyCollabVersionHistoryPanel
            session={session}
            sheetNameResolver={options.sheetNameResolver ?? null}
            createVersionStore={options.createVersionStore ?? options.createCollabVersioningStore}
          />
        </React.Suspense>,
      );
      return;
    }

    if (panelId === PanelIds.BRANCH_MANAGER) {
      const session = options.getCollabSession?.() ?? null;
      if (!session) {
        body.textContent = "Branch manager requires collaboration mode.";
        return;
      }
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyCollabBranchManagerPanel
            session={session}
            sheetNameResolver={options.sheetNameResolver ?? null}
            createBranchStore={options.createBranchStore ?? options.createCollabBranchStore}
          />
        </React.Suspense>,
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
        <React.Suspense fallback={<div className="collab-panel__message">Loading…</div>}>
          <LazyVbaMigratePanel
            key={workbookId ?? "default"}
            workbookId={workbookId}
            invoke={options.invoke}
            drainBackendSync={options.drainBackendSync}
            getMacroUiContext={options.getMacroUiContext}
          />
        </React.Suspense>,
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
