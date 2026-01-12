import React, { useEffect, useMemo, useState } from "react";
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
import type { CollabSession } from "@formula/collab-session";
import { buildVersionHistoryItems } from "./version-history/index.js";
import { BranchManagerPanel, type Actor as BranchActor } from "./branch-manager/BranchManagerPanel.js";
import { MergeBranchPanel } from "./branch-manager/MergeBranchPanel.js";
// Import the browser-safe branch store/service modules directly.
//
// `packages/versioning/branches/src/index.js` also re-exports a Node-only `SQLiteBranchStore`,
// which pulls `node:*` built-ins into the web bundle and breaks the desktop web shell / e2e.
import { BranchService } from "../../../../packages/versioning/branches/src/BranchService.js";
import { YjsBranchStore } from "../../../../packages/versioning/branches/src/store/YjsBranchStore.js";
import { applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "../../../../packages/versioning/branches/src/yjs/index.js";
import { BRANCHING_APPLY_ORIGIN } from "../collab/conflict-monitors.js";
import { getMarketplaceBaseUrl } from "./marketplace/getMarketplaceBaseUrl.ts";
import { verifyExtensionPackageV2Desktop } from "./marketplace/verifyExtensionPackageV2Desktop.ts";
import { showInputBox } from "../extensions/ui.js";
import * as nativeDialogs from "../tauri/nativeDialogs.js";

function formatVersionTimestamp(timestampMs: number): string {
  try {
    return new Date(timestampMs).toLocaleString();
  } catch {
    return String(timestampMs);
  }
}

function CollabVersionHistoryPanel({ session }: { session: CollabSession }) {
  // `@formula/collab-versioning` depends on the core versioning subsystem, which can pull in
  // Node-only modules (e.g. `node:events`). Avoid importing it at desktop shell startup so
  // split-view/grid e2e can boot without requiring those polyfills; load it lazily when the
  // panel is actually opened.
  const [collabVersioning, setCollabVersioning] = useState<any | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [versions, setVersions] = useState<any[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let disposed = false;
    let instance: any | null = null;

    void (async () => {
      try {
        setLoadError(null);
        setCollabVersioning(null);
        const mod = await import("../../../../packages/collab/versioning/src/index.js");
        if (disposed) return;
        const localPresence = session.presence?.localPresence ?? null;
        instance = mod.createCollabVersioning({
          session,
          user: localPresence ? { userId: localPresence.id, userName: localPresence.name } : undefined,
        });
        setCollabVersioning(instance);
      } catch (e) {
        if (disposed) return;
        setLoadError((e as Error).message);
      }
    })();

    return () => {
      disposed = true;
      instance?.destroy();
    };
  }, [session]);

  const refresh = async () => {
    try {
      setError(null);
      const manager = collabVersioning;
      if (!manager) return;
      const next = await manager.listVersions();
      setVersions(next);
      if (selectedId && !next.some((v) => v.id === selectedId)) setSelectedId(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  useEffect(() => {
    if (!collabVersioning) return;
    void refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [collabVersioning]);

  const items = useMemo(() => buildVersionHistoryItems(versions as any), [versions]);

  if (loadError) {
    return (
      <div className="collab-panel__message collab-panel__message--error">
        Version history is unavailable: {loadError}
      </div>
    );
  }

  if (!collabVersioning) {
    return <div className="collab-panel__message">Loading version history…</div>;
  }

  return (
    <div className="collab-version-history">
      <h3 className="collab-version-history__title">Version history</h3>

      {error ? <div className="collab-version-history__error">{error}</div> : null}

      <div className="collab-version-history__actions">
        <button
          disabled={busy}
          onClick={async () => {
            const name = await showInputBox({ prompt: "Checkpoint name?" });
            if (!name || !name.trim()) return;
            try {
              setBusy(true);
              setError(null);
              await collabVersioning.createCheckpoint({ name: name.trim() });
              await refresh();
            } catch (e) {
              setError((e as Error).message);
            } finally {
              setBusy(false);
            }
          }}
        >
          Create checkpoint
        </button>

        <button
          disabled={busy || !selectedId}
          onClick={async () => {
            const id = selectedId;
            if (!id) return;
            const ok = await nativeDialogs.confirm(
              "Restore this version? This will overwrite the current collaborative document state.",
            );
            if (!ok) return;
            try {
              setBusy(true);
              setError(null);
              await collabVersioning.restoreVersion(id);
              await refresh();
            } catch (e) {
              setError((e as Error).message);
            } finally {
              setBusy(false);
            }
          }}
        >
          Restore selected
        </button>

        <button disabled={busy} onClick={() => void refresh()}>
          Refresh
        </button>
      </div>

      {items.length === 0 ? (
        <div className="collab-version-history__empty">No versions yet.</div>
      ) : (
        <ul className="collab-version-history__list">
          {items.map((item) => {
            const selected = item.id === selectedId;
            return (
              <li
                key={item.id}
                className={
                  selected
                    ? "collab-version-history__item collab-version-history__item--selected"
                    : "collab-version-history__item"
                }
                onClick={() => setSelectedId(item.id)}
              >
                <input type="radio" checked={selected} onChange={() => setSelectedId(item.id)} />
                <div className="collab-version-history__item-content">
                  <div className="collab-version-history__item-title">{item.title}</div>
                  <div className="collab-version-history__item-meta">
                    {formatVersionTimestamp(item.timestampMs)} • {item.kind}
                    {item.locked ? " • locked" : ""}
                  </div>
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}

function CollabBranchManagerPanel({
  session,
  sheetNameResolver,
}: {
  session: CollabSession;
  sheetNameResolver?: SheetNameResolver | null;
}) {
  const actor = useMemo<BranchActor>(() => {
    const userId = session.presence?.localPresence?.id ?? "desktop";
    return { userId, role: "owner" };
  }, [session]);
  const docId = session.doc.guid;

  const store = useMemo(() => new YjsBranchStore({ ydoc: session.doc }), [session]);
  const branchService = useMemo(() => new BranchService({ docId, store }), [docId, store]);

  const [error, setError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [mergeSource, setMergeSource] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        setError(null);
        setReady(false);
        const initialState = yjsDocToDocumentState(session.doc);
        await branchService.init(actor as any, initialState as any);
        if (cancelled) return;
        setReady(true);
      } catch (e) {
        if (cancelled) return;
        setError((e as Error).message);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [actor, branchService, session]);

  const workflow = useMemo(() => {
    const commitCurrentState = async (message: string) => {
      const nextState = yjsDocToDocumentState(session.doc);
      await branchService.commit(actor as any, { nextState, message });
    };

    return {
      listBranches: () => branchService.listBranches(),
      createBranch: async (a: BranchActor, input: { name: string; description?: string }) => {
        await commitCurrentState("auto: create branch");
        return branchService.createBranch(a as any, input as any);
      },
      renameBranch: (a: BranchActor, input: { oldName: string; newName: string }) => branchService.renameBranch(a as any, input as any),
      deleteBranch: (a: BranchActor, input: { name: string }) => branchService.deleteBranch(a as any, input as any),
      checkoutBranch: async (a: BranchActor, input: { name: string }) => {
        await commitCurrentState("auto: checkout");
        const state = await branchService.checkoutBranch(a as any, input as any);
        // Branch checkout is a bulk "time travel" operation and must not be captured by
        // collaborative undo tracking. CollabSession also treats this origin as ignored
        // for conflict monitors so it doesn't surface spurious conflicts.
        applyDocumentStateToYjsDoc(session.doc, state as any, { origin: BRANCHING_APPLY_ORIGIN });
        return state;
      },
      previewMerge: async (a: BranchActor, input: { sourceBranch: string }) => {
        await commitCurrentState("auto: preview merge");
        return branchService.previewMerge(a as any, input as any);
      },
      merge: async (a: BranchActor, input: { sourceBranch: string; resolutions: any[]; message?: string }) => {
        await commitCurrentState("auto: merge");
        const result = await branchService.merge(a as any, input as any);
        // See checkoutBranch origin note above.
        applyDocumentStateToYjsDoc(session.doc, (result as any).state, { origin: BRANCHING_APPLY_ORIGIN });
        return result;
      },
    } as any;
  }, [actor, branchService, session]);

  if (error) {
    return (
      <div className="collab-panel__message collab-panel__message--error">
        {error}
      </div>
    );
  }

  if (!ready) {
    return <div className="collab-panel__message">Loading branches…</div>;
  }

  if (mergeSource) {
    return (
      <MergeBranchPanel
        actor={actor}
        branchService={workflow}
        sourceBranch={mergeSource}
        sheetNameResolver={sheetNameResolver ?? null}
        onClose={() => setMergeSource(null)}
      />
    );
  }

  return (
    <BranchManagerPanel
      actor={actor}
      branchService={workflow}
      onStartMerge={(sourceBranch) => setMergeSource(sourceBranch)}
    />
  );
}

export interface PanelBodyRendererOptions {
  getDocumentController: () => unknown;
  getActiveSheetId?: () => string;
  getSelection?: () => unknown;
  getSearchWorkbook?: () => unknown;
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
      const session = options.getCollabSession?.() ?? null;
      if (!session) {
        body.textContent = "Version history will appear here.";
        return;
      }
      makeBodyFillAvailableHeight(body);
      renderReactPanel(panelId, body, <CollabVersionHistoryPanel session={session} />);
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
        <CollabBranchManagerPanel session={session} sheetNameResolver={options.sheetNameResolver ?? null} />,
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
