import React, { useEffect, useMemo, useState } from "react";
import { createRoot, type Root } from "react-dom/client";

import { PanelIds } from "./panelRegistry.js";
import { AIChatPanelContainer } from "./ai-chat/AIChatPanelContainer.js";
import { DataQueriesPanelContainer } from "./data-queries/DataQueriesPanelContainer.js";
import { QueryEditorPanelContainer } from "./query-editor/QueryEditorPanelContainer.js";
import { createCollabVersioning, type CollabVersioning, type VersionRecord } from "@formula/collab-versioning";
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
import type { CollabSession } from "@formula/collab-session";
import { buildVersionHistoryItems } from "./version-history/index.js";
import { BranchManagerPanel, type Actor as BranchActor } from "./branch-manager/BranchManagerPanel.js";
import { MergeBranchPanel } from "./branch-manager/MergeBranchPanel.js";
import { BranchService, YjsBranchStore, applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "../../../../packages/versioning/branches/src/index.js";
import { getMarketplaceBaseUrl } from "../marketplace/getMarketplaceBaseUrl.js";
import { MarketplaceClient } from "../../../web/src/marketplace/MarketplaceClient.js";
import { WebExtensionManager } from "../../../web/src/marketplace/WebExtensionManager.js";

function formatVersionTimestamp(timestampMs: number): string {
  try {
    return new Date(timestampMs).toLocaleString();
  } catch {
    return String(timestampMs);
  }
}

function CollabVersionHistoryPanel({ session }: { session: CollabSession }) {
  const collabVersioning = useMemo<CollabVersioning>(() => createCollabVersioning({ session }), [session]);

  const [versions, setVersions] = useState<VersionRecord[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    return () => {
      collabVersioning.destroy();
    };
  }, [collabVersioning]);

  const refresh = async () => {
    try {
      setError(null);
      const next = await collabVersioning.listVersions();
      setVersions(next);
      if (selectedId && !next.some((v) => v.id === selectedId)) setSelectedId(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  useEffect(() => {
    void refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [collabVersioning]);

  const items = useMemo(() => buildVersionHistoryItems(versions as any), [versions]);

  return (
    <div style={{ padding: 12, fontFamily: "system-ui, sans-serif", overflow: "auto" }}>
      <h3 style={{ marginTop: 0 }}>Version history</h3>

      {error ? <div style={{ color: "var(--error)", marginBottom: 8 }}>{error}</div> : null}

      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <button
          disabled={busy}
          onClick={async () => {
            const name = window.prompt("Checkpoint name?");
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
            if (!window.confirm("Restore this version? This will overwrite the current collaborative document state.")) return;
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
        <div style={{ color: "var(--text-secondary)" }}>No versions yet.</div>
      ) : (
        <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
          {items.map((item) => {
            const selected = item.id === selectedId;
            return (
              <li
                key={item.id}
                style={{
                  display: "flex",
                  gap: 8,
                  padding: "6px 0",
                  borderBottom: "1px solid var(--border)",
                  cursor: "pointer",
                  // Use theme tokens only (no hardcoded colors).
                  background: selected ? "var(--bg-hover)" : "transparent",
                }}
                onClick={() => setSelectedId(item.id)}
              >
                <input type="radio" checked={selected} onChange={() => setSelectedId(item.id)} />
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 600 }}>{item.title}</div>
                  <div style={{ color: "var(--text-secondary)", fontSize: 12 }}>
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

function CollabBranchManagerPanel({ session }: { session: CollabSession }) {
  const actor = useMemo<BranchActor>(() => ({ userId: "desktop", role: "owner" }), []);
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
        applyDocumentStateToYjsDoc(session.doc, state as any, { origin: session.origin });
        return state;
      },
      previewMerge: async (a: BranchActor, input: { sourceBranch: string }) => {
        await commitCurrentState("auto: preview merge");
        return branchService.previewMerge(a as any, input as any);
      },
      merge: async (a: BranchActor, input: { sourceBranch: string; resolutions: any[]; message?: string }) => {
        await commitCurrentState("auto: merge");
        const result = await branchService.merge(a as any, input as any);
        applyDocumentStateToYjsDoc(session.doc, (result as any).state, { origin: session.origin });
        return result;
      },
    } as any;
  }, [actor, branchService, session]);

  if (error) {
    return (
      <div style={{ padding: 12, color: "var(--error)" }}>
        {error}
      </div>
    );
  }

  if (!ready) {
    return <div style={{ padding: 12 }}>Loading branches…</div>;
  }

  if (mergeSource) {
    return (
      <MergeBranchPanel
        actor={actor}
        branchService={workflow}
        sourceBranch={mergeSource}
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
  onExecuteExtensionCommand?: (commandId: string) => void;
  onOpenExtensionPanel?: (panelId: string) => void;
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
    if (!marketplaceClient) {
      marketplaceClient = new MarketplaceClient({ baseUrl: getMarketplaceBaseUrl() });
    }

    if (!marketplaceExtensionManager) {
      marketplaceExtensionManager = new WebExtensionManager({
        marketplaceClient,
        host: (options.extensionHostManager?.host as any) ?? null,
      });
    }

    if (!marketplaceExtensionHostManager) {
      marketplaceExtensionHostManager = {
        syncInstalledExtensions: async () => {
          const manager = marketplaceExtensionManager!;
          const installed = await manager.listInstalled();
          for (const item of installed) {
            // Best-effort: ignore failures to load individual extensions so the UI
            // can keep going.
            try {
              if (manager.isLoaded(item.id)) continue;
              // eslint-disable-next-line no-await-in-loop
              await manager.loadInstalled(item.id);
            } catch {
              // ignore
            }
          }
        },
        reloadExtension: async (id: string) => {
          const manager = marketplaceExtensionManager!;
          if (manager.isLoaded(id)) {
            await manager.unload(id);
          }
          await manager.loadInstalled(id);
        },
        unloadExtension: async (id: string) => {
          const manager = marketplaceExtensionManager!;
          await manager.unload(id);
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
      makeBodyFillAvailableHeight(body);
      renderReactPanel(
        panelId,
        body,
        <ExtensionsPanel
          manager={options.extensionHostManager}
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
        <DataQueriesPanelContainer getDocumentController={options.getDocumentController} workbookId={options.workbookId} />,
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
      renderReactPanel(panelId, body, <CollabBranchManagerPanel session={session} />);
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
