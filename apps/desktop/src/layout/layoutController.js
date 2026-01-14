import { normalizeLayout } from "./layoutNormalization.js";
import { clampZoom } from "@formula/grid/node";
import {
  activateDockedPanel,
  closePanel,
  dockPanel,
  floatPanel,
  openPanel,
  setActiveSplitPane,
  setDockCollapsed,
  setDockSize,
  setFloatingPanelMinimized,
  setFloatingPanelRect,
  setSplitDirection,
  setSplitPaneScroll,
  setSplitPaneSheet,
  setSplitPaneZoom,
  setSplitRatio,
  snapFloatingPanel,
} from "./layoutState.js";

export class LayoutController {
  /** @type {boolean} */
  #needsPersist = false;

  /**
   * @param {{ workbookId: string, workspaceManager: import("./layoutPersistence.js").LayoutWorkspaceManager, primarySheetId?: string | null, workspaceId?: string }} params
   */
  constructor({ workbookId, workspaceManager, primarySheetId = null, workspaceId }) {
    if (typeof workbookId !== "string" || workbookId.length === 0) {
      throw new Error("workbookId must be a non-empty string");
    }

    this.workbookId = workbookId;
    this.workspaceManager = workspaceManager;
    this.primarySheetId = primarySheetId;
    this.workspaceId = typeof workspaceId === "string" && workspaceId.length > 0 ? workspaceId : this.workspaceManager.getActiveWorkbookWorkspaceId(this.workbookId);

    /** @type {ReturnType<typeof import("./layoutState.js").createDefaultLayout>} */
    this.layout = this.workspaceManager.loadWorkbookLayoutForWorkspace(this.workbookId, this.workspaceId, {
      primarySheetId,
    });

    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
  }

  /**
   * @template {string} T
   * @param {T} event
   * @param {(payload: any) => void} listener
   * @returns {() => void}
   */
  on(event, listener) {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(listener);
    return () => set.delete(listener);
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  #emit(event, payload) {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }

  /**
   * @param {any} nextLayout
   * @param {{ persist?: boolean, emit?: boolean }} [options]
   */
  #commit(nextLayout, options = {}) {
    this.layout = normalizeLayout(nextLayout, {
      panelRegistry: this.workspaceManager.panelRegistry,
      primarySheetId: this.primarySheetId,
    });

    if (options.persist ?? true) {
      this.workspaceManager.saveWorkbookLayoutForWorkspace(this.workbookId, this.workspaceId, this.layout);
      this.#needsPersist = false;
    } else {
      this.#needsPersist = true;
    }

    if (options.emit ?? true) {
      this.#emit("change", { layout: this.layout });
    }
  }

  /**
   * Persist the current in-memory layout for the active workspace without emitting
   * an additional "change" event. Useful for debounced flushes after applying
   * ephemeral (non-persisted) updates at high frequency.
   */
  persistNow() {
    this.workspaceManager.saveWorkbookLayoutForWorkspace(this.workbookId, this.workspaceId, this.layout);
    this.#needsPersist = false;
  }

  /**
   * Alias for persistNow().
   */
  save() {
    this.persistNow();
  }

  /**
   * Reload the active workspace layout from persistence (discarding in-memory changes).
   */
  reload() {
    this.layout = this.workspaceManager.loadWorkbookLayoutForWorkspace(this.workbookId, this.workspaceId, {
      primarySheetId: this.primarySheetId,
    });
    this.#needsPersist = false;
    this.#emit("change", { layout: this.layout });
  }

  /**
   * @returns {string}
   */
  get activeWorkspaceId() {
    return this.workspaceId;
  }

  listWorkspaces() {
    return this.workspaceManager.listWorkbookWorkspaces(this.workbookId);
  }

  /**
   * @param {string} workspaceId
   */
  setActiveWorkspace(workspaceId) {
    if (this.#needsPersist) this.persistNow();
    this.workspaceManager.setActiveWorkbookWorkspace(this.workbookId, workspaceId);
    this.workspaceId = this.workspaceManager.getActiveWorkbookWorkspaceId(this.workbookId);
    this.reload();
    this.#emit("workspace", { workspaceId, layout: this.layout });
  }

  /**
   * Switch the controller to a specific workspace id without updating the workbook's "active workspace"
   * pointer. Useful for multi-window scenarios where two windows show the same workbook with different
   * workspace layouts.
   *
   * @param {string} workspaceId
   */
  setWorkspace(workspaceId) {
    if (this.#needsPersist) this.persistNow();
    this.workspaceId = typeof workspaceId === "string" && workspaceId.length > 0 ? workspaceId : "default";
    this.reload();
    this.#emit("workspace", { workspaceId: this.workspaceId, layout: this.layout });
  }

  /**
   * @param {string} workspaceId
   * @param {{ name?: string, makeActive?: boolean }} [options]
   */
  saveWorkspace(workspaceId, options = {}) {
    if (options.makeActive && this.#needsPersist) this.persistNow();
    this.workspaceManager.saveWorkbookWorkspace(this.workbookId, workspaceId, {
      name: options.name,
      layout: this.layout,
      makeActive: options.makeActive,
    });
    if (options.makeActive) {
      this.workspaceId = this.workspaceManager.getActiveWorkbookWorkspaceId(this.workbookId);
      this.reload();
      this.#emit("workspace", { workspaceId: this.workspaceId, layout: this.layout });
    }
  }

  /**
   * @param {string} workspaceId
   */
  deleteWorkspace(workspaceId) {
    // Deleting a workspace triggers a reload. If the deletion does not affect the
    // currently-active workspace, flush any pending ephemeral updates first so
    // we don't discard them during reload.
    if (this.#needsPersist && this.workspaceId !== workspaceId) this.persistNow();
    this.workspaceManager.deleteWorkbookWorkspace(this.workbookId, workspaceId);
    if (this.workspaceId === workspaceId) {
      this.workspaceId = this.workspaceManager.getActiveWorkbookWorkspaceId(this.workbookId);
    }
    this.reload();
    this.#emit("workspace", { workspaceId: this.workspaceId, layout: this.layout });
  }

  openPanel(panelId) {
    this.#commit(openPanel(this.layout, panelId, { panelRegistry: this.workspaceManager.panelRegistry }));
  }

  closePanel(panelId) {
    this.#commit(closePanel(this.layout, panelId));
  }

  dockPanel(panelId, side, options) {
    this.#commit(dockPanel(this.layout, panelId, side, options));
  }

  activateDockedPanel(panelId, side) {
    this.#commit(activateDockedPanel(this.layout, panelId, side));
  }

  floatPanel(panelId, rect, options) {
    this.#commit(floatPanel(this.layout, panelId, rect, options));
  }

  setFloatingPanelRect(panelId, rect) {
    this.#commit(setFloatingPanelRect(this.layout, panelId, rect));
  }

  setFloatingPanelMinimized(panelId, minimized) {
    this.#commit(setFloatingPanelMinimized(this.layout, panelId, minimized));
  }

  snapFloatingPanel(panelId, viewport, options) {
    this.#commit(snapFloatingPanel(this.layout, panelId, viewport, options));
  }

  setDockCollapsed(side, collapsed) {
    this.#commit(setDockCollapsed(this.layout, side, collapsed));
  }

  setDockSize(side, sizePx) {
    this.#commit(setDockSize(this.layout, side, sizePx));
  }

  setSplitDirection(direction, ratio, options) {
    // Backwards compatible signature + convenience: allow omitting ratio and passing options
    // as the second argument: setSplitDirection(direction, { persist: false }).
    if (ratio !== null && typeof ratio === "object") {
      options = ratio;
      ratio = undefined;
    }
    this.#commit(setSplitDirection(this.layout, direction, ratio), options);
  }

  setSplitRatio(ratio, options) {
    const persist = options?.persist ?? true;
    const emit = options?.emit ?? true;

    // Split ratio updates can be extremely high frequency (splitter drag). For ephemeral,
    // non-emitting updates we can safely apply the change directly without running full
    // normalization/persistence logic on every tick.
    if (!persist && !emit) {
      if (typeof ratio !== "number" || !Number.isFinite(ratio)) return;
      const clamped = Math.max(0.1, Math.min(0.9, ratio));
      if (this.layout?.splitView?.ratio === clamped) return;
      this.layout.splitView.ratio = clamped;
      this.#needsPersist = true;
      return;
    }

    this.#commit(setSplitRatio(this.layout, ratio), options);
  }

  setActiveSplitPane(pane) {
    this.#commit(setActiveSplitPane(this.layout, pane));
  }

  setSplitPaneSheet(pane, sheetId) {
    this.#commit(setSplitPaneSheet(this.layout, pane, sheetId));
  }

  setSplitPaneScroll(pane, scroll, options) {
    const persist = options?.persist ?? true;
    const emit = options?.emit ?? true;

    if (!persist && !emit) {
      if (pane !== "primary" && pane !== "secondary") {
        throw new Error(`Invalid split pane: ${pane}`);
      }
      const existing = this.layout?.splitView?.panes?.[pane];
      if (!existing) return;

      const clampScroll = (value) => Math.max(-1e12, Math.min(1e12, value));
      const nextX =
        typeof scroll?.scrollX === "number" && Number.isFinite(scroll.scrollX)
          ? clampScroll(scroll.scrollX)
          : existing.scrollX;
      const nextY =
        typeof scroll?.scrollY === "number" && Number.isFinite(scroll.scrollY)
          ? clampScroll(scroll.scrollY)
          : existing.scrollY;
      if (existing.scrollX === nextX && existing.scrollY === nextY) return;

      existing.scrollX = nextX;
      existing.scrollY = nextY;
      this.#needsPersist = true;
      return;
    }

    this.#commit(setSplitPaneScroll(this.layout, pane, scroll), options);
  }

  setSplitPaneZoom(pane, zoom, options) {
    const persist = options?.persist ?? true;
    const emit = options?.emit ?? true;

    if (!persist && !emit) {
      if (pane !== "primary" && pane !== "secondary") {
        throw new Error(`Invalid split pane: ${pane}`);
      }
      const existing = this.layout?.splitView?.panes?.[pane];
      if (!existing) return;

      const rawZoom = typeof zoom === "number" && Number.isFinite(zoom) ? zoom : 1;
      const clamped = clampZoom(rawZoom);
      if (existing.zoom === clamped) return;

      existing.zoom = clamped;
      this.#needsPersist = true;
      return;
    }

    this.#commit(setSplitPaneZoom(this.layout, pane, zoom), options);
  }

  saveAsGlobalDefault() {
    this.workspaceManager.saveGlobalDefaultLayout(this.layout);
    this.#emit("globalDefaultSaved", {});
  }
}
