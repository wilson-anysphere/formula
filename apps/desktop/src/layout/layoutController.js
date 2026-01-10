import { normalizeLayout } from "./layoutNormalization.js";
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
  setSplitRatio,
  snapFloatingPanel,
} from "./layoutState.js";

export class LayoutController {
  /**
   * @param {{ workbookId: string, workspaceManager: import("./layoutPersistence.js").LayoutWorkspaceManager, primarySheetId?: string | null }} params
   */
  constructor({ workbookId, workspaceManager, primarySheetId = null }) {
    if (typeof workbookId !== "string" || workbookId.length === 0) {
      throw new Error("workbookId must be a non-empty string");
    }

    this.workbookId = workbookId;
    this.workspaceManager = workspaceManager;
    this.primarySheetId = primarySheetId;

    /** @type {ReturnType<typeof import("./layoutState.js").createDefaultLayout>} */
    this.layout = this.workspaceManager.loadWorkbookLayout(this.workbookId, { primarySheetId });

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
   * @param {{ persist?: boolean }} [options]
   */
  #commit(nextLayout, options = {}) {
    this.layout = normalizeLayout(nextLayout, {
      panelRegistry: this.workspaceManager.panelRegistry,
      primarySheetId: this.primarySheetId,
    });

    if (options.persist ?? true) {
      this.workspaceManager.saveWorkbookLayout(this.workbookId, this.layout);
    }

    this.#emit("change", { layout: this.layout });
  }

  /**
   * Reload the active workspace layout from persistence (discarding in-memory changes).
   */
  reload() {
    this.layout = this.workspaceManager.loadWorkbookLayout(this.workbookId, { primarySheetId: this.primarySheetId });
    this.#emit("change", { layout: this.layout });
  }

  /**
   * @returns {string}
   */
  get activeWorkspaceId() {
    return this.workspaceManager.getActiveWorkbookWorkspaceId(this.workbookId);
  }

  listWorkspaces() {
    return this.workspaceManager.listWorkbookWorkspaces(this.workbookId);
  }

  /**
   * @param {string} workspaceId
   */
  setActiveWorkspace(workspaceId) {
    this.workspaceManager.setActiveWorkbookWorkspace(this.workbookId, workspaceId);
    this.reload();
    this.#emit("workspace", { workspaceId, layout: this.layout });
  }

  /**
   * @param {string} workspaceId
   * @param {{ name?: string, makeActive?: boolean }} [options]
   */
  saveWorkspace(workspaceId, options = {}) {
    this.workspaceManager.saveWorkbookWorkspace(this.workbookId, workspaceId, {
      name: options.name,
      layout: this.layout,
      makeActive: options.makeActive,
    });
    if (options.makeActive) {
      this.reload();
      this.#emit("workspace", { workspaceId: this.activeWorkspaceId, layout: this.layout });
    }
  }

  /**
   * @param {string} workspaceId
   */
  deleteWorkspace(workspaceId) {
    this.workspaceManager.deleteWorkbookWorkspace(this.workbookId, workspaceId);
    if (this.activeWorkspaceId === "default") {
      this.reload();
      this.#emit("workspace", { workspaceId: "default", layout: this.layout });
    }
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

  setSplitDirection(direction, ratio) {
    this.#commit(setSplitDirection(this.layout, direction, ratio));
  }

  setSplitRatio(ratio) {
    this.#commit(setSplitRatio(this.layout, ratio));
  }

  setActiveSplitPane(pane) {
    this.#commit(setActiveSplitPane(this.layout, pane));
  }

  setSplitPaneSheet(pane, sheetId) {
    this.#commit(setSplitPaneSheet(this.layout, pane, sheetId));
  }

  setSplitPaneScroll(pane, scroll) {
    this.#commit(setSplitPaneScroll(this.layout, pane, scroll));
  }

  saveAsGlobalDefault() {
    this.workspaceManager.saveGlobalDefaultLayout(this.layout);
    this.#emit("globalDefaultSaved", {});
  }
}
