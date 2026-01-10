import { LayoutController } from "./layoutController.js";

export class DesktopSessionController {
  /**
   * @param {{
   *   layoutWorkspaceManager: import("./layoutPersistence.js").LayoutWorkspaceManager,
   *   windowingController: import("./windowingController.js").WindowingController,
   *   primarySheetIdForWorkbook?: (workbookId: string) => string | null
   * }} params
   */
  constructor({ layoutWorkspaceManager, windowingController, primarySheetIdForWorkbook }) {
    this.layoutWorkspaceManager = layoutWorkspaceManager;
    this.windowingController = windowingController;
    this.primarySheetIdForWorkbook = primarySheetIdForWorkbook ?? (() => null);

    /** @type {Map<string, LayoutController>} */
    this.layoutsByWindowId = new Map();

    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();

    this.#syncFromWindowState();

    // Keep layout controllers aligned with window session state.
    this.windowingController.on("change", () => {
      this.#syncFromWindowState();
      this.#emit("change", { windows: this.windowingController.state.windows });
    });
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

  #syncFromWindowState() {
    const windows = this.windowingController.state.windows;
    const windowIds = new Set(windows.map((w) => w.id));

    // Remove controllers for closed windows.
    for (const [windowId, controller] of this.layoutsByWindowId.entries()) {
      if (windowIds.has(windowId)) continue;
      // Best-effort detach listeners by dropping the reference.
      this.layoutsByWindowId.delete(windowId);
      this.#emit("windowClosed", { windowId, controller });
    }

    // Ensure controllers exist and are up-to-date.
    for (const win of windows) {
      const existing = this.layoutsByWindowId.get(win.id);
      if (existing && existing.workbookId === win.workbookId && existing.workspaceId === win.workspaceId) {
        continue;
      }

      const controller = new LayoutController({
        workbookId: win.workbookId,
        workspaceManager: this.layoutWorkspaceManager,
        primarySheetId: this.primarySheetIdForWorkbook(win.workbookId),
        workspaceId: win.workspaceId,
      });

      this.layoutsByWindowId.set(win.id, controller);
      this.#emit("windowOpened", { windowId: win.id, controller });
    }
  }

  /**
   * @param {string} windowId
   */
  getLayoutController(windowId) {
    return this.layoutsByWindowId.get(windowId) ?? null;
  }

  openWorkbookWindow(workbookId, options = {}) {
    const windowId = this.windowingController.openWorkbookWindow(workbookId, options);
    // windowingController emits change -> sync will run, but do an immediate sync for callers.
    this.#syncFromWindowState();
    return windowId;
  }

  closeWindow(windowId) {
    this.windowingController.closeWindow(windowId);
    this.#syncFromWindowState();
  }

  focusWindow(windowId) {
    this.windowingController.focusWindow(windowId);
  }

  setWindowWorkspace(windowId, workspaceId) {
    this.windowingController.setWindowWorkspace(windowId, workspaceId);
    this.#syncFromWindowState();
  }
}

