import { deserializeWindowingState, serializeWindowingState } from "./windowingSerializer.js";
import {
  closeWindow,
  createDefaultWindowingState,
  focusWindow,
  getWindow,
  openWorkbookWindow,
  setWindowBounds,
  setWindowMaximized,
} from "./windowingState.js";

function clone(value) {
  return structuredClone(value);
}

export class WindowingController {
  /**
   * @param {{ sessionManager: import("./windowingPersistence.js").WindowingSessionManager }} params
   */
  constructor({ sessionManager }) {
    this.sessionManager = sessionManager;
    this.state = this.sessionManager.load();

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
   * @param {any} next
   * @param {{ persist?: boolean }} [options]
   */
  #commit(next, options = {}) {
    // Normalize by round-tripping through serializer logic (keeps in sync with future schema changes).
    this.state = deserializeWindowingState(serializeWindowingState(next));

    if (options.persist ?? true) {
      this.sessionManager.save(this.state);
    }

    this.#emit("change", { state: this.state });
  }

  reload() {
    this.state = this.sessionManager.load();
    this.#emit("change", { state: this.state });
  }

  clear() {
    this.sessionManager.clear();
    this.state = createDefaultWindowingState();
    this.#emit("change", { state: this.state });
  }

  /**
   * @param {string} workbookId
   * @param {{ workspaceId?: string, bounds?: any, maximized?: boolean, focus?: boolean }} [options]
   */
  openWorkbookWindow(workbookId, options = {}) {
    const next = openWorkbookWindow(this.state, workbookId, options);
    this.#commit(next);
    return this.state.focusedWindowId;
  }

  /**
   * @param {string} windowId
   */
  closeWindow(windowId) {
    this.#commit(closeWindow(this.state, windowId));
  }

  /**
   * @param {string} windowId
   */
  focusWindow(windowId) {
    this.#commit(focusWindow(this.state, windowId));
  }

  /**
   * @param {string} windowId
   * @param {{ x?: number | null, y?: number | null, width?: number, height?: number }} bounds
   */
  setWindowBounds(windowId, bounds) {
    this.#commit(setWindowBounds(this.state, windowId, bounds));
  }

  /**
   * @param {string} windowId
   * @param {boolean} maximized
   */
  setWindowMaximized(windowId, maximized) {
    this.#commit(setWindowMaximized(this.state, windowId, maximized));
  }

  /**
   * @param {string} windowId
   */
  getWindow(windowId) {
    const win = getWindow(this.state, windowId);
    return win ? clone(win) : null;
  }
}

