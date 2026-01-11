/**
 * Persistence hooks for `RefreshManager` scheduling.
 *
 * The state store is owned by the host application (desktop/web) so schedules can
 * survive across sessions (Excel parity: refresh on open + scheduled refresh).
 */

/**
 * @typedef {import("./model.js").RefreshPolicy} RefreshPolicy
 */

/**
 * @typedef {{ policy: RefreshPolicy, lastRunAtMs?: number }} RefreshStateEntry
 */

/**
 * @typedef {{ [queryId: string]: RefreshStateEntry }} RefreshState
 */

/**
 * @typedef {{
 *   load(): Promise<RefreshState>;
 *   save(state: RefreshState): Promise<void>;
 * }} RefreshStateStore
 */

function cloneState(state) {
  if (typeof globalThis.structuredClone === "function") return globalThis.structuredClone(state);
  return JSON.parse(JSON.stringify(state));
}

export class InMemoryRefreshStateStore {
  /**
   * @param {RefreshState} [initialState]
   */
  constructor(initialState) {
    /** @type {RefreshState} */
    this.state = initialState ? cloneState(initialState) : {};
  }

  async load() {
    return cloneState(this.state);
  }

  /**
   * @param {RefreshState} state
   */
  async save(state) {
    this.state = cloneState(state);
  }
}

