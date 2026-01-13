import { deserializePresenceState, serializePresenceState } from "./presence.js";
import { throttle } from "./throttle.js";

function normalizeRange(range) {
  if (!range || typeof range !== "object") return null;

  let startRow;
  let startCol;
  let endRow;
  let endCol;

  if (
    typeof range.startRow === "number" &&
    typeof range.startCol === "number" &&
    typeof range.endRow === "number" &&
    typeof range.endCol === "number"
  ) {
    startRow = range.startRow;
    startCol = range.startCol;
    endRow = range.endRow;
    endCol = range.endCol;
  } else if (
    range.start &&
    range.end &&
    typeof range.start.row === "number" &&
    typeof range.start.col === "number" &&
    typeof range.end.row === "number" &&
    typeof range.end.col === "number"
  ) {
    startRow = range.start.row;
    startCol = range.start.col;
    endRow = range.end.row;
    endCol = range.end.col;
  } else {
    return null;
  }

  const normalizedStartRow = Math.min(startRow, endRow);
  const normalizedEndRow = Math.max(startRow, endRow);
  const normalizedStartCol = Math.min(startCol, endCol);
  const normalizedEndCol = Math.max(startCol, endCol);

  return {
    startRow: Math.trunc(normalizedStartRow),
    startCol: Math.trunc(normalizedStartCol),
    endRow: Math.trunc(normalizedEndRow),
    endCol: Math.trunc(normalizedEndCol),
  };
}

function areRangesEqual(a, b) {
  return (
    a.startRow === b.startRow && a.startCol === b.startCol && a.endRow === b.endRow && a.endCol === b.endCol
  );
}

function areSelectionsEqual(a, b) {
  if (a === b) return true;
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (!areRangesEqual(a[i], b[i])) return false;
  }
  return true;
}

export class PresenceManager {
  constructor(awareness, options) {
    const {
      user,
      activeSheet,
      throttleMs = 100,
      staleAfterMs,
      now = () => Date.now(),
      setTimeout: setTimeoutFn,
      clearTimeout: clearTimeoutFn,
    } = options ?? {};

    if (!awareness) throw new Error("PresenceManager requires an awareness instance");
    if (!user?.id) throw new Error("PresenceManager requires user.id");
    if (!user?.name) throw new Error("PresenceManager requires user.name");
    if (!user?.color) throw new Error("PresenceManager requires user.color");
    if (!activeSheet) throw new Error("PresenceManager requires activeSheet");

    this.awareness = awareness;
    this.now = now;
    this.staleAfterMs =
      typeof staleAfterMs === "number" && Number.isFinite(staleAfterMs) && staleAfterMs >= 0 ? staleAfterMs : null;
    this._setTimeout = setTimeoutFn ?? globalThis.setTimeout;
    this._clearTimeout = clearTimeoutFn ?? globalThis.clearTimeout;
    this._staleEvictionTimeoutId = null;
    /** @type {Map<Function, { includeOtherSheets: boolean }>} */
    this._listeners = new Map();
    this._awarenessChangeHandler = (change) => {
      const localClientId = this.awareness.clientID;
      const added = Array.isArray(change?.added) ? change.added : [];
      const updated = Array.isArray(change?.updated) ? change.updated : [];
      const removed = Array.isArray(change?.removed) ? change.removed : [];
      const hasRemoteChange = [...added, ...updated, ...removed].some((clientId) => clientId !== localClientId);
      if (!hasRemoteChange) return;
      this._notify();
    };
    this._awarenessListenerAttached = false;

    this.localPresence = {
      id: user.id,
      name: user.name,
      color: user.color,
      activeSheet,
      cursor: null,
      selections: [],
      lastActive: now(),
    };

    this._broadcastNow();

    this._broadcastThrottled = throttle(
      () => {
        this._broadcastNow();
      },
      throttleMs,
      { now, setTimeout: setTimeoutFn, clearTimeout: clearTimeoutFn },
    );
  }

  _broadcastNow() {
    this.awareness.setLocalStateField("presence", serializePresenceState(this.localPresence));
  }

  _clearStaleEvictionTimer() {
    if (this._staleEvictionTimeoutId === null) return;
    this._clearTimeout(this._staleEvictionTimeoutId);
    this._staleEvictionTimeoutId = null;
  }

  _scheduleStaleEviction(presences) {
    this._clearStaleEvictionTimer();

    if (this._listeners.size === 0) return;
    if (this.staleAfterMs === null) return;
    if (!Array.isArray(presences) || presences.length === 0) return;

    const now = this.now();
    let nextEvictAt = Infinity;

    for (const presence of presences) {
      const lastActive = presence?.lastActive;
      if (typeof lastActive !== "number" || !Number.isFinite(lastActive)) continue;
      const evictAt = lastActive + this.staleAfterMs + 1;
      if (evictAt < nextEvictAt) nextEvictAt = evictAt;
    }

    if (!Number.isFinite(nextEvictAt)) return;

    const delayMs = Math.max(0, nextEvictAt - now);
    if (typeof this._setTimeout !== "function") return;

    this._staleEvictionTimeoutId = this._setTimeout(() => {
      this._staleEvictionTimeoutId = null;
      this._notify();
    }, delayMs);
    this._staleEvictionTimeoutId?.unref?.();
  }

  _getRemotePresenceSnapshot() {
    const allPresences = this.getRemotePresences({ includeOtherSheets: true });
    const activeSheet = this.localPresence.activeSheet;
    const presences = allPresences.filter((presence) => presence.activeSheet === activeSheet);
    return { allPresences, presences };
  }

  _notify() {
    if (this._listeners.size === 0) return;
    const { allPresences, presences } = this._getRemotePresenceSnapshot();
    // Schedule stale eviction across *all* remote presences (not just the current active
    // sheet). This ensures:
    // - `subscribe(..., { includeOtherSheets: true })` evicts stale users on non-active sheets
    // - legacy consumers calling `getRemotePresences({ includeOtherSheets: true })` inside a
    //   default subscription callback still see stale users removed, even when there are no
    //   active-sheet users.
    //
    // Schedule before calling listeners so a listener cannot affect eviction timing by
    // mutating the presences array.
    this._scheduleStaleEviction(allPresences);
    for (const [listener, opts] of this._listeners) listener(opts?.includeOtherSheets ? allPresences : presences);
  }

  /**
   * Subscribe to remote presence changes.
   *
   * The callback is called immediately with the current remote presences, and
   * then on any remote awareness update. Local cursor/selection updates are
   * ignored to avoid causing unnecessary re-renders during pointer movement.
   *
   * @param {(presences: any[]) => void} listener
   * @param {{ includeOtherSheets?: boolean }=} opts
   * @returns {() => void}
   */
  subscribe(listener, opts) {
    const includeOtherSheets = opts?.includeOtherSheets === true;
    this._listeners.set(listener, { includeOtherSheets });

    if (!this._awarenessListenerAttached && typeof this.awareness.on === "function") {
      this.awareness.on("change", this._awarenessChangeHandler);
      this._awarenessListenerAttached = true;
    }

    const { allPresences, presences } = this._getRemotePresenceSnapshot();
    this._scheduleStaleEviction(allPresences);
    listener(includeOtherSheets ? allPresences : presences);

    return () => {
      this._listeners.delete(listener);
      if (this._listeners.size > 0) return;
      if (this._awarenessListenerAttached && typeof this.awareness.off === "function") {
        this.awareness.off("change", this._awarenessChangeHandler);
      }
      this._awarenessListenerAttached = false;
      this._clearStaleEvictionTimer();
    };
  }

  destroy() {
    this._broadcastThrottled?.cancel?.();
    this._clearStaleEvictionTimer();
    if (this._awarenessListenerAttached && typeof this.awareness.off === "function") {
      this.awareness.off("change", this._awarenessChangeHandler);
      this._awarenessListenerAttached = false;
    }
    this._listeners.clear();
    if (typeof this.awareness.setLocalState === "function") {
      this.awareness.setLocalState(null);
      return;
    }
    this.awareness.setLocalStateField("presence", null);
  }

  setActiveSheet(activeSheet) {
    if (!activeSheet) return;
    if (this.localPresence.activeSheet === activeSheet) return;
    this.localPresence.activeSheet = activeSheet;
    this.localPresence.lastActive = this.now();
    this._broadcastNow();
    this._notify();
  }

  setUser(user) {
    if (!user?.id || !user?.name || !user?.color) return;
    this.localPresence.id = user.id;
    this.localPresence.name = user.name;
    this.localPresence.color = user.color;
    this.localPresence.lastActive = this.now();
    this._broadcastNow();
  }

  setCursor(cursor) {
    const nextCursor = cursor ? { row: Math.trunc(cursor.row), col: Math.trunc(cursor.col) } : null;
    const prevCursor = this.localPresence.cursor;

    if (prevCursor === null && nextCursor === null) return;
    if (prevCursor && nextCursor && prevCursor.row === nextCursor.row && prevCursor.col === nextCursor.col)
      return;

    this.localPresence.cursor = nextCursor;
    this.localPresence.lastActive = this.now();
    this._broadcastThrottled();
  }

  setSelections(selections) {
    const normalizedSelections = Array.isArray(selections)
      ? selections
          .map((range) => normalizeRange(range))
          .filter((range) => range !== null)
          .sort((a, b) => {
            if (a.startRow !== b.startRow) return a.startRow - b.startRow;
            if (a.startCol !== b.startCol) return a.startCol - b.startCol;
            if (a.endRow !== b.endRow) return a.endRow - b.endRow;
            return a.endCol - b.endCol;
          })
      : [];

    if (areSelectionsEqual(normalizedSelections, this.localPresence.selections)) return;

    this.localPresence.selections = normalizedSelections;
    this.localPresence.lastActive = this.now();
    this._broadcastThrottled();
  }

  getRemotePresences({ activeSheet, includeOtherSheets = false, staleAfterMs } = {}) {
    const targetSheet = activeSheet ?? this.localPresence.activeSheet;
    const states = this.awareness.getStates?.();
    if (!states) return [];

    const localClientId = this.awareness.clientID;
    const result = [];
    const staleAfterMsEffective = staleAfterMs ?? this.staleAfterMs;
    const cutoff =
      typeof staleAfterMsEffective === "number" && Number.isFinite(staleAfterMsEffective) && staleAfterMsEffective >= 0
        ? this.now() - staleAfterMsEffective
        : null;

    for (const [clientId, state] of states.entries()) {
      if (clientId === localClientId) continue;
      const presence = deserializePresenceState(state?.presence);
      if (!presence) continue;
      if (cutoff !== null && presence.lastActive < cutoff) continue;
      if (!includeOtherSheets && presence.activeSheet !== targetSheet) continue;
      result.push({ clientId, ...presence });
    }

    result.sort((a, b) => {
      if (a.id !== b.id) return a.id < b.id ? -1 : 1;
      return a.clientId - b.clientId;
    });

    return result;
  }
}
