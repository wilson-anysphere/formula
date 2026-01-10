import { deserializePresenceState, serializePresenceState } from "./presence.js";
import { throttle } from "./throttle.js";

export class PresenceManager {
  constructor(awareness, options) {
    const {
      user,
      activeSheet,
      throttleMs = 100,
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

  setActiveSheet(activeSheet) {
    if (!activeSheet) return;
    if (this.localPresence.activeSheet === activeSheet) return;
    this.localPresence.activeSheet = activeSheet;
    this.localPresence.lastActive = this.now();
    this._broadcastNow();
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
    if (!cursor) {
      this.localPresence.cursor = null;
      this.localPresence.lastActive = this.now();
      this._broadcastThrottled();
      return;
    }

    this.localPresence.cursor = { row: cursor.row, col: cursor.col };
    this.localPresence.lastActive = this.now();
    this._broadcastThrottled();
  }

  setSelections(selections) {
    this.localPresence.selections = Array.isArray(selections) ? selections : [];
    this.localPresence.lastActive = this.now();
    this._broadcastThrottled();
  }

  getRemotePresences({ activeSheet } = {}) {
    const targetSheet = activeSheet ?? this.localPresence.activeSheet;
    const states = this.awareness.getStates?.();
    if (!states) return [];

    const localClientId = this.awareness.clientID;
    const result = [];

    for (const [clientId, state] of states.entries()) {
      if (clientId === localClientId) continue;
      const presence = deserializePresenceState(state?.presence);
      if (!presence) continue;
      if (presence.activeSheet !== targetSheet) continue;
      result.push({ clientId, ...presence });
    }

    return result;
  }
}

