function createEmitter() {
  const listeners = new Set();
  return {
    on(listener) {
      listeners.add(listener);
    },
    off(listener) {
      listeners.delete(listener);
    },
    emit(change, origin) {
      for (const listener of listeners) listener(change, origin);
    },
  };
}

export class InMemoryAwarenessHub {
  constructor() {
    this._states = new Map();
    this._emitters = new Set();
    this._nextClientId = 1;
  }

  createAwareness(clientID = this._nextClientId++) {
    const emitter = createEmitter();
    this._emitters.add(emitter);

    const hub = this;
    const awareness = {
      clientID,
      getStates() {
        return hub._states;
      },
      getLocalState() {
        return hub._states.get(clientID) ?? null;
      },
      setLocalState(state) {
        if (state === null) {
          hub._update(clientID, null, awareness);
          return;
        }
        hub._update(clientID, { ...state }, awareness);
      },
      setLocalStateField(field, value) {
        const current = hub._states.get(clientID) ?? {};
        const next = { ...current, [field]: value };
        hub._update(clientID, next, awareness);
      },
      on(eventName, handler) {
        if (eventName !== "change") return;
        emitter.on(handler);
      },
      off(eventName, handler) {
        if (eventName !== "change") return;
        emitter.off(handler);
      },
      destroy() {
        hub._update(clientID, null, awareness);
        hub._emitters.delete(emitter);
      },
    };

    return awareness;
  }

  _update(clientID, nextState, origin) {
    const hadState = this._states.has(clientID);

    if (nextState === null) {
      if (!hadState) return;
      this._states.delete(clientID);
      const change = { added: [], updated: [], removed: [clientID] };
      for (const emitter of this._emitters) emitter.emit(change, origin);
      return;
    }

    this._states.set(clientID, nextState);
    const change = hadState
      ? { added: [], updated: [clientID], removed: [] }
      : { added: [clientID], updated: [], removed: [] };
    for (const emitter of this._emitters) emitter.emit(change, origin);
  }
}

