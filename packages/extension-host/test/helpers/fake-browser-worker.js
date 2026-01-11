function createWorkerCtor(scenarios, errors) {
  const scenarioQueue = Array.isArray(scenarios) ? [...scenarios] : [];
  return class FakeWorker {
    constructor(_url, _options) {
      this._listeners = new Map();
      this._terminated = false;
      this._scenario = scenarioQueue.shift() ?? {};
    }

    addEventListener(type, listener) {
      const key = String(type);
      if (!this._listeners.has(key)) this._listeners.set(key, new Set());
      this._listeners.get(key).add(listener);
    }

    removeEventListener(type, listener) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      set.delete(listener);
      if (set.size === 0) this._listeners.delete(String(type));
    }

    postMessage(message) {
      if (this._terminated) return;
      try {
        this._scenario.onPostMessage?.(message, this);
      } catch (err) {
        errors?.push(err);
        this._emit("error", { message: String(err?.message ?? err) });
      }
    }

    terminate() {
      this._terminated = true;
    }

    emitMessage(message) {
      if (this._terminated) return;
      this._emit("message", { data: message });
    }

    _emit(type, event) {
      const set = this._listeners.get(String(type));
      if (!set) return;
      for (const listener of [...set]) {
        try {
          listener(event);
        } catch {
          // ignore
        }
      }
    }
  };
}

module.exports = {
  installFakeWorker(t, scenarios) {
    const errors = [];
    const PrevWorker = globalThis.Worker;
    globalThis.Worker = createWorkerCtor(scenarios, errors);
    t.after(() => {
      globalThis.Worker = PrevWorker;
      if (errors.length > 0) {
        throw errors[0];
      }
    });
  }
};
