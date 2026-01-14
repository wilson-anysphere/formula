// Vitest shim for the `y-websocket` dependency.
//
// Some CI/dev environments can run with cached/stale `node_modules` where transitive deps are
// missing. `@formula/collab-session` imports `y-websocket` to provide a default `WebsocketProvider`
// implementation. For desktop/unit tests that don't actually establish WebSocket connections, a
// lightweight stub keeps the module graph loadable.

type Listener = (...args: any[]) => void;

class TinyEmitter {
  private readonly listeners = new Map<string, Set<Listener>>();

  on(event: string, cb: Listener): void {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(cb);
  }

  off(event: string, cb: Listener): void {
    this.listeners.get(event)?.delete(cb);
  }

  emit(event: string, ...args: any[]): void {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const cb of Array.from(set)) {
      try {
        cb(...args);
      } catch {
        // ignore
      }
    }
  }
}

function createStubAwareness(): any {
  const emitter = new TinyEmitter();
  const states = new Map<number, any>();
  return {
    getLocalState: () => states.get(0) ?? null,
    setLocalState: (state: any) => {
      states.set(0, state);
      emitter.emit("change", [], []);
    },
    setLocalStateField: (key: string, value: any) => {
      const current = states.get(0) ?? {};
      states.set(0, { ...(current as any), [key]: value });
      emitter.emit("change", [], []);
    },
    getStates: () => states,
    on: (event: string, cb: Listener) => emitter.on(event, cb),
    off: (event: string, cb: Listener) => emitter.off(event, cb),
  };
}

export class WebsocketProvider {
  readonly awareness: any;
  synced = false;

  // `@formula/collab-session` inspects `provider.ws` to attach close listeners.
  ws: any = null;

  private readonly emitter = new TinyEmitter();

  constructor(_serverUrl: string, _roomName: string, _doc: any, _opts?: any) {
    this.awareness = createStubAwareness();
  }

  on(event: string, cb: Listener): void {
    this.emitter.on(event, cb);
  }

  off(event: string, cb: Listener): void {
    this.emitter.off(event, cb);
  }

  connect(): void {
    // no-op
  }

  disconnect(): void {
    // no-op
  }

  destroy(): void {
    // no-op
  }
}

