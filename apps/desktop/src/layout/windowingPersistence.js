import { createDefaultWindowingState } from "./windowingState.js";
import { deserializeWindowingState, serializeWindowingState } from "./windowingSerializer.js";

export class WindowingSessionManager {
  /**
   * @param {{ storage: Pick<Storage, "getItem" | "setItem" | "removeItem">, key?: string }} params
   */
  constructor({ storage, key = "formula.windowing.v1" }) {
    this.storage = storage;
    this.key = key;
  }

  load() {
    const raw = this.storage.getItem(this.key);
    if (!raw) return createDefaultWindowingState();
    return deserializeWindowingState(raw);
  }

  save(state) {
    this.storage.setItem(this.key, serializeWindowingState(state));
  }

  clear() {
    this.storage.removeItem(this.key);
  }
}

