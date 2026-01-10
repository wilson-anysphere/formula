import { createDefaultLayout } from "./layoutState.js";
import { deserializeLayout, serializeLayout } from "./layoutSerializer.js";

export class MemoryStorage {
  /** @type {Map<string, string>} */
  #items = new Map();

  getItem(key) {
    return this.#items.has(key) ? this.#items.get(key) : null;
  }

  setItem(key, value) {
    this.#items.set(key, String(value));
  }

  removeItem(key) {
    this.#items.delete(key);
  }

  clear() {
    this.#items.clear();
  }
}

function encodeKeyPart(value) {
  return encodeURIComponent(String(value));
}

export class LayoutWorkspaceManager {
  /**
   * @param {{ storage: Pick<Storage, "getItem" | "setItem" | "removeItem">, panelRegistry?: Record<string, unknown>, keyPrefix?: string }} params
   */
  constructor({ storage, panelRegistry, keyPrefix = "formula.layout" }) {
    this.storage = storage;
    this.panelRegistry = panelRegistry;
    this.keyPrefix = keyPrefix;
  }

  globalKey() {
    return `${this.keyPrefix}.global.v1`;
  }

  workbookKey(workbookId) {
    return `${this.keyPrefix}.workbook.${encodeKeyPart(workbookId)}.v1`;
  }

  /**
   * @param {{ primarySheetId?: string | null }} [options]
   */
  loadGlobalDefaultLayout(options = {}) {
    const raw = this.storage.getItem(this.globalKey());
    if (!raw) return null;
    return deserializeLayout(raw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
  }

  saveGlobalDefaultLayout(layout) {
    this.storage.setItem(this.globalKey(), serializeLayout(layout, { panelRegistry: this.panelRegistry }));
  }

  deleteGlobalDefaultLayout() {
    this.storage.removeItem(this.globalKey());
  }

  /**
   * @param {string} workbookId
   * @param {{ primarySheetId?: string | null }} [options]
   */
  loadWorkbookLayout(workbookId, options = {}) {
    const raw = this.storage.getItem(this.workbookKey(workbookId));
    if (raw) {
      return deserializeLayout(raw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
    }

    const global = this.loadGlobalDefaultLayout({ primarySheetId: options.primarySheetId });
    if (global) return global;

    return createDefaultLayout({ primarySheetId: options.primarySheetId });
  }

  saveWorkbookLayout(workbookId, layout) {
    this.storage.setItem(
      this.workbookKey(workbookId),
      serializeLayout(layout, { panelRegistry: this.panelRegistry }),
    );
  }

  deleteWorkbookLayout(workbookId) {
    this.storage.removeItem(this.workbookKey(workbookId));
  }
}

