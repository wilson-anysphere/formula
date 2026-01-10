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

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function normalizeWorkspaceIndex(raw) {
  const active = typeof raw?.activeWorkspaceId === "string" ? raw.activeWorkspaceId : "default";
  const workspaces = raw?.workspaces && typeof raw.workspaces === "object" ? raw.workspaces : {};

  /** @type {Record<string, { name: string }>} */
  const normalizedWorkspaces = {};
  for (const [id, meta] of Object.entries(workspaces)) {
    if (typeof id !== "string" || id.length === 0 || id === "default") continue;
    const name = typeof meta?.name === "string" && meta.name.trim().length > 0 ? meta.name.trim() : id;
    normalizedWorkspaces[id] = { name };
  }

  const normalizedActive = active === "default" || Object.prototype.hasOwnProperty.call(normalizedWorkspaces, active)
    ? active
    : "default";

  return { schemaVersion: 1, activeWorkspaceId: normalizedActive, workspaces: normalizedWorkspaces };
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

  workbookWorkspaceIndexKey(workbookId) {
    return `${this.keyPrefix}.workbook.${encodeKeyPart(workbookId)}.workspaces.v1`;
  }

  workbookWorkspaceKey(workbookId, workspaceId) {
    return `${this.keyPrefix}.workbook.${encodeKeyPart(workbookId)}.workspace.${encodeKeyPart(workspaceId)}.v1`;
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
    const activeWorkspaceId = this.getActiveWorkbookWorkspaceId(workbookId);
    return this.loadWorkbookLayoutForWorkspace(workbookId, activeWorkspaceId, options);
  }

  saveWorkbookLayout(workbookId, layout) {
    const activeWorkspaceId = this.getActiveWorkbookWorkspaceId(workbookId);
    this.saveWorkbookLayoutForWorkspace(workbookId, activeWorkspaceId, layout);
  }

  deleteWorkbookLayout(workbookId) {
    this.storage.removeItem(this.workbookKey(workbookId));
  }

  /**
   * Load a specific workspace layout (without modifying the workbook's "active workspace" pointer).
   *
   * If the requested workspace does not exist yet, the default workspace (and then the global default)
   * is used as a fallback so new workspaces can start from a sensible base.
   *
   * @param {string} workbookId
   * @param {string} workspaceId
   * @param {{ primarySheetId?: string | null }} [options]
   */
  loadWorkbookLayoutForWorkspace(workbookId, workspaceId, options = {}) {
    const id = typeof workspaceId === "string" && workspaceId.length > 0 ? workspaceId : "default";
    const key = id === "default" ? this.workbookKey(workbookId) : this.workbookWorkspaceKey(workbookId, id);

    const raw = this.storage.getItem(key);
    if (raw) {
      return deserializeLayout(raw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
    }

    if (id !== "default") {
      const fallbackRaw = this.storage.getItem(this.workbookKey(workbookId));
      if (fallbackRaw) {
        return deserializeLayout(fallbackRaw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
      }
    }

    const global = this.loadGlobalDefaultLayout({ primarySheetId: options.primarySheetId });
    if (global) return global;

    return createDefaultLayout({ primarySheetId: options.primarySheetId });
  }

  /**
   * Save a layout to a specific workspace id (without changing the workbook's active workspace).
   *
   * @param {string} workbookId
   * @param {string} workspaceId
   * @param {any} layout
   */
  saveWorkbookLayoutForWorkspace(workbookId, workspaceId, layout) {
    const id = typeof workspaceId === "string" && workspaceId.length > 0 ? workspaceId : "default";

    if (id === "default") {
      this.storage.setItem(this.workbookKey(workbookId), serializeLayout(layout, { panelRegistry: this.panelRegistry }));
      return;
    }

    const index = this.loadWorkbookWorkspaceIndex(workbookId);
    const existingName = index.workspaces?.[id]?.name;

    this.saveWorkbookWorkspace(workbookId, id, {
      name: existingName,
      layout,
      makeActive: false,
    });
  }

  loadWorkbookWorkspaceIndex(workbookId) {
    const raw = this.storage.getItem(this.workbookWorkspaceIndexKey(workbookId));
    if (!raw) return normalizeWorkspaceIndex(null);

    return normalizeWorkspaceIndex(safeJsonParse(raw));
  }

  saveWorkbookWorkspaceIndex(workbookId, index) {
    this.storage.setItem(this.workbookWorkspaceIndexKey(workbookId), JSON.stringify(normalizeWorkspaceIndex(index)));
  }

  getActiveWorkbookWorkspaceId(workbookId) {
    return this.loadWorkbookWorkspaceIndex(workbookId).activeWorkspaceId;
  }

  /**
   * @param {string} workbookId
   * @returns {Array<{ id: string, name: string, active: boolean, isDefault: boolean }>}
   */
  listWorkbookWorkspaces(workbookId) {
    const index = this.loadWorkbookWorkspaceIndex(workbookId);
    const list = [
      { id: "default", name: "Default", active: index.activeWorkspaceId === "default", isDefault: true },
      ...Object.entries(index.workspaces).map(([id, meta]) => ({
        id,
        name: meta.name,
        active: index.activeWorkspaceId === id,
        isDefault: false,
      })),
    ];
    return list;
  }

  /**
   * Persist a named workspace layout for a workbook.
   *
   * @param {string} workbookId
   * @param {string} workspaceId
   * @param {{ name?: string, layout: any, makeActive?: boolean }} params
   */
  saveWorkbookWorkspace(workbookId, workspaceId, params) {
    const name = typeof params.name === "string" && params.name.trim().length > 0 ? params.name.trim() : workspaceId;
    const makeActive = Boolean(params.makeActive);

    if (workspaceId === "default") {
      this.storage.setItem(this.workbookKey(workbookId), serializeLayout(params.layout, { panelRegistry: this.panelRegistry }));
      if (makeActive) this.setActiveWorkbookWorkspace(workbookId, "default");
      return;
    }

    this.storage.setItem(
      this.workbookWorkspaceKey(workbookId, workspaceId),
      serializeLayout(params.layout, { panelRegistry: this.panelRegistry }),
    );

    const index = this.loadWorkbookWorkspaceIndex(workbookId);
    index.workspaces[workspaceId] = { name };
    if (makeActive) index.activeWorkspaceId = workspaceId;
    this.saveWorkbookWorkspaceIndex(workbookId, index);
  }

  /**
   * Load a named workspace layout.
   *
   * @param {string} workbookId
   * @param {string} workspaceId
   * @param {{ primarySheetId?: string | null }} [options]
   */
  loadWorkbookWorkspace(workbookId, workspaceId, options = {}) {
    if (workspaceId === "default") {
      const raw = this.storage.getItem(this.workbookKey(workbookId));
      if (raw) {
        return deserializeLayout(raw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
      }
      return null;
    }

    const raw = this.storage.getItem(this.workbookWorkspaceKey(workbookId, workspaceId));
    if (!raw) return null;
    return deserializeLayout(raw, { panelRegistry: this.panelRegistry, primarySheetId: options.primarySheetId });
  }

  /**
   * @param {string} workbookId
   * @param {string} workspaceId
   */
  setActiveWorkbookWorkspace(workbookId, workspaceId) {
    const index = this.loadWorkbookWorkspaceIndex(workbookId);

    if (workspaceId !== "default" && !Object.prototype.hasOwnProperty.call(index.workspaces, workspaceId)) {
      throw new Error(`Unknown workspace: ${workspaceId}`);
    }

    index.activeWorkspaceId = workspaceId;
    this.saveWorkbookWorkspaceIndex(workbookId, index);
  }

  /**
   * @param {string} workbookId
   * @param {string} workspaceId
   */
  deleteWorkbookWorkspace(workbookId, workspaceId) {
    if (workspaceId === "default") {
      this.deleteWorkbookLayout(workbookId);
      this.setActiveWorkbookWorkspace(workbookId, "default");
      return;
    }

    this.storage.removeItem(this.workbookWorkspaceKey(workbookId, workspaceId));
    const index = this.loadWorkbookWorkspaceIndex(workbookId);
    delete index.workspaces[workspaceId];
    if (index.activeWorkspaceId === workspaceId) index.activeWorkspaceId = "default";
    this.saveWorkbookWorkspaceIndex(workbookId, index);
  }
}
