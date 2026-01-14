import { t } from "../i18n/index.js";

export const PanelIds = Object.freeze({
  AI_CHAT: "aiChat",
  AI_AUDIT: "aiAudit",
  MACROS: "macros",
  VERSION_HISTORY: "versionHistory",
  FORMULA_DEBUGGER: "formulaDebugger",
  VBA_MIGRATE: "vbaMigrate",
  SCRIPT_EDITOR: "scriptEditor",
  PIVOT_BUILDER: "pivotBuilder",
  DATA_QUERIES: "dataQueries",
  QUERY_EDITOR: "queryEditor",
  PYTHON: "python",
  SOLVER: "solver",
  SCENARIO_MANAGER: "scenarioManager",
  MONTE_CARLO: "monteCarlo",
  MARKETPLACE: "marketplace",
  BRANCH_MANAGER: "branchManager",
  EXTENSIONS: "extensions",
});

/**
 * @typedef {{
 *  title?: string,
 *  titleKey?: string,
 *  icon?: string | null,
 *  defaultDock?: "left" | "right" | "bottom",
 *  defaultFloatingRect?: { x: number, y: number, width: number, height: number },
 *  source?: { kind: "builtin" } | { kind: "extension", extensionId: string, contributed: boolean }
 * }} PanelDefinition
 */

/**
 * Runtime panel registry (builtin + extensions).
 *
 * The layout persistence layer consults this registry during normalize/serialize/deserialize
 * so extension panel ids are not dropped from persisted layouts.
 */
export class PanelRegistry {
  /** @type {Map<string, PanelDefinition>} */
  #panels = new Map();
  /** @type {Map<string, string>} */
  #owners = new Map();
  /** @type {Set<() => void>} */
  #listeners = new Set();

  /**
   * @param {string} panelId
   */
  has(panelId) {
    return this.#panels.has(String(panelId));
  }

  /**
   * @param {string} panelId
   * @returns {PanelDefinition | undefined}
   */
  get(panelId) {
    return this.#panels.get(String(panelId));
  }

  /**
   * @returns {string[]}
   */
  listPanelIds() {
    return [...this.#panels.keys()];
  }

  /**
   * @param {() => void} listener
   * @returns {() => void}
   */
  onDidChange(listener) {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  #emitChange() {
    for (const listener of [...this.#listeners]) {
      try {
        listener();
      } catch {
        // ignore
      }
    }
  }

  /**
   * Register or update a panel definition.
   *
   * @param {string} panelId
   * @param {PanelDefinition} definition
   * @param {{ owner?: string, overwrite?: boolean }} [options]
   */
  registerPanel(panelId, definition, options = {}) {
    const id = String(panelId);
    const owner = typeof options.owner === "string" && options.owner.length > 0 ? options.owner : "builtin";
    const existingOwner = this.#owners.get(id);
    if (existingOwner && existingOwner !== owner) {
      throw new Error(`Panel id already registered: ${id} (owned by ${existingOwner})`);
    }
    if (existingOwner && existingOwner === owner && !(options.overwrite ?? true)) {
      return;
    }
    this.#owners.set(id, owner);
    this.#panels.set(id, definition);
    this.#emitChange();
  }

  /**
   * @param {string} panelId
   * @param {{ owner?: string }} [options]
   */
  unregisterPanel(panelId, options = {}) {
    const id = String(panelId);
    const owner = typeof options.owner === "string" && options.owner.length > 0 ? options.owner : null;
    const existingOwner = this.#owners.get(id);
    if (!existingOwner) return;
    if (owner && existingOwner !== owner) return;
    this.#owners.delete(id);
    this.#panels.delete(id);
    this.#emitChange();
  }

  /**
   * Remove all panels owned by a given extension id.
   *
   * @param {string} owner
   */
  unregisterOwner(owner) {
    const id = String(owner);
    let changed = false;
    for (const [panelId, panelOwner] of this.#owners.entries()) {
      if (panelOwner !== id) continue;
      this.#owners.delete(panelId);
      this.#panels.delete(panelId);
      changed = true;
    }
    if (changed) this.#emitChange();
  }

  /**
   * @param {string} panelId
   */
  getTitle(panelId) {
    const def = this.get(panelId);
    if (!def) return panelId;
    if (def.titleKey) return t(def.titleKey);
    if (def.title) return def.title;
    return panelId;
  }
}

export const panelRegistry = new PanelRegistry();

// Built-in panels registered on module init.
panelRegistry.registerPanel(
  PanelIds.AI_CHAT,
  {
    titleKey: "chat.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 120, y: 120, width: 480, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.AI_AUDIT,
  {
    title: "Audit Log",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 640, height: 720 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.MACROS,
  {
    title: "Macros",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 480, height: 420 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.VERSION_HISTORY,
  {
    titleKey: "panels.versionHistory.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 160, y: 160, width: 480, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.FORMULA_DEBUGGER,
  {
    titleKey: "panels.formulaDebugger.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 180, y: 180, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.VBA_MIGRATE,
  {
    title: "Migrate Macros",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.SCRIPT_EDITOR,
  {
    titleKey: "panels.scriptEditor.title",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 420 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.PIVOT_BUILDER,
  {
    titleKey: "panels.pivotBuilder.title",
    defaultDock: "left",
    defaultFloatingRect: { x: 100, y: 100, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);

panelRegistry.registerPanel(
  PanelIds.DATA_QUERIES,
  {
    title: "Data / Queries",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 560 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);

panelRegistry.registerPanel(
  PanelIds.QUERY_EDITOR,
  {
    titleKey: "panels.queryEditor.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 640, height: 720 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.PYTHON,
  {
    titleKey: "panels.python.title",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 120, y: 120, width: 760, height: 460 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.SOLVER,
  {
    titleKey: "panels.solver.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 180, y: 160, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.SCENARIO_MANAGER,
  {
    titleKey: "whatIf.scenario.title",
    defaultDock: "left",
    defaultFloatingRect: { x: 120, y: 160, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.MONTE_CARLO,
  {
    titleKey: "whatIf.monteCarlo.title",
    defaultDock: "left",
    defaultFloatingRect: { x: 140, y: 160, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.MARKETPLACE,
  {
    title: "Marketplace",
    defaultDock: "right",
    defaultFloatingRect: { x: 160, y: 120, width: 560, height: 680 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.BRANCH_MANAGER,
  {
    title: "Branches",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 160, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);
panelRegistry.registerPanel(
  PanelIds.EXTENSIONS,
  {
    title: "Extensions",
    defaultDock: "left",
    defaultFloatingRect: { x: 120, y: 120, width: 520, height: 640 },
    source: { kind: "builtin" },
  },
  { owner: "builtin" },
);

export function isPanelId(panelId) {
  return panelRegistry.has(panelId);
}

export function getPanelTitle(panelId) {
  return panelRegistry.getTitle(panelId);
}

// Back-compat export: existing code referenced `PANEL_REGISTRY` as a constant object.
// We now expose the live registry instance so callers can pass it into layout APIs.
export const PANEL_REGISTRY = panelRegistry;
