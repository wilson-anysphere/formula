import { t } from "../i18n/index.js";

export const PanelIds = Object.freeze({
  AI_CHAT: "aiChat",
  MACROS: "macros",
  VERSION_HISTORY: "versionHistory",
  FORMULA_DEBUGGER: "formulaDebugger",
  VBA_MIGRATE: "vbaMigrate",
  SCRIPT_EDITOR: "scriptEditor",
  PIVOT_BUILDER: "pivotBuilder",
  QUERY_EDITOR: "queryEditor",
  PYTHON: "python",
  SOLVER: "solver",
  SCENARIO_MANAGER: "scenarioManager",
  MARKETPLACE: "marketplace",
  BRANCH_MANAGER: "branchManager",
});

export const PANEL_REGISTRY = Object.freeze({
  [PanelIds.AI_CHAT]: {
    titleKey: "chat.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 120, y: 120, width: 480, height: 640 },
  },
  [PanelIds.MACROS]: {
    title: "Macros",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 480, height: 420 },
  },
  [PanelIds.VERSION_HISTORY]: {
    titleKey: "panels.versionHistory.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 160, y: 160, width: 480, height: 640 },
  },
  [PanelIds.FORMULA_DEBUGGER]: {
    titleKey: "panels.formulaDebugger.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 180, y: 180, width: 520, height: 640 },
  },
  [PanelIds.VBA_MIGRATE]: {
    title: "Migrate Macros",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 640 },
  },
  [PanelIds.SCRIPT_EDITOR]: {
    titleKey: "panels.scriptEditor.title",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 420 },
  },
  [PanelIds.PIVOT_BUILDER]: {
    titleKey: "panels.pivotBuilder.title",
    defaultDock: "left",
    defaultFloatingRect: { x: 100, y: 100, width: 520, height: 640 },
  },
  [PanelIds.QUERY_EDITOR]: {
    titleKey: "panels.queryEditor.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 640, height: 720 },
  },
  [PanelIds.PYTHON]: {
    titleKey: "panels.python.title",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 120, y: 120, width: 760, height: 460 },
  },
  [PanelIds.SOLVER]: {
    titleKey: "panels.solver.title",
    defaultDock: "right",
    defaultFloatingRect: { x: 180, y: 160, width: 520, height: 640 },
  },
  [PanelIds.SCENARIO_MANAGER]: {
    titleKey: "whatIf.scenario.title",
    defaultDock: "left",
    defaultFloatingRect: { x: 120, y: 160, width: 520, height: 640 },
  },
  [PanelIds.MARKETPLACE]: {
    title: "Marketplace",
    defaultDock: "right",
    defaultFloatingRect: { x: 160, y: 120, width: 560, height: 680 },
  },
  [PanelIds.BRANCH_MANAGER]: {
    title: "Branches",
    defaultDock: "left",
    defaultFloatingRect: { x: 140, y: 160, width: 520, height: 640 },
  },
});

export function isPanelId(panelId) {
  return Object.prototype.hasOwnProperty.call(PANEL_REGISTRY, panelId);
}

export function getPanelTitle(panelId) {
  const def = PANEL_REGISTRY[panelId];
  if (!def) return panelId;
  if (def.titleKey) return t(def.titleKey);
  if (def.title) return def.title;
  return panelId;
}
