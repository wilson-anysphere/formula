export const PanelIds = Object.freeze({
  AI_CHAT: "aiChat",
  VERSION_HISTORY: "versionHistory",
  SCRIPT_EDITOR: "scriptEditor",
  PIVOT_BUILDER: "pivotBuilder",
  QUERY_EDITOR: "queryEditor",
  PYTHON: "python",
  SOLVER: "solver",
  SCENARIO_MANAGER: "scenarioManager",
});

export const PANEL_REGISTRY = Object.freeze({
  [PanelIds.AI_CHAT]: {
    title: "AI Assistant",
    defaultDock: "right",
    defaultFloatingRect: { x: 120, y: 120, width: 480, height: 640 },
  },
  [PanelIds.VERSION_HISTORY]: {
    title: "Version History",
    defaultDock: "right",
    defaultFloatingRect: { x: 160, y: 160, width: 480, height: 640 },
  },
  [PanelIds.SCRIPT_EDITOR]: {
    title: "Script Editor",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 140, y: 140, width: 720, height: 420 },
  },
  [PanelIds.PIVOT_BUILDER]: {
    title: "Pivot Builder",
    defaultDock: "left",
    defaultFloatingRect: { x: 100, y: 100, width: 520, height: 640 },
  },
  [PanelIds.QUERY_EDITOR]: {
    title: "Query Editor",
    defaultDock: "right",
    defaultFloatingRect: { x: 140, y: 140, width: 640, height: 720 },
  },
  [PanelIds.PYTHON]: {
    title: "Python",
    defaultDock: "bottom",
    defaultFloatingRect: { x: 120, y: 120, width: 760, height: 460 },
  },
  [PanelIds.SOLVER]: {
    title: "Solver",
    defaultDock: "right",
    defaultFloatingRect: { x: 180, y: 160, width: 520, height: 640 },
  },
  [PanelIds.SCENARIO_MANAGER]: {
    title: "Scenario Manager",
    defaultDock: "left",
    defaultFloatingRect: { x: 120, y: 160, width: 520, height: 640 },
  },
});

export function isPanelId(panelId) {
  return Object.prototype.hasOwnProperty.call(PANEL_REGISTRY, panelId);
}
