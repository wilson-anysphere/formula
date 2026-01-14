export const PanelIds: Readonly<{
  AI_CHAT: string;
  AI_AUDIT: string;
  MACROS: string;
  VERSION_HISTORY: string;
  FORMULA_DEBUGGER: string;
  VBA_MIGRATE: string;
  SCRIPT_EDITOR: string;
  PIVOT_BUILDER: string;
  DATA_QUERIES: string;
  QUERY_EDITOR: string;
  PYTHON: string;
  SOLVER: string;
  SCENARIO_MANAGER: string;
  MONTE_CARLO: string;
  MARKETPLACE: string;
  BRANCH_MANAGER: string;
  EXTENSIONS: string;
}>;

export type PanelDefinition = {
  title?: string;
  titleKey?: string;
  icon?: string | null;
  defaultDock?: "left" | "right" | "bottom";
  defaultFloatingRect?: { x: number; y: number; width: number; height: number };
  source?: { kind: "builtin" } | { kind: "extension"; extensionId: string; contributed: boolean };
};

export class PanelRegistry {
  has(panelId: string): boolean;
  get(panelId: string): PanelDefinition | undefined;
  listPanelIds(): string[];
  onDidChange(listener: () => void): () => void;
  registerPanel(panelId: string, definition: PanelDefinition, options?: { owner?: string; overwrite?: boolean }): void;
  unregisterPanel(panelId: string, options?: { owner?: string }): void;
  unregisterOwner(owner: string): void;
  getTitle(panelId: string): string;
}

export const panelRegistry: PanelRegistry;

export function isPanelId(panelId: string): boolean;
export function getPanelTitle(panelId: string): string;
export const PANEL_REGISTRY: PanelRegistry;
