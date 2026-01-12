export class LayoutController {
  workbookId: string;
  workspaceManager: any;
  primarySheetId: string | null;
  layout: any;

  constructor(params: { workbookId: string; workspaceManager: any; primarySheetId?: string | null; workspaceId?: string });

  on(event: string, listener: (payload: any) => void): () => void;
  persistNow(): void;
  save(): void;
  reload(): void;

  readonly activeWorkspaceId: string;
  listWorkspaces(): any[];
  setActiveWorkspace(workspaceId: string): void;
  setWorkspace(workspaceId: string): void;
  saveWorkspace(workspaceId: string, options?: { name?: string; makeActive?: boolean }): void;
  deleteWorkspace(workspaceId: string): void;

  openPanel(panelId: string): void;
  closePanel(panelId: string): void;
  dockPanel(panelId: string, side: any, options?: any): void;
  activateDockedPanel(panelId: string, side: any): void;
  floatPanel(panelId: string, rect: any, options?: any): void;

  setFloatingPanelRect(panelId: string, rect: any): void;
  setFloatingPanelMinimized(panelId: string, minimized: boolean): void;
  snapFloatingPanel(panelId: string, viewport: any, options?: any): void;

  setDockCollapsed(side: any, collapsed: boolean): void;
  setDockSize(side: any, sizePx: number): void;

  setSplitDirection(direction: any, options?: { persist?: boolean }): void;
  setSplitDirection(direction: any, ratio?: number, options?: { persist?: boolean }): void;
  setSplitRatio(ratio: number, options?: { persist?: boolean }): void;
  setActiveSplitPane(pane: any): void;
  setSplitPaneSheet(pane: any, sheetId: string): void;
  setSplitPaneScroll(pane: any, scroll: any, options?: { persist?: boolean }): void;
  setSplitPaneZoom(pane: any, zoom: any, options?: { persist?: boolean }): void;

  saveAsGlobalDefault(): void;
}
