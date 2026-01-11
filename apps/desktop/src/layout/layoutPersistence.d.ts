export class MemoryStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
  clear(): void;
}

export class LayoutWorkspaceManager {
  storage: Pick<Storage, "getItem" | "setItem" | "removeItem">;
  panelRegistry?: any;
  keyPrefix: string;

  constructor(params: { storage: Pick<Storage, "getItem" | "setItem" | "removeItem">; panelRegistry?: any; keyPrefix?: string });

  loadWorkbookLayout(workbookId: string, options?: { primarySheetId?: string | null }): any;
  saveWorkbookLayout(workbookId: string, layout: any): void;

  loadWorkbookLayoutForWorkspace(workbookId: string, workspaceId: string, options?: { primarySheetId?: string | null }): any;
  saveWorkbookLayoutForWorkspace(workbookId: string, workspaceId: string, layout: any): void;

  getActiveWorkbookWorkspaceId(workbookId: string): string;
  setActiveWorkbookWorkspace(workbookId: string, workspaceId: string): void;

  listWorkbookWorkspaces(workbookId: string): any[];
  saveWorkbookWorkspace(workbookId: string, workspaceId: string, params: { name?: string; layout: any; makeActive?: boolean }): void;
  deleteWorkbookWorkspace(workbookId: string, workspaceId: string): void;

  saveGlobalDefaultLayout(layout: any): void;
  loadGlobalDefaultLayout(options?: { primarySheetId?: string | null }): any;
}

