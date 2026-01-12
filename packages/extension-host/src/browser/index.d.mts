export const API_PERMISSIONS: Record<string, string[]>;

export type ContributedCommand = {
  extensionId: string;
  command: string;
  title: string;
  category: string | null;
  icon: string | null;
  description: string | null;
  keywords: string[] | null;
};

export type ClipboardApi = {
  readText: () => Promise<string>;
  writeText: (text: string) => Promise<void>;
};

export type TaintedRange = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export type ClipboardWriteGuard = (params: { extensionId: string; taintedRanges: TaintedRange[] }) => Promise<void> | void;

export class BrowserExtensionHost {
  constructor(options: {
    spreadsheetApi: any;
    engineVersion?: string;
    uiApi?: any;
    permissionPrompt?: (...args: any[]) => unknown;
    permissionStorage?: any;
    permissionStorageKey?: string;
    clipboardApi?: ClipboardApi;
    clipboardWriteGuard?: ClipboardWriteGuard;
    storageApi?: any;
    activationTimeoutMs?: number;
    commandTimeoutMs?: number;
    customFunctionTimeoutMs?: number;
    dataConnectorTimeoutMs?: number;
    sandbox?: any;
  });

  loadExtensionFromUrl(manifestUrl: string): Promise<string>;
  loadExtension(args: {
    extensionId: string;
    extensionPath: string;
    manifest: Record<string, any>;
    mainUrl: string;
  }): Promise<string>;
  reloadExtension(extensionId: string): Promise<void>;
  unloadExtension(extensionId: string): Promise<void | boolean>;
  updateExtension(args: {
    extensionId: string;
    extensionPath: string;
    manifest: Record<string, any>;
    mainUrl: string;
  }): Promise<string>;

  /**
   * Clears persisted state owned by an extension (permission grants + extension storage).
   * Intended for uninstall flows so a reinstall behaves like a clean install.
   */
  resetExtensionState(extensionId: string): Promise<void>;
  startup(): Promise<void>;
  startupExtension(extensionId: string): Promise<void>;
  activateView(viewId: string): Promise<void>;
  activateCustomFunction(functionName: string): Promise<void>;

  getContributedCommands(): ContributedCommand[];
  getContributedPanels(): any[];
  getContributedKeybindings(): any[];
  getContributedMenu(menuId: string): any[];

  listExtensions(): any[];

  getGrantedPermissions(extensionId: string): Promise<any>;
  revokePermissions(extensionId: string, permissions?: string[]): Promise<void>;
  resetPermissions(extensionId: string): Promise<void>;
  resetAllPermissions(): Promise<void>;

  clearExtensionStorage(extensionId: string): Promise<void>;

  executeCommand(commandId: string, ...args: any[]): Promise<any>;
  invokeCustomFunction(functionName: string, ...args: any[]): Promise<any>;
  invokeDataConnector(connectorId: string, method: string, ...args: any[]): Promise<any>;

  getMessages(): Array<{ message: string; type: string }>;
  dispose(): Promise<void>;
}
