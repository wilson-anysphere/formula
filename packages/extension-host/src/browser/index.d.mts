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

export class BrowserExtensionHost {
  constructor(options: {
    engineVersion: string;
    spreadsheetApi: any;
    uiApi?: any;
    permissionPrompt?: (...args: any[]) => unknown;
    permissionStorage?: any;
    permissionStorageKey?: string;
    clipboardApi?: ClipboardApi;
    storageApi?: any;
    activationTimeoutMs?: number;
    commandTimeoutMs?: number;
    customFunctionTimeoutMs?: number;
    dataConnectorTimeoutMs?: number;
    sandbox?: any;
  });

  loadExtensionFromUrl(manifestUrl: string): Promise<string>;
  startup(): Promise<void>;

  getContributedCommands(): ContributedCommand[];
  getContributedPanels(): any[];
  getContributedKeybindings(): any[];
  getContributedMenu(menuId: string): any[];

  listExtensions(): any[];

  getGrantedPermissions(extensionId: string): Promise<any>;
  revokePermissions(extensionId: string, permissions?: string[]): Promise<void>;
  resetPermissions(extensionId: string): Promise<void>;
  resetAllPermissions(): Promise<void>;

  executeCommand(commandId: string, ...args: any[]): Promise<any>;
}
