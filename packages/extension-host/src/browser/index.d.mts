export const API_PERMISSIONS: Record<string, string[]>;

export class BrowserExtensionHost {
  constructor(options: {
    engineVersion: string;
    spreadsheetApi: any;
    uiApi: any;
    permissionPrompt?: (...args: any[]) => unknown;
    sandboxOptions?: any;
    storage?: any;
  });

  loadExtensionFromUrl(manifestUrl: string): Promise<void>;
  startup(): Promise<void>;

  getContributedCommands(): any[];
  getContributedPanels(): any[];
  getContributedKeybindings(): any[];
  getContributedMenu(menuId: string): any[];

  executeCommand(commandId: string, ...args: any[]): Promise<any>;
}

