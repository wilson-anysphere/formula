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

  getContributedCommands(): ContributedCommand[];
  getContributedPanels(): any[];
  getContributedKeybindings(): any[];
  getContributedMenu(menuId: string): any[];

  listExtensions(): any[];

  executeCommand(commandId: string, ...args: any[]): Promise<any>;
}
