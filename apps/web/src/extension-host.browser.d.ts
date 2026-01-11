declare module "@formula/extension-host/browser" {
  export class BrowserExtensionHost {
    constructor(options: any);

    loadExtension(args: {
      extensionId: string;
      extensionPath: string;
      manifest: Record<string, any>;
      mainUrl: string;
    }): Promise<string>;

    unloadExtension(extensionId: string): Promise<boolean>;

    updateExtension(args: {
      extensionId: string;
      extensionPath: string;
      manifest: Record<string, any>;
      mainUrl: string;
    }): Promise<string>;

    listExtensions(): Array<{
      id: string;
      path: string;
      active: boolean;
      manifest: Record<string, any>;
    }>;

    executeCommand(commandId: string, ...args: any[]): Promise<any>;

    getMessages(): Array<{ message: string; type: string }>;

    dispose(): Promise<void>;
  }

  export const API_PERMISSIONS: Record<string, string[]>;
}

