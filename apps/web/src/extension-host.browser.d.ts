declare module "@formula/extension-host/browser" {
  export class BrowserExtensionHost {
    constructor(options: any);

    loadExtension(args: {
      extensionId: string;
      extensionPath: string;
      manifest: Record<string, any>;
      mainUrl: string;
    }): Promise<string>;

    unloadExtension(extensionId: string): Promise<void | boolean>;

    resetExtensionState(extensionId: string): Promise<void>;

    updateExtension(args: {
      extensionId: string;
      extensionPath: string;
      manifest: Record<string, any>;
      mainUrl: string;
    }): Promise<string>;

    startup(): Promise<void>;

    startupExtension(extensionId: string): Promise<void>;

    listExtensions(): Array<{
      id: string;
      path: string;
      active: boolean;
      manifest: Record<string, any>;
    }>;

    getGrantedPermissions(extensionId: string): Promise<any>;
    revokePermissions(extensionId: string, permissions?: string[]): Promise<void>;
    resetPermissions(extensionId: string): Promise<void>;
    resetAllPermissions(): Promise<void>;

    executeCommand(commandId: string, ...args: any[]): Promise<any>;

    getMessages(): Array<{ message: string; type: string }>;

    dispose(): Promise<void>;
  }

  export const API_PERMISSIONS: Record<string, string[]>;
}
