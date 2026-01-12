export type MarketplacePanelDeps = {
  container: HTMLElement;
  marketplaceClient: {
    search: (params: any) => Promise<any>;
    getExtension: (id: string) => Promise<any>;
  };
  extensionManager: {
    getInstalled: (id: string) => Promise<any>;
    install: (
      id: string,
      version?: string | null,
      options?: {
        scanPolicy?: "enforce" | "allow" | "ignore" | string;
        confirm?: (warning: { kind: string; message: string; scanStatus?: string | null }) => Promise<boolean> | boolean;
      },
    ) => Promise<any>;
    uninstall: (id: string) => Promise<any>;
    checkForUpdates: () => Promise<any>;
    update: (id: string) => Promise<any>;
    repair?: (id: string) => Promise<any>;
    scanPolicy?: "enforce" | "allow" | "ignore" | string;
  };
  extensionHostManager?: {
    syncInstalledExtensions?: () => Promise<void> | void;
    reloadExtension?: (id: string) => Promise<void> | void;
    unloadExtension?: (id: string) => Promise<void> | void;
    resetExtensionState?: (id: string) => Promise<void> | void;
  } | null;
};

export function createMarketplacePanel(args: MarketplacePanelDeps): { dispose: () => void };
