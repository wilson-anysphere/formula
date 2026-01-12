export type MarketplacePanelDeps = {
  container: HTMLElement;
  marketplaceClient: {
    search: (params: any) => Promise<any>;
  };
  extensionManager: {
    getInstalled: (id: string) => Promise<any>;
    install: (id: string) => Promise<any>;
    uninstall: (id: string) => Promise<any>;
    checkForUpdates: () => Promise<any>;
    update: (id: string) => Promise<any>;
  };
  extensionHostManager?: {
    syncInstalledExtensions?: () => Promise<void> | void;
    reloadExtension?: (id: string) => Promise<void> | void;
    unloadExtension?: (id: string) => Promise<void> | void;
    resetExtensionState?: (id: string) => Promise<void> | void;
  } | null;
};

export function createMarketplacePanel(args: MarketplacePanelDeps): { dispose: () => void };
