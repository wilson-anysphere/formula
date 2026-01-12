// Desktop uses Vite `resolve.alias` entries (see `apps/desktop/vite.config.ts`) to load several
// workspace-only packages that are not present in `apps/desktop/node_modules` in all environments
// (e.g. when pnpm workspace links are missing/stale).
//
// TypeScript's `tsc` does not read Vite aliases, so provide permissive ambient module shims to keep
// desktop `tsc` runs usable without requiring a fully linked workspace install.

declare module "@formula/collab-session" {
  export type CollabSession = any;
  export type CollabSessionConnectionOptions = any;
  export type DocumentControllerBinder = any;

  export const createCollabSession: (...args: any[]) => any;
  export const bindCollabSessionToDocumentController: (...args: any[]) => any;
}

declare module "@formula/collab-undo" {
  export type UndoService = any;

  export const REMOTE_ORIGIN: any;
  export const createUndoService: (...args: any[]) => any;
}

declare module "@formula/collab-persistence/indexeddb" {
  export const IndexedDbCollabPersistence: any;
}

declare module "@formula/extension-marketplace" {
  // Keep these intentionally permissive: desktop builds resolve the real implementation via Vite
  // aliases, but `tsc` doesn't. Treat the surface as `any` to avoid dragging the entire marketplace
  // package (and its transitive deps) into the desktop TS program.
  export type MarketplaceClient = any;
  export const MarketplaceClient: any;

  export type WebExtensionManager = any;
  export const WebExtensionManager: any;
}

declare module "@formula/marketplace-shared/extension-package/v2-browser.mjs" {
  export type ReadExtensionPackageV2Result = any;
  export type VerifiedExtensionPackageV2 = any;

  export const readExtensionPackageV2: (...args: any[]) => any;
  export const verifyExtensionPackageV2Browser: (...args: any[]) => any;
}

declare module "@formula/marketplace-shared/extension-manifest/index.mjs" {
  export const validateExtensionManifest: (...args: any[]) => any;
}
