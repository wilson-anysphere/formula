// `vitestWorkspaceAliases.test.ts` verifies that Vitest/Vite `resolve.alias` entries allow importing
// workspace packages that may not exist in `packages/engine/node_modules` (e.g. when workspace links
// are missing/stale in CI caches).
//
// `tsc` does not read `vitest.config.ts`, so provide minimal ambient module shims to keep
// `pnpm -w typecheck` green while the runtime test continues to exercise the real imports.

declare module "@formula/fill-engine" {
  export const computeFillEdits: (...args: any[]) => any;
}

declare module "@formula/grid/node" {
  export const DEFAULT_GRID_FONT_FAMILY: string;
  export class LruCache {
    constructor(...args: any[]);
  }
}

