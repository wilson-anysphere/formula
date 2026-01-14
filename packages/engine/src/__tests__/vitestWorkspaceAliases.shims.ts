// Vitest resolves several workspace-only packages via `resolve.alias` entries in `vitest.config.ts`
// so tests still run in environments with stale/missing pnpm workspace links.
//
// TypeScript's `tsc` does not read Vite/Vitest aliases, so provide permissive ambient module shims
// for the packages imported by `vitestWorkspaceAliases.test.ts`. These shims keep `pnpm typecheck`
// usable without requiring a fully linked workspace `node_modules`.

declare module "@formula/fill-engine" {
  export const computeFillEdits: (...args: any[]) => any;
}

declare module "@formula/grid/node" {
  export const DEFAULT_GRID_FONT_FAMILY: string;
  export const LruCache: any;
}

