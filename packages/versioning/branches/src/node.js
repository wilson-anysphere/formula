// Node-only entrypoint for branch/versioning helpers that depend on Node built-ins.
//
// The main `packages/versioning/branches/src/index.js` entrypoint is safe to import in
// browser/Vite runtimes (desktop app + Playwright). Node-specific stores should be
// imported from this module instead.

export * from "./index.js";

export { SQLiteBranchStore } from "./store/SQLiteBranchStore.js";
