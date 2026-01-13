// Node-only entrypoint for branching helpers that depend on Node built-ins.
//
// Browser/Vite runtimes should import from `packages/versioning/branches/src/browser.js`
// (a browser-safe subset that excludes Node-only stores like `SQLiteBranchStore`).
//
// This module re-exports the full browser-safe surface from `./index.js` and adds
// Node-only stores on top.

export * from "./index.js";

export { SQLiteBranchStore } from "./store/SQLiteBranchStore.js";
