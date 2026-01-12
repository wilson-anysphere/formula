// Node-only entrypoint for versioning helpers that depend on Node built-ins.
//
// The main `packages/versioning/src/index.js` entrypoint is safe to import in
// browser/Vite runtimes (desktop app + Playwright). Node-specific stores should be
// imported from this module instead.

export * from "./index.js";

export { FileVersionStore } from "./store/fileVersionStore.js";
export { SQLiteVersionStore } from "./store/sqliteVersionStore.js";
