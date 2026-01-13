// Browser-safe entrypoint for branching helpers.
//
// This module intentionally avoids exporting Node-only stores (e.g. SQLiteBranchStore)
// so UI/bundled runtimes (Vite, WebView, web workers) can import branching
// functionality without accidentally pulling `node:*` built-ins into the bundle.

export { BranchService } from "./BranchService.js";
export { YjsBranchStore } from "./store/YjsBranchStore.js";

// Yjs adapters.
export { yjsDocToDocumentState, applyDocumentStateToYjsDoc, rowColToA1, a1ToRowCol } from "./yjs/index.js";
export { branchStateFromYjsDoc, applyBranchStateToYjsDoc } from "./yjs/branchStateAdapter.js";

// Pure helpers (safe in browser/Vite bundles). These are optional conveniences for
// consumers that want branching data manipulation without reaching for deep imports.
export { mergeDocumentStates, applyConflictResolutions } from "./merge.js";
export { diffDocumentStates, applyPatch } from "./patch.js";
export { emptyDocumentState, normalizeDocumentState } from "./state.js";
