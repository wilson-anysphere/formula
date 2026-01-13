// Browser-safe entrypoint for branching helpers.
//
// This module intentionally avoids exporting Node-only stores (e.g. SQLiteBranchStore)
// so UI/bundled runtimes (Vite, WebView, web workers) can import branching
// functionality without accidentally pulling `node:*` built-ins into the bundle.

export { BranchService } from "./BranchService.js";
export { YjsBranchStore } from "./store/YjsBranchStore.js";

// Yjs adapters.
export { yjsDocToDocumentState, applyDocumentStateToYjsDoc } from "./yjs/index.js";
export { branchStateFromYjsDoc, applyBranchStateToYjsDoc } from "./yjs/branchStateAdapter.js";

