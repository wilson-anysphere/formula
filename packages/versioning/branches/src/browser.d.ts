export { BranchService } from "./BranchService.js";
export { YjsBranchStore } from "./store/YjsBranchStore.js";

export { yjsDocToDocumentState, applyDocumentStateToYjsDoc } from "./yjs/index.js";
export { branchStateFromYjsDoc, applyBranchStateToYjsDoc } from "./yjs/branchStateAdapter.js";

export { mergeDocumentStates, applyConflictResolutions } from "./merge.js";
export { diffDocumentStates, applyPatch } from "./patch.js";
export { emptyDocumentState, normalizeDocumentState } from "./state.js";
