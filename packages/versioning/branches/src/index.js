export { mergeDocumentStates, applyConflictResolutions } from "./merge.js";
export { diffDocumentStates, applyPatch } from "./patch.js";
export { BranchService } from "./BranchService.js";
export { InMemoryBranchStore } from "./store/InMemoryBranchStore.js";
export { SQLiteBranchStore } from "./store/SQLiteBranchStore.js";
export { YjsBranchStore } from "./store/YjsBranchStore.js";
export { yjsDocToDocumentState, applyDocumentStateToYjsDoc } from "./yjs/index.js";
