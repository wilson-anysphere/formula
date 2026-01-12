export { DocumentController } from "./documentController.js";
export { MockEngine } from "./engine.js";
export { installUndoRedoShortcuts, isRedoKeyboardEvent, isUndoKeyboardEvent } from "./shortcuts.js";

export function installUnsavedChangesPrompt(target: any, controller: any, options?: any): () => void;

export { parseA1, formatA1, parseRangeA1 } from "./coords.js";
