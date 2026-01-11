export type UndoRedoKeyboardEvent = {
  key?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  preventDefault?: () => void;
};

export function isUndoKeyboardEvent(event: UndoRedoKeyboardEvent): boolean;
export function isRedoKeyboardEvent(event: UndoRedoKeyboardEvent): boolean;

export function installUndoRedoShortcuts(
  target: {
    addEventListener: (type: string, listener: (event: any) => void) => void;
    removeEventListener: (type: string, listener: (event: any) => void) => void;
  },
  controller: { undo: () => boolean; redo: () => boolean }
): () => void;

