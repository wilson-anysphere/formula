import type * as Y from "yjs";

export type UndoMode = "single" | "collab";

export interface UndoService {
  mode: UndoMode;
  canUndo: () => boolean;
  canRedo: () => boolean;
  undo: () => void;
  redo: () => void;
  stopCapturing: () => void;
  perform: (change: { redo: () => void; undo?: () => void }) => void;
  transact?: (fn: () => void) => void;
  localOrigins?: Set<any>;
}

export function createUndoService(opts: {
  mode: UndoMode;
  doc?: Y.Doc;
  scope?: Y.AbstractType<any> | Array<Y.AbstractType<any>>;
  captureTimeoutMs?: number;
  origin?: object;
}): UndoService;

export const REMOTE_ORIGIN: object;

export function createCollabUndoService(opts: {
  doc: Y.Doc;
  scope: Y.AbstractType<any> | Array<Y.AbstractType<any>>;
  captureTimeoutMs?: number;
  origin?: object;
}): {
  mode: "collab";
  undoManager: Y.UndoManager;
  origin: object;
  localOrigins: Set<any>;
  canUndo: () => boolean;
  canRedo: () => boolean;
  undo: () => void;
  redo: () => void;
  stopCapturing: () => void;
  transact: (fn: () => void) => void;
};

export class LocalUndoStack {
  constructor(opts?: { captureTimeoutMs?: number });
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): void;
  redo(): void;
  stopCapturing(): void;
  perform(change: { redo: () => void; undo: () => void }): void;
}
