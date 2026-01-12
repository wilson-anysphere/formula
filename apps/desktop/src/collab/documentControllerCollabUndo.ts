import { bindCollabSessionToDocumentController, type CollabSession, type DocumentControllerBinder } from "@formula/collab-session";
import { getCommentsRoot } from "@formula/collab-comments";
import { createUndoService, type UndoService } from "@formula/collab-undo";

export type DocumentControllerCollabUndoBinding = {
  binder: DocumentControllerBinder;
  undoService: UndoService;
  /**
   * Origin used for DocumentController→Yjs transactions (distinct from `session.origin`).
   *
   * This origin is tracked by the undo service so only DocumentController-origin changes
   * are undoable.
   */
  binderOrigin: object;
};

export async function bindDocumentControllerWithCollabUndo(options: {
  session: CollabSession;
  documentController: any;
  defaultSheetId?: string;
  userId?: string | null;
}): Promise<DocumentControllerCollabUndoBinding> {
  const { session, documentController, defaultSheetId, userId } = options ?? ({} as any);
  if (!session) throw new Error("bindDocumentControllerWithCollabUndo requires { session }");
  if (!documentController) throw new Error("bindDocumentControllerWithCollabUndo requires { documentController }");

  // Intentionally distinct from `session.origin`. Collab conflict detection uses `session.origin`,
  // and those writes must still propagate through the binder (i.e. must *not* be treated as binder-local).
  const binderOrigin = { type: "document-controller:binder" };
  if ((session as any).origin === binderOrigin) {
    // This should never happen (fresh object), but keep the invariant obvious if callers
    // replace `session.origin` with a shared constant.
    throw new Error("bindDocumentControllerWithCollabUndo requires binderOrigin !== session.origin");
  }

  const scope = [
    session.cells,
    session.sheets,
    session.metadata,
    session.namedRanges,
  ];
  // Include comments root when present. Avoid instantiating `doc.getMap("comments")`
  // pre-hydration because older docs may still use an Array-backed schema.
  try {
    if (session.doc.share.get("comments")) {
      const root = getCommentsRoot(session.doc);
      scope.push(root.kind === "map" ? root.map : root.array);
    }
  } catch {
    // Best-effort; avoid breaking binder setup due to comment schema issues.
  }

  const undoService = createUndoService({
    mode: "collab",
    doc: session.doc,
    scope,
    origin: binderOrigin,
  }) as UndoService & { origin?: any };

  // The binder uses `undoService.origin` for echo suppression. `createUndoService` doesn't
  // currently expose it, so attach it explicitly.
  undoService.origin = binderOrigin;

  // Ensure undo/redo transactions are treated as "local" by conflict monitors.
  for (const origin of undoService.localOrigins ?? []) {
    session.localOrigins.add(origin);
  }

  const binder = await bindCollabSessionToDocumentController({
    session,
    documentController,
    undoService,
    defaultSheetId,
    userId,
  });

  return { binder, undoService, binderOrigin };
}
