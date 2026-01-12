import { bindCollabSessionToDocumentController, type CollabSession, type DocumentControllerBinder } from "@formula/collab-session";
import { getCommentsRoot } from "@formula/collab-comments";
import { createUndoService, type UndoService } from "@formula/collab-undo";
import * as Y from "yjs";

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

function isYUndoManager(value: unknown): value is Y.UndoManager {
  if (value instanceof Y.UndoManager) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = value as any;
  // Bundlers can rename constructors and pnpm workspaces can load multiple `yjs`
  // module instances (ESM + CJS). Avoid relying on `constructor.name`; prefer a
  // structural check instead.
  return (
    typeof maybe.addToScope === "function" &&
    typeof maybe.undo === "function" &&
    typeof maybe.redo === "function" &&
    typeof maybe.stopCapturing === "function"
  );
}

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

  // Ensure comments participate in the binder-origin undo scope once it's safe to
  // instantiate the `comments` root (older docs may still use an Array-backed schema).
  const undoManager: Y.UndoManager | null = (() => {
    for (const origin of undoService.localOrigins ?? []) {
      if (isYUndoManager(origin)) return origin as Y.UndoManager;
    }
    return null;
  })();

  const ensureCommentsUndoScope = () => {
    if (!undoManager) return;

    // Avoid clobbering legacy docs by instantiating a Map root before the provider
    // has hydrated the document.
    const provider = session.provider;
    const providerSynced =
      provider && typeof (provider as any).on === "function" ? Boolean((provider as any).synced) : true;
    if (!providerSynced && !session.doc.share.get("comments")) return;

    try {
      const root = getCommentsRoot(session.doc);
      undoManager.addToScope(root.kind === "map" ? root.map : root.array);
    } catch {
      // Best-effort.
    }
  };

  const provider = session.provider;
  if (provider && typeof provider.on === "function") {
    const onSync = (isSynced: boolean) => {
      if (!isSynced) return;
      provider.off?.("sync", onSync);
      ensureCommentsUndoScope();
    };
    provider.on("sync", onSync);
    if ((provider as any).synced) onSync(true);

    // If the comments root already exists before the provider reports `sync=true`
    // (e.g. offline persistence hydration, or an early local comment), attempt to
    // add it to scope immediately so comment edits are undoable without waiting
    // for provider sync.
    try {
      if (session.doc.share.get("comments")) ensureCommentsUndoScope();
    } catch {
      // Best-effort.
    }

    // Offline persistence can hydrate the Y.Doc before provider sync. Ensure we
    // add comments to the undo scope as soon as offline hydration completes.
    if (typeof (session as any).offline?.whenLoaded === "function") {
      void (session as any).offline
        .whenLoaded()
        .then(() => {
          try {
            if (session.doc.share.get("comments")) ensureCommentsUndoScope();
          } catch {
            // Best-effort.
          }
        })
        .catch(() => {
          // ignore
        });
    }
  } else {
    ensureCommentsUndoScope();
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
