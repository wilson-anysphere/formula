import * as Y from "yjs";
import { useEffect, useMemo, useState } from "react";

import { CommentManager, createCommentManagerForDoc, getCommentsRoot } from "@formula/collab-comments";
import type { Comment } from "@formula/collab-comments";

export function useComments(
  doc: Y.Doc,
  cellRef: string | null,
  opts: { transact?: (fn: () => void) => void; canComment?: () => boolean } = {},
): {
  manager: CommentManager;
  comments: Comment[];
} {
  const manager = useMemo(() => {
    return typeof opts.transact === "function"
      ? createCommentManagerForDoc({ doc, transact: opts.transact, canComment: opts.canComment })
      : new CommentManager(doc, { canComment: opts.canComment });
  }, [doc, opts.transact, opts.canComment]);
  const [comments, setComments] = useState<Comment[]>([]);

  useEffect(() => {
    const update = (): void => {
      // Pre-hydration safety: avoid instantiating the comments root (which would
      // pick a Map constructor and could clobber legacy Array-backed docs) until
      // we know the root exists.
      if (!doc.share.get("comments")) {
        setComments([]);
        return;
      }
      const all = manager.listAll();
      setComments(cellRef ? all.filter((comment) => comment.cellRef === cellRef) : []);
    };

    update();

    const attach = (): (() => void) | null => {
      if (!doc.share.get("comments")) return null;
      const root = getCommentsRoot(doc);
      if (root.kind === "map") {
        root.map.observeDeep(update);
        return () => root.map.unobserveDeep(update);
      }
      root.array.observeDeep(update);
      return () => root.array.unobserveDeep(update);
    };

    let detach = attach();
    if (detach) return () => detach?.();

    // Root doesn't exist yet. Subscribe to doc updates until it does, then
    // switch over to observing the comments root directly.
    const onUpdate = (): void => {
      if (detach) return;
      detach = attach();
      if (detach) {
        doc.off("update", onUpdate);
        update();
      }
    };
    doc.on("update", onUpdate);
    return () => {
      doc.off("update", onUpdate);
      detach?.();
    };
  }, [doc, cellRef, manager]);

  return { manager, comments };
}
