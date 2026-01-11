import * as Y from "yjs";
import { useEffect, useMemo, useState } from "react";

import { CommentManager, getCommentsRoot } from "@formula/collab-comments";
import type { Comment } from "@formula/collab-comments";

export function useComments(doc: Y.Doc, cellRef: string | null): {
  manager: CommentManager;
  comments: Comment[];
} {
  const manager = useMemo(() => new CommentManager(doc), [doc]);
  const [comments, setComments] = useState<Comment[]>([]);

  useEffect(() => {
    const update = (): void => {
      const all = manager.listAll();
      setComments(cellRef ? all.filter((comment) => comment.cellRef === cellRef) : []);
    };

    update();

    const root = getCommentsRoot(doc);
    if (root.kind === "map") {
      root.map.observe(update);
      return () => {
        root.map.unobserve(update);
      };
    }

    root.array.observe(update);
    return () => {
      root.array.unobserve(update);
    };
  }, [doc, cellRef, manager]);

  return { manager, comments };
}
