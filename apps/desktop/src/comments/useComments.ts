import * as Y from "yjs";
import { useEffect, useMemo, useState } from "react";

import { CommentManager, getCommentsMap, yCommentToComment } from "@formula/collab-comments";
import type { Comment } from "@formula/collab-comments";

export function useComments(doc: Y.Doc, cellRef: string | null): {
  manager: CommentManager;
  comments: Comment[];
} {
  const manager = useMemo(() => new CommentManager(doc), [doc]);
  const [comments, setComments] = useState<Comment[]>([]);

  useEffect(() => {
    const map = getCommentsMap(doc);
    const update = (): void => {
      const all = Array.from(map.values()).map(yCommentToComment);
      setComments(cellRef ? all.filter((comment) => comment.cellRef === cellRef) : []);
    };

    update();

    map.observe(update);
    return () => {
      map.unobserve(update);
    };
  }, [doc, cellRef]);

  return { manager, comments };
}

