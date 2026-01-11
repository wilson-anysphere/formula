import type * as Y from "yjs";

import { CommentManager } from "./manager.ts";

export * from "./types.ts";
export * from "./yjs.ts";
export * from "./manager.ts";
export * from "./persistence.ts";

export function createCommentManagerForSession(session: { doc: Y.Doc; transactLocal: (fn: () => void) => void }): CommentManager {
  return new CommentManager(session.doc, { transact: (fn) => session.transactLocal(fn) });
}
