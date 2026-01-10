import * as Y from "yjs";

import type { Comment, CommentAuthor, CommentKind } from "./types";
import { createYComment, createYReply, getCommentsMap, yCommentToComment } from "./yjs";

export class CommentManager {
  constructor(private readonly doc: Y.Doc) {}

  listAll(): Comment[] {
    const map = getCommentsMap(this.doc);
    return Array.from(map.values()).map(yCommentToComment);
  }

  listForCell(cellRef: string): Comment[] {
    return this.listAll().filter((comment) => comment.cellRef === cellRef);
  }

  addComment(input: {
    cellRef: string;
    kind: CommentKind;
    content: string;
    author: CommentAuthor;
    now?: number;
    id?: string;
  }): string {
    const map = getCommentsMap(this.doc);
    const id = input.id ?? createId();
    const now = input.now ?? Date.now();

    this.doc.transact(() => {
      map.set(
        id,
        createYComment({
          id,
          cellRef: input.cellRef,
          kind: input.kind,
          author: input.author,
          now,
          content: input.content,
        }),
      );
    });

    return id;
  }

  addReply(input: {
    commentId: string;
    content: string;
    author: CommentAuthor;
    now?: number;
    id?: string;
  }): string {
    const map = getCommentsMap(this.doc);
    const yComment = map.get(input.commentId);
    if (!yComment) {
      throw new Error(`Comment not found: ${input.commentId}`);
    }

    const replies = yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined;
    if (!replies) {
      throw new Error(`Comment replies missing: ${input.commentId}`);
    }

    const id = input.id ?? createId();
    const now = input.now ?? Date.now();

    this.doc.transact(() => {
      replies.push([
        createYReply({
          id,
          author: input.author,
          now,
          content: input.content,
        }),
      ]);
      yComment.set("updatedAt", now);
    });

    return id;
  }

  setResolved(input: { commentId: string; resolved: boolean; now?: number }): void {
    const map = getCommentsMap(this.doc);
    const yComment = map.get(input.commentId);
    if (!yComment) {
      throw new Error(`Comment not found: ${input.commentId}`);
    }
    const now = input.now ?? Date.now();

    this.doc.transact(() => {
      yComment.set("resolved", input.resolved);
      yComment.set("updatedAt", now);
    });
  }
}

function createId(): string {
  const globalCrypto = (globalThis as any).crypto as Crypto | undefined;
  if (globalCrypto?.randomUUID) {
    return globalCrypto.randomUUID();
  }
  return `c_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

