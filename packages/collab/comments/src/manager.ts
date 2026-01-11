import * as Y from "yjs";

import type { Comment, CommentAuthor, CommentKind } from "./types.ts";
import { createYComment, createYReply, getCommentsMap, yCommentToComment } from "./yjs.ts";

export interface CommentManagerOptions {
  transact?: (fn: () => void) => void;
}

export class CommentManager {
  private readonly doc: Y.Doc;
  private readonly transact: (fn: () => void) => void;

  constructor(doc: Y.Doc, options: CommentManagerOptions = {}) {
    this.doc = doc;
    this.transact = options.transact ?? ((fn) => doc.transact(fn));
  }

  listAll(): Comment[] {
    const map = getCommentsMap(this.doc);
    const comments = Array.from(map.values()).map(yCommentToComment);
    comments.sort((a, b) => {
      if (a.createdAt !== b.createdAt) return a.createdAt - b.createdAt;
      return a.id.localeCompare(b.id);
    });
    return comments;
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

    this.transact(() => {
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

    this.transact(() => {
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

    this.transact(() => {
      yComment.set("resolved", input.resolved);
      yComment.set("updatedAt", now);
    });
  }

  setCommentContent(input: { commentId: string; content: string; now?: number }): void {
    const map = getCommentsMap(this.doc);
    const yComment = map.get(input.commentId);
    if (!yComment) {
      throw new Error(`Comment not found: ${input.commentId}`);
    }
    const now = input.now ?? Date.now();

    this.transact(() => {
      yComment.set("content", input.content);
      yComment.set("updatedAt", now);
    });
  }

  setReplyContent(input: { commentId: string; replyId: string; content: string; now?: number }): void {
    const map = getCommentsMap(this.doc);
    const yComment = map.get(input.commentId);
    if (!yComment) {
      throw new Error(`Comment not found: ${input.commentId}`);
    }

    const replies = yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined;
    if (!replies) {
      throw new Error(`Comment replies missing: ${input.commentId}`);
    }

    const now = input.now ?? Date.now();

    const replyIndex = replies
      .toArray()
      .findIndex((reply) => String(reply.get("id") ?? "") === input.replyId);
    if (replyIndex < 0) {
      throw new Error(`Reply not found: ${input.replyId}`);
    }

    const yReply = replies.get(replyIndex) as Y.Map<unknown> | undefined;
    if (!yReply) {
      throw new Error(`Reply missing at index ${replyIndex}: ${input.replyId}`);
    }

    this.transact(() => {
      yReply.set("content", input.content);
      yReply.set("updatedAt", now);
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
