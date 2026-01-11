import * as Y from "yjs";

import type { Comment, CommentAuthor, CommentKind, Reply } from "./types";

export interface CommentManagerOptions {
  transact?: (fn: () => void) => void;
}

export type YCommentsMap = Y.Map<Y.Map<unknown>>;

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

export function createCommentManagerForSession(session: { doc: Y.Doc; transactLocal: (fn: () => void) => void }): CommentManager {
  return new CommentManager(session.doc, { transact: (fn) => session.transactLocal(fn) });
}

function createId(): string {
  const globalCrypto = (globalThis as any).crypto as Crypto | undefined;
  if (globalCrypto?.randomUUID) {
    return globalCrypto.randomUUID();
  }
  return `c_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}

export function getCommentsMap(doc: Y.Doc): YCommentsMap {
  return doc.getMap("comments");
}

export function yCommentToComment(yComment: Y.Map<unknown>): Comment {
  const replies = (yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined)?.toArray().map(yReplyToReply) ?? [];

  return {
    id: String(yComment.get("id") ?? ""),
    cellRef: String(yComment.get("cellRef") ?? ""),
    kind: (yComment.get("kind") as CommentKind) ?? "threaded",
    author: {
      id: String(yComment.get("authorId") ?? ""),
      name: String(yComment.get("authorName") ?? ""),
    },
    createdAt: Number(yComment.get("createdAt") ?? 0),
    updatedAt: Number(yComment.get("updatedAt") ?? 0),
    resolved: Boolean(yComment.get("resolved") ?? false),
    content: String(yComment.get("content") ?? ""),
    mentions: Array.isArray(yComment.get("mentions")) ? (yComment.get("mentions") as any) : [],
    replies,
  };
}

export function yReplyToReply(yReply: Y.Map<unknown>): Reply {
  return {
    id: String(yReply.get("id") ?? ""),
    author: {
      id: String(yReply.get("authorId") ?? ""),
      name: String(yReply.get("authorName") ?? ""),
    },
    createdAt: Number(yReply.get("createdAt") ?? 0),
    updatedAt: Number(yReply.get("updatedAt") ?? 0),
    content: String(yReply.get("content") ?? ""),
    mentions: Array.isArray(yReply.get("mentions")) ? (yReply.get("mentions") as any) : [],
  };
}

export function createYComment(input: {
  id: string;
  cellRef: string;
  kind: CommentKind;
  author: CommentAuthor;
  now: number;
  content: string;
}): Y.Map<unknown> {
  const yComment = new Y.Map<unknown>();
  yComment.set("id", input.id);
  yComment.set("cellRef", input.cellRef);
  yComment.set("kind", input.kind);
  yComment.set("authorId", input.author.id);
  yComment.set("authorName", input.author.name);
  yComment.set("createdAt", input.now);
  yComment.set("updatedAt", input.now);
  yComment.set("resolved", false);
  yComment.set("content", input.content);
  yComment.set("mentions", []);
  yComment.set("replies", new Y.Array<Y.Map<unknown>>());
  return yComment;
}

export function createYReply(input: {
  id: string;
  author: CommentAuthor;
  now: number;
  content: string;
}): Y.Map<unknown> {
  const yReply = new Y.Map<unknown>();
  yReply.set("id", input.id);
  yReply.set("authorId", input.author.id);
  yReply.set("authorName", input.author.name);
  yReply.set("createdAt", input.now);
  yReply.set("updatedAt", input.now);
  yReply.set("content", input.content);
  yReply.set("mentions", []);
  return yReply;
}
