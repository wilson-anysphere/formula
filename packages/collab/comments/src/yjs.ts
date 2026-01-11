import * as Y from "yjs";

import type { Comment, CommentAuthor, CommentKind, Reply } from "./types.ts";

export type YCommentsMap = Y.Map<Y.Map<unknown>>;

export function getCommentsMap(doc: Y.Doc): YCommentsMap {
  return doc.getMap("comments");
}

export function yCommentToComment(yComment: Y.Map<unknown>): Comment {
  const replies = (yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined)
    ?.toArray()
    .map(yReplyToReply) ?? [];

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
