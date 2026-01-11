import * as Y from "yjs";

import type { Comment, CommentAuthor, CommentKind, Reply } from "./types";

export interface CommentManagerOptions {
  transact?: (fn: () => void) => void;
}

export type YCommentsMap = Y.Map<Y.Map<unknown>>;
export type YCommentsArray = Y.Array<Y.Map<unknown>>;

export type CommentsRoot =
  | { kind: "map"; map: YCommentsMap }
  | { kind: "array"; array: YCommentsArray };

/**
 * Safely determine whether the `comments` root is a `Y.Map` (current schema) or a
 * legacy `Y.Array` (older docs).
 *
 * Important: never call `doc.getMap("comments")` until you're sure the root is a
 * Map (or doesn't exist yet). If a legacy Array-backed doc is instantiated as a
 * Map, Yjs will refuse to later instantiate it as an Array, making the array
 * content inaccessible.
 *
 * We reuse the heuristic from `workbookStateFromYjsDoc` in `packages/versioning`:
 * inspect the `doc.share.get("comments")` placeholder (`_start`, `_map`) to infer
 * the intended kind without clobbering it.
 */
export function getCommentsRoot(doc: Y.Doc): CommentsRoot {
  const existing = doc.share.get("comments");
  if (!existing) {
    // Canonical schema is a Map keyed by comment id.
    return { kind: "map", map: doc.getMap("comments") as YCommentsMap };
  }

  if (existing instanceof Y.Map) {
    return { kind: "map", map: existing as YCommentsMap };
  }
  if (existing instanceof Y.Array) {
    return { kind: "array", array: existing as YCommentsArray };
  }

  // Root types may be a generic `AbstractType` placeholder until a constructor is
  // chosen. Peek at its internal structure before choosing a constructor.
  const placeholder = existing as any;
  const hasStart = placeholder?._start != null; // sequence item => likely array
  const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
  const kind: CommentsRoot["kind"] = hasStart && mapSize === 0 ? "array" : "map";

  if (kind === "array") {
    return { kind: "array", array: doc.getArray("comments") as YCommentsArray };
  }
  return { kind: "map", map: doc.getMap("comments") as YCommentsMap };
}

export class CommentManager {
  private readonly doc: Y.Doc;
  private readonly transact: (fn: () => void) => void;

  constructor(doc: Y.Doc, options: CommentManagerOptions = {}) {
    this.doc = doc;
    this.transact = options.transact ?? ((fn) => doc.transact(fn));
  }

  listAll(): Comment[] {
    const root = getCommentsRoot(this.doc);
    const comments =
      root.kind === "map"
        ? Array.from(root.map.values()).map(yCommentToComment)
        : root.array.toArray().map(yCommentToComment);
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
    const root = getCommentsRoot(this.doc);
    const id = input.id ?? createId();
    const now = input.now ?? Date.now();

    this.transact(() => {
      const yComment = createYComment({
        id,
        cellRef: input.cellRef,
        kind: input.kind,
        author: input.author,
        now,
        content: input.content,
      });

      if (root.kind === "map") {
        root.map.set(id, yComment);
      } else {
        root.array.push([yComment]);
      }
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
    const yComment = this.getYComment(input.commentId);

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
    const yComment = this.getYComment(input.commentId);
    const now = input.now ?? Date.now();

    this.transact(() => {
      yComment.set("resolved", input.resolved);
      yComment.set("updatedAt", now);
    });
  }

  setCommentContent(input: { commentId: string; content: string; now?: number }): void {
    const yComment = this.getYComment(input.commentId);
    const now = input.now ?? Date.now();

    this.transact(() => {
      yComment.set("content", input.content);
      yComment.set("updatedAt", now);
    });
  }

  setReplyContent(input: { commentId: string; replyId: string; content: string; now?: number }): void {
    const yComment = this.getYComment(input.commentId);

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

  private getYComment(commentId: string): Y.Map<unknown> {
    const root = getCommentsRoot(this.doc);
    if (root.kind === "map") {
      const yComment = root.map.get(commentId);
      if (!yComment) {
        throw new Error(`Comment not found: ${commentId}`);
      }
      return yComment;
    }

    for (const item of root.array.toArray()) {
      if (String(item.get("id") ?? "") === commentId) {
        return item;
      }
    }

    throw new Error(`Comment not found: ${commentId}`);
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

/**
 * Get the canonical comments Map (`Y.Map<id, Y.Map<comment>>`).
 *
 * This will throw if the current doc uses the legacy Array schema; in that case
 * use `getCommentsRoot` / `CommentManager` (which supports both), or migrate
 * first with `migrateCommentsArrayToMap`.
 */
export function getCommentsMap(doc: Y.Doc): YCommentsMap {
  const root = getCommentsRoot(doc);
  if (root.kind !== "map") {
    throw new Error(
      'Comments root "comments" is a Y.Array (legacy schema). Use getCommentsRoot/CommentManager or call migrateCommentsArrayToMap(doc) before getCommentsMap().',
    );
  }
  return root.map;
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

/**
 * Migrate legacy Array-backed comments (`Y.Array<Y.Map>`) to the canonical
 * Map-backed schema (`Y.Map<string, Y.Map>`).
 *
 * Strategy: rename the legacy `comments` Array root to `comments_legacy*` (so we
 * don't lose access to its content), then create the canonical Map root under
 * the original `comments` name and copy entries keyed by comment id.
 *
 * This should be called immediately after loading/applying the document
 * snapshot, before any code calls `doc.getMap("comments")` directly. If
 * `getMap("comments")` is called first on an Array-backed doc, the legacy array
 * content can become inaccessible.
 */
export function migrateCommentsArrayToMap(doc: Y.Doc, opts: { origin?: unknown } = {}): boolean {
  if (!doc.share.has("comments")) return false;
  const root = getCommentsRoot(doc);
  if (root.kind !== "array") return false;

  const legacy = root.array;
  doc.transact(
    () => {
      /** @type {Array<[string, Y.Map<unknown>]>} */
      const entries: Array<[string, Y.Map<unknown>]> = [];
      for (const item of legacy.toArray()) {
        if (!(item instanceof Y.Map)) continue;
        const id = String(item.get("id") ?? "");
        if (!id) continue;
        entries.push([id, cloneYjsValue(item) as Y.Map<unknown>]);
      }

      // Root types are schema-defined by name; once "comments" is instantiated as
      // an Array, Yjs will refuse to create a Map root with the same name.
      //
      // To migrate in-place, we "tombstone" the legacy array under a new root
      // name and then create the canonical Map root at "comments".
      const legacyRootName = findAvailableLegacyRootName(doc);
      doc.share.set(legacyRootName, legacy);
      doc.share.delete("comments");

      const map = doc.getMap("comments") as YCommentsMap;
      for (const [id, value] of entries) {
        map.set(id, value);
      }
    },
    opts.origin ?? "comments-migrate-array-to-map",
  );

  return true;
}

function findAvailableLegacyRootName(doc: Y.Doc): string {
  const base = "comments_legacy";
  if (!doc.share.has(base)) return base;
  for (let i = 1; i < 1000; i += 1) {
    const name = `${base}_${i}`;
    if (!doc.share.has(name)) return name;
  }
  return `${base}_${Date.now()}`;
}

function cloneYjsValue(value: any): any {
  if (value instanceof Y.Map) {
    const out = new Y.Map();
    value.forEach((v: any, k: string) => {
      out.set(k, cloneYjsValue(v));
    });
    return out;
  }

  if (value instanceof Y.Array) {
    const out = new Y.Array();
    for (const item of value.toArray()) {
      out.push([cloneYjsValue(item)]);
    }
    return out;
  }

  if (value instanceof Y.Text) {
    const out = new Y.Text();
    out.applyDelta(structuredClone(value.toDelta()));
    return out;
  }

  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}
