import * as Y from "yjs";
import { getYArray, getYMap, getYText, isYAbstractType, replaceForeignRootType } from "@formula/collab-yjs-utils";

import type { Comment, CommentAuthor, CommentKind, Reply } from "./types.ts";

export interface CommentManagerOptions {
  transact?: (fn: () => void) => void;
  /**
   * Optional permission guard for comment mutations.
   *
   * When provided and it returns `false`, all mutating APIs will throw:
   * `Error("Permission denied: cannot comment")`.
   *
   * Read-only APIs (`listAll`, `listForCell`) remain available regardless of
   * permissions so viewers can still read existing threads.
   */
  canComment?: () => boolean;
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

  const existingMap = getYMap(existing);
  if (existingMap) return { kind: "map", map: existingMap as YCommentsMap };

  const existingArray = getYArray(existing);
  if (existingArray) return { kind: "array", array: existingArray as YCommentsArray };

  // Root types may be a generic `AbstractType` placeholder until a constructor is
  // chosen. Peek at its internal structure before choosing a constructor.
  const placeholder = existing as any;
  const hasStart = placeholder?._start != null; // sequence item => likely array
  const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
  const kind: CommentsRoot["kind"] = hasStart && mapSize === 0 ? "array" : "map";

  if (kind === "array") {
    // If another Yjs module instance called `Doc.prototype.get(name)` (defaulting
    // to `AbstractType`) on this doc, `doc.share.get("comments")` can be a foreign
    // `AbstractType` placeholder. Calling `doc.getArray("comments")` from this
    // module would then throw "different constructor".
    //
    // In that case, re-wrap the placeholder into this module's constructor
    // directly (mirrors Yjs' `Doc.get()` conversion logic).
    // A foreign `AbstractType` placeholder can be patched to pass
    // `instanceof Y.AbstractType` checks (see undo service prototype patching),
    // so use constructor identity to detect foreign placeholders.
    if (doc instanceof Y.Doc && isYAbstractType(existing) && (existing as any).constructor !== Y.AbstractType) {
      return {
        kind: "array",
        array: replaceForeignRootType({
          doc,
          name: "comments",
          existing: placeholder,
          create: () => new Y.Array(),
        }) as YCommentsArray,
      };
    }
    return { kind: "array", array: doc.getArray("comments") as YCommentsArray };
  }

  // Prefer preserving foreign Yjs constructors (CJS vs ESM) when the doc was
  // hydrated by a foreign build. If we instantiate the root as a local `Y.Map`
  // while its entries were created by a foreign Yjs instance, we end up with a
  // mixed-module tree that breaks constructor checks in Yjs UndoManager and can
  // cause `doc.getMap("comments")` to throw later.
  const placeholderMap = placeholder?._map;
  if (placeholderMap instanceof Map) {
    for (const item of placeholderMap.values()) {
      if (!item || item.deleted) continue;
      const content = item.content?.getContent?.() ?? [];
      const value = content[content.length - 1];
      const yValueMap = getYMap(value);
      if (!yValueMap) continue;
      if (yValueMap instanceof Y.Map) break; // already local; fall back to `doc.getMap`

      const MapCtor = (yValueMap as any).constructor as new () => any;
      if (typeof MapCtor === "function") {
        const map = replaceForeignRootType({
          doc,
          name: "comments",
          existing: placeholder,
          create: () => new MapCtor(),
        }) as YCommentsMap;
        return { kind: "map", map };
      }
      break;
    }
  }

  // Fallback: choose the canonical map schema. Avoid calling `doc.getMap` when
  // the existing placeholder is a foreign `AbstractType` instance, which would
  // throw "different constructor" on local docs.
  if (doc instanceof Y.Doc && isYAbstractType(existing) && (existing as any).constructor !== Y.AbstractType) {
    const map = replaceForeignRootType({
      doc,
      name: "comments",
      existing: placeholder,
      create: () => new Y.Map(),
    }) as YCommentsMap;
    return { kind: "map", map };
  }

  return { kind: "map", map: doc.getMap("comments") as YCommentsMap };
}

function cloneToLocalYjsValue(value: unknown): unknown {
  const map = getYMap(value);
  if (map) {
    const out = new Y.Map();
    map.forEach((v, k) => out.set(k, cloneToLocalYjsValue(v)));
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new Y.Array();
    for (const item of array.toArray()) {
      out.push([cloneToLocalYjsValue(item)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const out = new Y.Text();
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}

function normalizeCommentsRootToLocalTypes(doc: Y.Doc): void {
  const root = getCommentsRoot(doc);
  if (root.kind !== "map") return;

  // Only normalize when the root itself is from the local Yjs module instance.
  // If the root was created by a foreign build (CJS vs ESM) we cannot safely
  // mix local Yjs types into it.
  if (!(root.map instanceof Y.Map)) return;

  const replacements: Array<{ key: string; cloned: Y.Map<unknown> }> = [];
  root.map.forEach((value, key) => {
    const map = getYMap(value);
    if (!map) return;
    // Already local.
    if (map instanceof Y.Map) return;
    replacements.push({ key, cloned: cloneToLocalYjsValue(map) as Y.Map<unknown> });
  });
  if (replacements.length === 0) return;

  // Use a non-local origin so collaborative undo managers don't capture this
  // normalization as a user edit.
  doc.transact(() => {
    for (const { key, cloned } of replacements) {
      root.map.set(key, cloned);
    }
  });
}

export class CommentManager {
  private readonly doc: Y.Doc;
  private readonly transact: (fn: () => void) => void;
  private readonly canComment: () => boolean;

  constructor(doc: Y.Doc, options: CommentManagerOptions = {}) {
    this.doc = doc;
    this.transact = options.transact ?? ((fn) => doc.transact(fn));
    this.canComment = options.canComment ?? (() => true);
  }

  listAll(): Comment[] {
    // Pre-hydration safety: avoid instantiating the comments root when it doesn't
    // exist yet. Older documents may still use a legacy Array-backed schema; if
    // we create a Map root too early, the legacy array content can become
    // inaccessible once remote updates arrive.
    if (!this.doc.share.get("comments")) return [];
    const root = getCommentsRoot(this.doc);
    const entries = this.getAllYComments(root);
    const comments = entries.map(({ yComment, mapKey }) => yCommentToComment(yComment, mapKey));
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
    this.assertCanComment();
    const root = getCommentsRoot(this.doc);
    const id = input.id ?? createId();
    const now = input.now ?? Date.now();

    this.transact(() => {
      const mapConstructor = (() => {
        if (root.kind === "map") return (root.map as any).constructor as new () => any;

        // Local Yjs instance.
        if (root.array instanceof Y.Array) return Y.Map;

        // Foreign Yjs instance: try to reuse an existing comment map constructor
        // (if the array already has comments), otherwise derive it from the doc.
        const existingComment = this.getAllYComments(root)[0]?.yComment;
        if (existingComment) return (existingComment as any).constructor as new () => any;
        return getDocMapConstructor(this.doc);
      })();

      const arrayConstructor = (() => {
        // For local (ESM) Y.Maps, always use the local Y.Array constructor.
        // Mixing module instances here causes `typeMapSet` to throw.
        if (mapConstructor === Y.Map) return Y.Array;

        // Foreign Yjs instance: prefer the replies array constructor from any
        // existing comment, otherwise derive it from the doc.
        const existingComment = this.getAllYComments(root)[0]?.yComment;
        const existingReplies = existingComment ? getYArray(existingComment.get("replies")) : null;
        if (existingReplies) return existingReplies.constructor as new () => any;

        if (root.kind === "array") return (root.array as any).constructor as new () => any;
        return getDocArrayConstructor(this.doc);
      })();

      const yComment = createYComment({
        id,
        cellRef: input.cellRef,
        kind: input.kind,
        author: input.author,
        now,
        content: input.content,
      }, {
        mapConstructor,
        arrayConstructor,
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
    this.assertCanComment();
    const yComment = this.getYCommentForWrite(input.commentId);

    const replies = yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined;
    if (!replies) {
      throw new Error(`Comment replies missing: ${input.commentId}`);
    }

    const id = input.id ?? createId();
    const now = input.now ?? Date.now();
    const mapConstructor = (() => {
      // Local replies array => create local Y.Maps.
      if ((replies as any) instanceof Y.Array) return undefined;

      // Prefer the constructor of any existing reply item.
      const existingReply = replies.toArray().find((reply) => getYMap(reply));
      if (existingReply) return (existingReply as any).constructor as any;

      // Fallback: comment map is usually from the same Yjs module instance as
      // its `replies` array.
      const commentMap = getYMap(yComment);
      if (commentMap && !(commentMap instanceof Y.Map)) return commentMap.constructor as any;
      return undefined;
    })();

    this.transact(() => {
      replies.push([
        createYReply({
          id,
          author: input.author,
          now,
          content: input.content,
        }, mapConstructor ? { mapConstructor } : undefined),
      ]);
      yComment.set("updatedAt", now);
    });

    return id;
  }

  setResolved(input: { commentId: string; resolved: boolean; now?: number }): void {
    this.assertCanComment();
    const yComment = this.getYCommentForWrite(input.commentId);
    const now = input.now ?? Date.now();

    this.transact(() => {
      yComment.set("resolved", input.resolved);
      yComment.set("updatedAt", now);
    });
  }

  setCommentContent(input: { commentId: string; content: string; now?: number }): void {
    this.assertCanComment();
    const yComment = this.getYCommentForWrite(input.commentId);
    const now = input.now ?? Date.now();

    this.transact(() => {
      yComment.set("content", input.content);
      yComment.set("updatedAt", now);
    });
  }

  setReplyContent(input: { commentId: string; replyId: string; content: string; now?: number }): void {
    this.assertCanComment();
    const yComment = this.getYCommentForWrite(input.commentId);

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

  /**
   * Ensure the returned comment map (and its nested Yjs types) are from this module's
   * Yjs instance before mutating them.
   *
   * UndoManager (from this module's `yjs` import) cannot reliably undo edits that
   * were applied through a different Yjs module instance (e.g. CJS `applyUpdate`),
   * because map overwrites delete the previous item and later need to "redo" it.
   * If the previous item is owned by a foreign Yjs build, UndoManager will skip it.
   *
   * We avoid that by cloning foreign comment trees into local types in an *untracked*
   * transaction (no origin), and then performing the actual edit in the session's
   * tracked transaction.
   */
  private getYCommentForWrite(commentId: string): Y.Map<unknown> {
    const root = getCommentsRoot(this.doc);
    const yComment = this.getYComment(commentId);

    // If we can't insert local types into the root (e.g. the root itself was
    // created by a foreign Yjs instance), fall back to the original comment map.
    // This keeps cross-instance CommentManager usage working even when undo isn't
    // involved.
    if (root.kind === "map" && !(root.map instanceof Y.Map)) return yComment;
    if (root.kind === "array" && !(root.array instanceof Y.Array)) return yComment;

    if (!hasForeignYjsTypes(yComment)) return yComment;

    if (root.kind === "array") {
      const index = root.array
        .toArray()
        .findIndex((item) => item === yComment || String(item.get("id") ?? "") === commentId);
      if (index < 0) return yComment;

      const local = cloneYjsValueToLocal(yComment) as Y.Map<unknown>;
      this.doc.transact(() => {
        root.array.delete(index, 1);
        root.array.insert(index, [local]);
      });
      return local;
    }

    const local = cloneYjsValueToLocal(yComment) as Y.Map<unknown>;
    this.doc.transact(() => {
      root.map.set(commentId, local);
    });
    return local;
  }

  private getYComment(commentId: string): Y.Map<unknown> {
    const root = getCommentsRoot(this.doc);
    if (root.kind === "map") {
      const yComment = getYMap(root.map.get(commentId));
      if (!yComment) {
        // Some historical docs used a legacy array schema and were later
        // "clobbered" by instantiating the root as a Map. In that case the
        // underlying list items still exist, but are not present in `map.get`.
        const legacy = findLegacyListCommentById(root.map, commentId);
        if (!legacy) {
          throw new Error(`Comment not found: ${commentId}`);
        }
        return legacy;
      }
      return yComment;
    }

    const mapEntry = findMapEntryCommentByKey(root.array, commentId);
    if (mapEntry) return mapEntry;

    for (const item of root.array.toArray()) {
      if (String(item.get("id") ?? "") === commentId) {
        return item;
      }
    }

    throw new Error(`Comment not found: ${commentId}`);
  }

  private getAllYComments(root: CommentsRoot): Array<{ yComment: Y.Map<unknown>; mapKey?: string }> {
    if (root.kind === "array") {
      const byId = new Map<string, { yComment: Y.Map<unknown>; mapKey?: string }>();
      for (const { key, value } of iterMapEntryComments(root.array)) {
        byId.set(key, { yComment: value, mapKey: key });
      }
      for (const yComment of root.array.toArray()) {
        const id = String(yComment.get("id") ?? "");
        if (!id) continue;
        if (byId.has(id)) continue;
        byId.set(id, { yComment });
      }
      return Array.from(byId.values());
    }

    // Canonical Map entries (keyed by id).
    const byId = new Map<string, { yComment: Y.Map<unknown>; mapKey?: string }>();
    root.map.forEach((value, key) => {
      const map = getYMap(value);
      if (map) byId.set(key, { yComment: map, mapKey: key });
    });

    // Legacy array items that ended up inside a Map root (see comment in
    // `getYComment` above).
    for (const yComment of iterLegacyListComments(root.map)) {
      const id = String(yComment.get("id") ?? "");
      if (!id) continue;
      if (byId.has(id)) continue;
      byId.set(id, { yComment });
    }

    return Array.from(byId.values());
  }

  private assertCanComment(): void {
    let allowed = false;
    try {
      allowed = this.canComment();
    } catch {
      // Fail closed. Do not leak permission/token details from guard callbacks.
      allowed = false;
    }
    if (allowed) return;
    throw new Error("Permission denied: cannot comment");
  }
}

export function createCommentManagerForDoc(params: (
  | { doc: Y.Doc; transactLocal: (fn: () => void) => void; canComment?: () => boolean }
  | { doc: Y.Doc; transact: (fn: () => void) => void; canComment?: () => boolean }
)): CommentManager {
  // In Node environments, remote updates can be applied using a different Yjs
  // module instance (CJS vs ESM). This can leave nested comment maps (the values
  // inside the `comments` root map) backed by foreign Item constructors, which
  // breaks collaborative undo (Y.UndoManager relies on strict `instanceof` checks
  // when applying undo/redo).
  //
  // Normalize any foreign nested types into local Yjs instances so comment edits
  // are undoable even when the doc was hydrated by a foreign Yjs build.
  //
  // Important: callers may create a CommentManager before the doc has been
  // hydrated by a provider. In that case the `comments` root may not exist yet.
  // We must not instantiate it eagerly (it could clobber legacy Array-backed docs
  // by fixing the root kind too early).
  let didAttemptNormalize = false;
  const maybeNormalize = () => {
    if (didAttemptNormalize) return;
    if (!params.doc.share.get("comments")) return;
    // View-only clients should not generate Yjs updates (including best-effort
    // normalization transactions). If permissions are dynamic, we intentionally
    // *do not* mark normalization as attempted so a later role upgrade (or other
    // permission change) can enable normalization on the next write.
    if (typeof params.canComment === "function") {
      try {
        if (!params.canComment()) return;
      } catch {
        return;
      }
    }
    // Only "finish" normalization once the comments root is known to be safe to
    // normalize (either it's a local `Y.Map`, or it's an Array schema where map
    // normalization doesn't apply).
    //
    // If the root exists but is a foreign `Y.Map` constructor (cross-module Yjs),
    // normalization cannot safely insert local types into it. In that case we
    // keep trying on subsequent transactions in case some other code replaces the
    // root with a local constructor (e.g. a session/binder undo-scope helper).
    try {
      const root = getCommentsRoot(params.doc);
      if (root.kind === "map" && !(root.map instanceof Y.Map)) return;
      didAttemptNormalize = true;
    } catch {
      didAttemptNormalize = true;
      return;
    }
    try {
      normalizeCommentsRootToLocalTypes(params.doc);
    } catch {
      // Best-effort; never block comment usage on normalization.
    }
  };
  // Normalize immediately if possible, and otherwise defer until the first local
  // comment transaction after the root exists.
  maybeNormalize();

  // Be careful to preserve the caller's `this` binding.
  //
  // `CollabSession.transactLocal` is a class method, so extracting it into a local
  // variable (or passing it through) would lose `this` and crash under strict-mode
  // invocation. Always invoke via the originating object.
  return new CommentManager(params.doc, {
    transact: (fn) => {
      // Ensure the comments root is normalized to local Yjs constructors (when
      // possible) before applying the user's tracked transaction.
      maybeNormalize();
      if ("transactLocal" in params) params.transactLocal(fn);
      else params.transact(fn);
    },
    // Preserve the caller's `this` binding (e.g. `CollabSession.canComment`).
    canComment: typeof params.canComment === "function" ? () => params.canComment() : undefined,
  });
}

export function createCommentManagerForSession(session: { doc: Y.Doc; transactLocal: (fn: () => void) => void }): CommentManager {
  return createCommentManagerForDoc(session);
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

export function yCommentToComment(yComment: Y.Map<unknown>, mapKey?: string): Comment {
  const replies = (yComment.get("replies") as Y.Array<Y.Map<unknown>> | undefined)?.toArray().map(yReplyToReply) ?? [];

  return {
    // For the canonical Map schema, the map key is authoritative.
    id: String(mapKey ?? yComment.get("id") ?? ""),
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

function iterLegacyListComments(type: any): Y.Map<unknown>[] {
  const out: Y.Map<unknown>[] = [];
  let item = type?._start ?? null;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      const content = item.content?.getContent?.() ?? [];
      for (const value of content) {
        const map = getYMap(value);
        if (map) out.push(map);
      }
    }
    item = item.right;
  }
  return out;
}

function findLegacyListCommentById(type: any, commentId: string): Y.Map<unknown> | null {
  for (const yComment of iterLegacyListComments(type)) {
    if (String(yComment.get("id") ?? "") === commentId) return yComment;
  }
  return null;
}

function iterMapEntryComments(type: any): Array<{ key: string; value: Y.Map<unknown> }> {
  const out: Array<{ key: string; value: Y.Map<unknown> }> = [];
  const map = type?._map;
  if (!(map instanceof Map)) return out;
  for (const [key, item] of map.entries()) {
    if (!item || item.deleted) continue;
    const content = item.content?.getContent?.() ?? [];
    const value = content[content.length - 1];
    const yMap = getYMap(value);
    if (yMap) out.push({ key, value: yMap });
  }
  out.sort((a, b) => a.key.localeCompare(b.key));
  return out;
}

function findMapEntryCommentByKey(type: any, key: string): Y.Map<unknown> | null {
  const map = type?._map;
  if (!(map instanceof Map)) return null;
  const item = map.get(key);
  if (!item || item.deleted) return null;
  const content = item.content?.getContent?.() ?? [];
  const value = content[content.length - 1];
  return getYMap(value);
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
}, opts?: { mapConstructor?: new () => any; arrayConstructor?: new () => any }): Y.Map<unknown>;

export function createYComment(
  input: {
    id: string;
    cellRef: string;
    kind: CommentKind;
    author: CommentAuthor;
    now: number;
    content: string;
  },
  opts: { mapConstructor?: new () => any; arrayConstructor?: new () => any } = {},
): Y.Map<unknown> {
  const MapCtor = opts.mapConstructor ?? Y.Map;
  const ArrayCtor = opts.arrayConstructor ?? Y.Array;
  const yComment = new MapCtor();
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
  yComment.set("replies", new ArrayCtor());
  return yComment;
}

export function createYReply(input: {
  id: string;
  author: CommentAuthor;
  now: number;
  content: string;
}, opts?: { mapConstructor?: new () => any }): Y.Map<unknown>;

export function createYReply(
  input: {
    id: string;
    author: CommentAuthor;
    now: number;
    content: string;
  },
  opts: { mapConstructor?: new () => any } = {},
): Y.Map<unknown> {
  const MapCtor = opts.mapConstructor ?? Y.Map;
  const yReply = new MapCtor();
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
  const legacyList =
    root.kind === "array"
      ? root.array.toArray()
      : // Some historical docs were instantiated as a Map even though their
        // content is a legacy list (array) of comment maps.
        iterLegacyListComments(root.map);

  const hasLegacyListContent = legacyList.length > 0;
  if (root.kind !== "array" && !hasLegacyListContent) return false;

  const legacyRoot: Y.AbstractType<any> = root.kind === "array" ? root.array : root.map;
  doc.transact(
    () => {
      const ctors = getDocTypeConstructors(doc);
      /** @type {Map<string, Y.Map<unknown>>} */
      const entries = new Map<string, Y.Map<unknown>>();

      // Canonical Map entries (if present).
      if (root.kind === "map") {
        root.map.forEach((value, key) => {
          const map = getYMap(value);
          if (!map) return;
          entries.set(key, cloneYjsValue(map, ctors) as Y.Map<unknown>);
        });
      }

      // Map entries stored on an Array root (mixed-schema docs).
      if (root.kind === "array") {
        for (const { key, value } of iterMapEntryComments(root.array)) {
          if (entries.has(key)) continue;
          entries.set(key, cloneYjsValue(value, ctors) as Y.Map<unknown>);
        }
      }

      // Legacy list entries (array schema, or clobbered Map schema).
      for (const item of legacyList) {
        const map = getYMap(item);
        if (!map) continue;
        const id = String(map.get("id") ?? "");
        if (!id) continue;
        if (entries.has(id)) continue;
        entries.set(id, cloneYjsValue(map, ctors) as Y.Map<unknown>);
      }

      // Root types are schema-defined by name; once "comments" is instantiated as
      // an Array, Yjs will refuse to create a Map root with the same name.
      //
      // To migrate in-place, we "tombstone" the legacy array under a new root
      // name and then create the canonical Map root at "comments".
      const legacyRootName = findAvailableLegacyRootName(doc);
      doc.share.set(legacyRootName, legacyRoot);
      doc.share.delete("comments");

      const map = doc.getMap("comments") as YCommentsMap;
      for (const [id, value] of entries.entries()) {
        map.set(id, value);
      }
    },
    opts.origin ?? "comments-migrate-array-to-map",
  );

  return true;
}

function findAvailableLegacyRootName(doc: Y.Doc): string {
  return findAvailableRootName(doc, "comments_legacy");
}

function findAvailableRootName(doc: Y.Doc, base: string): string {
  if (!doc.share.has(base)) return base;
  for (let i = 1; i < 1000; i += 1) {
    const name = `${base}_${i}`;
    if (!doc.share.has(name)) return name;
  }
  return `${base}_${Date.now()}`;
}

function getDocArrayConstructor(doc: any): new () => any {
  const name = findAvailableRootName(doc, "__comments_tmp_array");
  const tmp = doc.getArray(name);
  const ctor = tmp.constructor as new () => any;
  doc.share.delete(name);
  return ctor;
}

function getDocMapConstructor(doc: any): new () => any {
  const name = findAvailableRootName(doc, "__comments_tmp_map");
  const tmp = doc.getMap(name);
  const ctor = tmp.constructor as new () => any;
  doc.share.delete(name);
  return ctor;
}

type DocTypeConstructors = {
  MapCtor: new () => any;
  ArrayCtor: new () => any;
  TextCtor: new () => any;
};

function getDocTypeConstructors(doc: any): DocTypeConstructors {
  return {
    MapCtor: getDocMapConstructor(doc),
    ArrayCtor: getDocArrayConstructor(doc),
    TextCtor: getDocTextConstructor(doc),
  };
}

function getDocTextConstructor(doc: any): new () => any {
  const name = findAvailableRootName(doc, "__comments_tmp_text");
  const tmp = doc.getText(name);
  const ctor = tmp.constructor as new () => any;
  doc.share.delete(name);
  return ctor;
}

function cloneYjsValue(value: any, ctors: DocTypeConstructors): any {
  const map = getYMap(value);
  if (map) {
    const out = new ctors.MapCtor();
    map.forEach((v: any, k: string) => {
      out.set(k, cloneYjsValue(v, ctors));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new ctors.ArrayCtor();
    for (const item of array.toArray()) {
      out.push([cloneYjsValue(item, ctors)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const out = new ctors.TextCtor();
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  if (Array.isArray(value)) {
    return value.map((v) => cloneYjsValue(v, ctors));
  }

  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}

function hasForeignYjsTypes(value: unknown, seen: Set<any> = new Set()): boolean {
  if (!value || typeof value !== "object") return false;
  if (seen.has(value)) return false;
  seen.add(value);

  const map = getYMap(value);
  if (map) {
    if (!(map instanceof Y.Map)) return true;
    let foreign = false;
    map.forEach((v: unknown) => {
      if (hasForeignYjsTypes(v, seen)) foreign = true;
    });
    return foreign;
  }

  const array = getYArray(value);
  if (array) {
    if (!(array instanceof Y.Array)) return true;
    for (const item of array.toArray()) {
      if (hasForeignYjsTypes(item, seen)) return true;
    }
    return false;
  }

  const text = getYText(value);
  if (text) {
    return !(text instanceof Y.Text);
  }

  return false;
}

function cloneYjsValueToLocal(value: any, seen: Map<any, any> = new Map()): any {
  if (value && typeof value === "object") {
    const cached = seen.get(value);
    if (cached) return cached;
  }

  const map = getYMap(value);
  if (map) {
    const out = new Y.Map();
    seen.set(value, out);
    map.forEach((v: any, k: string) => {
      out.set(k, cloneYjsValueToLocal(v, seen));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new Y.Array();
    seen.set(value, out);
    for (const item of array.toArray()) {
      out.push([cloneYjsValueToLocal(item, seen)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const out = new Y.Text();
    seen.set(value, out);
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  if (value && typeof value === "object") {
    return structuredClone(value);
  }

  return value;
}
