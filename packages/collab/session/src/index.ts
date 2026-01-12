import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import { attachOfflinePersistence } from "@formula/collab-offline";
import { PresenceManager } from "@formula/collab-presence";
import { createUndoService, type UndoService } from "@formula/collab-undo";
import {
  CellConflictMonitor,
  CellStructuralConflictMonitor,
  FormulaConflictMonitor,
  type CellConflict,
  type CellStructuralConflict,
  type FormulaConflict,
} from "@formula/collab-conflicts";
import { ensureWorkbookSchema, getWorkbookRoots } from "@formula/collab-workbook";
import {
  decryptCellPlaintext,
  encryptCellPlaintext,
  isEncryptedCellPayload,
  type CellEncryptionKey,
  type CellPlaintext,
} from "@formula/collab-encryption";
import type { CollabPersistence, CollabPersistenceBinding } from "@formula/collab-persistence";

import { assertValidRole, getCellPermissions, maskCellValue } from "../../permissions/index.js";
import {
  makeCellKey as makeCellKeyImpl,
  normalizeCellKey as normalizeCellKeyImpl,
  parseCellKey as parseCellKeyImpl,
} from "./cell-key.js";

function getYMap(value: unknown): any | null {
  if (value instanceof Y.Map) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value as any;
  if (maybe.constructor?.name !== "YMap") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.keys !== "function") return null;
  if (typeof maybe.forEach !== "function") return null;
  return maybe;
}

function getYArray(value: unknown): any | null {
  if (value instanceof Y.Array) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value as any;
  if (maybe.constructor?.name !== "YArray") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.toArray !== "function") return null;
  if (typeof maybe.push !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return maybe;
}

function replaceForeignRootType<T extends Y.AbstractType<any>>(params: {
  doc: Y.Doc;
  name: string;
  existing: any;
  create: () => T;
}): T {
  const { doc, name, existing, create } = params;
  const t = create();

  // Copy the internal linked-list structures that hold the CRDT content. This mirrors
  // Yjs' own `Doc.get()` logic when converting a root placeholder (`AbstractType`)
  // into a concrete type, but also supports the case where the existing root was
  // created by a different Yjs module instance (e.g. CJS `applyUpdate`).
  (t as any)._map = existing?._map;
  (t as any)._start = existing?._start;
  (t as any)._length = existing?._length;

  // Update parent pointers so future updates can resolve the correct root key via
  // `findRootTypeKey` when encoding.
  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item: any) => {
      for (let n = item; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t as any);
  (t as any)._integrate(doc as any, null);
  return t;
}

function getCommentsRootForUndoScope(doc: Y.Doc): Y.AbstractType<any> {
  // Yjs root types are schema-defined: you must know whether a key is a Map or
  // Array. When applying updates into a fresh Doc, root types can temporarily
  // appear as a generic `AbstractType` until a constructor is chosen.
  //
  // Importantly, calling `doc.getMap("comments")` on an Array-backed root can
  // define it as a Map and make the array content inaccessible. To support both
  // historical schemas (Map or Array) we peek at the underlying state before
  // choosing a constructor.
  const existing = doc.share.get("comments");
  let root: Y.AbstractType<any>;

  if (!existing) {
    root = doc.getMap("comments");
  } else {
    const existingMap = getYMap(existing);
    if (existingMap) {
      root =
        existingMap instanceof Y.Map
          ? existingMap
          : replaceForeignRootType({ doc, name: "comments", existing: existingMap, create: () => new Y.Map() });
    } else {
      const existingArray = getYArray(existing);
      if (existingArray) {
        root =
          existingArray instanceof Y.Array
            ? existingArray
            : replaceForeignRootType({
                doc,
                name: "comments",
                existing: existingArray,
                create: () => new Y.Array(),
              });
      } else {
        const placeholder = existing as any;
        const hasStart = placeholder?._start != null; // sequence item => likely array
        const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
        const kind = hasStart && mapSize === 0 ? "array" : "map";
        root = kind === "array" ? doc.getArray("comments") : doc.getMap("comments");
      }
    }
  }

  // If updates were applied using a different Yjs module instance (e.g. y-websocket
  // applying updates via CommonJS `require("yjs")` while the app uses ESM imports),
  // the `comments` root can contain nested Y.Maps/Y.Arrays whose constructors do
  // not match this module instance.
  //
  // Yjs UndoManager relies on constructor checks, so undo may fail to revert
  // comment edits unless we normalize those nested types up front.
  normalizeCommentsForUndoScope(doc, root);

  return root;
}

export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

function normalizeCommentsForUndoScope(doc: Y.Doc, root: Y.AbstractType<any>): void {
  if (root instanceof Y.Map) {
    /** @type {Array<[string, any]>} */
    const foreignComments: Array<[string, any]> = [];
    root.forEach((value, key) => {
      const yComment = getYMap(value);
      if (!yComment) return;
      // Foreign Yjs instances fail `instanceof` checks, but we still want to
      // normalize them to local constructors before UndoManager is created.
      if (yComment instanceof Y.Map) return;
      foreignComments.push([String(key), yComment]);
    });

    if (foreignComments.length === 0) return;

    doc.transact(() => {
      for (const [key, yComment] of foreignComments) {
        root.set(key, cloneForeignCommentToLocal(yComment));
      }
    });
    return;
  }

  if (!(root instanceof Y.Array)) return;

  const items = root.toArray();
  /** @type {Array<{ index: number, yComment: any }>} */
  const replacements: Array<{ index: number; yComment: any }> = [];
  for (let i = 0; i < items.length; i += 1) {
    const yComment = getYMap(items[i]);
    if (!yComment) continue;
    if (yComment instanceof Y.Map) continue;
    replacements.push({ index: i, yComment });
  }

  if (replacements.length === 0) return;

  doc.transact(() => {
    // Replace from back-to-front so indices remain stable.
    for (let i = replacements.length - 1; i >= 0; i -= 1) {
      const replacement = replacements[i];
      if (!replacement) continue;
      const cloned = cloneForeignCommentToLocal(replacement.yComment);
      root.delete(replacement.index, 1);
      root.insert(replacement.index, [cloned]);
    }
  });
}

function cloneForeignCommentToLocal(comment: any): Y.Map<unknown> {
  const out = new Y.Map<unknown>();

  if (comment && typeof comment.forEach === "function") {
    comment.forEach((value: any, key: string) => {
      if (key === "replies") return;
      out.set(key, normalizeForeignScalar(value));
    });
  }

  const replies = comment?.get?.("replies");
  out.set("replies", cloneForeignRepliesToLocal(replies));
  return out;
}

function cloneForeignRepliesToLocal(replies: any): Y.Array<Y.Map<unknown>> {
  const out = new Y.Array<Y.Map<unknown>>();
  const yReplies = getYArray(replies);
  if (!yReplies) return out;

  const items = typeof yReplies.toArray === "function" ? yReplies.toArray() : [];
  for (const item of items) {
    const yReply = getYMap(item);
    if (!yReply) continue;

    const reply = new Y.Map<unknown>();
    yReply.forEach((value: any, key: string) => {
      reply.set(key, normalizeForeignScalar(value));
    });
    out.push([reply]);
  }

  return out;
}

function normalizeForeignScalar(value: any): any {
  if (value == null) return value;

  const yMap = getYMap(value);
  if (yMap && !(yMap instanceof Y.Map)) {
    /** @type {Record<string, any>} */
    const obj = {};
    yMap.forEach((v: any, k: string) => {
      obj[String(k)] = normalizeForeignScalar(v);
    });
    return obj;
  }

  const yArray = getYArray(value);
  if (yArray && !(yArray instanceof Y.Array)) {
    const arr = typeof yArray.toArray === "function" ? yArray.toArray() : [];
    return arr.map((v: any) => normalizeForeignScalar(v));
  }

  if (value instanceof Y.Text) return value.toString();
  if (value && typeof value === "object" && value.constructor?.name === "YText" && typeof value.toString === "function") {
    return value.toString();
  }

  return value;
}

export interface CellAddress {
  sheetId: string;
  row: number;
  col: number;
}

export interface SessionPermissions {
  role: DocumentRole;
  rangeRestrictions: unknown[];
  userId: string | null;
}

export interface SessionPresenceOptions {
  user: { id: string; name: string; color: string };
  activeSheet: string;
  throttleMs?: number;
  staleAfterMs?: number;
  now?: () => number;
  setTimeout?: typeof setTimeout;
  clearTimeout?: typeof clearTimeout;
}

export interface CollabSessionConnectionOptions {
  wsUrl: string;
  docId: string;
  token?: string;
  WebSocketPolyfill?: any;
  disableBc?: boolean;
  params?: Record<string, string>;
}

export type CollabSessionProvider = {
  awareness?: unknown;
  connect?: () => void;
  disconnect?: () => void;
  destroy?: () => void;
  on?: (event: string, cb: (...args: any[]) => void) => void;
  off?: (event: string, cb: (...args: any[]) => void) => void;
  synced?: boolean;
};

export interface CollabSessionOptions {
  /**
   * Stable identifier for this document. Required when `persistence` is enabled
   * and `connection` is not provided.
   */
  docId?: string;
  doc?: Y.Doc;
  /**
   * Convenience option to construct a y-websocket provider for this session.
   * When provided, `session.provider` will be a `WebsocketProvider` instance.
   */
  connection?: CollabSessionConnectionOptions;
  /**
   * Optional offline-first local persistence for this document.
   *
   * When combined with `connection`, CollabSession will ensure the persisted
   * state is loaded before connecting the sync provider so offline edits are
   * present when syncing.
   */
  persistence?: CollabPersistence;
  /**
   * Optional sync provider (e.g. y-websocket's WebsocketProvider). If provided,
   * we will use `provider.awareness` when constructing a PresenceManager.
   */
  provider?: CollabSessionProvider | null;
  /**
   * Awareness instance used for presence. Overrides `provider.awareness` when provided.
   */
  awareness?: unknown;
  /**
   * If provided, the session will create a PresenceManager and expose it via
   * `session.presence`.
   */
  presence?: SessionPresenceOptions;
  /**
   * Default sheet id used when parsing cell keys that omit a sheet identifier.
   */
  defaultSheetId?: string;
  /**
   * Workbook schema initialization options.
   *
   * When enabled (default), the session ensures the workbook schema roots exist
   * and creates a default sheet when the document has no sheets.
   *
   * When using a sync provider (e.g. y-websocket), initialization is deferred
   * until the first `sync=true` event so we don't create local default sheets
   * before hydration (which would later show up as duplicates).
   */
  schema?: { autoInit?: boolean; defaultSheetId?: string; defaultSheetName?: string };
  /**
   * Optional offline persistence configuration. When enabled, the session will
   * load/store Yjs updates locally so edits survive reload/crash and merge when
   * connectivity returns.
   */
  offline?: {
    mode: "indexeddb" | "file";
    /**
     * Storage key / namespace. Defaults to `connection.docId` (when provided)
     * and otherwise `doc.guid`.
     */
    key?: string;
    /**
     * File path for Node/desktop persistence when `mode: "file"` is used.
     */
    filePath?: string;
    /**
     * When false, offline persistence is only started when `session.offline.whenLoaded()`
     * is first called. Defaults to true.
     */
    autoLoad?: boolean;
    /**
     * When true (default when offline is enabled alongside `connection`), the
     * session delays WebSocket provider connection until offline state has
     * finished loading.
     */
    autoConnectAfterLoad?: boolean;
  };
  /**
   * When enabled, the session tracks local edits using Yjs' UndoManager so undo/redo
   * only affects this client's changes (never remote users' edits).
   */
  undo?: {
    captureTimeoutMs?: number;
    origin?: object;
    /**
     * Additional Yjs root map names to include in the collaborative undo scope.
     *
     * These roots are created eagerly (via `doc.getMap(name)`) when undo is
     * enabled so edits remain undoable even if the root is normally created
     * lazily after session construction.
     */
    scopeNames?: string[];
    /**
     * Advanced escape hatch to include additional Yjs root types in the undo
     * scope (e.g. `doc.getArray("foo")`).
     */
    includeRoots?: (doc: Y.Doc) => Array<Y.AbstractType<any>>;
  };
  /**
   * When enabled, the session monitors formula updates for true conflicts
   * (offline/concurrent same-cell edits) and surfaces them via `onConflict`.
   *
    * In `"formula+value"` mode, the monitor also detects concurrent value edits and
    * formula-vs-value "content" conflicts (e.g. one user writes a formula while
    * another concurrently writes a literal value).
   *
    * Note: `"formula+value"` overlaps with `cellValueConflicts` (both can surface
    * value conflicts). Prefer one or the other to avoid duplicated/confusing UX.
    *
    * Note: `remoteUserId` in emitted conflicts is best-effort and may be an empty
    * string when the writer does not update `modifiedBy` (legacy/edge clients).
    */
  formulaConflicts?: {
    localUserId: string;
    onConflict: (conflict: FormulaConflict) => void;
    /**
     * Deprecated/ignored. Former wall-clock heuristic for inferring concurrency.
     *
     * Conflict detection is now causal (Yjs-based) and works across long offline periods.
     *
     * @deprecated
     */
    concurrencyWindowMs?: number;
    mode?: "formula" | "formula+value";
    includeValueConflicts?: boolean;
  };
  /**
   * When enabled, the session monitors structural operations (moves / deletes)
   * for true offline conflicts and surfaces them via `onConflict`.
   *
   * Note: `remoteUserId` in emitted conflicts is best-effort and may be an empty
   * string when the remote op record does not include a `userId`.
   */
  cellConflicts?: {
    localUserId: string;
    onConflict: (conflict: CellStructuralConflict) => void;
    /**
     * Maximum number of structural op records to keep per user in the shared
     * `cellStructuralOps` Yjs log.
     *
     * Higher values improve conflict detection fidelity over longer offline
     * periods; lower values cap document growth.
     */
    maxOpRecordsPerUser?: number;
  };
  /**
   * When enabled, the session monitors cell value updates for true conflicts
   * (offline/concurrent same-cell edits) and surfaces them via `onConflict`.
   *
   * Note: if you enable `formulaConflicts` with `mode: "formula+value"`, value
   * conflicts are already covered there (and it can also surface formula-vs-value
   * "content" conflicts). Prefer one monitor to avoid redundant conflict UX.
   *
   * Note: `remoteUserId` in emitted conflicts is best-effort and may be an empty
   * string when the writer does not update `modifiedBy` (legacy/edge clients).
   */
  cellValueConflicts?: {
    localUserId: string;
    onConflict: (conflict: CellConflict) => void;
  };
  /**
   * Optional end-to-end encryption configuration for protecting specific cells.
   *
   * When enabled, cell values/formulas are encrypted *before* they are written
   * into the Yjs CRDT so unauthorized collaborators (and the sync server) cannot
   * read protected content.
   */
  encryption?: {
    keyForCell: (cell: CellAddress) => CellEncryptionKey | null;
    /**
     * Optional override for deciding whether a cell should be encrypted.
     * Defaults to `true` when `keyForCell` returns a non-null key.
     */
    shouldEncryptCell?: (cell: CellAddress) => boolean;
  };
}

export interface CollabCell {
  value: unknown;
  formula: string | null;
  modified: number | null;
  modifiedBy: string | null;
  /**
   * True when the cell is stored encrypted and this session cannot decrypt it
   * (so `value` is masked and `formula` is null).
   */
  encrypted?: boolean;
}

export function makeCellKey(cell: CellAddress): string {
  return makeCellKeyImpl(cell);
}

export function parseCellKey(
  key: string,
  options: { defaultSheetId?: string } = {}
): CellAddress | null {
  const parsed = parseCellKeyImpl(key, options);
  if (!parsed) return null;
  return { sheetId: parsed.sheetId, row: parsed.row, col: parsed.col };
}

export function normalizeCellKey(
  key: string,
  options: { defaultSheetId?: string } = {}
): string | null {
  return normalizeCellKeyImpl(key, options);
}

function getYMapCell(cellData: unknown): Y.Map<unknown> | null {
  if (cellData instanceof Y.Map) return cellData;

  // In some environments (notably pnpm workspaces + Node), it's possible to end up with
  // multiple `yjs` module instances (e.g. one loaded via ESM import and another via CJS require).
  // When that happens, `instanceof Y.Map` checks fail even though the value is a valid Yjs map.
  //
  // Use a small duck-type check so CollabSession APIs keep working regardless of module loader.
  if (!cellData || typeof cellData !== "object") return null;
  const maybe = cellData as any;
  if (maybe.constructor?.name !== "YMap") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return maybe as Y.Map<unknown>;
}

export class CollabSession {
  readonly doc: Y.Doc;
  readonly cells: Y.Map<unknown>;
  readonly sheets: Y.Array<Y.Map<unknown>>;
  readonly metadata: Y.Map<unknown>;
  readonly namedRanges: Y.Map<unknown>;

  readonly provider: CollabSessionProvider | null;
  readonly awareness: unknown;
  readonly presence: PresenceManager | null;
  /**
   * Optional offline persistence status + controls. Present when `options.offline`
   * is provided.
   */
  readonly offline?: {
    whenLoaded: () => Promise<void>;
    isLoaded: boolean;
    destroy: () => void;
    clear: () => Promise<void>;
  };

  /**
   * Origin token used for local transactions. Exposed for downstream consumers
   * (e.g. conflict monitors) that need to distinguish local vs remote writes.
   */
  readonly origin: object;
  /**
   * Origins that are considered "local" for this session. When undo is enabled,
   * this includes both `origin` and the underlying Yjs UndoManager instance.
   */
  readonly localOrigins: Set<any>;
  /**
   * Collaborative undo/redo. Only present when `options.undo` is provided.
   */
  readonly undo: UndoService | null;
  /**
   * Formula conflict monitor. Only present when `options.formulaConflicts` is provided.
   */
  readonly formulaConflictMonitor: FormulaConflictMonitor | null;
  /**
   * Structural cell conflict monitor. Only present when `options.cellConflicts` is provided.
   */
  readonly cellConflictMonitor: CellStructuralConflictMonitor | null;
  /**
   * Cell value conflict monitor. Only present when `options.cellValueConflicts` is provided.
   */
  readonly cellValueConflictMonitor: CellConflictMonitor | null;

  private readonly persistence: CollabPersistence | null;
  private readonly persistenceDocId: string | null;
  private persistenceBinding: CollabPersistenceBinding | null = null;
  private readonly localPersistenceLoaded: Promise<void>;

  private permissions: SessionPermissions | null = null;
  private readonly defaultSheetId: string;
  private readonly encryption:
    | {
        keyForCell: (cell: CellAddress) => CellEncryptionKey | null;
        shouldEncryptCell?: (cell: CellAddress) => boolean;
      }
    | null;
  private readonly docIdForEncryption: string;
  private schemaSyncHandler: ((isSynced: boolean) => void) | null = null;
  private sheetsSchemaObserver: ((event: any, transaction: Y.Transaction) => void) | null = null;
  private ensuringSchema = false;
  private readonly offlineAutoConnectAfterLoad: boolean;
  private readonly formulaConflictsIncludeValueConflicts: boolean;
  private isDestroyed = false;
  private providerConnectScheduled = false;

  private ensureLocalCellMapForWrite(cellKey: string): void {
    const existingCell = getYMapCell(this.cells.get(cellKey));
    if (!existingCell || existingCell instanceof Y.Map) return;

    // Normalize foreign nested Y.Maps (e.g. created by a different `yjs` module
    // instance via CJS `applyUpdate`) into local types before mutating them.
    //
    // This conversion is intentionally performed in an *untracked* transaction
    // (no origin) so collaborative undo only captures the user's actual edit.
    this.doc.transact(() => {
      const cellData = this.cells.get(cellKey);
      const cell = getYMapCell(cellData);
      if (!cell || cell instanceof Y.Map) return;

      const local = new Y.Map();
      cell.forEach((v: any, k: string) => {
        local.set(k, v);
      });
      this.cells.set(cellKey, local);
    });
  }

  private getEncryptedPayloadForCell(cell: CellAddress): unknown | undefined {
    // Cells may exist under historical key encodings (`${sheetId}:${row},${col}`) or
    // the unit-test convenience form (`r{row}c{col}`). Treat any `enc` marker under
    // those aliases as authoritative so we don't accidentally allow plaintext reads/writes.
    const keys: string[] = [makeCellKey(cell), `${cell.sheetId}:${cell.row},${cell.col}`];
    if (cell.sheetId === this.defaultSheetId) {
      keys.push(`r${cell.row}c${cell.col}`);
    }

    for (const key of keys) {
      const cellData = this.cells.get(key);
      const ycell = getYMapCell(cellData);
      if (!ycell) continue;
      const encRaw = ycell.get("enc");
      if (encRaw !== undefined) return encRaw;
    }

    return undefined;
  }

  constructor(options: CollabSessionOptions = {}) {
    const guid = options.connection?.docId ?? options.docId;
    this.doc = options.doc ?? new Y.Doc(guid ? { guid } : undefined);

    const persistence = options.persistence ?? null;
    const persistenceDocId = persistence ? options.connection?.docId ?? options.docId : null;
    if (persistence && !persistenceDocId) {
      throw new Error(
        "CollabSession persistence requires a stable docId (options.docId or options.connection.docId)"
      );
    }
    this.persistence = persistence;
    this.persistenceDocId = persistenceDocId;

    this.localPersistenceLoaded = persistence
      ? this.initLocalPersistence(persistenceDocId!, persistence)
      : Promise.resolve();
    // Avoid unhandled rejections when callers don't explicitly await persistence readiness.
    if (persistence) void this.localPersistenceLoaded.catch(() => {});

    if (options.connection && options.provider) {
      throw new Error("CollabSession cannot be constructed with both `connection` and `provider` options");
    }

    let onOfflineLoaded: (() => void) | null = null;

    const offlineEnabled = options.offline != null;
    const offlineAutoLoad = options.offline?.autoLoad ?? true;
    const offlineAutoConnectAfterLoad =
      offlineEnabled && options.connection ? (options.offline?.autoConnectAfterLoad ?? true) : false;
    this.offlineAutoConnectAfterLoad = offlineAutoConnectAfterLoad;

    const delayProviderConnect = Boolean(
      options.connection && (persistence || offlineAutoConnectAfterLoad)
    );
    this.provider =
      options.provider ??
      (options.connection
        ? new WebsocketProvider(options.connection.wsUrl, options.connection.docId, this.doc, {
            connect: !delayProviderConnect,
            WebSocketPolyfill: options.connection.WebSocketPolyfill,
            disableBc: options.connection.disableBc,
            params: {
              ...(options.connection.params ?? {}),
              ...(options.connection.token !== undefined ? { token: options.connection.token } : {}),
            },
          })
         : null);
    this.awareness = options.awareness ?? this.provider?.awareness ?? null;

    if (offlineEnabled) {
      const key = options.offline?.key ?? options.connection?.docId ?? this.doc.guid;
      const handle = attachOfflinePersistence(this.doc, {
        mode: options.offline!.mode,
        key,
        filePath: options.offline!.filePath,
        autoLoad: options.offline!.autoLoad,
      });

      const state = {
        isLoaded: false,
        whenLoaded: async () => {
          try {
            await handle.whenLoaded();
          } finally {
            state.isLoaded = true;
            onOfflineLoaded?.();
          }
        },
        destroy: () => handle.destroy(),
        clear: () => handle.clear(),
      };

      this.offline = state;

      if (offlineAutoLoad) {
        void state.whenLoaded().catch(() => {
          // Consumers can observe the failure by awaiting `session.offline.whenLoaded()`.
        });
      }
    }

    if (delayProviderConnect) {
      this.scheduleProviderConnectAfterHydration();
    }

    const schemaAutoInit = options.schema?.autoInit ?? true;
    const schemaDefaultSheetId = options.schema?.defaultSheetId ?? options.defaultSheetId ?? "Sheet1";
    const schemaDefaultSheetName = options.schema?.defaultSheetName ?? schemaDefaultSheetId;
    this.defaultSheetId = schemaDefaultSheetId;

    const roots = getWorkbookRoots(this.doc);
    this.cells = roots.cells;
    this.sheets = roots.sheets;
    this.metadata = roots.metadata;
    this.namedRanges = roots.namedRanges;

    this.encryption = options.encryption ?? null;
    // Bind AAD to the document id so ciphertext cannot be replayed between docs.
    this.docIdForEncryption = options.connection?.docId ?? this.doc.guid;

    // Stable origin token for local edits. This must be unique per-session; if
    // multiple clients share an origin, collaborative undo would treat all edits
    // as local and revert other users' changes.
    this.origin = options.undo?.origin ?? { type: "collab-session-local" };
    this.localOrigins = new Set([this.origin]);
    this.undo = null;
    this.formulaConflictMonitor = null;
    this.cellConflictMonitor = null;
    this.cellValueConflictMonitor = null;
    const formulaConflictMode =
      options.formulaConflicts?.mode ??
      (options.formulaConflicts?.includeValueConflicts ? "formula+value" : "formula");
    this.formulaConflictsIncludeValueConflicts =
      options.formulaConflicts != null && formulaConflictMode === "formula+value";

    if (schemaAutoInit) {
      const provider = this.provider;
      let providerSynced = !(provider && typeof provider.on === "function");
      if (provider && typeof provider.on === "function") {
        providerSynced = Boolean(provider.synced);
      }

      // When offline persistence is enabled and configured to auto-load (or
      // auto-connect after loading), schema initialization must wait for offline
      // state to load to avoid creating default sheets that race with persisted
      // document state.
      //
      // Note: we *do not* trigger offline hydration here. When `offline.autoLoad`
      // is false, callers must call `session.offline.whenLoaded()` themselves.
      // We still wait for that load to complete before creating the default
      // sheet, otherwise we'd race with persisted state.
      const shouldWaitForOffline = this.offline != null;
      let offlineReady = !shouldWaitForOffline;

      const shouldWaitForLocalPersistence = this.persistence != null;
      let localPersistenceReady = !shouldWaitForLocalPersistence;

      let ensureDefaultSheetScheduled = false;

      const ensureSchema = (transaction?: Y.Transaction) => {
        if (this.isDestroyed) return;
        // Avoid mutating the workbook schema while a sync provider is still in
        // the middle of initial hydration. In particular, sheets can be created
        // incrementally (e.g. map inserted before its `id` field is applied),
        // and eagerly inserting a default sheet during that window can create
        // spurious extra sheets.
        if (!providerSynced) return;
        if (!offlineReady) return;
        if (!localPersistenceReady) return;
        if (this.ensuringSchema) return;
        this.ensuringSchema = true;
        try {
          const isLocalTx = !transaction || this.localOrigins.has(transaction.origin);
          ensureWorkbookSchema(this.doc, {
            defaultSheetId: schemaDefaultSheetId,
            defaultSheetName: schemaDefaultSheetName,
            createDefaultSheet: isLocalTx,
          });
        } finally {
          this.ensuringSchema = false;
        }

        // If a remote-origin transaction temporarily leaves the sheets array
        // empty (e.g. during a merge where duplicate Sheet1 entries are pruned),
        // avoid creating a default sheet synchronously. In tests we forward Yjs
        // updates synchronously, and mutating the doc during applyUpdate can
        // cause recursive update ping-pong and stack overflows.
        //
        // Instead, schedule a microtask to re-run schema init once the merge
        // settles. Local transactions still create a default sheet immediately.
        if (transaction && !this.localOrigins.has(transaction.origin) && this.sheets.length === 0 && !ensureDefaultSheetScheduled) {
          ensureDefaultSheetScheduled = true;
          queueMicrotask(() => {
            ensureDefaultSheetScheduled = false;
            ensureSchema();
          });
        }
      };

      if (shouldWaitForOffline) {
        const markOfflineReady = () => {
          if (offlineReady) return;
          offlineReady = true;
          onOfflineLoaded = null;
          ensureSchema();
        };

        onOfflineLoaded = markOfflineReady;
        // Handle the case where offline finished loading before we registered the callback.
        if (this.offline!.isLoaded) markOfflineReady();
      }

      if (shouldWaitForLocalPersistence) {
        void this.localPersistenceLoaded
          .catch(() => {
            // Ignore load errors; we'll still continue with schema init so online
            // sessions remain usable.
          })
          .finally(() => {
            if (this.isDestroyed) return;
            if (localPersistenceReady) return;
            localPersistenceReady = true;
            ensureSchema();
          });
      }

      // Keep the sheets array well-formed over time (e.g. remove duplicate ids).
      // This primarily protects against concurrent schema initialization when two
      // clients join a brand new document at the same time.
      this.sheetsSchemaObserver = (_event, transaction) => ensureSchema(transaction);
      this.sheets.observe(this.sheetsSchemaObserver);

      if (provider && typeof provider.on === "function" && !providerSynced) {
        const handler = (isSynced: boolean) => {
          providerSynced = Boolean(isSynced);
          if (!isSynced) return;
          if (typeof provider.off === "function") provider.off("sync", handler);
          this.schemaSyncHandler = null;
          ensureSchema();
        };
        this.schemaSyncHandler = handler;
        provider.on("sync", handler);
        if (provider.synced) handler(true);
      } else {
        ensureSchema();
      }
    }

    if (options.undo) {
      const scope = new Set<Y.AbstractType<any>>([this.cells, this.sheets, this.metadata, this.namedRanges]);

      // Include comments in the undo scope deterministically.
      //
      // Callers typically create the comments root lazily (e.g. via
      // `doc.getMap("comments")` in CommentManager). If the session builds its
      // undo scope before that happens, comment edits won't be undoable.
      scope.add(getCommentsRootForUndoScope(this.doc));

      // Root names that are either already part of the built-in undo scope, or
      // should never be added via `undo.scopeNames`.
      //
      // `cellStructuralOps` is an internal log used by CellStructuralConflictMonitor.
      // It is intentionally excluded from undo tracking so conflict detection
      // metadata is never undone (which would break future conflict detection).
      const builtInScopeNames = new Set(["cells", "sheets", "metadata", "namedRanges", "comments", "cellStructuralOps"]);
      for (const name of options.undo.scopeNames ?? []) {
        if (!name || builtInScopeNames.has(name)) continue;
        scope.add(this.doc.getMap(name));
      }

      if (typeof options.undo.includeRoots === "function") {
        for (const root of options.undo.includeRoots(this.doc) ?? []) {
          if (root) scope.add(root);
        }
      }

      const undo = createUndoService({
        mode: "collab",
        doc: this.doc,
        scope: Array.from(scope),
        captureTimeoutMs: options.undo.captureTimeoutMs,
        origin: this.origin,
      });

      this.undo = undo;
      if (undo.localOrigins) this.localOrigins = undo.localOrigins;
    }

    if (options.formulaConflicts) {
      this.formulaConflictMonitor = new FormulaConflictMonitor({
        doc: this.doc,
        cells: this.cells,
        localUserId: options.formulaConflicts.localUserId,
        origin: this.origin,
        localOrigins: this.localOrigins,
        onConflict: options.formulaConflicts.onConflict,
        concurrencyWindowMs: options.formulaConflicts.concurrencyWindowMs,
        mode: options.formulaConflicts.mode,
        includeValueConflicts: options.formulaConflicts.includeValueConflicts,
      });
    }

    if (options.cellConflicts) {
      this.cellConflictMonitor = new CellStructuralConflictMonitor({
        doc: this.doc,
        cells: this.cells,
        localUserId: options.cellConflicts.localUserId,
        origin: this.origin,
        localOrigins: this.localOrigins,
        onConflict: options.cellConflicts.onConflict,
        maxOpRecordsPerUser: options.cellConflicts.maxOpRecordsPerUser,
      });
    }

    if (options.cellValueConflicts) {
      this.cellValueConflictMonitor = new CellConflictMonitor({
        doc: this.doc,
        cells: this.cells,
        localUserId: options.cellValueConflicts.localUserId,
        origin: this.origin,
        localOrigins: this.localOrigins,
        onConflict: options.cellValueConflicts.onConflict,
      });
    }

    if (options.presence) {
      if (!this.awareness) {
        throw new Error("CollabSession presence requires an awareness instance (options.awareness or provider.awareness)");
      }
      this.presence = new PresenceManager(this.awareness, options.presence);
    } else {
      this.presence = null;
    }
  }

  private async initLocalPersistence(docId: string, persistence: CollabPersistence): Promise<void> {
    try {
      await persistence.load(docId, this.doc);
    } finally {
      // Bind even if load fails so future edits still persist.
      if (this.isDestroyed) return;
      const binding = persistence.bind(docId, this.doc);
      if (this.isDestroyed) {
        void binding.destroy().catch(() => {});
        return;
      }
      this.persistenceBinding = binding;
    }
  }

  private scheduleProviderConnectAfterHydration(): void {
    if (this.providerConnectScheduled) return;
    const provider = this.provider;
    if (!provider || typeof provider.connect !== "function") return;

    this.providerConnectScheduled = true;

    const gates: Promise<void>[] = [];
    if (this.persistence) {
      gates.push(
        this.localPersistenceLoaded.catch(() => {
          // Even if local persistence fails, allow the provider to connect so the
          // session still works online.
        })
      );
    }

    if (this.offline && this.offlineAutoConnectAfterLoad && !this.offline.isLoaded) {
      gates.push(
        this.offline.whenLoaded().catch(() => {
          // Ignore offline load errors. Callers can await `session.offline.whenLoaded()`.
        })
      );
    }

    void Promise.all(gates).finally(() => {
      if (this.isDestroyed) return;
      provider.connect?.();
    });
  }

  destroy(): void {
    if (this.isDestroyed) return;
    this.isDestroyed = true;
    if (this.sheetsSchemaObserver) {
      this.sheets.unobserve(this.sheetsSchemaObserver);
      this.sheetsSchemaObserver = null;
    }
    if (this.schemaSyncHandler && this.provider && typeof this.provider.off === "function") {
      this.provider.off("sync", this.schemaSyncHandler);
      this.schemaSyncHandler = null;
    }
    this.formulaConflictMonitor?.dispose();
    this.cellConflictMonitor?.dispose();
    this.cellValueConflictMonitor?.dispose();
    this.presence?.destroy();
    this.offline?.destroy();
    this.provider?.destroy?.();
    void this.persistenceBinding?.destroy().catch(() => {});
  }

  connect(): void {
    if (this.isDestroyed) return;
    if (!this.provider?.connect) return;

    if (
      this.persistence ||
      (this.offline && this.offlineAutoConnectAfterLoad && !this.offline.isLoaded)
    ) {
      this.scheduleProviderConnectAfterHydration();
      return;
    }

    this.provider.connect();
  }

  disconnect(): void {
    this.provider?.disconnect?.();
  }

  whenLocalPersistenceLoaded(): Promise<void> {
    return this.localPersistenceLoaded;
  }

  async flushLocalPersistence(): Promise<void> {
    const persistence = this.persistence;
    const docId = this.persistenceDocId;
    if (!persistence || !docId || typeof persistence.flush !== "function") return;

    await this.localPersistenceLoaded.catch(() => {
      // If load failed, flushing may still be useful for subsequent updates.
    });
    await persistence.flush(docId);
  }

  whenSynced(timeoutMs: number = 10_000): Promise<void> {
    const provider = this.provider;
    if (!provider || typeof provider.on !== "function") return Promise.resolve();
    if (provider.synced) return Promise.resolve();

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        if (typeof provider.off === "function") provider.off("sync", handler);
        reject(new Error("Timed out waiting for provider sync"));
      }, timeoutMs);
      (timeout as any).unref?.();

      const handler = (isSynced: boolean) => {
        if (!isSynced) return;
        clearTimeout(timeout);
        if (typeof provider.off === "function") provider.off("sync", handler);
        resolve();
      };

      provider.on("sync", handler);

      if (provider.synced) handler(true);
    });
  }

  setPermissions(permissions: SessionPermissions): void {
    assertValidRole(permissions.role);
    this.permissions = {
      role: permissions.role,
      rangeRestrictions: Array.isArray(permissions.rangeRestrictions) ? permissions.rangeRestrictions : [],
      userId: permissions.userId ?? null,
    };
  }

  canEditCell(cell: CellAddress): boolean {
    const canEditByPermissions = this.permissions
      ? getCellPermissions({
          role: this.permissions.role,
          restrictions: this.permissions.rangeRestrictions,
          userId: this.permissions.userId,
          cell,
        }).canEdit
      : true;
    if (!canEditByPermissions) return false;

    // If the cell is already encrypted in Yjs, only allow edits when a key is available
    // so we can preserve encryption invariants (never write plaintext into an encrypted cell).
    const encRaw = this.getEncryptedPayloadForCell(cell);
    if (encRaw !== undefined) {
      const key = this.encryption?.keyForCell(cell) ?? null;
      return Boolean(key);
    }

    if (this.encryption) {
      const key = this.encryption.keyForCell(cell);
      const shouldEncrypt =
        typeof this.encryption.shouldEncryptCell === "function"
          ? this.encryption.shouldEncryptCell(cell)
          : key != null;
      if (shouldEncrypt && !key) return false;
    }

    return true;
  }

  canReadCell(cell: CellAddress): boolean {
    const canReadByPermissions = this.permissions
      ? getCellPermissions({
          role: this.permissions.role,
          restrictions: this.permissions.rangeRestrictions,
          userId: this.permissions.userId,
          cell,
        }).canRead
      : true;
    if (!canReadByPermissions) return false;

    // Encrypted cells require an encryption key to read. Treat malformed payloads as unreadable.
    const encRaw = this.getEncryptedPayloadForCell(cell);
    if (encRaw !== undefined) {
      const key = this.encryption?.keyForCell(cell) ?? null;
      if (!key) return false;
      if (!isEncryptedCellPayload(encRaw)) return false;
      if (key.keyId !== encRaw.keyId) return false;
    }

    return true;
  }

  getEncryptionConfig(): CollabSessionOptions["encryption"] | null {
    return this.encryption;
  }

  maskValueIfUnreadable<T>({
    sheetId,
    row,
    col,
    value,
  }: {
    sheetId: string;
    row: number;
    col: number;
    value: T;
  }): T | string {
    if (this.canReadCell({ sheetId, row, col })) return value;
    return maskCellValue(value);
  }

  async getCell(cellKey: string): Promise<CollabCell | null> {
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });

    // Cell keys can be stored under legacy encodings (`${sheetId}:${row},${col}` or
    // `r{row}c{col}`). Prefer the canonical key but fall back to legacy keys so
    // callers using canonical keys can still read documents with historical data.
    const keys = parsed
      ? Array.from(
          new Set([
            makeCellKey(parsed),
            `${parsed.sheetId}:${parsed.row},${parsed.col}`,
            ...(parsed.sheetId === this.defaultSheetId ? [`r${parsed.row}c${parsed.col}`] : []),
          ])
        )
      : [cellKey];

    /** @type {Y.Map<unknown> | null} */
    let cell: Y.Map<unknown> | null = null;

    // If any key for this coordinate is encrypted, treat the cell as encrypted
    // and do not fall back to plaintext duplicates.
    for (const key of keys) {
      const cellData = this.cells.get(key);
      const candidate = getYMapCell(cellData);
      if (!candidate) continue;
      if (candidate.get("enc") !== undefined) {
        cell = candidate;
        break;
      }
    }

    if (!cell) {
      for (const key of keys) {
        const cellData = this.cells.get(key);
        const candidate = getYMapCell(cellData);
        if (!candidate) continue;
        cell = candidate;
        break;
      }
    }

    if (!cell) return null;

    const encRaw = cell.get("enc");
    if (encRaw !== undefined) {
      if (isEncryptedCellPayload(encRaw)) {
        if (!parsed) {
          return {
            value: maskCellValue(null),
            formula: null,
            modified: (cell.get("modified") ?? null) as number | null,
            modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
            encrypted: true,
          };
        }

        const key = this.encryption?.keyForCell(parsed) ?? null;
        if (key && key.keyId === encRaw.keyId) {
          try {
            const plaintext = (await decryptCellPlaintext({
              encrypted: encRaw,
              key,
              context: {
                docId: this.docIdForEncryption,
                sheetId: parsed.sheetId,
                row: parsed.row,
                col: parsed.col,
              },
            })) as CellPlaintext;

            return {
              value: (plaintext as any)?.value ?? null,
              formula: typeof (plaintext as any)?.formula === "string" ? (plaintext as any).formula : null,
              modified: (cell.get("modified") ?? null) as number | null,
              modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
            };
          } catch {
            // Decryption failed (wrong key, tampered payload, or AAD mismatch).
          }
        }
      }

      // `enc` is present but we can't decrypt (missing key, wrong key id, corrupt payload, etc).
      return {
        value: maskCellValue(null),
        formula: null,
        modified: (cell.get("modified") ?? null) as number | null,
        modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
        encrypted: true,
      };
    }

    return {
      value: cell.get("value") ?? null,
      formula: (cell.get("formula") ?? null) as string | null,
      modified: (cell.get("modified") ?? null) as number | null,
      modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
    };
  }

  transactLocal(fn: () => void): void {
    const undoTransact = this.undo?.transact;
    if (typeof undoTransact === "function") {
      undoTransact(fn);
      return;
    }
    this.doc.transact(fn, this.origin);
  }

  async setCellValue(cellKey: string, value: unknown): Promise<void> {
    const userId = this.permissions?.userId ?? null;

    const cellData = this.cells.get(cellKey);
    const existingCell = getYMapCell(cellData);
    const parsedMaybe = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    const existingEnc = existingCell?.get("enc") ?? (parsedMaybe ? this.getEncryptedPayloadForCell(parsedMaybe) : undefined);

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parsedMaybe : null;
    if (needsCellAddress && !parsed) throw new Error(`Invalid cellKey: ${cellKey}`);

    const key = parsed && this.encryption ? this.encryption.keyForCell(parsed) : null;
    const shouldEncrypt =
      existingEnc !== undefined ||
      (parsed
        ? (typeof this.encryption?.shouldEncryptCell === "function" ? this.encryption.shouldEncryptCell(parsed) : key != null)
        : false);

    const cellValueMonitor = this.cellValueConflictMonitor;
    if (cellValueMonitor && !shouldEncrypt) {
      this.transactLocal(() => {
        // Ensure we never write plaintext value updates into an encrypted cell
        // (old clients could otherwise overwrite encrypted content).
        let cellData = this.cells.get(cellKey);
        let cell = getYMapCell(cellData);
        if (cell && cell.get("enc") !== undefined) {
          throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
        }

        cellValueMonitor.setLocalValue(cellKey, value ?? null);
      });
      return;
    }

    let encryptedPayload = null;
    if (shouldEncrypt) {
      if (!key) throw new Error(`Missing encryption key for cell ${cellKey}`);
      encryptedPayload = await encryptCellPlaintext({
        plaintext: { value: value ?? null, formula: null },
        key,
        context: {
          docId: this.docIdForEncryption,
          sheetId: parsed!.sheetId,
          row: parsed!.row,
          col: parsed!.col,
        },
      });
    }

    const monitor = this.formulaConflictMonitor;
    this.ensureLocalCellMapForWrite(cellKey);
    this.transactLocal(() => {
      const modified = Date.now();

      // Re-check inside the transaction to avoid racing with remote updates that
      // may have encrypted this cell while we were preparing a plaintext write.
      const existing = getYMapCell(this.cells.get(cellKey));
      if (!encryptedPayload && existing?.get("enc") !== undefined) {
        throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
      }

      if (!encryptedPayload && monitor && this.formulaConflictsIncludeValueConflicts) {
        monitor.setLocalValue(cellKey, value ?? null);
        return;
      }

      if (encryptedPayload) {
        const keysToUpdate = parsed
          ? Array.from(
              new Set([
                cellKey,
                makeCellKey(parsed),
                `${parsed.sheetId}:${parsed.row},${parsed.col}`,
                ...(parsed.sheetId === this.defaultSheetId ? [`r${parsed.row}c${parsed.col}`] : []),
              ])
            )
          : [cellKey];

        for (const key of keysToUpdate) {
          if (key !== cellKey && !this.cells.has(key)) continue;
          let ycell = getYMapCell(this.cells.get(key));
          if (!ycell) {
            ycell = new Y.Map();
            this.cells.set(key, ycell);
          }
          ycell.set("enc", encryptedPayload);
          ycell.delete("value");
          ycell.delete("formula");
          ycell.set("modified", modified);
          if (userId) ycell.set("modifiedBy", userId);
        }
        return;
      }

      let cell = existing;
      if (!cell) {
        cell = new Y.Map();
        this.cells.set(cellKey, cell);
      }

      cell.delete("enc");
      cell.set("value", value ?? null);
      cell.delete("formula");
      cell.set("modified", modified);
      if (userId) cell.set("modifiedBy", userId);
    });
  }

  async setCellFormula(cellKey: string, formula: string | null): Promise<void> {
    const cellData = this.cells.get(cellKey);
    const existingCell = getYMapCell(cellData);
    const parsedMaybe = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    const existingEnc = existingCell?.get("enc") ?? (parsedMaybe ? this.getEncryptedPayloadForCell(parsedMaybe) : undefined);

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parsedMaybe : null;
    if (needsCellAddress && !parsed) throw new Error(`Invalid cellKey: ${cellKey}`);

    const key = parsed && this.encryption ? this.encryption.keyForCell(parsed) : null;
    const wantsEncryption =
      existingEnc !== undefined ||
      (parsed
        ? (typeof this.encryption?.shouldEncryptCell === "function" ? this.encryption.shouldEncryptCell(parsed) : key != null)
        : false);

    const nextFormula = (formula ?? "").trim();
    if (wantsEncryption) {
      if (!key) throw new Error(`Missing encryption key for cell ${cellKey}`);

      const encryptedPayload = await encryptCellPlaintext({
        plaintext: { value: null, formula: nextFormula || null },
        key,
        context: {
          docId: this.docIdForEncryption,
          sheetId: parsed.sheetId,
          row: parsed.row,
          col: parsed.col,
        },
      });

      const userId = this.permissions?.userId ?? null;
      this.ensureLocalCellMapForWrite(cellKey);
      this.transactLocal(() => {
        const modified = Date.now();
        const keysToUpdate = Array.from(
          new Set([
            cellKey,
            makeCellKey(parsed),
            `${parsed.sheetId}:${parsed.row},${parsed.col}`,
            ...(parsed.sheetId === this.defaultSheetId ? [`r${parsed.row}c${parsed.col}`] : []),
          ])
        );

        for (const key of keysToUpdate) {
          if (key !== cellKey && !this.cells.has(key)) continue;
          let cell = getYMapCell(this.cells.get(key));
          if (!cell) {
            cell = new Y.Map();
            this.cells.set(key, cell);
          }
          cell.set("enc", encryptedPayload);
          cell.delete("value");
          cell.delete("formula");
          cell.set("modified", modified);
          if (userId) cell.set("modifiedBy", userId);
        }
      });
      return;
    }

    const monitor = this.formulaConflictMonitor;
    if (monitor) {
      this.transactLocal(() => {
        // Ensure we never write plaintext formula updates into an encrypted cell
        // (old clients could otherwise overwrite encrypted content).
        let cellData = this.cells.get(cellKey);
        let cell = getYMapCell(cellData);
        if (!cell) {
          cell = new Y.Map();
          this.cells.set(cellKey, cell);
        }

        if (cell.get("enc") !== undefined) {
          throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
        }

        monitor.setLocalFormula(cellKey, formula ?? "");
      });
      return;
    }

    const userId = this.permissions?.userId ?? null;

    this.ensureLocalCellMapForWrite(cellKey);
    this.transactLocal(() => {
      let cellData = this.cells.get(cellKey);
      let cell = getYMapCell(cellData);
      if (!cell) {
        cell = new Y.Map();
        this.cells.set(cellKey, cell);
      }

      // Re-check inside the transaction to avoid racing with remote updates that
      // may have encrypted this cell while we were preparing a plaintext write.
      if (cell.get("enc") !== undefined) {
        throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
      }

      cell.delete("enc");

      if (nextFormula) cell.set("formula", nextFormula);
      else cell.delete("formula");

      cell.set("value", null);
      cell.set("modified", Date.now());
      if (userId) cell.set("modifiedBy", userId);
    });
  }

  async safeSetCellValue(cellKey: string, value: unknown): Promise<boolean> {
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${cellKey}`);
    if (!this.canEditCell(parsed)) return false;

    await this.setCellValue(makeCellKey(parsed), value);
    return true;
  }

  async safeSetCellFormula(cellKey: string, formula: string | null): Promise<boolean> {
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${cellKey}`);
    if (!this.canEditCell(parsed)) return false;

    await this.setCellFormula(makeCellKey(parsed), formula);
    return true;
  }
}

export function createCollabSession(options: CollabSessionOptions = {}): CollabSession {
  return new CollabSession(options);
}

// Backwards-compatible alias (Task 133 naming).
export const createSession = createCollabSession;

export type DocumentControllerBinder = { destroy: () => void };

export async function bindCollabSessionToDocumentController(options: {
  session: CollabSession;
  documentController: any;
  undoService?: { transact?: (fn: () => void) => void; origin?: any } | null;
  defaultSheetId?: string;
  userId?: string | null;
}): Promise<DocumentControllerBinder> {
  const { session, documentController, undoService, defaultSheetId, userId } = options ?? ({} as any);
  if (!session) throw new Error("bindCollabSessionToDocumentController requires { session }");
  if (!documentController)
    throw new Error("bindCollabSessionToDocumentController requires { documentController }");

  // Avoid importing the Node-oriented binder (and its encryption dependencies)
  // unless a consumer explicitly opts into DocumentController wiring.
  const { bindYjsToDocumentController } = await import("../../binder/index.js");

  return bindYjsToDocumentController({
    ydoc: session.doc,
    documentController,
    // Intentionally do not default to `session.undo` here: CollabSession's undo origin
    // is the same as `session.origin`, and wiring it into the binder would cause
    // direct `session.setCell*` calls to be treated as "local" and ignored by the
    // Yjs→DocumentController observer. Callers that want Yjs UndoManager semantics
    // can pass an explicit `undoService`.
    undoService: undoService ?? null,
    defaultSheetId,
    userId,
    encryption: session.getEncryptionConfig(),
    canReadCell: (cell) => session.canReadCell(cell),
    canEditCell: (cell) => session.canEditCell(cell),
    // Use the standard enterprise mask. The binder also uses this hook for
    // encrypted cells that cannot be decrypted.
    maskCellValue: (value) => maskCellValue(value),
  }) as DocumentControllerBinder;
}
