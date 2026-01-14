import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import { PresenceManager } from "@formula/collab-presence";
import { createUndoService, type UndoService } from "@formula/collab-undo";
import { getCommentsRoot, migrateCommentsArrayToMap } from "@formula/collab-comments";
import {
  getArrayRoot,
  getMapRoot,
  getYArray,
  getYMap,
  getYText,
  isYAbstractType,
  yjsValueToJson,
} from "@formula/collab-yjs-utils";
import {
  CellConflictMonitor,
  CellStructuralConflictMonitor,
  FormulaConflictMonitor,
  type CellConflict,
  type CellStructuralConflict,
  type FormulaConflict,
} from "@formula/collab-conflicts";
import {
  createMetadataManagerForSessionWithPermissions as createMetadataManagerForSessionWithPermissionsImpl,
  createNamedRangeManagerForSessionWithPermissions as createNamedRangeManagerForSessionWithPermissionsImpl,
  createSheetManagerForSessionWithPermissions as createSheetManagerForSessionWithPermissionsImpl,
  ensureWorkbookSchema,
  getWorkbookRoots,
} from "@formula/collab-workbook";
import {
  decryptCellPlaintext,
  encryptCellPlaintext,
  isEncryptedCellPayload,
  type CellEncryptionKey,
  type CellPlaintext,
} from "@formula/collab-encryption";
import type {
  CollabPersistence,
  CollabPersistenceBinding,
  CollabPersistenceFlushOptions,
} from "@formula/collab-persistence";

import {
  assertValidRole,
  getCellPermissions,
  maskCellValue,
  normalizeRestriction,
  roleCanComment,
  roleCanEdit,
  roleCanShare,
} from "../../permissions/index.js";
import {
  makeCellKey as makeCellKeyImpl,
  normalizeCellKey as normalizeCellKeyImpl,
  parseCellKey as parseCellKeyImpl,
} from "./cell-key.js";

function getCommentsRootForUndoScope(doc: Y.Doc): Y.AbstractType<any> {
  // Yjs root types are schema-defined: you must know whether a key is a Map or
  // Array. When applying updates into a fresh Doc, root types can temporarily
  // appear as a generic `AbstractType` placeholder until a constructor is chosen.
  //
  // Importantly, calling `doc.getMap("comments")` on an Array-backed doc can
  // permanently define it as a Map and make legacy array content inaccessible.
  // Use the shared comment-root detection logic to determine whether "comments"
  // is a Map or legacy Array before instantiating it.
  const existing = doc.share.get("comments");
  const kind = existing ? getCommentsRoot(doc).kind : "map";
  const root = kind === "array" ? getArrayRoot(doc, "comments") : getMapRoot(doc, "comments");

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

  const text = getYText(value);
  if (text) return yjsValueToJson(text);

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

export type CollabSessionSyncState = { connected: boolean; synced: boolean };

export type CollabSessionUpdateStats = {
  lastUpdateBytes: number;
  maxRecentBytes: number;
  avgRecentBytes: number;
};

export type CollabSessionPersistenceState = {
  enabled: boolean;
  loaded: boolean;
  lastFlushedAt: number | null;
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
   *
   * @deprecated Use `options.persistence` (see `@formula/collab-persistence`) instead.
   */
  offline?: {
    mode: "indexeddb" | "file";
    /**
     * Storage key / namespace.
     *
     * - `mode: "indexeddb"`: defaults to `connection.docId` / `options.docId` / `doc.guid`.
     * - `mode: "file"`: the persistence doc id is derived from `connection.docId` / `options.docId`
     *   (and otherwise falls back to `filePath`, then `key`, then `doc.guid`).
     */
    key?: string;
    /**
     * Legacy file path for Node/desktop persistence when `mode: "file"` is used.
     *
     * Internally, CollabSession maps this to `new FileCollabPersistence(dirname(filePath))`
     * and uses a hashed per-doc file name. If a legacy `.yjslog` file exists at
     * `filePath`, CollabSession will copy it into the new hashed file on first run
     * (best-effort migration).
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
    /**
     * Optional age-based pruning window (in milliseconds) for records in the shared
     * `cellStructuralOps` Yjs log. When set, records older than `Date.now() - maxOpRecordAgeMs`
     * may be deleted by any client (best-effort).
     *
     * Pruning is additionally conservative relative to the local op log queue:
     * records are only pruned when they are older than both the age cutoff and
     * the oldest local op record (queue head). This avoids deleting records that
     * may still be needed to compare against local ops that are in-flight (e.g.
     * long offline periods).
     *
     * Pruning is conservative: records are not deleted in the same op-log transaction
     * they are added, so late-arriving/offline records have a chance to be ingested
     * before being removed.
     *
     * Note: Pruning is incremental and may run over multiple passes for very large
     * logs.
     *
     * Defaults to null/disabled.
     */
    maxOpRecordAgeMs?: number | null;
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
     *
     * Tip: to drive encryption policy from shared workbook metadata (so clients
     * without keys can still refuse plaintext writes), use
     * `createEncryptionPolicyFromDoc(doc)` from `@formula/collab-encrypted-ranges`
     * and pass its `shouldEncryptCell` here.
     */
    shouldEncryptCell?: (cell: CellAddress) => boolean;
    /**
     * When true, cell formatting (`format`) is included in the encrypted plaintext payload
     * and removed from the shared Yjs cell map. Defaults to false for backwards compatibility.
     */
    encryptFormat?: boolean;
  };

  /**
   * Comments configuration.
   */
  comments?: {
    /**
     * When enabled, CollabSession will attempt to migrate legacy Array-backed
     * comments roots (`Y.Array<Y.Map>`) into the canonical Map-backed schema
     * (`Y.Map<string, Y.Map>`).
     *
     * Migration runs best-effort after initial document hydration (local
     * persistence load and/or provider initial sync).
     */
    migrateLegacyArrayToMap?: boolean;
  };
}

export interface CollabCell {
  value: unknown;
  formula: string | null;
  modified: number | null;
  modifiedBy: string | null;
  /**
   * Optional per-cell style object when `options.encryption.encryptFormat=true` and the cell
   * is decrypted successfully. Omitted otherwise.
   */
  format?: unknown | null;
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
  return getYMap(cellData) as Y.Map<unknown> | null;
}

const RECENT_OUTGOING_UPDATE_BYTES_LIMIT = 20;

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
   *
   * @deprecated `options.offline` is deprecated. Prefer `options.persistence` and
   * `session.whenLocalPersistenceLoaded()` / `persistence.clear(docId)` instead.
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

  private persistence: CollabPersistence | null;
  private readonly persistenceDocId: string | null;
  private persistenceBinding: CollabPersistenceBinding | null = null;
  private readonly hasLocalPersistence: boolean;
  private readonly shouldGateProviderConnectOnLocalPersistence: boolean;
  private localPersistenceStarted = false;
  private localPersistenceStartPromise: Promise<void> | null = null;
  private readonly localPersistenceLoaded: Promise<void>;
  private readonly resolveLocalPersistenceLoaded: (() => void) | null;
  private readonly rejectLocalPersistenceLoaded: ((err: unknown) => void) | null;
  private readonly localPersistenceFactory: (() => Promise<CollabPersistence>) | null;
  private persistenceDetached = false;
  private readonly legacyOfflineFilePath: string | null;
  private readonly localPersistenceBindBeforeLoad: boolean;

  private permissions: SessionPermissions | null = null;
  private readonly permissionsListeners = new Set<(permissions: SessionPermissions | null) => void>();
  private readonly defaultSheetId: string;
  private readonly encryption:
    | {
        keyForCell: (cell: CellAddress) => CellEncryptionKey | null;
        shouldEncryptCell?: (cell: CellAddress) => boolean;
        encryptFormat?: boolean;
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

  private syncState: CollabSessionSyncState = { connected: false, synced: false };
  private readonly statusListeners = new Set<(state: CollabSessionSyncState) => void>();
  private providerStatusListener: ((event: any) => void) | null = null;
  private providerSyncListener: ((isSynced: boolean) => void) | null = null;
  private commentsMigrationPermissionsUnsubscribe: (() => void) | null = null;
  private commentsMigrationSyncHandler: ((isSynced: boolean) => void) | null = null;
  private commentsUndoScopePermissionsUnsubscribe: (() => void) | null = null;
  private commentsUndoScopeSyncHandler: ((isSynced: boolean) => void) | null = null;

  private readonly recentOutgoingUpdateBytes: number[] = [];
  private docUpdateListener: ((update: Uint8Array, origin: any) => void) | null = null;

  private localPersistenceLoadedFlag = false;
  private lastLocalPersistenceFlushAt: number | null = null;

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

    if (options.connection && options.provider) {
      throw new Error("CollabSession cannot be constructed with both `connection` and `provider` options");
    }

    const explicitPersistence = options.persistence ?? null;

    const offlineEnabled = options.offline != null;
    const offlineAutoLoad = options.offline?.autoLoad ?? true;
    const offlineAutoConnectAfterLoad =
      offlineEnabled && !explicitPersistence && options.connection ? (options.offline?.autoConnectAfterLoad ?? true) : false;
    this.offlineAutoConnectAfterLoad = offlineAutoConnectAfterLoad;

    const offlineShouldConfigurePersistence = offlineEnabled && !explicitPersistence;
    this.legacyOfflineFilePath =
      offlineShouldConfigurePersistence && options.offline?.mode === "file"
        ? options.offline.filePath ?? null
        : null;
    this.localPersistenceBindBeforeLoad = offlineShouldConfigurePersistence;

    const persistenceDocId =
      explicitPersistence
        ? options.connection?.docId ?? options.docId
        : offlineShouldConfigurePersistence
          ? this.getDocIdForOfflinePersistence(options.offline!, {
              connectionDocId: options.connection?.docId,
              explicitDocId: options.docId,
            })
          : null;

    if (explicitPersistence && !persistenceDocId) {
      throw new Error(
        "CollabSession persistence requires a stable docId (options.docId or options.connection.docId)"
      );
    }

    this.persistence = explicitPersistence;
    this.persistenceDocId = persistenceDocId;

    this.hasLocalPersistence = Boolean(explicitPersistence || offlineShouldConfigurePersistence);
    this.shouldGateProviderConnectOnLocalPersistence = Boolean(
      explicitPersistence || (offlineShouldConfigurePersistence && offlineAutoConnectAfterLoad)
    );

    this.localPersistenceFactory = offlineShouldConfigurePersistence
      ? () => this.createPersistenceFromOfflineOptions(options.offline!)
      : explicitPersistence
        ? async () => explicitPersistence
        : null;

    if (!this.hasLocalPersistence) {
      this.localPersistenceLoaded = Promise.resolve();
      this.resolveLocalPersistenceLoaded = null;
      this.rejectLocalPersistenceLoaded = null;
    } else {
      let resolveLocalPersistenceLoaded: (() => void) | null = null;
      let rejectLocalPersistenceLoaded: ((err: unknown) => void) | null = null;
      this.localPersistenceLoaded = new Promise<void>((resolve, reject) => {
        resolveLocalPersistenceLoaded = resolve;
        rejectLocalPersistenceLoaded = reject;
      });
      this.resolveLocalPersistenceLoaded = resolveLocalPersistenceLoaded;
      this.rejectLocalPersistenceLoaded = rejectLocalPersistenceLoaded;

      // Eagerly start persistence for explicit `options.persistence`. For legacy
      // `options.offline`, respect `offline.autoLoad`.
      if (explicitPersistence || (offlineShouldConfigurePersistence && offlineAutoLoad)) {
        this.startLocalPersistence();
        // Avoid unhandled rejections when callers don't explicitly await persistence readiness.
        void this.localPersistenceLoaded.catch(() => {});
      }
    }

    const delayProviderConnect = Boolean(
      options.connection && (explicitPersistence || offlineAutoConnectAfterLoad)
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
      const state = {
        isLoaded: false,
        whenLoaded: async () => {
          try {
            await this.whenLocalPersistenceLoaded();
          } finally {
            state.isLoaded = true;
          }
        },
        destroy: () => {
          // Match legacy `@formula/collab-offline` behavior: if persistence hasn't
          // started yet (e.g. `offline.autoLoad: false` and the caller never
          // awaited `whenLoaded()`), destroy is a no-op.
          if (!this.localPersistenceStarted) return;
          this.detachLocalPersistence();
        },
        clear: async () => {
          await this.clearLocalPersistence();
        },
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

      const shouldWaitForLocalPersistence = this.hasLocalPersistence;
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

      // Comments root selection is special: historical documents may store the
      // root under a legacy Array schema. Calling `doc.getMap("comments")` on a
      // fresh Doc (before provider/persistence hydration) can permanently define
      // the root as a Map and make legacy array content inaccessible.
      //
      // To avoid clobbering legacy docs, we add the comments root to the undo
      // scope lazily once hydration is complete (or once the root already exists).
      //
      // See regression test: `CollabSession undo does not clobber legacy comments root when undo is enabled before sync`.
      const shouldWaitForLocalPersistence = this.hasLocalPersistence;
      let localPersistenceReady = !shouldWaitForLocalPersistence;
      const providerForCommentsScope = this.provider;
      let providerHydrated = !(providerForCommentsScope && typeof providerForCommentsScope.on === "function");
      if (providerForCommentsScope && typeof providerForCommentsScope.on === "function") {
        providerHydrated = Boolean((providerForCommentsScope as any).synced);
      }

      // Root names that are either already part of the built-in undo scope, or
      // should never be added via `undo.scopeNames`.
      //
      // `cellStructuralOps` is an internal log used by CellStructuralConflictMonitor.
      // It is intentionally excluded from undo tracking so conflict detection
      // metadata is never undone (which would break future conflict detection).
      const builtInScopeNames = new Set(["cells", "sheets", "metadata", "namedRanges", "comments", "cellStructuralOps"]);
      for (const name of options.undo.scopeNames ?? []) {
        if (!name || builtInScopeNames.has(name)) continue;
        scope.add(getMapRoot(this.doc, name));
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

      const ensureCommentsUndoScope = () => {
        if (this.isDestroyed) return;

        // Only attempt to normalize/extend undo scope when the role allows comment
        // writes. This avoids viewer clients generating best-effort Yjs updates
        // during normalization/re-wrapping of foreign comment types.
        let allowed = false;
        try {
          allowed = this.canComment();
        } catch {
          allowed = false;
        }
        if (!allowed) return;

        const hasRoot = Boolean(this.doc.share.get("comments"));

        // If the doc is still hydrating and the root doesn't exist yet, do not
        // instantiate it (it could clobber legacy Array-backed docs).
        if (!hasRoot && (!providerHydrated || !localPersistenceReady)) return;

        let root: Y.AbstractType<any> | null = null;
        try {
          root = getCommentsRootForUndoScope(this.doc);
        } catch {
          root = null;
        }
        if (!root) return;

        // Add to all known UndoManagers (session + any downstream binder-origin
        // managers) so comment edits remain undoable regardless of origin.
        const undoManagers = Array.from(this.localOrigins).filter((value) => {
          if (value instanceof Y.UndoManager) return true;
          if (!value || typeof value !== "object") return false;
          const maybe = value as any;
          return (
            typeof maybe.addToScope === "function" &&
            typeof maybe.undo === "function" &&
            typeof maybe.redo === "function" &&
            typeof maybe.stopCapturing === "function"
          );
        });
        for (const undoManager of undoManagers) {
          try {
            (undoManager as any).addToScope(root);
          } catch {
            // ignore
          }
        }
      };

      // Re-attempt once permissions are configured (and on future role upgrades).
      if (this.commentsUndoScopePermissionsUnsubscribe) {
        try {
          this.commentsUndoScopePermissionsUnsubscribe();
        } catch {
          // ignore
        }
        this.commentsUndoScopePermissionsUnsubscribe = null;
      }
      this.commentsUndoScopePermissionsUnsubscribe = this.onPermissionsChanged(() => {
        ensureCommentsUndoScope();
      });

      // Re-attempt once local persistence hydration completes (if enabled).
      if (shouldWaitForLocalPersistence) {
        void this.localPersistenceLoaded
          .catch(() => {
            // Ignore load failures; allow undo-scope setup to continue based on
            // provider hydration alone.
          })
          .finally(() => {
            localPersistenceReady = true;
            ensureCommentsUndoScope();
          });
      }

      // Re-attempt once the sync provider reports hydration complete.
      const provider = providerForCommentsScope;
      if (provider && typeof provider.on === "function") {
        const handler = (isSynced: boolean) => {
          providerHydrated = Boolean(isSynced);
          if (!isSynced) return;
          try {
            provider.off?.("sync", handler);
          } catch {
            // ignore
          }
          if (this.commentsUndoScopeSyncHandler === handler) {
            this.commentsUndoScopeSyncHandler = null;
          }
          ensureCommentsUndoScope();
        };
        this.commentsUndoScopeSyncHandler = handler;
        try {
          provider.on("sync", handler);
        } catch {
          // ignore
        }
        if ((provider as any).synced) handler(true);
      } else {
        // No sync provider: the caller-provided doc is already in memory.
        ensureCommentsUndoScope();
      }
    }

    // Certain transactions are intentional, bulk "time travel" operations (e.g.
    // version restores) and should not participate in conflict detection or
    // local-edit tracking inside conflict monitors.
    const ignoredConflictOrigins = new Set<any>(["versioning-restore", "branching-apply"]);

    if (options.formulaConflicts) {
      this.formulaConflictMonitor = new FormulaConflictMonitor({
        doc: this.doc,
        cells: this.cells,
        localUserId: options.formulaConflicts.localUserId,
        origin: this.origin,
        localOrigins: this.localOrigins,
        ignoredOrigins: ignoredConflictOrigins,
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
        ignoredOrigins: ignoredConflictOrigins,
        onConflict: options.cellConflicts.onConflict,
        maxOpRecordsPerUser: options.cellConflicts.maxOpRecordsPerUser,
        maxOpRecordAgeMs: options.cellConflicts.maxOpRecordAgeMs,
      });
    }

    if (options.cellValueConflicts) {
      this.cellValueConflictMonitor = new CellConflictMonitor({
        doc: this.doc,
        cells: this.cells,
        localUserId: options.cellValueConflicts.localUserId,
        origin: this.origin,
        localOrigins: this.localOrigins,
        ignoredOrigins: ignoredConflictOrigins,
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

    // --- Observability hooks ---
    // Sync state: track connected/synced transitions from the provider.
    const providerAny = this.provider as any;
    this.syncState = {
      connected:
        typeof providerAny?.wsconnected === "boolean"
          ? providerAny.wsconnected
          : typeof providerAny?.connected === "boolean"
            ? providerAny.connected
            : false,
      synced: Boolean(providerAny?.synced),
    };

    if (this.provider && typeof this.provider.on === "function") {
      const handleStatus = (event: any) => {
        const prev = this.syncState;
        let connected: boolean | null = null;
        const status = typeof event === "string" ? event : event?.status;
        if (status === "connected") connected = true;
        else if (status === "disconnected") connected = false;
        else if (typeof event?.connected === "boolean") connected = event.connected;

        if (connected === null) return;
        const next: CollabSessionSyncState = connected
          ? { connected: true, synced: Boolean((this.provider as any)?.synced) }
          : { connected: false, synced: false };
        if (prev.connected === next.connected && prev.synced === next.synced) return;
        this.syncState = next;
        this.emitStatusChange();
      };
      const handleSync = (isSynced: boolean) => {
        const next: CollabSessionSyncState = {
          connected: isSynced ? true : this.syncState.connected,
          synced: Boolean(isSynced),
        };
        if (!next.connected) next.synced = false;
        if (this.syncState.connected === next.connected && this.syncState.synced === next.synced) return;
        this.syncState = next;
        this.emitStatusChange();
      };

      // Not all providers implement all events (or accept arbitrary event names).
      // Observability should be best-effort and never prevent session startup.
      try {
        this.provider.on("status", handleStatus);
        this.providerStatusListener = handleStatus;
      } catch {
        // ignore
      }
      try {
        this.provider.on("sync", handleSync);
        this.providerSyncListener = handleSync;
      } catch {
        // ignore
      }
    }

    // Update size tracking: record sizes for local-origin updates only.
    const handleDocUpdate = (update: Uint8Array, origin: any) => {
      if (!this.localOrigins.has(origin)) return;
      const bytes = typeof (update as any)?.length === "number" ? (update as any).length : 0;
      if (bytes <= 0) return;
      this.recentOutgoingUpdateBytes.push(bytes);
      if (this.recentOutgoingUpdateBytes.length > RECENT_OUTGOING_UPDATE_BYTES_LIMIT) {
        this.recentOutgoingUpdateBytes.splice(
          0,
          this.recentOutgoingUpdateBytes.length - RECENT_OUTGOING_UPDATE_BYTES_LIMIT
        );
      }
    };
    this.docUpdateListener = handleDocUpdate;
    this.doc.on("update", handleDocUpdate);

    void this.localPersistenceLoaded
      .then(
        () => {
          this.localPersistenceLoadedFlag = true;
        },
        () => {
          // If persistence fails to load, keep the loaded flag false so callers can
          // treat it as an unhealthy persistence state.
        }
      )
      .catch(() => {
        // Best-effort: avoid unhandled rejections if the `.then` bookkeeping callback throws.
      });

    this.scheduleCommentsMigration(options.comments);
  }

  private scheduleCommentsMigration(opts: CollabSessionOptions["comments"] | undefined): void {
    if (!opts?.migrateLegacyArrayToMap) return;

    // Migration mutates the shared Y.Doc. In collab mode, viewers should never
    // generate Yjs updates (they would be rejected by server-side access control
    // anyway). Gate migration on comment permissions and re-attempt after a role
    // upgrade (e.g. viewer → commenter).
    const provider = this.provider;
    const providerUsesSyncEvents = Boolean(provider && typeof provider.on === "function");
    // Like workbook schema init, only migrate after the doc is hydrated. When both a
    // sync provider and local persistence are present, require *both* to settle so
    // we don't run migrations against a partially-hydrated doc.
    let providerHydrated = !providerUsesSyncEvents;
    if (providerUsesSyncEvents) {
      providerHydrated = Boolean((provider as any).synced);
    }
    let localPersistenceHydrated = !this.hasLocalPersistence;

    const tryMigrate = () => {
      if (this.isDestroyed) return;
      if (!providerHydrated) return;
      if (!localPersistenceHydrated) return;

      let canComment = false;
      try {
        canComment = this.canComment();
      } catch {
        canComment = false;
      }
      if (!canComment) return;

      let didMigrate = false;
      try {
        didMigrate = migrateCommentsArrayToMap(this.doc, { origin: "comments-migrate" });
      } catch {
        // Best-effort: never block session usage on comment schema migration.
        return;
      }

      if (!didMigrate) return;

      // Migration replaces the `comments` root type, which can leave any
      // existing UndoManager scopes pointing at the old root. Ensure all known
      // UndoManagers track the new canonical root so comment edits remain
      // undoable.
      try {
        const root = this.doc.share.get("comments");
        if (!root || !isYAbstractType(root)) return;

        const undoManagers = Array.from(this.localOrigins).filter((value) => {
          if (value instanceof Y.UndoManager) return true;
          if (!value || typeof value !== "object") return false;
          const maybe = value as any;
          return (
            typeof maybe.addToScope === "function" &&
            typeof maybe.undo === "function" &&
            typeof maybe.redo === "function" &&
            typeof maybe.stopCapturing === "function"
          );
        });

        for (const undoManager of undoManagers) {
          try {
            (undoManager as any).addToScope(root);
          } catch {
            // ignore
          }
        }
      } catch {
        // ignore
      }
    };

    // If permissions are applied after session construction, retry migration once
    // hydration is complete and the role allows comment writes.
    if (this.commentsMigrationPermissionsUnsubscribe) {
      try {
        this.commentsMigrationPermissionsUnsubscribe();
      } catch {
        // ignore
      }
      this.commentsMigrationPermissionsUnsubscribe = null;
    }
    this.commentsMigrationPermissionsUnsubscribe = this.onPermissionsChanged(() => {
      // Defer to a microtask so we don't do heavy Yjs work directly inside the
      // caller's `setPermissions()` stack.
      queueMicrotask(tryMigrate);
    });

    // Run after local persistence hydration (if enabled).
    if (this.hasLocalPersistence) {
      void this.localPersistenceLoaded
        .catch(() => {
          // Ignore load failures; we still might be able to migrate based on
          // remote/provider hydration.
        })
        .finally(() => {
          localPersistenceHydrated = true;
          queueMicrotask(tryMigrate);
        });
    }

    if (providerUsesSyncEvents) {
      const handler = (isSynced: boolean) => {
        providerHydrated = Boolean(isSynced);
        if (!isSynced) return;
        try {
          if (typeof provider.off === "function") provider.off("sync", handler);
        } catch {
          // ignore
        }
        if (this.commentsMigrationSyncHandler === handler) {
          this.commentsMigrationSyncHandler = null;
        }
        queueMicrotask(tryMigrate);
      };
      this.commentsMigrationSyncHandler = handler;
      try {
        provider.on("sync", handler);
      } catch {
        // ignore; if the provider refuses the listener we'll still fall back to persistence/no-provider logic.
      }
      if (provider.synced) handler(true);
    } else if (!this.hasLocalPersistence) {
      // No persistence + no sync provider: the caller-provided doc is already in memory.
      queueMicrotask(tryMigrate);
    }
  }

  private startLocalPersistence(): void {
    if (!this.hasLocalPersistence) return;
    if (this.localPersistenceStarted) return;
    this.localPersistenceStarted = true;

    const factory = this.localPersistenceFactory;
    const docId = this.persistenceDocId;
    if (!factory || !docId) {
      this.rejectLocalPersistenceLoaded?.(new Error("Internal error: persistence is configured but missing factory/docId"));
      return;
    }

    this.localPersistenceStartPromise = (async () => {
      const persistence = this.persistence ?? (await factory());
      this.persistence = persistence;

      let binding: CollabPersistenceBinding | null = null;

      // Legacy offline persistence attached listeners immediately and buffered
      // writes during the initial load. To preserve that behavior, bind before
      // loading when the persistence implementation is synthesized from the
      // deprecated `options.offline` config.
      if (this.localPersistenceBindBeforeLoad && !this.isDestroyed && !this.persistenceDetached) {
        binding = persistence.bind(docId, this.doc);
        if (this.isDestroyed || this.persistenceDetached) {
          await binding.destroy().catch(() => {});
          return;
        }
        this.persistenceBinding = binding;
      }

      // If persistence was detached while we were setting up (e.g. legacy
      // `session.offline.destroy()` called during async imports), do not apply
      // persisted state into the doc.
      if (this.isDestroyed || this.persistenceDetached) return;

      try {
        await persistence.load(docId, this.doc);
      } finally {
        if (binding) {
          // If we detached/destroyed while loading, ensure the pre-bound binding
          // does not linger.
          if (this.isDestroyed || this.persistenceDetached) {
            if (this.persistenceBinding === binding) {
              this.persistenceBinding = null;
              await binding.destroy().catch(() => {});
            }
          }
        } else {
          // Bind even if load fails so future edits still persist.
          if (!this.isDestroyed && !this.persistenceDetached) {
            const nextBinding = persistence.bind(docId, this.doc);
            if (this.isDestroyed || this.persistenceDetached) {
              void nextBinding.destroy().catch(() => {});
            } else {
              this.persistenceBinding = nextBinding;
            }
          }
        }
      }
    })();

    void this.localPersistenceStartPromise
      .then(
        () => this.resolveLocalPersistenceLoaded?.(),
        (err) => this.rejectLocalPersistenceLoaded?.(err)
      )
      .catch(() => {
        // Best-effort: avoid unhandled rejections if the resolve/reject hooks throw.
      });
  }

  private async createPersistenceFromOfflineOptions(
    offline: NonNullable<CollabSessionOptions["offline"]>
  ): Promise<CollabPersistence> {
    if (offline.mode === "indexeddb") {
      const { IndexedDbCollabPersistence } = await import("@formula/collab-persistence/indexeddb");
      return new IndexedDbCollabPersistence();
    }
    if (offline.mode === "file") {
      if (!offline.filePath) {
        // Match the legacy `@formula/collab-offline` error message for easier
        // migration (some callers may assert on it).
        throw new Error('Offline persistence mode "file" requires opts.filePath');
      }
      // Avoid a top-level import of the Node-only persistence implementation so
      // this module can still be bundled for browser environments.
      // Use a computed specifier so browser bundlers don't try to resolve this
      // Node-only module unless it's actually used at runtime.
      const specifier = ["@formula/collab-persistence", "file"].join("/");
      let FileCollabPersistence: any = null;
      try {
        // eslint-disable-next-line no-undef
        ({ FileCollabPersistence } = await import(
          // eslint-disable-next-line no-undef
          /* @vite-ignore */ specifier
        ));
      } catch {
        throw new Error('Offline persistence mode "file" is only supported in Node environments');
      }
      // Use node:path.dirname for correctness (handles POSIX roots, Windows drive
      // roots, etc) without a static Node import that would break browser bundlers.
      const pathSpecifier = ["node", "path"].join(":");
      let path: any = null;
      try {
        // eslint-disable-next-line no-undef
        path = await import(
          // eslint-disable-next-line no-undef
          /* @vite-ignore */ pathSpecifier
        );
      } catch {
        throw new Error('Offline persistence mode "file" is only supported in Node environments');
      }
      const dir = typeof path.dirname === "function" ? path.dirname(offline.filePath) : this.dirnameForOfflineFilePath(offline.filePath);

      // Best-effort migration: older `@formula/collab-offline` used `offline.filePath`
      // as the *actual* append-only log file. `FileCollabPersistence` stores one file
      // per doc inside a directory, so users upgrading would otherwise "lose" their
      // local offline state until the next sync.
      //
      // If the legacy file exists and the new persistence file does not, copy the
      // legacy log bytes over so `FileCollabPersistence.load()` can replay them.
      await this.migrateLegacyOfflineFileLogIfNeeded({
        legacyFilePath: offline.filePath,
        dir,
        docId: this.persistenceDocId ?? offline.filePath,
      });

      return new FileCollabPersistence(dir);
    }
    throw new Error(`Unsupported offline persistence mode: ${String((offline as any).mode)}`);
  }

  private async migrateLegacyOfflineFileLogIfNeeded(opts: {
    legacyFilePath: string;
    dir: string;
    docId: string;
  }): Promise<void> {
    // Only relevant in Node environments. Keep this logic self-contained so
    // browser bundlers can tree-shake it.
    try {
      const fsSpecifier = ["node", "fs"].join(":");
      const cryptoSpecifier = ["node", "crypto"].join(":");
      const pathSpecifier = ["node", "path"].join(":");
      // eslint-disable-next-line no-undef
      const { promises: fs } = await import(
        // eslint-disable-next-line no-undef
        /* @vite-ignore */ fsSpecifier
      );
      // eslint-disable-next-line no-undef
      const { createHash } = await import(
        // eslint-disable-next-line no-undef
        /* @vite-ignore */ cryptoSpecifier
      );
      // eslint-disable-next-line no-undef
      const path = await import(
        // eslint-disable-next-line no-undef
        /* @vite-ignore */ pathSpecifier
      );

      const docHash = createHash("sha256").update(opts.docId).digest("hex");
      const nextFilePath = path.join(opts.dir, `${docHash}.yjs`);

      // Avoid self-copy (or re-migrating once the new file exists).
      if (nextFilePath === opts.legacyFilePath) return;

      const [legacyStat, nextStat] = await Promise.all([
        fs.stat(opts.legacyFilePath).catch(() => null),
        fs.stat(nextFilePath).catch(() => null),
      ]);
      if (!legacyStat || !legacyStat.isFile()) return;
      if (nextStat && nextStat.isFile()) return;

      await fs.mkdir(opts.dir, { recursive: true }).catch(() => {});
      const bytes = await fs.readFile(opts.legacyFilePath).catch(() => null);
      if (!bytes || bytes.length === 0) return;

      await fs.writeFile(nextFilePath, bytes, { mode: 0o600, flag: "wx" }).catch((err) => {
        const code = (err as any)?.code;
        if (code === "EEXIST") return;
        throw err;
      });
    } catch {
      // Best-effort: migration failure should not prevent the session from starting.
    }
  }

  private getDocIdForOfflinePersistence(
    offline: NonNullable<CollabSessionOptions["offline"]>,
    ctx: { connectionDocId?: string; explicitDocId?: string }
  ): string {
    if (offline.mode === "file") {
      // The legacy `offline.filePath` mode identified documents purely by file
      // path. Preserve that behavior by falling back to `filePath` as the
      // persistence doc id when callers don't provide a stable `docId`.
      return (
        ctx.connectionDocId ??
        ctx.explicitDocId ??
        offline.filePath ??
        offline.key ??
        this.doc.guid
      );
    }
    // Preserve legacy semantics: `offline.key` (when provided) overrides any
    // derived doc id so callers can control the persistence namespace.
    return offline.key ?? ctx.connectionDocId ?? ctx.explicitDocId ?? this.doc.guid;
  }

  private dirnameForOfflineFilePath(filePath: string): string {
    // `@formula/collab-offline` historically accepted an explicit file path. Our
    // file persistence implementation stores one file per doc in a directory, so
    // we map `offline.filePath` to its parent directory.
    const slash = filePath.lastIndexOf("/");
    const backslash = filePath.lastIndexOf("\\");
    const idx = Math.max(slash, backslash);
    if (idx <= 0) return ".";
    return filePath.slice(0, idx);
  }

  private detachLocalPersistence(): void {
    this.persistenceDetached = true;
    const binding = this.persistenceBinding;
    this.persistenceBinding = null;
    if (binding) {
      void binding.destroy().catch(() => {});
    }
  }

  private async clearLocalPersistence(): Promise<void> {
    this.persistenceDetached = true;
    const docId = this.persistenceDocId;
    if (!docId) return;

    const binding = this.persistenceBinding;
    this.persistenceBinding = null;
    if (binding) {
      await binding.destroy().catch(() => {});
    }

    const persistence =
      this.persistence ??
      (this.localPersistenceFactory ? await this.localPersistenceFactory() : null);
    if (!persistence) return;
    this.persistence = persistence;
    if (typeof persistence.clear === "function") {
      await persistence.clear(docId);
    }

    // If we migrated from the legacy `@formula/collab-offline` file log format,
    // also clear the original `offline.filePath` so the migration doesn't
    // resurrect cleared state on the next session start.
    if (this.legacyOfflineFilePath) {
      try {
        const fsSpecifier = ["node", "fs"].join(":");
        // eslint-disable-next-line no-undef
        const { promises: fs } = await import(
          // eslint-disable-next-line no-undef
          /* @vite-ignore */ fsSpecifier
        );
        await fs.rm(this.legacyOfflineFilePath, { force: true });
      } catch {
        // Best-effort cleanup.
      }
    }
  }

  private scheduleProviderConnectAfterHydration(): void {
    if (this.providerConnectScheduled) return;
    const provider = this.provider;
    if (!provider || typeof provider.connect !== "function") return;

    this.providerConnectScheduled = true;

    const gates: Promise<void>[] = [];
    if (this.shouldGateProviderConnectOnLocalPersistence) {
      const gate = this.offline && this.offlineAutoConnectAfterLoad ? this.offline.whenLoaded() : this.whenLocalPersistenceLoaded();
      gates.push(
        gate.catch(() => {
          // Even if local persistence fails, allow the provider to connect so the
          // session still works online.
        })
      );
    }

    void Promise.all(gates)
      .finally(() => {
        if (this.isDestroyed) return;
        provider.connect?.();
      })
      .catch(() => {
        // Best-effort: gates should already swallow errors, but avoid any unexpected
        // unhandled rejection from the `.finally` bookkeeping chain.
      });
  }

  private emitStatusChange(): void {
    if (this.isDestroyed) return;
    if (this.statusListeners.size === 0) return;
    const state = this.getSyncState();
    for (const cb of Array.from(this.statusListeners)) {
      try {
        cb({ ...state });
      } catch {
        // Ignore observer errors so one listener cannot break the session.
      }
    }
  }

  destroy(): void {
    if (this.isDestroyed) return;
    this.isDestroyed = true;
    if (this.commentsMigrationPermissionsUnsubscribe) {
      try {
        this.commentsMigrationPermissionsUnsubscribe();
      } catch {
        // ignore
      }
      this.commentsMigrationPermissionsUnsubscribe = null;
    }
    if (this.provider && this.commentsMigrationSyncHandler && typeof this.provider.off === "function") {
      try {
        this.provider.off("sync", this.commentsMigrationSyncHandler);
      } catch {
        // ignore
      }
      this.commentsMigrationSyncHandler = null;
    } else {
      this.commentsMigrationSyncHandler = null;
    }
    if (this.commentsUndoScopePermissionsUnsubscribe) {
      try {
        this.commentsUndoScopePermissionsUnsubscribe();
      } catch {
        // ignore
      }
      this.commentsUndoScopePermissionsUnsubscribe = null;
    }
    if (this.provider && this.commentsUndoScopeSyncHandler && typeof this.provider.off === "function") {
      try {
        this.provider.off("sync", this.commentsUndoScopeSyncHandler);
      } catch {
        // ignore
      }
      this.commentsUndoScopeSyncHandler = null;
    } else {
      this.commentsUndoScopeSyncHandler = null;
    }
    if (this.sheetsSchemaObserver) {
      this.sheets.unobserve(this.sheetsSchemaObserver);
      this.sheetsSchemaObserver = null;
    }
    if (this.schemaSyncHandler && this.provider && typeof this.provider.off === "function") {
      this.provider.off("sync", this.schemaSyncHandler);
      this.schemaSyncHandler = null;
    }
    if (this.docUpdateListener) {
      this.doc.off("update", this.docUpdateListener);
      this.docUpdateListener = null;
    }
    if (this.provider && typeof this.provider.off === "function") {
      if (this.providerStatusListener) {
        this.provider.off("status", this.providerStatusListener);
        this.providerStatusListener = null;
      }
      if (this.providerSyncListener) {
        this.provider.off("sync", this.providerSyncListener);
        this.providerSyncListener = null;
      }
    }
    this.statusListeners.clear();
    this.formulaConflictMonitor?.dispose();
    this.cellConflictMonitor?.dispose();
    this.cellValueConflictMonitor?.dispose();

    // Collaborative undo uses Yjs' UndoManager which registers `afterTransaction`
    // handlers on the Y.Doc. CollabSession does not own the doc instance (callers
    // can pass their own), so we must explicitly destroy the UndoManager on
    // teardown to avoid leaking doc observers and retaining deleted structs via
    // `keepItem`.
    //
    // Note: the UndoManager instance is stored inside `localOrigins` (see
    // @formula/collab-undo).
    try {
      const isYUndoManager = (value: unknown): value is { destroy: () => void } => {
        if (value instanceof Y.UndoManager) return true;
        if (!value || typeof value !== "object") return false;
        const maybe = value as any;
        // Bundlers can rename constructors and pnpm workspaces can load multiple `yjs`
        // module instances (ESM + CJS). Avoid relying on `instanceof` and prefer a
        // structural check so teardown still cleans up doc observers.
        return (
          typeof maybe.addToScope === "function" &&
          typeof maybe.undo === "function" &&
          typeof maybe.redo === "function" &&
          typeof maybe.stopCapturing === "function" &&
          typeof maybe.destroy === "function"
        );
      };

      const undoManagers = Array.from(this.localOrigins).filter(isYUndoManager);
      for (const undoManager of undoManagers) undoManager.destroy();
    } catch {
      // ignore
    }

    this.presence?.destroy();
    this.offline?.destroy();
    this.provider?.destroy?.();
    this.detachLocalPersistence();
  }

  connect(): void {
    if (this.isDestroyed) return;
    if (!this.provider?.connect) return;

    if (this.shouldGateProviderConnectOnLocalPersistence) {
      this.scheduleProviderConnectAfterHydration();
      return;
    }

    this.provider.connect();
  }

  disconnect(): void {
    this.provider?.disconnect?.();
  }

  getSyncState(): CollabSessionSyncState {
    const providerAny = this.provider as any;
    const connected =
      typeof providerAny?.wsconnected === "boolean"
        ? providerAny.wsconnected
        : typeof providerAny?.connected === "boolean"
          ? providerAny.connected
          : this.syncState.connected;
    // Prefer the value we derive from provider events (sync/status) so callers
    // can use lightweight provider mocks that don't update `.synced` eagerly.
    // Fall back to `.synced` only when no listener is installed.
    const synced =
      this.providerSyncListener != null
        ? this.syncState.synced
        : typeof providerAny?.synced === "boolean"
          ? providerAny.synced
          : this.syncState.synced;
    return connected ? { connected: true, synced: Boolean(synced) } : { connected: false, synced: false };
  }

  onStatusChange(cb: (state: CollabSessionSyncState) => void): () => void {
    if (this.isDestroyed) return () => {};
    this.statusListeners.add(cb);
    return () => {
      this.statusListeners.delete(cb);
    };
  }

  getUpdateStats(): CollabSessionUpdateStats {
    if (this.recentOutgoingUpdateBytes.length === 0) {
      return { lastUpdateBytes: 0, maxRecentBytes: 0, avgRecentBytes: 0 };
    }

    let max = 0;
    let sum = 0;
    for (const bytes of this.recentOutgoingUpdateBytes) {
      sum += bytes;
      if (bytes > max) max = bytes;
    }
    const last = this.recentOutgoingUpdateBytes[this.recentOutgoingUpdateBytes.length - 1] ?? 0;
    return { lastUpdateBytes: last, maxRecentBytes: max, avgRecentBytes: sum / this.recentOutgoingUpdateBytes.length };
  }

  getLocalPersistenceState(): CollabSessionPersistenceState {
    return {
      enabled: this.hasLocalPersistence,
      loaded: this.hasLocalPersistence ? this.localPersistenceLoadedFlag : false,
      lastFlushedAt: this.lastLocalPersistenceFlushAt,
    };
  }

  whenLocalPersistenceLoaded(): Promise<void> {
    this.startLocalPersistence();
    return this.localPersistenceLoaded;
  }

  async flushLocalPersistence(opts?: CollabPersistenceFlushOptions): Promise<void> {
    this.startLocalPersistence();
    const docId = this.persistenceDocId;
    if (!docId) return;

    await this.localPersistenceLoaded.catch(() => {
      // If load failed, flushing may still be useful for subsequent updates.
    });
    const persistence = this.persistence;
    if (!persistence || typeof persistence.flush !== "function") return;
    await persistence.flush(docId, opts);
    this.lastLocalPersistenceFlushAt = Date.now();
  }

  whenSynced(timeoutMs: number = 10_000): Promise<void> {
    const provider = this.provider;
    if (!provider || typeof provider.on !== "function") return Promise.resolve();
    // Some lightweight provider mocks emit `sync=true` events without updating a
    // `.synced` property. Prefer the session's observed sync state when
    // available, and fall back to `provider.synced` for real providers like
    // y-websocket.
    const alreadySynced =
      this.providerSyncListener != null
        ? this.syncState.synced
        : typeof (provider as any)?.synced === "boolean"
          ? Boolean((provider as any).synced)
          : false;
    if (alreadySynced) return Promise.resolve();

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        if (typeof provider.off === "function") {
          try {
            provider.off("sync", handler);
          } catch {
            // ignore
          }
        }
        reject(new Error("Timed out waiting for provider sync"));
      }, timeoutMs);
      (timeout as any).unref?.();

      const handler = (isSynced: boolean) => {
        if (!isSynced) return;
        clearTimeout(timeout);
        if (typeof provider.off === "function") {
          try {
            provider.off("sync", handler);
          } catch {
            // ignore
          }
        }
        resolve();
      };

      try {
        provider.on("sync", handler);
      } catch {
        // If the provider refuses sync listeners, fall back to a best-effort
        // behavior and resolve immediately (mirrors the fail-open behavior when
        // no provider is present).
        clearTimeout(timeout);
        resolve();
        return;
      }

      // After registering, double-check whether the provider has already synced
      // (either via the `.synced` property or our observed event-based state).
      const syncedNow =
        this.providerSyncListener != null
          ? this.syncState.synced
          : typeof (provider as any)?.synced === "boolean"
            ? Boolean((provider as any).synced)
            : false;
      if (syncedNow) handler(true);
    });
  }

  setPermissions(permissions: SessionPermissions): void {
    assertValidRole(permissions.role);

    let normalizedRestrictions: unknown[] = [];
    const rawRestrictions = permissions.rangeRestrictions;
    if (rawRestrictions !== undefined) {
      if (!Array.isArray(rawRestrictions)) {
        throw new Error("rangeRestrictions must be an array");
      }

      normalizedRestrictions = [];
      for (let i = 0; i < rawRestrictions.length; i += 1) {
        try {
          normalizedRestrictions.push(normalizeRestriction(rawRestrictions[i] as any));
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          throw new Error(`rangeRestrictions[${i}] invalid: ${message}`);
        }
      }
    }

    this.permissions = {
      role: permissions.role,
      rangeRestrictions: normalizedRestrictions,
      userId: permissions.userId ?? null,
    };
    const snapshot = this.getPermissions();
    for (const listener of [...this.permissionsListeners]) {
      try {
        listener(snapshot);
      } catch {
        // ignore listener errors
      }
    }
  }

  onPermissionsChanged(listener: (permissions: SessionPermissions | null) => void): () => void {
    this.permissionsListeners.add(listener);
    try {
      listener(this.getPermissions());
    } catch {
      // ignore listener errors
    }
    return () => {
      this.permissionsListeners.delete(listener);
    };
  }

  /**
   * Returns the current session permissions, or null when `setPermissions()` has
   * never been called.
   *
   * Note: This returns a defensive copy of the permissions object so consumers
   * can't mutate the session's internal `rangeRestrictions` array by accident.
   */
  getPermissions(): SessionPermissions | null {
    const permissions = this.permissions;
    if (!permissions) return null;
    return {
      role: permissions.role,
      userId: permissions.userId ?? null,
      rangeRestrictions: Array.isArray(permissions.rangeRestrictions) ? [...permissions.rangeRestrictions] : [],
    };
  }

  /**
   * Convenience accessor for the current document role, or null when
   * `setPermissions()` has never been called.
   */
  getRole(): DocumentRole | null {
    return this.permissions?.role ?? null;
  }

  /**
   * Returns true when the current role can create comments/annotations.
   * Defaults to false when permissions are unset.
   */
  canComment(): boolean {
    const role = this.permissions?.role ?? null;
    if (!role) return false;
    return roleCanComment(role);
  }

  /**
   * Returns true when the current role can share/manage access (typically
   * owner/admin only). Defaults to false when permissions are unset.
   */
  canShare(): boolean {
    const role = this.permissions?.role ?? null;
    if (!role) return false;
    return roleCanShare(role);
  }

  /**
   * Returns true when permissions are configured and the current role is
   * read-only (i.e. cannot edit).
   */
  isReadOnly(): boolean {
    const role = this.permissions?.role ?? null;
    if (!role) return false;
    return !roleCanEdit(role);
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
    // so we can preserve encryption invariants (never write plaintext into an encrypted cell),
    // and only when the key id matches the ciphertext payload (so we never clobber encrypted
    // content we cannot decrypt, including future/unknown payload schemas).
    const encRaw = this.getEncryptedPayloadForCell(cell);
    if (encRaw !== undefined) {
      const key = this.encryption?.keyForCell(cell) ?? null;
      if (!key) return false;
      if (!isEncryptedCellPayload(encRaw)) return false;
      if (key.keyId !== encRaw.keyId) return false;
      return true;
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
      const encryptFormat = Boolean(this.encryption?.encryptFormat);
      if (isEncryptedCellPayload(encRaw)) {
        if (!parsed) {
          return {
            value: maskCellValue(null),
            formula: null,
            modified: (cell.get("modified") ?? null) as number | null,
            modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
            ...(encryptFormat ? { format: null } : {}),
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
              ...(encryptFormat ? { format: (plaintext as any)?.format ?? null } : {}),
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
        ...(Boolean(this.encryption?.encryptFormat) ? { format: null } : {}),
        encrypted: true,
      };
    }

    const rawValue = cell.get("value");
    let value: any = rawValue ?? null;
    if (value && typeof value === "object" && (value as any).t === "blank") {
      value = null;
    }

    const rawFormula = cell.get("formula");
    let formula: string | null = null;
    if (typeof rawFormula === "string") {
      const trimmed = rawFormula.trim();
      formula = trimmed ? trimmed : null;
    } else if (rawFormula != null) {
      const trimmed = String(rawFormula).trim();
      formula = trimmed ? trimmed : null;
    }

    // Treat marker-only cells (value/formula both null) as empty UI state.
    // This keeps `getCell()` semantics aligned with other read paths (e.g. the
    // DocumentController binder) even though we preserve the underlying Y.Map so
    // conflict monitors can reason about causal history.
    if (value === null && formula === null) return null;

    return {
      value,
      formula,
      modified: (cell.get("modified") ?? null) as number | null,
      modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
    };
  }

  compactCells(opts?: {
    origin?: unknown;
    maxCellsScanned?: number;
    dryRun?: boolean;
    pruneMarkerOnly?: boolean;
  }): { scanned: number; deleted: number } {
    const maxCellsScanned = opts?.maxCellsScanned ?? Number.POSITIVE_INFINITY;
    if (maxCellsScanned !== Number.POSITIVE_INFINITY && (!Number.isFinite(maxCellsScanned) || maxCellsScanned <= 0)) {
      return { scanned: 0, deleted: 0 };
    }
    const dryRun = Boolean(opts?.dryRun);

    // Marker-only cells (formula=null + no value) are preserved by some writers so
    // conflict monitors can reason about delete-vs-overwrite causality. When any
    // conflict monitors are enabled, default to *not* pruning those marker-only
    // entries unless the caller explicitly opts in.
    const conflictMonitorsEnabled = Boolean(
      this.formulaConflictMonitor || this.cellConflictMonitor || this.cellValueConflictMonitor
    );
    const pruneMarkerOnly = opts?.pruneMarkerOnly ?? !conflictMonitorsEnabled;

    const keysToDelete: string[] = [];
    let scanned = 0;

    for (const cellKey of this.cells.keys()) {
      if (scanned >= maxCellsScanned) break;
      scanned += 1;

      const cellData = this.cells.get(cellKey);
      const ycell = getYMapCell(cellData);
      const recordCell =
        !ycell && cellData && typeof cellData === "object"
          ? (Object.getPrototypeOf(cellData) === Object.prototype || Object.getPrototypeOf(cellData) === null
              ? (cellData as Record<string, any>)
              : null)
          : null;

      const get = ycell ? (k: string) => ycell.get(k) : recordCell ? (k: string) => recordCell[k] : null;
      if (!get) continue;

      const keys: string[] = [];
      if (ycell) {
        ycell.forEach((_: any, k: any) => keys.push(String(k)));
      } else if (recordCell) {
        keys.push(...Object.keys(recordCell));
      }
      const keySet = new Set(keys);

      // Encrypted cells should never be pruned automatically.
      if (keySet.has("enc")) continue;

      // Format-only cells (or explicit format clears) should not be pruned.
      // `style` is a legacy alias for `format`.
      if (keySet.has("format") || keySet.has("style")) continue;

      const rawValue = get("value");
      const valueEmpty =
        rawValue == null || (rawValue && typeof rawValue === "object" && (rawValue as any).t === "blank");
      if (!valueEmpty) continue;

      const rawFormula = get("formula");
      const formulaEmpty =
        rawFormula == null ||
        (typeof rawFormula === "string" ? rawFormula.trim() === "" : String(rawFormula).trim() === "");
      if (!formulaEmpty) continue;

      const isMarkerOnly = keySet.has("formula") && formulaEmpty && valueEmpty;
      if (isMarkerOnly && !pruneMarkerOnly) continue;

      // Only prune entries that contain no other known meaningful keys.
      // We intentionally allow best-effort metadata like `modified` / `modifiedBy`
      // so marker-only cells created by clears remain prunable when the app opts in.
      let hasOtherKeys = false;
      for (const key of keySet) {
        if (key === "value" || key === "formula" || key === "modified" || key === "modifiedBy") continue;
        hasOtherKeys = true;
        break;
      }
      if (hasOtherKeys) continue;

      keysToDelete.push(cellKey);
    }

    const deleted = keysToDelete.length;
    if (deleted === 0 || dryRun) return { scanned, deleted };

    const origin = opts?.origin ?? "cells-compact";
    this.doc.transact(() => {
      for (const key of keysToDelete) {
        this.cells.delete(key);
      }
    }, origin);

    return { scanned, deleted };
  }

  transactLocal(fn: () => void): void {
    const undoTransact = this.undo?.transact;
    if (typeof undoTransact === "function") {
      undoTransact(fn);
      return;
    }
    this.doc.transact(fn, this.origin);
  }

  /**
   * Sets a cell value directly in the shared Yjs document.
   *
   * When session permissions are configured via `setPermissions()`, this method
   * enforces edit permissions and throws when the caller cannot edit the target
   * cell.
   *
   * Note: when permissions are configured and `ignorePermissions` is not set,
   * `cellKey` must be parseable (e.g. `${sheetId}:${row}:${col}` or supported
   * legacy forms). Unparseable keys are rejected to avoid bypassing restrictions.
   *
   * Escape hatch: internal tooling (e.g. migrations/admin repair scripts) may
   * bypass permission checks by passing `{ ignorePermissions: true }`. This
   * does *not* bypass encryption invariants.
   */
  async setCellValue(cellKey: string, value: unknown, options?: { ignorePermissions?: boolean }): Promise<void> {
    const ignorePermissions = options?.ignorePermissions === true;
    if (typeof cellKey !== "string" || cellKey.length === 0) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }
    const userId = this.permissions?.userId ?? null;
    const parsedMaybe = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });

    // Fail closed: when permissions are configured, we must not allow callers to
    // bypass permission checks by providing an unparseable key.
    if (this.permissions && !ignorePermissions && !parsedMaybe) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }

    if (this.permissions && !ignorePermissions && parsedMaybe && !this.canEditCell(parsedMaybe)) {
      throw new Error(`Permission denied: cannot edit cell ${makeCellKey(parsedMaybe)}`);
    }

    const cellData = this.cells.get(cellKey);
    const existingCell = getYMapCell(cellData);
    const directEnc = existingCell?.get("enc");
    const existingEnc = directEnc !== undefined ? directEnc : (parsedMaybe ? this.getEncryptedPayloadForCell(parsedMaybe) : undefined);

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parsedMaybe : null;
    if (needsCellAddress && !parsed) {
      throw new Error(`Invalid cellKey "${String(cellKey)}": expected "SheetId:row:col"`);
    }

    const key = parsed && this.encryption ? this.encryption.keyForCell(parsed) : null;
    // If the cell already has an encrypted payload, only allow overwriting it when
    // the resolver provides a *matching* key id (and the payload schema is supported).
    // This prevents clobbering ciphertext authored by newer clients or written with a
    // different key id.
    if (existingEnc !== undefined && key) {
      if (!isEncryptedCellPayload(existingEnc)) {
        throw new Error(`Unsupported encrypted cell payload for cell ${cellKey}`);
      }
      if (key.keyId !== existingEnc.keyId) {
        throw new Error(
          `Encryption key id mismatch for cell ${cellKey} (payload=${existingEnc.keyId}, resolver=${key.keyId})`
        );
      }
    }
    const encryptFormat = Boolean(this.encryption?.encryptFormat);
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

      let formatToEncrypt: any | null = null;
      if (encryptFormat && parsed) {
        // Preserve existing formatting while migrating it into the encrypted payload.
        // Prefer decrypting the current payload when possible, otherwise fall back to any
        // plaintext `format` key stored under legacy aliases.
        if (existingEnc !== undefined && isEncryptedCellPayload(existingEnc) && key.keyId === existingEnc.keyId) {
          try {
            const plaintext = (await decryptCellPlaintext({
              encrypted: existingEnc,
              key,
              context: {
                docId: this.docIdForEncryption,
                sheetId: parsed.sheetId,
                row: parsed.row,
                col: parsed.col,
              },
            })) as CellPlaintext;
            const candidate = (plaintext as any)?.format ?? null;
            if (candidate && typeof candidate === "object" && Object.keys(candidate as any).length === 0) {
              formatToEncrypt = null;
            } else if (candidate != null) {
              formatToEncrypt = candidate;
            }
          } catch {
            // Ignore: fall back to plaintext `format` lookup below.
          }
        }

        if (formatToEncrypt == null) {
          const keysToCheck = Array.from(
            new Set([
              cellKey,
              makeCellKey(parsed),
              `${parsed.sheetId}:${parsed.row},${parsed.col}`,
              ...(parsed.sheetId === this.defaultSheetId ? [`r${parsed.row}c${parsed.col}`] : []),
            ])
          );
          for (const key of keysToCheck) {
            const ycell = getYMapCell(this.cells.get(key));
            if (!ycell) continue;
            const raw = ycell.get("format");
            const rawOrStyle = raw === undefined ? ycell.get("style") : raw;
            if (rawOrStyle === undefined) continue;
            if (rawOrStyle && typeof rawOrStyle === "object" && Object.keys(rawOrStyle as any).length === 0) {
              formatToEncrypt = null;
            } else {
              formatToEncrypt = rawOrStyle ?? null;
            }
            break;
          }
        }
      }

      /** @type {any} */
      const plaintext: any = { value: value ?? null, formula: null };
      if (encryptFormat && formatToEncrypt != null) {
        plaintext.format = formatToEncrypt;
      }
      encryptedPayload = await encryptCellPlaintext({
        plaintext,
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
          if (encryptFormat) {
            ycell.delete("format");
            ycell.delete("style");
          }
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
      // Preserve explicit formula clear markers (`formula=null`) so other
      // clients can reason about delete-vs-overwrite causality via Yjs Item
      // origin ids (Y.Map deletes do not create a new Item).
      cell.set("formula", null);
      cell.set("modified", modified);
      if (userId) cell.set("modifiedBy", userId);
    });
  }

  /**
   * Sets a cell formula directly in the shared Yjs document.
   *
   * When session permissions are configured via `setPermissions()`, this method
   * enforces edit permissions and throws when the caller cannot edit the target
   * cell.
   *
   * Note: when permissions are configured and `ignorePermissions` is not set,
   * `cellKey` must be parseable (e.g. `${sheetId}:${row}:${col}` or supported
   * legacy forms). Unparseable keys are rejected to avoid bypassing restrictions.
   *
   * Escape hatch: internal tooling (e.g. migrations/admin repair scripts) may
   * bypass permission checks by passing `{ ignorePermissions: true }`. This
   * does *not* bypass encryption invariants.
   */
  async setCellFormula(
    cellKey: string,
    formula: string | null,
    options?: { ignorePermissions?: boolean }
  ): Promise<void> {
    const ignorePermissions = options?.ignorePermissions === true;
    if (typeof cellKey !== "string" || cellKey.length === 0) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }
    const parsedMaybe = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });

    // Fail closed: when permissions are configured, we must not allow callers to
    // bypass permission checks by providing an unparseable key.
    if (this.permissions && !ignorePermissions && !parsedMaybe) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }

    if (this.permissions && !ignorePermissions && parsedMaybe && !this.canEditCell(parsedMaybe)) {
      throw new Error(`Permission denied: cannot edit cell ${makeCellKey(parsedMaybe)}`);
    }

    const cellData = this.cells.get(cellKey);
    const existingCell = getYMapCell(cellData);
    const directEnc = existingCell?.get("enc");
    const existingEnc = directEnc !== undefined ? directEnc : (parsedMaybe ? this.getEncryptedPayloadForCell(parsedMaybe) : undefined);

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parsedMaybe : null;
    if (needsCellAddress && !parsed) {
      throw new Error(`Invalid cellKey "${String(cellKey)}": expected "SheetId:row:col"`);
    }

    const key = parsed && this.encryption ? this.encryption.keyForCell(parsed) : null;
    // If the cell already has an encrypted payload, only allow overwriting it when
    // the resolver provides a *matching* key id (and the payload schema is supported).
    if (existingEnc !== undefined && key) {
      if (!isEncryptedCellPayload(existingEnc)) {
        throw new Error(`Unsupported encrypted cell payload for cell ${cellKey}`);
      }
      if (key.keyId !== existingEnc.keyId) {
        throw new Error(
          `Encryption key id mismatch for cell ${cellKey} (payload=${existingEnc.keyId}, resolver=${key.keyId})`
        );
      }
    }
    const encryptFormat = Boolean(this.encryption?.encryptFormat);
    const wantsEncryption =
      existingEnc !== undefined ||
      (parsed
        ? (typeof this.encryption?.shouldEncryptCell === "function" ? this.encryption.shouldEncryptCell(parsed) : key != null)
        : false);

    const nextFormula = (formula ?? "").trim();
    if (wantsEncryption) {
      if (!key) throw new Error(`Missing encryption key for cell ${cellKey}`);

      let formatToEncrypt: any | null = null;
      if (encryptFormat) {
        // Preserve existing formatting while migrating it into the encrypted payload.
        if (existingEnc !== undefined && isEncryptedCellPayload(existingEnc) && key.keyId === existingEnc.keyId) {
          try {
            const plaintext = (await decryptCellPlaintext({
              encrypted: existingEnc,
              key,
              context: {
                docId: this.docIdForEncryption,
                sheetId: parsed.sheetId,
                row: parsed.row,
                col: parsed.col,
              },
            })) as CellPlaintext;
            const candidate = (plaintext as any)?.format ?? null;
            if (candidate && typeof candidate === "object" && Object.keys(candidate as any).length === 0) {
              formatToEncrypt = null;
            } else if (candidate != null) {
              formatToEncrypt = candidate;
            }
          } catch {
            // Ignore: fall back to plaintext `format` lookup below.
          }
        }

        if (formatToEncrypt == null) {
          const keysToCheck = Array.from(
            new Set([
              cellKey,
              makeCellKey(parsed),
              `${parsed.sheetId}:${parsed.row},${parsed.col}`,
              ...(parsed.sheetId === this.defaultSheetId ? [`r${parsed.row}c${parsed.col}`] : []),
            ])
          );
          for (const key of keysToCheck) {
            const ycell = getYMapCell(this.cells.get(key));
            if (!ycell) continue;
            const raw = ycell.get("format");
            const rawOrStyle = raw === undefined ? ycell.get("style") : raw;
            if (rawOrStyle === undefined) continue;
            if (rawOrStyle && typeof rawOrStyle === "object" && Object.keys(rawOrStyle as any).length === 0) {
              formatToEncrypt = null;
            } else {
              formatToEncrypt = rawOrStyle ?? null;
            }
            break;
          }
        }
      }

      /** @type {any} */
      const plaintext: any = { value: null, formula: nextFormula || null };
      if (encryptFormat && formatToEncrypt != null) {
        plaintext.format = formatToEncrypt;
      }

      const encryptedPayload = await encryptCellPlaintext({
        plaintext,
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
          if (encryptFormat) {
            cell.delete("format");
            cell.delete("style");
          }
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
      // Store a null marker rather than deleting the key so subsequent writes
      // can causally reference this deletion via Item.origin. Yjs map deletes
      // do not create a new Item, which makes delete-vs-overwrite concurrency
      // ambiguous without an explicit marker.
      else cell.set("formula", null);

      cell.set("value", null);
      cell.set("modified", Date.now());
      if (userId) cell.set("modifiedBy", userId);
    });
  }

  async setCells(
    updates: Array<{ cellKey: string; value?: unknown; formula?: string | null }>,
    opts: { ignorePermissions?: boolean } = {}
  ): Promise<void> {
    if (!Array.isArray(updates) || updates.length === 0) return;
    const ignorePermissions = opts?.ignorePermissions === true;

    const planned: Array<{
      cellKey: string;
      parsed: CellAddress | null;
      kind: "value" | "formula";
      nextValue: unknown | null;
      nextFormula: string;
      shouldEncrypt: boolean;
      encryptionCell: CellAddress | null;
      encryptedPayload: unknown | null;
    }> = [];

    for (const update of updates) {
      const cellKey = (update as any)?.cellKey;
      if (typeof cellKey !== "string" || cellKey.length === 0) {
        throw new Error(`Invalid cellKey: ${String(cellKey)}`);
      }

      // Formula wins when both are provided.
      const kind: "value" | "formula" = update.formula !== undefined ? "formula" : "value";

      const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
      planned.push({
        cellKey,
        parsed,
        kind,
        nextValue: kind === "value" ? (update.value ?? null) : null,
        nextFormula: kind === "formula" ? (update.formula ?? "").trim() : "",
        shouldEncrypt: false,
        encryptionCell: null,
        encryptedPayload: null,
      });
    }

    // Permission gate: reject the entire batch before encrypting/writing anything.
    if (this.permissions && !ignorePermissions) {
      for (const update of planned) {
        if (!update.parsed) throw new Error(`Invalid cellKey: ${update.cellKey}`);
        if (!this.canEditCell(update.parsed)) {
          throw new Error(`Permission denied: cannot edit cell ${makeCellKey(update.parsed)}`);
        }
      }
    }

    /** @type {Array<Promise<void>>} */
    const encryptions: Array<Promise<void>> = [];
    const encryptFormat = Boolean(this.encryption?.encryptFormat);

    for (const update of planned) {
      const cellData = this.cells.get(update.cellKey);
      const existingCell = getYMapCell(cellData);
      const directEnc = existingCell?.get("enc");
      const existingEnc =
        directEnc !== undefined ? directEnc : (update.parsed ? this.getEncryptedPayloadForCell(update.parsed) : undefined);

      const needsCellAddress = this.encryption != null || existingEnc !== undefined;
      const encryptionCell = needsCellAddress ? update.parsed : null;
      if (needsCellAddress && !encryptionCell) throw new Error(`Invalid cellKey: ${update.cellKey}`);

      const key = encryptionCell && this.encryption ? this.encryption.keyForCell(encryptionCell) : null;
      const shouldEncrypt =
        existingEnc !== undefined ||
        (encryptionCell
          ? (typeof this.encryption?.shouldEncryptCell === "function" ? this.encryption.shouldEncryptCell(encryptionCell) : key != null)
          : false);

      update.shouldEncrypt = shouldEncrypt;
      update.encryptionCell = encryptionCell;

      if (!shouldEncrypt) continue;
      if (!key) throw new Error(`Missing encryption key for cell ${update.cellKey}`);
      if (existingEnc !== undefined) {
        if (!isEncryptedCellPayload(existingEnc)) {
          throw new Error(`Unsupported encrypted cell payload for cell ${update.cellKey}`);
        }
        if (key.keyId !== existingEnc.keyId) {
          throw new Error(
            `Encryption key id mismatch for cell ${update.cellKey} (payload=${existingEnc.keyId}, resolver=${key.keyId})`
          );
        }
      }

      /** @type {CellPlaintext} */
      const plaintext: CellPlaintext =
        update.kind === "formula"
          ? { value: null, formula: update.nextFormula || null }
          : { value: update.nextValue ?? null, formula: null };

      if (encryptFormat && encryptionCell) {
        let formatToEncrypt: any | null = null;

        // Prefer decrypting the existing payload to preserve encrypted formatting when possible.
        if (existingEnc !== undefined && isEncryptedCellPayload(existingEnc) && key.keyId === existingEnc.keyId) {
          try {
            const prev = (await decryptCellPlaintext({
              encrypted: existingEnc,
              key,
              context: {
                docId: this.docIdForEncryption,
                sheetId: encryptionCell.sheetId,
                row: encryptionCell.row,
                col: encryptionCell.col,
              },
            })) as CellPlaintext;
            const candidate = (prev as any)?.format ?? null;
            if (candidate && typeof candidate === "object" && Object.keys(candidate as any).length === 0) {
              formatToEncrypt = null;
            } else if (candidate != null) {
              formatToEncrypt = candidate;
            }
          } catch {
            // Ignore and fall back to plaintext `format`.
          }
        }

        // Back-compat migration: preserve any plaintext `format` stored under legacy aliases.
        if (formatToEncrypt == null) {
          const keysToCheck = Array.from(
            new Set([
              update.cellKey,
              makeCellKey(encryptionCell),
              `${encryptionCell.sheetId}:${encryptionCell.row},${encryptionCell.col}`,
              ...(encryptionCell.sheetId === this.defaultSheetId
                ? [`r${encryptionCell.row}c${encryptionCell.col}`]
                : []),
            ])
          );
          for (const key of keysToCheck) {
            const ycell = getYMapCell(this.cells.get(key));
            if (!ycell) continue;
            const raw = ycell.get("format");
            const rawOrStyle = raw === undefined ? ycell.get("style") : raw;
            if (rawOrStyle === undefined) continue;
            if (rawOrStyle && typeof rawOrStyle === "object" && Object.keys(rawOrStyle as any).length === 0) {
              formatToEncrypt = null;
            } else {
              formatToEncrypt = rawOrStyle ?? null;
            }
            break;
          }
        }

        if (formatToEncrypt != null) {
          (plaintext as any).format = formatToEncrypt;
        }
      }

      encryptions.push(
        encryptCellPlaintext({
          plaintext,
          key,
          context: {
            docId: this.docIdForEncryption,
            sheetId: encryptionCell!.sheetId,
            row: encryptionCell!.row,
            col: encryptionCell!.col,
          },
        }).then((payload) => {
          update.encryptedPayload = payload;
        })
      );
    }

    if (encryptions.length > 0) {
      await Promise.all(encryptions);
    }

    // Normalize foreign nested Y.Maps before mutating them (same semantics as setCellValue/setCellFormula).
    for (const update of planned) {
      if (update.kind === "value") {
        // `setCellValue` skips this normalization when delegating to CellConflictMonitor.
        if (!update.shouldEncrypt && this.cellValueConflictMonitor) continue;
        this.ensureLocalCellMapForWrite(update.cellKey);
        continue;
      }

      // `setCellFormula` skips this normalization when delegating to FormulaConflictMonitor.
      if (!update.shouldEncrypt && this.formulaConflictMonitor) continue;
      this.ensureLocalCellMapForWrite(update.cellKey);
    }

    const userId = this.permissions?.userId ?? null;
    const formulaMonitor = this.formulaConflictMonitor;
    const cellValueMonitor = this.cellValueConflictMonitor;

    this.transactLocal(() => {
      // Preflight: ensure we never write plaintext into encrypted cells. Validate
      // every update *before* applying any, so failures leave the doc unchanged.
      for (const update of planned) {
        if (update.shouldEncrypt) continue;
        if (!update.parsed) continue;
        const directEnc = getYMapCell(this.cells.get(update.cellKey))?.get("enc");
        const encRaw = directEnc !== undefined ? directEnc : this.getEncryptedPayloadForCell(update.parsed);
        if (encRaw !== undefined) {
          throw new Error(`Refusing to write plaintext to encrypted cell ${update.cellKey}`);
        }
      }

      for (const update of planned) {
        if (update.shouldEncrypt) {
          const modified = Date.now();
          const enc = update.encryptedPayload;
          if (!enc) throw new Error(`Missing encrypted payload for cell ${update.cellKey}`);

          const cell = update.encryptionCell;
          const keysToUpdate = cell
            ? Array.from(
                new Set([
                  update.cellKey,
                  makeCellKey(cell),
                  `${cell.sheetId}:${cell.row},${cell.col}`,
                  ...(cell.sheetId === this.defaultSheetId ? [`r${cell.row}c${cell.col}`] : []),
                ])
              )
            : [update.cellKey];

          for (const key of keysToUpdate) {
            if (key !== update.cellKey && !this.cells.has(key)) continue;
            let ycell = getYMapCell(this.cells.get(key));
            if (!ycell) {
              ycell = new Y.Map();
              this.cells.set(key, ycell);
            }
            ycell.set("enc", enc);
            ycell.delete("value");
            ycell.delete("formula");
            if (encryptFormat) {
              ycell.delete("format");
              ycell.delete("style");
            }
            ycell.set("modified", modified);
            if (userId) ycell.set("modifiedBy", userId);
          }
          continue;
        }

        if (update.kind === "formula") {
          if (formulaMonitor) {
            // FormulaConflictMonitor implements the correct null-marker semantics.
            formulaMonitor.setLocalFormula(update.cellKey, update.nextFormula);
            continue;
          }

          let cell = getYMapCell(this.cells.get(update.cellKey));
          if (!cell) {
            cell = new Y.Map();
            this.cells.set(update.cellKey, cell);
          }

          cell.delete("enc");
          if (update.nextFormula) cell.set("formula", update.nextFormula);
          else cell.set("formula", null);
          cell.set("value", null);
          cell.set("modified", Date.now());
          if (userId) cell.set("modifiedBy", userId);
          continue;
        }

        // Value update.
        if (cellValueMonitor) {
          cellValueMonitor.setLocalValue(update.cellKey, update.nextValue ?? null);
          continue;
        }

        if (formulaMonitor && this.formulaConflictsIncludeValueConflicts) {
          formulaMonitor.setLocalValue(update.cellKey, update.nextValue ?? null);
          continue;
        }

        let cell = getYMapCell(this.cells.get(update.cellKey));
        if (!cell) {
          cell = new Y.Map();
          this.cells.set(update.cellKey, cell);
        }

        cell.delete("enc");
        cell.set("value", update.nextValue ?? null);
        // Preserve explicit formula clear markers (`formula=null`) so other clients
        // can reason about delete-vs-overwrite causality via Yjs Item origin ids.
        cell.set("formula", null);
        cell.set("modified", Date.now());
        if (userId) cell.set("modifiedBy", userId);
      }
    });
  }

  async safeSetCellValue(cellKey: string, value: unknown): Promise<boolean> {
    if (typeof cellKey !== "string" || cellKey.length === 0) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    if (!this.canEditCell(parsed)) return false;

    await this.setCellValue(makeCellKey(parsed), value);
    return true;
  }

  async safeSetCellFormula(cellKey: string, formula: string | null): Promise<boolean> {
    if (typeof cellKey !== "string" || cellKey.length === 0) {
      throw new Error(`Invalid cellKey: ${String(cellKey)}`);
    }
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${String(cellKey)}`);
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

export function createSheetManagerForSessionWithPermissions(session: CollabSession) {
  return createSheetManagerForSessionWithPermissionsImpl(session);
}

export function createMetadataManagerForSessionWithPermissions(session: CollabSession) {
  return createMetadataManagerForSessionWithPermissionsImpl(session);
}

export function createNamedRangeManagerForSessionWithPermissions(session: CollabSession) {
  return createNamedRangeManagerForSessionWithPermissionsImpl(session);
}

export type DocumentControllerBinder = {
  destroy: () => void;
  rehydrate?: () => void;
  whenIdle?: () => Promise<void>;
};

export async function bindCollabSessionToDocumentController(options: {
  session: CollabSession;
  documentController: any;
  undoService?: { transact?: (fn: () => void) => void; origin?: any; localOrigins?: Set<any> } | null;
  defaultSheetId?: string;
  userId?: string | null;
  /**
   * When true, suppress per-cell formatting for masked cells (unreadable due to
   * permissions, or encrypted without an available key). Defaults to false to
   * preserve existing formatting semantics.
   */
  maskCellFormat?: boolean;
  /**
   * Called when the underlying binder rejects a local edit (permissions, missing encryption key, etc).
   *
   * The binder will revert the local `DocumentController` state to keep it consistent with the shared
   * Yjs document; this hook allows UIs to surface a user-facing message instead of silently "snapping
   * back".
   */
  onEditRejected?: (rejected: any[]) => void;
  /**
   * Opt-in binder write semantics needed for `FormulaConflictMonitor` to reliably
   * detect true offline/concurrent conflicts when edits flow through the desktop
   * UI path (DocumentController → binder → Yjs).
   */
  formulaConflictsMode?: "off" | "formula" | "formula+value";
}): Promise<DocumentControllerBinder> {
  const { session, documentController, undoService, defaultSheetId, userId, maskCellFormat, onEditRejected, formulaConflictsMode } =
    options ?? ({} as any);
  if (!session) throw new Error("bindCollabSessionToDocumentController requires { session }");
  if (!documentController)
    throw new Error("bindCollabSessionToDocumentController requires { documentController }");

  // Avoid importing the Node-oriented binder (and its encryption dependencies)
  // unless a consumer explicitly opts into DocumentController wiring.
  const { bindYjsToDocumentController } = await import("../../binder/index.js");

  // Default to running DocumentController-driven writes through the session's
  // local transaction helper so edits use the session's origin token. This keeps
  // collaborative undo and conflict monitors consistent across both
  // DocumentController edits and direct `session.setCell*` calls.
  //
  // Callers can override this by providing an explicit `undoService` (or `null`
  // to disable the wrapper).
  const defaultUndoService =
    undoService === undefined
      ? {
          transact: (fn: () => void) => session.transactLocal(fn),
          origin: session.origin,
          localOrigins: session.localOrigins,
        }
      : undoService;

  // Ensure the binder's origin is treated as "local" for downstream monitors
  // (e.g. FormulaConflictMonitor), even when callers provide a custom undo service.
  //
  // Note: the binder uses an `applyingLocal` guard for echo suppression, so it is
  // safe (and preferable) for DocumentController-driven writes to share the
  // session's origin token by default.
  const binderOrigin = defaultUndoService?.origin ?? { type: "document-controller:binder" };
  session.localOrigins.add(binderOrigin);
  const normalizedUndoService = defaultUndoService ? { ...defaultUndoService, origin: binderOrigin } : { origin: binderOrigin };

  const sessionPermissions = session.getPermissions();

  return bindYjsToDocumentController({
    ydoc: session.doc,
    documentController,
    undoService: normalizedUndoService,
    defaultSheetId,
    userId,
    encryption: session.getEncryptionConfig(),
    canReadCell: (cell) => session.canReadCell(cell),
    canEditCell: (cell) => session.canEditCell(cell),
    // In non-edit collab roles (viewer/commenter), shared-state writes (sheet view state,
    // sheet-level formatting defaults, etc) must not be written into the shared Yjs
    // document. Some UI paths may still emit these deltas (startup races, programmatic
    // calls, etc); this guard is defense-in-depth.
    canWriteSharedState: () => !session.isReadOnly(),
    permissions: sessionPermissions
      ? {
          role: sessionPermissions.role,
          restrictions: sessionPermissions.rangeRestrictions,
          userId: sessionPermissions.userId,
        }
      : null,
    // Use the standard enterprise mask. The binder also uses this hook for
    // encrypted cells that cannot be decrypted.
    maskCellValue: (value) => maskCellValue(value),
    maskCellFormat,
    onEditRejected,
    formulaConflictsMode,
  }) as DocumentControllerBinder;
}

export {
  migrateLegacyCellKeys,
  type CellKeyMigrationConflictStrategy,
  type MigrateLegacyCellKeysOptions,
  type MigrateLegacyCellKeysResult,
} from "./migrations/migrateCellKeys.js";
