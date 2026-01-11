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
import { ensureWorkbookSchema } from "@formula/collab-workbook";
import {
  decryptCellPlaintext,
  encryptCellPlaintext,
  isEncryptedCellPayload,
  type CellEncryptionKey,
  type CellPlaintext,
} from "@formula/collab-encryption";

import { assertValidRole, getCellPermissions, maskCellValue } from "../../permissions/index.js";
import {
  makeCellKey as makeCellKeyImpl,
  normalizeCellKey as normalizeCellKeyImpl,
  parseCellKey as parseCellKeyImpl,
} from "./cell-key.js";

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
  if (!existing) return doc.getMap("comments");
  if (existing instanceof Y.Map) return existing;
  if (existing instanceof Y.Array) return existing;
  const placeholder = existing as any;
  const hasStart = placeholder?._start != null; // sequence item => likely array
  const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
  const kind = hasStart && mapSize === 0 ? "array" : "map";
  return kind === "array" ? doc.getArray("comments") : doc.getMap("comments");
}

export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

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
  doc?: Y.Doc;
  /**
   * Convenience option to construct a y-websocket provider for this session.
   * When provided, `session.provider` will be a `WebsocketProvider` instance.
   */
  connection?: CollabSessionConnectionOptions;
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
   */
  formulaConflicts?: {
    localUserId: string;
    onConflict: (conflict: FormulaConflict) => void;
    concurrencyWindowMs?: number;
  };
  /**
   * When enabled, the session monitors structural operations (moves / deletes)
   * for true offline conflicts and surfaces them via `onConflict`.
   */
  cellConflicts?: {
    localUserId: string;
    onConflict: (conflict: CellStructuralConflict) => void;
    maxOpRecordsPerUser?: number;
  };
  /**
   * When enabled, the session monitors cell value updates for true conflicts
   * (offline/concurrent same-cell edits) and surfaces them via `onConflict`.
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
  private sheetsSchemaObserver: (() => void) | null = null;
  private ensuringSchema = false;
  private readonly offlineAutoConnectAfterLoad: boolean;

  constructor(options: CollabSessionOptions = {}) {
    // When connecting to a sync provider, use the provider document id as the
    // Y.Doc guid to make encryption AAD stable across clients by default.
    this.doc =
      options.doc ??
      new Y.Doc(options.connection?.docId ? { guid: options.connection.docId } : undefined);

    if (options.connection && options.provider) {
      throw new Error("CollabSession cannot be constructed with both `connection` and `provider` options");
    }

    const offlineEnabled = options.offline != null;
    const offlineAutoLoad = options.offline?.autoLoad ?? true;
    const offlineAutoConnectAfterLoad =
      offlineEnabled && options.connection ? (options.offline?.autoConnectAfterLoad ?? true) : false;
    this.offlineAutoConnectAfterLoad = offlineAutoConnectAfterLoad;

    this.provider =
      options.provider ??
      (options.connection
        ? new WebsocketProvider(options.connection.wsUrl, options.connection.docId, this.doc, {
            WebSocketPolyfill: options.connection.WebSocketPolyfill,
            disableBc: options.connection.disableBc,
            params: {
               ...(options.connection.params ?? {}),
               ...(options.connection.token !== undefined ? { token: options.connection.token } : {}),
             },
             connect: offlineAutoConnectAfterLoad ? false : true,
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
          await handle.whenLoaded();
          state.isLoaded = true;
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

      if (offlineAutoConnectAfterLoad && this.provider?.connect) {
        void state
          .whenLoaded()
          .catch(() => {
            // Ignore offline load errors when auto-connecting. Callers can
            // observe the rejection by awaiting `session.offline.whenLoaded()`.
          })
          .then(() => this.provider?.connect?.());
      }
    }

    const schemaAutoInit = options.schema?.autoInit ?? true;
    const schemaDefaultSheetId = options.schema?.defaultSheetId ?? options.defaultSheetId ?? "Sheet1";
    const schemaDefaultSheetName = options.schema?.defaultSheetName ?? schemaDefaultSheetId;
    this.defaultSheetId = schemaDefaultSheetId;

    this.cells = this.doc.getMap<unknown>("cells");
    this.sheets = this.doc.getArray<Y.Map<unknown>>("sheets");
    this.metadata = this.doc.getMap<unknown>("metadata");
    this.namedRanges = this.doc.getMap<unknown>("namedRanges");

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

    if (schemaAutoInit) {
      const provider = this.provider;
      let providerSynced = !(provider && typeof provider.on === "function");
      if (provider && typeof provider.on === "function") {
        providerSynced = Boolean(provider.synced);
      }

      // When offline persistence is enabled, schema initialization must wait for
      // offline state to load to avoid creating default sheets that race with
      // persisted document state.
      const shouldWaitForOffline = this.offline != null;
      let offlineReady = !shouldWaitForOffline;

      const ensureSchema = () => {
        // Avoid mutating the workbook schema while a sync provider is still in
        // the middle of initial hydration. In particular, sheets can be created
        // incrementally (e.g. map inserted before its `id` field is applied),
        // and eagerly inserting a default sheet during that window can create
        // spurious extra sheets.
        if (!providerSynced) return;
        if (!offlineReady) return;
        if (this.ensuringSchema) return;
        this.ensuringSchema = true;
        try {
          ensureWorkbookSchema(this.doc, {
            defaultSheetId: schemaDefaultSheetId,
            defaultSheetName: schemaDefaultSheetName,
          });
        } finally {
          this.ensuringSchema = false;
        }
      };

      if (shouldWaitForOffline) {
        void this.offline!
          .whenLoaded()
          .catch(() => {
            // Schema init should still run even if offline persistence fails to load.
          })
          .then(() => {
            offlineReady = true;
            ensureSchema();
          });
      }

      // Keep the sheets array well-formed over time (e.g. remove duplicate ids).
      // This primarily protects against concurrent schema initialization when two
      // clients join a brand new document at the same time.
      this.sheetsSchemaObserver = () => ensureSchema();
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

      const builtInScopeNames = new Set(["cells", "sheets", "metadata", "namedRanges", "comments"]);
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

  destroy(): void {
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
  }

  connect(): void {
    if (!this.provider?.connect) return;

    if (this.offline && this.offlineAutoConnectAfterLoad && !this.offline.isLoaded) {
      void this.offline
        .whenLoaded()
        .catch(() => {
          // Ignore offline load errors. Callers can await `session.offline.whenLoaded()`.
        })
        .then(() => this.provider?.connect?.());
      return;
    }

    this.provider.connect();
  }

  disconnect(): void {
    this.provider?.disconnect?.();
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
    if (!this.permissions) return true;
    return getCellPermissions({
      role: this.permissions.role,
      restrictions: this.permissions.rangeRestrictions,
      userId: this.permissions.userId,
      cell,
    }).canEdit;
  }

  canReadCell(cell: CellAddress): boolean {
    if (!this.permissions) return true;
    return getCellPermissions({
      role: this.permissions.role,
      restrictions: this.permissions.rangeRestrictions,
      userId: this.permissions.userId,
      cell,
    }).canRead;
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
    const cellData = this.cells.get(cellKey);
    const cell = getYMapCell(cellData);
    if (!cell) return null;

    const encRaw = cell.get("enc");
    if (encRaw !== undefined) {
      if (isEncryptedCellPayload(encRaw)) {
        const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
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
    const existingEnc = existingCell?.get("enc");

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId }) : null;
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

    this.transactLocal(() => {
      let cellData = this.cells.get(cellKey);
      let cell = getYMapCell(cellData);
      if (!cell) {
        cell = new Y.Map();
        this.cells.set(cellKey, cell);
      }

      // Re-check inside the transaction to avoid racing with remote updates that
      // may have encrypted this cell while we were preparing a plaintext write.
      if (!encryptedPayload && cell.get("enc") !== undefined) {
        throw new Error(`Refusing to write plaintext to encrypted cell ${cellKey}`);
      }

      if (encryptedPayload) {
        cell.set("enc", encryptedPayload);
        cell.delete("value");
        cell.delete("formula");
      } else {
        cell.delete("enc");
        cell.set("value", value ?? null);
        cell.delete("formula");
      }
      cell.set("modified", Date.now());
      if (userId) cell.set("modifiedBy", userId);
    });
  }

  async setCellFormula(cellKey: string, formula: string | null): Promise<void> {
    const cellData = this.cells.get(cellKey);
    const existingCell = getYMapCell(cellData);
    const existingEnc = existingCell?.get("enc");

    const needsCellAddress = this.encryption != null || existingEnc !== undefined;
    const parsed = needsCellAddress ? parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId }) : null;
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
      this.transactLocal(() => {
        let cellData = this.cells.get(cellKey);
        let cell = getYMapCell(cellData);
        if (!cell) {
          cell = new Y.Map();
          this.cells.set(cellKey, cell);
        }

        cell.set("enc", encryptedPayload);
        cell.delete("value");
        cell.delete("formula");

        cell.set("modified", Date.now());
        if (userId) cell.set("modifiedBy", userId);
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
        // We don't sync calculated values. Clearing `value` marks the cell dirty
        // for the local formula engine to recompute.
        cell.set("value", null);
      });
      return;
    }

    const userId = this.permissions?.userId ?? null;

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
    canReadCell: (cell) => session.canReadCell(cell),
    canEditCell: (cell) => session.canEditCell(cell),
    // Use the standard enterprise mask. The binder also uses this hook for
    // encrypted cells that cannot be decrypted.
    maskCellValue: (value) => maskCellValue(value),
  }) as DocumentControllerBinder;
}
