import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import { PresenceManager } from "@formula/collab-presence";

import { assertValidRole, getCellPermissions, maskCellValue } from "../../permissions/index.js";

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
}

export interface CollabCell {
  value: unknown;
  formula: string | null;
  modified: number | null;
  modifiedBy: string | null;
}

export function makeCellKey(cell: CellAddress): string {
  return `${cell.sheetId}:${cell.row}:${cell.col}`;
}

export function parseCellKey(
  key: string,
  options: { defaultSheetId?: string } = {}
): CellAddress | null {
  const defaultSheetId = options.defaultSheetId ?? "Sheet1";
  if (typeof key !== "string" || key.length === 0) return null;

  const parts = key.split(":");
  if (parts.length === 3) {
    const sheetId = parts[0] || defaultSheetId;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  // Some internal modules use `${sheetId}:${row},${col}`.
  if (parts.length === 2) {
    const sheetId = parts[0] || defaultSheetId;
    const m = parts[1].match(/^(\d+),(\d+)$/);
    if (m) {
      return { sheetId, row: Number(m[1]), col: Number(m[2]) };
    }
  }

  const m = key.match(/^r(\d+)c(\d+)$/);
  if (m) {
    return { sheetId: defaultSheetId, row: Number(m[1]), col: Number(m[2]) };
  }

  return null;
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

  readonly provider: CollabSessionProvider | null;
  readonly awareness: unknown;
  readonly presence: PresenceManager | null;

  private permissions: SessionPermissions | null = null;
  private readonly defaultSheetId: string;

  constructor(options: CollabSessionOptions = {}) {
    this.doc = options.doc ?? new Y.Doc();
    this.cells = this.doc.getMap("cells");

    if (options.connection && options.provider) {
      throw new Error("CollabSession cannot be constructed with both `connection` and `provider` options");
    }

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
          })
        : null);
    this.awareness = options.awareness ?? this.provider?.awareness ?? null;
    this.defaultSheetId = options.defaultSheetId ?? "Sheet1";

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
    this.presence?.destroy();
    this.provider?.destroy?.();
  }

  connect(): void {
    this.provider?.connect?.();
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

  getCell(cellKey: string): CollabCell | null {
    const cellData = this.cells.get(cellKey);
    const cell = getYMapCell(cellData);
    if (!cell) return null;

    return {
      value: cell.get("value") ?? null,
      formula: (cell.get("formula") ?? null) as string | null,
      modified: (cell.get("modified") ?? null) as number | null,
      modifiedBy: (cell.get("modifiedBy") ?? null) as string | null,
    };
  }

  setCellValue(cellKey: string, value: unknown): void {
    const userId = this.permissions?.userId ?? null;

    this.doc.transact(() => {
      let cellData = this.cells.get(cellKey);
      let cell = getYMapCell(cellData);
      if (!cell) {
        cell = new Y.Map();
        this.cells.set(cellKey, cell);
      }

      cell.set("value", value ?? null);
      cell.delete("formula");
      cell.set("modified", Date.now());
      if (userId) cell.set("modifiedBy", userId);
    }, "collab-session:setCellValue");
  }

  setCellFormula(cellKey: string, formula: string | null): void {
    const userId = this.permissions?.userId ?? null;

    this.doc.transact(() => {
      let cellData = this.cells.get(cellKey);
      let cell = getYMapCell(cellData);
      if (!cell) {
        cell = new Y.Map();
        this.cells.set(cellKey, cell);
      }

      const nextFormula = (formula ?? "").trim();
      if (nextFormula) cell.set("formula", nextFormula);
      else cell.delete("formula");

      cell.set("value", null);
      cell.set("modified", Date.now());
      if (userId) cell.set("modifiedBy", userId);
    }, "collab-session:setCellFormula");
  }

  safeSetCellValue(cellKey: string, value: unknown): boolean {
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${cellKey}`);
    if (!this.canEditCell(parsed)) return false;

    this.setCellValue(makeCellKey(parsed), value);
    return true;
  }

  safeSetCellFormula(cellKey: string, formula: string | null): boolean {
    const parsed = parseCellKey(cellKey, { defaultSheetId: this.defaultSheetId });
    if (!parsed) throw new Error(`Invalid cellKey: ${cellKey}`);
    if (!this.canEditCell(parsed)) return false;

    this.setCellFormula(makeCellKey(parsed), formula);
    return true;
  }
}

export function createCollabSession(options: CollabSessionOptions = {}): CollabSession {
  return new CollabSession(options);
}

// Backwards-compatible alias (Task 133 naming).
export const createSession = createCollabSession;
