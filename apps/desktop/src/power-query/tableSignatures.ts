import { hashValue } from "../../../../packages/power-query/src/cache/key.js";

import type { DocumentController } from "../document/documentController.js";
import type { TableInfo } from "../tauri/workbookBackend";

export type TableRectangle = {
  sheetId: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

type TableRegistryEntry = {
  name: string;
  rectangle: TableRectangle;
  definitionHash: string;
  version: number;
};

function computeDefinitionHash(rectangle: TableRectangle, columns: string[]): string {
  // This hash should be stable across app sessions for unchanged table definitions so
  // cached query results survive reloads while still invalidating when tables resize
  // or their header schema changes.
  return hashValue({
    sheetId: rectangle.sheetId,
    startRow: rectangle.startRow,
    startCol: rectangle.startCol,
    endRow: rectangle.endRow,
    endCol: rectangle.endCol,
    columns,
  });
}

function containsCell(rectangle: TableRectangle, cell: { row: number; col: number }): boolean {
  return (
    cell.row >= rectangle.startRow &&
    cell.row <= rectangle.endRow &&
    cell.col >= rectangle.startCol &&
    cell.col <= rectangle.endCol
  );
}

/**
 * In-memory registry of workbook table definitions + versions.
 *
 * `version` is a monotonically-increasing counter that bumps whenever a cell edit
 * lands inside the table's rectangle. The `definitionHash` changes when the table
 * definition changes (resize / header rename), and definition changes also bump
 * `version`.
 *
 * QueryEngine uses the combined `${definitionHash}:${version}` signature to safely
 * cache table-source queries.
 */
export class TableSignatureRegistry {
  #tablesByName = new Map<string, TableRegistryEntry>();
  #tablesBySheetId = new Map<string, TableRegistryEntry[]>();
  #unsubscribe: (() => void) | null = null;

  constructor(doc: DocumentController) {
    if (doc && typeof (doc as any).on === "function") {
      this.#unsubscribe = (doc as any).on("change", (payload: any) => {
        const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
        if (deltas.length === 0) return;
        this.applyCellDeltas(deltas, { source: payload?.source });
      });
    }
  }

  dispose(): void {
    this.#unsubscribe?.();
    this.#unsubscribe = null;
  }

  /**
   * Refresh the registry from backend `list_tables` results.
   */
  refreshFromTables(tables: TableInfo[]): void {
    const next = new Map<string, TableRegistryEntry>();

    for (const table of tables) {
      const name = typeof (table as any)?.name === "string" ? String((table as any).name) : "";
      if (!name) continue;

      const sheetId = typeof (table as any)?.sheet_id === "string" ? String((table as any).sheet_id) : "";
      if (!sheetId) continue;

      const startRow = Number((table as any).start_row);
      const startCol = Number((table as any).start_col);
      const endRow = Number((table as any).end_row);
      const endCol = Number((table as any).end_col);
      if (![startRow, startCol, endRow, endCol].every((n) => Number.isFinite(n))) continue;

      const rectangle: TableRectangle = {
        sheetId,
        startRow,
        startCol,
        endRow,
        endCol,
      };
      const columns = Array.isArray((table as any)?.columns) ? (table as any).columns.map(String) : [];
      const definitionHash = computeDefinitionHash(rectangle, columns);

      const existing = this.#tablesByName.get(name);
      if (existing) {
        const changed = existing.definitionHash !== definitionHash;
        const version = changed ? existing.version + 1 : existing.version;
        next.set(name, { ...existing, rectangle, definitionHash, version });
      } else {
        next.set(name, { name, rectangle, definitionHash, version: 0 });
      }
    }

    this.#tablesByName = next;
    this.rebuildSheetIndex();
  }

  getTableSignature(tableName: string): string | undefined {
    const entry = this.#tablesByName.get(tableName);
    if (!entry) return undefined;
    return `${entry.definitionHash}:${entry.version}`;
  }

  /**
   * Apply document cell deltas and bump versions for any tables touched.
   *
   * This bumps each table at most once per change event, even if many cells in
   * the table changed.
   */
  applyCellDeltas(deltas: Array<{ sheetId: string; row: number; col: number }>, options?: { source?: string }): void {
    // `applyState` can generate huge delta lists; avoid double-scanning by tracking
    // which tables we've already bumped for this batch.
    const bumped = new Set<string>();

    for (const delta of deltas) {
      const sheetId = typeof (delta as any)?.sheetId === "string" ? String((delta as any).sheetId) : "";
      if (!sheetId) continue;
      const row = Number((delta as any).row);
      const col = Number((delta as any).col);
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;

      const candidates = this.#tablesBySheetId.get(sheetId);
      if (!candidates || candidates.length === 0) continue;

      for (const entry of candidates) {
        if (bumped.has(entry.name)) continue;
        if (!containsCell(entry.rectangle, { row, col })) continue;
        bumped.add(entry.name);
        const current = this.#tablesByName.get(entry.name);
        if (!current) continue;
        this.#tablesByName.set(entry.name, { ...current, version: current.version + 1 });
      }
    }

    if (bumped.size > 0) {
      // Update sheet index references for bumped entries (version is stored in the
      // canonical `#tablesByName` map; sheet index entries hold stale references).
      //
      // Rebuilds are cheap (tables per workbook are small) and keep lookup fast.
      this.rebuildSheetIndex();
    }
  }

  private rebuildSheetIndex(): void {
    const bySheet = new Map<string, TableRegistryEntry[]>();
    for (const entry of this.#tablesByName.values()) {
      const list = bySheet.get(entry.rectangle.sheetId);
      if (list) list.push(entry);
      else bySheet.set(entry.rectangle.sheetId, [entry]);
    }
    this.#tablesBySheetId = bySheet;
  }
}

const REGISTRY_BY_DOCUMENT = new WeakMap<object, TableSignatureRegistry>();

export function getTableSignatureRegistry(doc: DocumentController): TableSignatureRegistry {
  const key = doc as unknown as object;
  const existing = REGISTRY_BY_DOCUMENT.get(key);
  if (existing) return existing;
  const created = new TableSignatureRegistry(doc);
  REGISTRY_BY_DOCUMENT.set(key, created);
  return created;
}

export function refreshTableSignaturesFromBackend(doc: DocumentController, tables: TableInfo[]): void {
  getTableSignatureRegistry(doc).refreshFromTables(tables);
}

