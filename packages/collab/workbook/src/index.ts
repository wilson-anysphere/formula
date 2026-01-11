import * as Y from "yjs";

export interface WorkbookSchemaOptions {
  defaultSheetName?: string;
  defaultSheetId?: string;
}

export type WorkbookSchemaRoots = {
  cells: Y.Map<unknown>;
  sheets: Y.Array<Y.Map<unknown>>;
  metadata: Y.Map<unknown>;
  namedRanges: Y.Map<unknown>;
};

export function ensureWorkbookSchema(doc: Y.Doc, options: WorkbookSchemaOptions = {}): WorkbookSchemaRoots {
  const cells = doc.getMap<unknown>("cells");
  const sheets = doc.getArray<Y.Map<unknown>>("sheets");
  const metadata = doc.getMap<unknown>("metadata");
  const namedRanges = doc.getMap<unknown>("namedRanges");

  const defaultSheetId = options.defaultSheetId ?? "Sheet1";
  const defaultSheetName = options.defaultSheetName ?? defaultSheetId;

  // `sheets` is a Y.Array of sheet metadata maps (with at least `{ id, name }`).
  // In practice we may see duplicate sheet ids when two clients concurrently
  // initialize an empty workbook. Treat ids as unique and prune duplicates so
  // downstream sheet lookups remain deterministic.
  const shouldNormalize = (() => {
    const seen = new Set<string>();
    for (const entry of sheets.toArray()) {
      const maybe = entry as any;
      const id = coerceString(maybe?.get?.("id") ?? maybe?.id);
      if (!id) return true;
      if (seen.has(id)) return true;
      seen.add(id);
    }
    return sheets.length === 0;
  })();

  if (shouldNormalize) {
    doc.transact(() => {
      const seen = new Set<string>();
      const deleteIndices: number[] = [];

      for (let i = 0; i < sheets.length; i++) {
        const entry = sheets.get(i) as any;
        const id = coerceString(entry?.get?.("id") ?? entry?.id);
        if (!id) {
          deleteIndices.push(i);
          continue;
        }
        if (seen.has(id)) {
          deleteIndices.push(i);
          continue;
        }
        seen.add(id);
      }

      for (let i = deleteIndices.length - 1; i >= 0; i--) {
        sheets.delete(deleteIndices[i], 1);
      }

      if (sheets.length === 0) {
        const sheet = new Y.Map<unknown>();
        sheet.set("id", defaultSheetId);
        sheet.set("name", defaultSheetName);
        sheets.push([sheet]);
      }
    });
  }

  return { cells, sheets, metadata, namedRanges };
}

export type WorkbookTransact = (fn: () => void) => void;

function defaultTransact(doc: Y.Doc): WorkbookTransact {
  return (fn) => {
    doc.transact(fn);
  };
}

function coerceString(value: unknown): string | null {
  if (value instanceof Y.Text) return value.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

export class SheetManager {
  readonly sheets: Y.Array<Y.Map<unknown>>;
  private readonly transact: WorkbookTransact;

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    this.sheets = opts.doc.getArray<Y.Map<unknown>>("sheets");
    this.transact = opts.transact ?? defaultTransact(opts.doc);
  }

  list(): Array<{ id: string; name: string | null }> {
    const out: Array<{ id: string; name: string | null }> = [];
    for (const entry of this.sheets.toArray()) {
      const id = coerceString(entry?.get("id"));
      if (!id) continue;
      const name = coerceString(entry.get("name"));
      out.push({ id, name });
    }
    return out;
  }

  getById(id: string): Y.Map<unknown> | null {
    const index = this.indexOf(id);
    if (index < 0) return null;
    return this.sheets.get(index) ?? null;
  }

  addSheet(input: { id: string; name?: string | null; index?: number }): void {
    const id = input.id;
    const name = input.name ?? id;

    this.transact(() => {
      if (this.indexOf(id) >= 0) {
        throw new Error(`Sheet already exists: ${id}`);
      }

      const sheet = new Y.Map<unknown>();
      sheet.set("id", id);
      sheet.set("name", name);

      const idx =
        typeof input.index === "number" && Number.isFinite(input.index)
          ? Math.max(0, Math.min(Math.floor(input.index), this.sheets.length))
          : this.sheets.length;

      this.sheets.insert(idx, [sheet]);
    });
  }

  renameSheet(id: string, name: string): void {
    this.transact(() => {
      const sheet = this.getById(id);
      if (!sheet) throw new Error(`Sheet not found: ${id}`);
      sheet.set("name", name);
    });
  }

  removeSheet(id: string): void {
    this.transact(() => {
      const idx = this.indexOf(id);
      if (idx < 0) throw new Error(`Sheet not found: ${id}`);
      // Workbooks must always have at least one sheet. Match common spreadsheet
      // semantics by preventing callers from deleting the last remaining sheet.
      if (this.countSheetEntriesWithIds() <= 1) {
        throw new Error("Cannot delete the last remaining sheet");
      }
      this.sheets.delete(idx, 1);
    });
  }

  moveSheet(id: string, toIndex: number): void {
    this.transact(() => {
      const fromIndex = this.indexOf(id);
      if (fromIndex < 0) throw new Error(`Sheet not found: ${id}`);

      const maxIndex = Math.max(0, this.sheets.length - 1);
      const targetIndex = Math.max(0, Math.min(Math.floor(toIndex), maxIndex));
      if (fromIndex === targetIndex) return;

      const sheet = this.sheets.get(fromIndex);
      if (!sheet) throw new Error(`Sheet missing at index ${fromIndex}: ${id}`);

      this.sheets.delete(fromIndex, 1);
      this.sheets.insert(targetIndex, [sheet]);
    });
  }

  private indexOf(id: string): number {
    const entries = this.sheets.toArray();
    for (let i = 0; i < entries.length; i++) {
      const entryId = coerceString(entries[i]?.get("id"));
      if (entryId === id) return i;
    }
    return -1;
  }

  private countSheetEntriesWithIds(): number {
    let count = 0;
    for (const entry of this.sheets.toArray()) {
      const id = coerceString(entry?.get("id"));
      if (id) count += 1;
    }
    return count;
  }
}

export class NamedRangeManager {
  readonly namedRanges: Y.Map<unknown>;
  private readonly transact: WorkbookTransact;

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    this.namedRanges = opts.doc.getMap<unknown>("namedRanges");
    this.transact = opts.transact ?? defaultTransact(opts.doc);
  }

  get(name: string): unknown {
    return this.namedRanges.get(name);
  }

  set(name: string, value: unknown): void {
    this.transact(() => {
      this.namedRanges.set(name, value);
    });
  }

  delete(name: string): void {
    this.transact(() => {
      this.namedRanges.delete(name);
    });
  }
}

export class MetadataManager {
  readonly metadata: Y.Map<unknown>;
  private readonly transact: WorkbookTransact;

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    this.metadata = opts.doc.getMap<unknown>("metadata");
    this.transact = opts.transact ?? defaultTransact(opts.doc);
  }

  get(key: string): unknown {
    return this.metadata.get(key);
  }

  set(key: string, value: unknown): void {
    this.transact(() => {
      this.metadata.set(key, value);
    });
  }

  delete(key: string): void {
    this.transact(() => {
      this.metadata.delete(key);
    });
  }
}

export function createSheetManagerForSession(session: {
  doc: Y.Doc;
  transactLocal: (fn: () => void) => void;
}): SheetManager {
  return new SheetManager({ doc: session.doc, transact: (fn) => session.transactLocal(fn) });
}

export function createNamedRangeManagerForSession(session: {
  doc: Y.Doc;
  transactLocal: (fn: () => void) => void;
}): NamedRangeManager {
  return new NamedRangeManager({ doc: session.doc, transact: (fn) => session.transactLocal(fn) });
}

export function createMetadataManagerForSession(session: {
  doc: Y.Doc;
  transactLocal: (fn: () => void) => void;
}): MetadataManager {
  return new MetadataManager({ doc: session.doc, transact: (fn) => session.transactLocal(fn) });
}
