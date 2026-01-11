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

  if (sheets.length === 0) {
    doc.transact(() => {
      if (sheets.length !== 0) return;
      const sheet = new Y.Map<unknown>();
      sheet.set("id", defaultSheetId);
      sheet.set("name", defaultSheetName);
      sheets.push([sheet]);
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
