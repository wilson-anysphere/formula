import * as Y from "yjs";
import { getSheetNameValidationErrorMessage } from "@formula/workbook-backend";
import { getArrayRoot, getMapRoot, getYArray, getYMap, getYText } from "@formula/collab-yjs-utils";

export interface WorkbookSchemaOptions {
  defaultSheetName?: string;
  defaultSheetId?: string;
  /**
   * Whether to create a default sheet when the workbook has no sheets.
   * Defaults to true.
   */
  createDefaultSheet?: boolean;
}

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export type WorkbookSchemaRoots = {
  cells: Y.Map<unknown>;
  sheets: Y.Array<Y.Map<unknown>>;
  metadata: Y.Map<unknown>;
  namedRanges: Y.Map<unknown>;
};

export function getWorkbookRoots(doc: Y.Doc): WorkbookSchemaRoots {
  return {
    cells: getMapRoot<unknown>(doc, "cells"),
    sheets: getArrayRoot<Y.Map<unknown>>(doc, "sheets") as Y.Array<Y.Map<unknown>>,
    metadata: getMapRoot<unknown>(doc, "metadata"),
    namedRanges: getMapRoot<unknown>(doc, "namedRanges"),
  };
}

export function ensureWorkbookSchema(doc: Y.Doc, options: WorkbookSchemaOptions = {}): WorkbookSchemaRoots {
  const { cells, sheets, metadata, namedRanges } = getWorkbookRoots(doc);
  const YMapCtor = cells.constructor as unknown as { new (): Y.Map<unknown> };

  const defaultSheetId = options.defaultSheetId ?? "Sheet1";
  const defaultSheetName = options.defaultSheetName ?? defaultSheetId;
  const createDefaultSheet = options.createDefaultSheet ?? true;

  // `sheets` is a Y.Array of sheet metadata maps (with at least `{ id, name }`).
  // Sheet metadata may also include:
  //   - visibility: "visible" | "hidden" | "veryHidden"
  //   - tabColor: ARGB hex string (e.g. "FFFF0000")
  // In practice we may see duplicate sheet ids when two clients concurrently
  // initialize an empty workbook. Treat ids as unique and prune duplicates so
  // downstream sheet lookups remain deterministic.
  const shouldNormalize = (() => {
    if (sheets.length === 0) return createDefaultSheet;
    const seen = new Set<string>();
    let hasSheetWithId = false;
    let hasVisibleSheet = false;
    for (const entry of sheets.toArray()) {
      const maybe = entry as any;
      const id = coerceString(maybe?.get?.("id") ?? maybe?.id);
      if (!id) continue;
      hasSheetWithId = true;
      if (seen.has(id)) return true;
      seen.add(id);

      const visibilityRaw = maybe?.get?.("visibility") ?? maybe?.visibility;
      const visibilityStr = coerceString(visibilityRaw);
      const visibility = coerceSheetVisibility(visibilityStr) ?? "visible";
      if (visibilityStr !== visibility || typeof visibilityRaw !== "string") return true;
      if (visibility === "visible") hasVisibleSheet = true;

      const tabColorRaw = maybe?.get?.("tabColor") ?? maybe?.tabColor;
      if (tabColorRaw != null) {
        const normalized = coerceTabColor(tabColorRaw);
        if (!normalized) return true;
        const tabColorStr = coerceString(tabColorRaw);
        if (tabColorStr !== normalized || typeof tabColorRaw !== "string") return true;
      }
    }

    // Workbooks should always have at least one visible sheet.
    if (hasSheetWithId && !hasVisibleSheet) return true;

    return createDefaultSheet && !hasSheetWithId;
  })();

  if (shouldNormalize) {
    doc.transact(() => {
      /** @type {Map<string, number[]>} */
      const indicesById = new Map<string, number[]>();
      let hasSheetWithId = false;

      for (let i = 0; i < sheets.length; i += 1) {
        const entry = sheets.get(i) as any;
        const id = coerceString(entry?.get?.("id") ?? entry?.id);
        if (!id) continue;
        hasSheetWithId = true;
        const existing = indicesById.get(id);
        if (existing) existing.push(i);
        else indicesById.set(id, [i]);
      }

      const deleteIndices: number[] = [];
      for (const indices of indicesById.values()) {
        if (indices.length <= 1) continue;

        // Deterministic pruning: keep exactly one surviving entry by index.
        //
        // Importantly, the choice must be stable across clients. Using `doc.clientID`
        // to prefer "non-local" entries can lead to divergence when two clients
        // concurrently initialize a brand new workbook (each sees their own Sheet1
        // as local and the other as non-local, so both delete their own and the
        // workbook ends up with *no* Sheet1).
        //
        // Keeping the last entry by index is deterministic in Yjs and tends to
        // preserve the sheet order/metadata that arrives later (e.g. from merges,
        // restores, or persistence hydration).
        const remaining = indices.slice().sort((a, b) => a - b);
        const winnerIndex = remaining[remaining.length - 1]!;

        // Before deleting duplicates, opportunistically merge any non-default
        // metadata from the losing entries into the winner. This helps avoid
        // losing canonical sheet metadata if a placeholder entry happened to win
        // the index tie-breaker.
        const winner = sheets.get(winnerIndex) as any;
        if (winner && typeof winner.get === "function" && typeof winner.set === "function") {
          for (const index of remaining) {
            if (index === winnerIndex) continue;
            const entry = sheets.get(index) as any;
            if (!entry || typeof entry.get !== "function") continue;

            const id = coerceString(winner.get("id"));
            if (!id) continue;

            const winnerName = coerceString(winner.get("name"));
            const entryName = coerceString(entry.get("name"));
            // Prefer a non-default display name when the winner has a blank/default name.
            if ((!winnerName || winnerName === id) && entryName && entryName !== id) {
              winner.set("name", entryName);
            }

            const winnerVis = coerceSheetVisibility(coerceString(winner.get("visibility"))) ?? "visible";
            const entryVis = coerceSheetVisibility(coerceString(entry.get("visibility")));
            // Prefer explicit non-visible visibility over default "visible" when deduping.
            if (winnerVis === "visible" && entryVis && entryVis !== "visible") {
              winner.set("visibility", entryVis);
            }

            const winnerTab = coerceTabColor(winner.get("tabColor"));
            const entryTab = coerceTabColor(entry.get("tabColor"));
            if (!winnerTab && entryTab) {
              winner.set("tabColor", entryTab);
            }
          }
        }

        for (let i = 0; i < remaining.length - 1; i += 1) {
          deleteIndices.push(remaining[i]!);
        }
      }

      deleteIndices.sort((a, b) => b - a);
      for (const index of deleteIndices) {
        sheets.delete(index, 1);
      }

      if (createDefaultSheet && sheets.length === 0) {
        const sheet = new YMapCtor();
        sheet.set("id", defaultSheetId);
        sheet.set("name", defaultSheetName);
        sheet.set("visibility", "visible");
        sheets.push([sheet]);
        hasSheetWithId = true;
      }

      // If the workbook has sheets but none are valid (no `id` field), salvage
      // the first entry by assigning it the default sheet id/name. This keeps
      // the sheet list stable even if a client created the first sheet in
      // multiple transactions (e.g. insert map, then set id later).
      if (createDefaultSheet && !hasSheetWithId && sheets.length > 0) {
        const first: any = sheets.get(0);
        if (first && typeof first.get === "function" && typeof first.set === "function") {
          const existingId = coerceString(first.get("id"));
          if (!existingId) first.set("id", defaultSheetId);
          const existingName = coerceString(first.get("name"));
          if (!existingName) first.set("name", defaultSheetName);
          hasSheetWithId = true;
        } else {
          const sheet = new YMapCtor();
          sheet.set("id", defaultSheetId);
          sheet.set("name", defaultSheetName);
          sheet.set("visibility", "visible");
          sheets.insert(0, [sheet]);
          hasSheetWithId = true;
        }
      }

      // Normalize per-sheet metadata now that ids are stable.
      let hasVisibleSheet = false;
      const entries = sheets.toArray() as any[];
      for (const entry of entries) {
        const id = coerceString(entry?.get?.("id") ?? entry?.id);
        if (!id) continue;

        if (!entry || typeof entry.get !== "function" || typeof entry.set !== "function") continue;

        const currentVisibility = entry.get("visibility");
        const visibilityStr = coerceString(currentVisibility);
        const visibility = coerceSheetVisibility(visibilityStr) ?? "visible";
        if (visibilityStr !== visibility || typeof currentVisibility !== "string") {
          entry.set("visibility", visibility);
        }
        if (visibility === "visible") hasVisibleSheet = true;

        const currentTabColor = entry.get("tabColor");
        if (currentTabColor != null) {
          const normalized = coerceTabColor(currentTabColor);
          if (!normalized) {
            if (typeof entry.delete === "function") entry.delete("tabColor");
          } else {
            const tabColorStr = coerceString(currentTabColor);
            if (tabColorStr !== normalized || typeof currentTabColor !== "string") {
              entry.set("tabColor", normalized);
            }
          }
        }
      }

      // Ensure the workbook always has at least one visible sheet.
      if (hasSheetWithId && !hasVisibleSheet) {
        for (const entry of entries) {
          const id = coerceString(entry?.get?.("id") ?? entry?.id);
          if (!id) continue;
          if (!entry || typeof entry.set !== "function") continue;
          entry.set("visibility", "visible");
          break;
        }
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
  const text = getYText(value);
  if (text) return text.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

function coerceSheetVisibility(value: unknown): SheetVisibility | null {
  if (value === "visible" || value === "hidden" || value === "veryHidden") return value;
  return null;
}

function coerceTabColor(value: unknown): string | null {
  const str = coerceString(value);
  if (!str) return null;
  if (!/^[0-9A-Fa-f]{8}$/.test(str)) return null;
  // Canonicalize to uppercase so equality checks are stable across clients.
  return str.toUpperCase();
}

type DocTypeConstructors = {
  MapCtor: new () => any;
  ArrayCtor: new () => any;
  TextCtor: new () => any;
};

function cloneYjsValueWithCtors(value: any, ctors: DocTypeConstructors): any {
  const map = getYMap(value);
  if (map) {
    const out = new ctors.MapCtor();
    map.forEach((v: any, k: string) => {
      out.set(k, cloneYjsValueWithCtors(v, ctors));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new ctors.ArrayCtor();
    for (const item of array.toArray()) {
      out.push([cloneYjsValueWithCtors(item, ctors)]);
    }
    return out;
  }

  const text = getYText(value);
  if (text) {
    const out = new ctors.TextCtor();
    out.applyDelta(structuredClone(text.toDelta()));
    return out;
  }

  if (value && typeof value === "object") {
    return structuredClone(value);
  }
  return value;
}

function findAvailableRootName(doc: Y.Doc, base: string): string {
  if (!doc.share.has(base)) return base;
  for (let i = 1; i < 1000; i += 1) {
    const name = `${base}_${i}`;
    if (!doc.share.has(name)) return name;
  }
  return `${base}_${Date.now()}`;
}

function getDocTextConstructor(doc: any): new () => any {
  const name = findAvailableRootName(doc, "__workbook_tmp_text");
  const tmp = doc.getText(name);
  const ctor = tmp.constructor as new () => any;
  doc.share.delete(name);
  return ctor;
}

export class SheetManager {
  readonly sheets: Y.Array<Y.Map<unknown>>;
  private readonly transact: WorkbookTransact;
  private readonly YMapCtor: { new (): Y.Map<unknown> };
  private readonly YArrayCtor: { new (): Y.Array<any> };
  private readonly YTextCtor: { new (): Y.Text };

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    const cells = getMapRoot<unknown>(opts.doc, "cells");
    this.sheets = getArrayRoot<Y.Map<unknown>>(opts.doc, "sheets") as Y.Array<Y.Map<unknown>>;
    this.transact = opts.transact ?? defaultTransact(opts.doc);
    this.YMapCtor = cells.constructor as unknown as { new (): Y.Map<unknown> };
    this.YArrayCtor = this.sheets.constructor as unknown as { new (): Y.Array<any> };
    this.YTextCtor = getDocTextConstructor(opts.doc) as unknown as { new (): Y.Text };
  }

  list(): Array<{ id: string; name: string | null; visibility?: SheetVisibility; tabColor?: string | null }> {
    const out: Array<{ id: string; name: string | null; visibility?: SheetVisibility; tabColor?: string | null }> = [];
    for (const entry of this.sheets.toArray()) {
      const id = coerceString(entry?.get("id"));
      if (!id) continue;
      const name = coerceString(entry.get("name"));

      const visibilityRaw = entry.get("visibility");
      const visibility = coerceSheetVisibility(coerceString(visibilityRaw)) ?? "visible";

      const tabColorRaw = entry.get("tabColor");
      const tabColor = tabColorRaw == null ? null : coerceTabColor(tabColorRaw);

      out.push({ id, name, visibility, tabColor });
    }
    return out;
  }

  listVisible(): Array<{ id: string; name: string | null; visibility: SheetVisibility; tabColor: string | null }> {
    const out: Array<{ id: string; name: string | null; visibility: SheetVisibility; tabColor: string | null }> = [];
    for (const entry of this.sheets.toArray()) {
      const id = coerceString(entry?.get("id"));
      if (!id) continue;
      const name = coerceString(entry.get("name"));

      const visibilityRaw = entry.get("visibility");
      const visibility = coerceSheetVisibility(coerceString(visibilityRaw)) ?? "visible";
      if (visibility !== "visible") continue;

      const tabColorRaw = entry.get("tabColor");
      const tabColor = tabColorRaw == null ? null : coerceTabColor(tabColorRaw);

      out.push({ id, name, visibility, tabColor });
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

      const existingNames: string[] = [];
      for (const sheet of this.list()) {
        // Prefer the display name but fall back to id so uniqueness checks match how other
        // parts of the stack treat missing/legacy names.
        const existing = sheet.name ?? sheet.id;
        if (existing) existingNames.push(existing);
      }
      const nameError = getSheetNameValidationErrorMessage(name, { existingNames });
      if (nameError) throw new Error(nameError);

      const sheet = new this.YMapCtor();
      sheet.set("id", id);
      sheet.set("name", name);
      sheet.set("visibility", "visible");

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

      const existingNames: string[] = [];
      for (const entry of this.list()) {
        if (entry.id === id) continue;
        const existing = entry.name ?? entry.id;
        if (existing) existingNames.push(existing);
      }
      const nameError = getSheetNameValidationErrorMessage(name, { existingNames });
      if (nameError) throw new Error(nameError);

      sheet.set("name", name);
    });
  }

  setVisibility(id: string, visibility: SheetVisibility): void {
    this.transact(() => {
      const sheet = this.getById(id);
      if (!sheet) throw new Error(`Sheet not found: ${id}`);

      const next = visibility;
      if (next !== "visible" && next !== "hidden" && next !== "veryHidden") {
        throw new Error(`Invalid sheet visibility: ${String(next)}`);
      }

      const current = coerceSheetVisibility(coerceString(sheet.get("visibility"))) ?? "visible";
      if (current === next) return;

      // Match common spreadsheet semantics by preventing callers from hiding the
      // last visible sheet.
      if (current === "visible" && next !== "visible") {
        const visibleCount = this.countVisibleSheets();
        if (visibleCount <= 1) {
          throw new Error("Cannot hide the last visible sheet");
        }
      }

      sheet.set("visibility", next);
    });
  }

  hideSheet(id: string): void {
    this.setVisibility(id, "hidden");
  }

  unhideSheet(id: string): void {
    this.setVisibility(id, "visible");
  }

  setTabColor(id: string, tabColor: string | null): void {
    this.transact(() => {
      const sheet = this.getById(id);
      if (!sheet) throw new Error(`Sheet not found: ${id}`);

      if (tabColor == null) {
        sheet.delete("tabColor");
        return;
      }

      const normalized = coerceTabColor(tabColor);
      if (!normalized) {
        throw new Error(`Invalid tabColor (expected 8-digit ARGB hex): ${tabColor}`);
      }
      sheet.set("tabColor", normalized);
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

      const sheetClone = cloneYjsValueWithCtors(sheet, {
        MapCtor: this.YMapCtor,
        ArrayCtor: this.YArrayCtor,
        TextCtor: this.YTextCtor,
      });
      this.sheets.delete(fromIndex, 1);
      this.sheets.insert(targetIndex, [sheetClone]);
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

  private countVisibleSheets(): number {
    let count = 0;
    for (const entry of this.sheets.toArray()) {
      const id = coerceString(entry?.get("id"));
      if (!id) continue;
      const visibility = coerceSheetVisibility(coerceString(entry?.get("visibility"))) ?? "visible";
      if (visibility === "visible") count += 1;
    }
    return count;
  }
}

export class NamedRangeManager {
  readonly namedRanges: Y.Map<unknown>;
  private readonly transact: WorkbookTransact;

  constructor(opts: { doc: Y.Doc; transact?: WorkbookTransact }) {
    this.namedRanges = getMapRoot<unknown>(opts.doc, "namedRanges");
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
    this.metadata = getMapRoot<unknown>(opts.doc, "metadata");
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

export type PermissionAwareWorkbookSession = {
  doc: Y.Doc;
  transactLocal: (fn: () => void) => void;
  isReadOnly: () => boolean;
};

function assertSessionCanMutateWorkbook(session: { isReadOnly: () => boolean }): void {
  if (!session.isReadOnly()) return;
  throw new Error("Permission denied: cannot mutate workbook in a read-only session");
}

function transactLocalWithWorkbookPermissions(session: PermissionAwareWorkbookSession): WorkbookTransact {
  return (fn) => {
    assertSessionCanMutateWorkbook(session);
    session.transactLocal(fn);
  };
}

function createPermissionAwareYMapProxy<T>(params: {
  map: Y.Map<T>;
  session: PermissionAwareWorkbookSession;
  /**
   * Optional wrapper for values returned by read APIs like `get()` / `values()` /
   * `entries()`.
   *
   * This is used to return permission-guarded nested Yjs types (e.g. `metadata`
   * values like `encryptedRanges` which are often stored as nested Y.Arrays/Y.Maps).
   */
  wrapValue?: (value: unknown) => unknown;
  /**
   * Optional cache so we can preserve referential equality when returning nested
   * Y.Maps (e.g. sheet entries inside the `sheets` array).
   */
  cache?: WeakMap<object, Y.Map<T>>;
  /**
   * Tracks the proxies created by this helper so we can avoid proxy-wrapping a
   * proxy (which would otherwise break referential equality and add overhead).
   */
  proxySet?: WeakSet<object>;
}): Y.Map<T> {
  const { map, session, cache, proxySet, wrapValue } = params;
  if (proxySet?.has(map as any)) return map;
  const cached = cache?.get(map as any);
  if (cached) return cached;

  const proxy = new Proxy(map as any, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, target);
      if (typeof value !== "function") return value;
      if (prop === "constructor") return value;

      // Guard the primary mutators so callers can't accidentally bypass the
      // manager APIs by writing through the exposed Y.Map.
      if (prop === "set" || prop === "delete" || prop === "clear") {
        return (...args: any[]) => {
          assertSessionCanMutateWorkbook(session);
          return Reflect.apply(value, target, args);
        };
      }

      if (wrapValue && prop === "get") {
        return (...args: any[]) => {
          const out = Reflect.apply(value, target, args);
          return wrapValue(out) as any;
        };
      }

      if (wrapValue && prop === "forEach") {
        return (cb: (...args: any[]) => void, thisArg?: any) => {
          return Reflect.apply(value, target, [
            (v: any, k: any) => cb.call(thisArg, wrapValue(v), k, receiver),
            thisArg,
          ]);
        };
      }

      if (wrapValue && (prop === "values" || prop === "entries" || prop === Symbol.iterator)) {
        return (...args: any[]) => {
          const iter = Reflect.apply(value, target, args) as IterableIterator<any>;
          const wrapEntry = (entry: any) => {
            // values() yields the value, entries()/iterator yield [key, value]
            if (prop === "values") return wrapValue(entry);
            if (Array.isArray(entry) && entry.length >= 2) return [entry[0], wrapValue(entry[1])];
            return entry;
          };

          return {
            [Symbol.iterator]() {
              return this;
            },
            next() {
              const { value: v, done } = iter.next();
              return done ? { value: v, done } : { value: wrapEntry(v), done };
            },
          } as IterableIterator<any>;
        };
      }

      return value.bind(target);
    },
  }) as Y.Map<T>;

  cache?.set(map as any, proxy);
  proxySet?.add(proxy as any);
  return proxy;
}

function createPermissionAwareYArrayProxy<T>(params: {
  array: Y.Array<T>;
  session: PermissionAwareWorkbookSession;
  /**
   * Optional cache so we can preserve referential equality when returning nested
   * Y.Arrays.
   */
  cache?: WeakMap<object, Y.Array<T>>;
  /**
   * Tracks the proxies created by this helper so we can avoid proxy-wrapping a
   * proxy (which would otherwise break referential equality and add overhead).
   */
  proxySet?: WeakSet<object>;
  /**
   * Optional wrapper for values returned by read APIs like `get()` / `toArray()`.
   * Useful for returning permission-guarded nested Y.Maps (e.g. sheet entries).
   */
  wrapValue?: (value: unknown) => unknown;
}): Y.Array<T> {
  const { array, session, wrapValue, cache, proxySet } = params;
  if (proxySet?.has(array as any)) return array;
  const cached = cache?.get(array as any);
  if (cached) return cached;

  const proxy = new Proxy(array as any, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, target);
      if (typeof value !== "function") return value;
      if (prop === "constructor") return value;

      // Guard the primary mutators so callers can't accidentally bypass the
      // manager APIs by writing through the exposed Y.Array.
      if (prop === "insert" || prop === "push" || prop === "unshift" || prop === "delete") {
        return (...args: any[]) => {
          assertSessionCanMutateWorkbook(session);
          return Reflect.apply(value, target, args);
        };
      }

      if (wrapValue && prop === "get") {
        return (...args: any[]) => {
          const out = Reflect.apply(value, target, args);
          return wrapValue(out) as any;
        };
      }

      if (wrapValue && prop === "toArray") {
        return (...args: any[]) => {
          const out = Reflect.apply(value, target, args);
          return Array.isArray(out) ? out.map((v) => wrapValue(v) as any) : out;
        };
      }

      if (wrapValue && prop === "slice") {
        return (...args: any[]) => {
          const out = Reflect.apply(value, target, args);
          return Array.isArray(out) ? out.map((v) => wrapValue(v) as any) : out;
        };
      }

      if (wrapValue && prop === "map") {
        return (cb: (...args: any[]) => any) => {
          return Reflect.apply(value, target, [
            (v: any, idx: number) => cb(wrapValue(v), idx, receiver),
          ]);
        };
      }

      if (wrapValue && prop === "forEach") {
        return (cb: (...args: any[]) => void, thisArg?: any) => {
          return Reflect.apply(value, target, [
            (v: any, idx: number) => cb.call(thisArg, wrapValue(v), idx, receiver),
            thisArg,
          ]);
        };
      }

      if (wrapValue && prop === Symbol.iterator) {
        return (...args: any[]) => {
          const iter = Reflect.apply(value, target, args) as IterableIterator<any>;
          return {
            [Symbol.iterator]() {
              return this;
            },
            next() {
              const { value: v, done } = iter.next();
              return done ? { value: v, done } : { value: wrapValue(v), done };
            },
          } as IterableIterator<any>;
        };
      }

      return value.bind(target);
    },
  }) as Y.Array<T>;

  cache?.set(array as any, proxy);
  proxySet?.add(proxy as any);
  return proxy;
}

function createPermissionAwareYTextProxy(params: {
  text: Y.Text;
  session: PermissionAwareWorkbookSession;
  cache?: WeakMap<object, Y.Text>;
  proxySet?: WeakSet<object>;
}): Y.Text {
  const { text, session, cache, proxySet } = params;
  if (proxySet?.has(text as any)) return text;
  const cached = cache?.get(text as any);
  if (cached) return cached;

  const proxy = new Proxy(text as any, {
    get(target, prop) {
      const value = Reflect.get(target, prop, target);
      if (typeof value !== "function") return value;
      if (prop === "constructor") return value;

      if (prop === "insert" || prop === "delete" || prop === "applyDelta" || prop === "format") {
        return (...args: any[]) => {
          assertSessionCanMutateWorkbook(session);
          return Reflect.apply(value, target, args);
        };
      }

      return value.bind(target);
    },
  }) as Y.Text;

  cache?.set(text as any, proxy);
  proxySet?.add(proxy as any);
  return proxy;
}

export function createSheetManagerForSessionWithPermissions(session: PermissionAwareWorkbookSession): SheetManager {
  const mgr = new SheetManager({ doc: session.doc, transact: transactLocalWithWorkbookPermissions(session) });

  // Also guard the exposed Yjs roots so callers can't bypass the permission-aware
  // manager methods by mutating `mgr.sheets` directly.
  const mapCache = new WeakMap<object, Y.Map<any>>();
  const mapProxies = new WeakSet<object>();
  const arrayCache = new WeakMap<object, Y.Array<any>>();
  const arrayProxies = new WeakSet<object>();
  const textCache = new WeakMap<object, Y.Text>();
  const textProxies = new WeakSet<object>();

  const wrapYjsValue = (value: unknown): unknown => {
    if (mapProxies.has(value as any) || arrayProxies.has(value as any) || textProxies.has(value as any)) return value;
    const map = getYMap(value);
    if (map) {
      return createPermissionAwareYMapProxy({
        map,
        session,
        cache: mapCache,
        proxySet: mapProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const array = getYArray(value);
    if (array) {
      return createPermissionAwareYArrayProxy({
        array,
        session,
        cache: arrayCache,
        proxySet: arrayProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const text = getYText(value);
    if (text) {
      return createPermissionAwareYTextProxy({ text, session, cache: textCache, proxySet: textProxies });
    }
    return value;
  };

  (mgr as any).sheets = createPermissionAwareYArrayProxy({
    array: mgr.sheets,
    session,
    cache: arrayCache,
    proxySet: arrayProxies,
    wrapValue: wrapYjsValue,
  });
  return mgr;
}

export function createNamedRangeManagerForSessionWithPermissions(
  session: PermissionAwareWorkbookSession
): NamedRangeManager {
  const mgr = new NamedRangeManager({ doc: session.doc, transact: transactLocalWithWorkbookPermissions(session) });
  const mapCache = new WeakMap<object, Y.Map<any>>();
  const mapProxies = new WeakSet<object>();
  const arrayCache = new WeakMap<object, Y.Array<any>>();
  const arrayProxies = new WeakSet<object>();
  const textCache = new WeakMap<object, Y.Text>();
  const textProxies = new WeakSet<object>();

  const wrapYjsValue = (value: unknown): unknown => {
    if (mapProxies.has(value as any) || arrayProxies.has(value as any) || textProxies.has(value as any)) return value;
    const map = getYMap(value);
    if (map) {
      return createPermissionAwareYMapProxy({
        map,
        session,
        cache: mapCache,
        proxySet: mapProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const array = getYArray(value);
    if (array) {
      return createPermissionAwareYArrayProxy({
        array,
        session,
        cache: arrayCache,
        proxySet: arrayProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const text = getYText(value);
    if (text) {
      return createPermissionAwareYTextProxy({ text, session, cache: textCache, proxySet: textProxies });
    }
    return value;
  };

  (mgr as any).namedRanges = createPermissionAwareYMapProxy({
    map: mgr.namedRanges as any,
    session,
    cache: mapCache,
    proxySet: mapProxies,
    wrapValue: wrapYjsValue,
  });
  return mgr;
}

export function createMetadataManagerForSessionWithPermissions(session: PermissionAwareWorkbookSession): MetadataManager {
  const mgr = new MetadataManager({ doc: session.doc, transact: transactLocalWithWorkbookPermissions(session) });
  const mapCache = new WeakMap<object, Y.Map<any>>();
  const mapProxies = new WeakSet<object>();
  const arrayCache = new WeakMap<object, Y.Array<any>>();
  const arrayProxies = new WeakSet<object>();
  const textCache = new WeakMap<object, Y.Text>();
  const textProxies = new WeakSet<object>();

  const wrapYjsValue = (value: unknown): unknown => {
    if (mapProxies.has(value as any) || arrayProxies.has(value as any) || textProxies.has(value as any)) return value;
    const map = getYMap(value);
    if (map) {
      return createPermissionAwareYMapProxy({
        map,
        session,
        cache: mapCache,
        proxySet: mapProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const array = getYArray(value);
    if (array) {
      return createPermissionAwareYArrayProxy({
        array,
        session,
        cache: arrayCache,
        proxySet: arrayProxies,
        wrapValue: wrapYjsValue,
      });
    }
    const text = getYText(value);
    if (text) {
      return createPermissionAwareYTextProxy({ text, session, cache: textCache, proxySet: textProxies });
    }
    return value;
  };

  (mgr as any).metadata = createPermissionAwareYMapProxy({
    map: mgr.metadata as any,
    session,
    cache: mapCache,
    proxySet: mapProxies,
    wrapValue: wrapYjsValue,
  });
  return mgr;
}
