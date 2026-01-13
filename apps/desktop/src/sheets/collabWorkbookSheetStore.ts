import type { CollabSession } from "@formula/collab-session";
import * as Y from "yjs";
import { getYArray, getYMap, getYText, isYAbstractType } from "@formula/collab-yjs-utils";

import type { SheetMeta, SheetVisibility, TabColor } from "./workbookSheetStore";
import { WorkbookSheetStore } from "./workbookSheetStore";

type CollabSessionLike = Pick<CollabSession, "sheets" | "transactLocal">;

export type CollabSheetsKeyRef = { value: string };

function coerceCollabSheetField(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  const text = getYText(value);
  if (text) {
    try {
      return text.toString();
    } catch {
      return null;
    }
  }
  return null;
}

function coerceSheetVisibility(value: unknown): SheetVisibility | null {
  if (value === "visible" || value === "hidden" || value === "veryHidden") return value;
  return null;
}

/**
 * Normalize a user-provided or remote-provided tabColor into an 8-digit ARGB hex
 * string (uppercase), matching the @formula/collab-workbook schema.
 */
function coerceTabColorArgb(value: unknown): string | null {
  let raw = coerceCollabSheetField(value);
  if (!raw && value && typeof value === "object") {
    // Be tolerant of legacy/non-canonical shapes (e.g. `{ rgb: "FFFF0000" }`) that may
    // show up in older snapshots or malformed remote writes.
    const maybe: any = value;
    raw = coerceCollabSheetField(maybe?.rgb ?? maybe?.argb);
  }
  if (!raw) return null;

  let str = raw.trim();
  if (!str) return null;
  if (str.startsWith("#")) str = str.slice(1);

  // Allow 6-digit RGB hex for convenience by assuming opaque alpha.
  if (/^[0-9A-Fa-f]{6}$/.test(str)) {
    str = `FF${str}`;
  }

  if (!/^[0-9A-Fa-f]{8}$/.test(str)) return null;
  return str.toUpperCase();
}

export function listSheetsFromCollabSession(session: Pick<CollabSessionLike, "sheets">): SheetMeta[] {
  const out: SheetMeta[] = [];
  const seen = new Set<string>();
  const entries = session?.sheets?.toArray?.() ?? [];
  for (const entry of entries) {
    const map: any = entry;
    const id = coerceCollabSheetField(map?.get?.("id") ?? map?.id);
    if (!id) continue;
    const trimmed = id.trim();
    if (!trimmed || seen.has(trimmed)) continue;
    seen.add(trimmed);

    const name = coerceCollabSheetField(map?.get?.("name") ?? map?.name) ?? trimmed;

    const visibilityRaw = map?.get?.("visibility") ?? map?.visibility;
    const visibility = coerceSheetVisibility(coerceCollabSheetField(visibilityRaw)) ?? "visible";

    const tabColorRaw = map?.get?.("tabColor") ?? map?.tabColor;
    const tabColorArgb = tabColorRaw == null ? null : coerceTabColorArgb(tabColorRaw);
    const tabColor = tabColorArgb ? ({ rgb: tabColorArgb } satisfies TabColor) : undefined;

    out.push({ id: trimmed, name, visibility, tabColor });
  }

  return out.length > 0 ? out : [{ id: "Sheet1", name: "Sheet1", visibility: "visible" }];
}

export function computeCollabSheetsKey(sheets: ReadonlyArray<Pick<SheetMeta, "id" | "name" | "visibility" | "tabColor">>): string {
  // Intentionally ignore unknown keys in the underlying Y.Map entries. We only
  // rebuild the desktop sheet store for tab-relevant metadata changes.
  return JSON.stringify(sheets.map((s) => [s.id, s.name, s.visibility, s.tabColor?.rgb ?? null]));
}

export function findCollabSheetIndexById(session: Pick<CollabSessionLike, "sheets">, sheetId: string): number {
  const query = String(sheetId ?? "").trim();
  if (!query) return -1;
  for (let i = 0; i < session.sheets.length; i += 1) {
    const entry: any = session.sheets.get(i);
    const id = coerceCollabSheetField(entry?.get?.("id") ?? entry?.id);
    if (id && id.trim() === query) return i;
  }
  return -1;
}

function cloneCollabSheetMetaValue(value: unknown): unknown {
  if (value == null) return value;
  if (typeof value !== "object") return value;

  // Clone nested Yjs types into local constructors so they can be safely re-inserted
  // into the document (Yjs types cannot be "re-parented" after deletion).
  const text = getYText(value);
  if (text) {
    const out = new Y.Text();
    const delta = text.toDelta();
    const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
    try {
      out.applyDelta(typeof structuredCloneFn === "function" ? structuredCloneFn(delta) : JSON.parse(JSON.stringify(delta)));
    } catch {
      // Best-effort: fall back to inserting the string representation.
      try {
        out.insert(0, text.toString());
      } catch {
        // ignore
      }
    }
    return out;
  }

  const map = getYMap(value);
  if (map) {
    const out = new Y.Map<unknown>();
    map.forEach((v: unknown, k: string) => {
      out.set(String(k), cloneCollabSheetMetaValue(v));
    });
    return out;
  }

  const array = getYArray(value);
  if (array) {
    const out = new Y.Array<unknown>();
    for (const item of array.toArray()) {
      out.push([cloneCollabSheetMetaValue(item)]);
    }
    return out;
  }

  // Avoid copying other Yjs types directly (e.g. Xml) since we don't have a safe
  // cloning strategy for them here.
  if (isYAbstractType(value)) return undefined;

  const structuredCloneFn = (globalThis as any).structuredClone as ((input: unknown) => unknown) | undefined;
  if (typeof structuredCloneFn === "function") {
    try {
      return structuredCloneFn(value);
    } catch {
      // Fall through to JSON clone below.
    }
  }

  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function cloneCollabSheetMap(entry: unknown): Y.Map<unknown> {
  const out = new Y.Map<unknown>();
  const map: any = entry;

  if (map && typeof map.forEach === "function") {
    map.forEach((value: unknown, key: string) => {
      const k = String(key ?? "");
      if (!k) return;
      if (k === "id") return;
      if (k === "name") return;
      const cloned = cloneCollabSheetMetaValue(value);
      if (cloned === undefined) return;
      out.set(k, cloned);
    });
  }

  const id = coerceCollabSheetField(map?.get?.("id") ?? map?.id);
  if (id) out.set("id", id.trim());

  const hasName = typeof map?.has === "function" ? Boolean(map.has("name")) : map?.get?.("name") !== undefined;
  if (hasName) {
    const nameRaw = map?.get?.("name") ?? map?.name;
    const name = coerceCollabSheetField(nameRaw);
    if (name != null) out.set("name", name);
  }

  return out;
}

export class CollabWorkbookSheetStore extends WorkbookSheetStore {
  private readonly canEditWorkbook: () => boolean;

  constructor(
    private readonly session: CollabSessionLike,
    initialSheets: ConstructorParameters<typeof WorkbookSheetStore>[0],
    private readonly keyRef: CollabSheetsKeyRef,
    opts?: { canEditWorkbook?: () => boolean },
  ) {
    super(initialSheets);
    this.canEditWorkbook = opts?.canEditWorkbook ?? (() => true);
  }

  private refreshKeyFromSession(): void {
    this.keyRef.value = computeCollabSheetsKey(listSheetsFromCollabSession(this.session));
  }

  override rename(id: string, newName: string): void {
    if (!this.canEditWorkbook()) return;
    const before = this.getName(id);
    super.rename(id, newName);
    const after = this.getName(id);
    if (!after || after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;
      entry.set("name", after);
      // This update originated locally; update the cached key so our observer
      // doesn't unnecessarily rebuild the sheet store instance.
      this.refreshKeyFromSession();
    });
  }

  override move(id: string, toIndex: number): void {
    if (!this.canEditWorkbook()) return;
    const before = this.listAll().map((s) => s.id).join("|");
    super.move(id, toIndex);
    const after = this.listAll().map((s) => s.id).join("|");
    if (after === before) return;

    this.session.transactLocal(() => {
      const fromIndex = findCollabSheetIndexById(this.session, id);
      if (fromIndex < 0) return;

      const entry: any = this.session.sheets.get(fromIndex);
      if (!entry) return;

      const clone = cloneCollabSheetMap(entry);
      this.session.sheets.delete(fromIndex, 1);
      this.session.sheets.insert(toIndex, [clone as any]);

      // This update originated locally; update the cached key so our observer
      // doesn't unnecessarily rebuild the sheet store instance.
      this.refreshKeyFromSession();
    });
  }

  override remove(id: string): void {
    if (!this.canEditWorkbook()) return;
    super.remove(id);

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      this.session.sheets.delete(idx, 1);
      this.refreshKeyFromSession();
    });
  }

  override hide(id: string): void {
    if (!this.canEditWorkbook()) return;
    const before = this.getById(id)?.visibility;
    super.hide(id);
    const after = this.getById(id)?.visibility;
    if (!after || after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;
      entry.set("visibility", after);
      this.refreshKeyFromSession();
    });
  }

  override unhide(id: string): void {
    if (!this.canEditWorkbook()) return;
    const before = this.getById(id)?.visibility;
    super.unhide(id);
    const after = this.getById(id)?.visibility;
    if (!after || after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;
      entry.set("visibility", after);
      this.refreshKeyFromSession();
    });
  }

  override setVisibility(id: string, visibility: SheetVisibility): void {
    if (!this.canEditWorkbook()) return;
    const before = this.getById(id)?.visibility;
    super.setVisibility(id, visibility);
    const after = this.getById(id)?.visibility;
    if (!after || after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;
      entry.set("visibility", after);
      this.refreshKeyFromSession();
    });
  }

  override setTabColor(id: string, color: TabColor | undefined): void {
    if (!this.canEditWorkbook()) return;
    // Collab sheet metadata stores tab colors as 8-digit ARGB hex strings (e.g. "FFFF0000").
    // Normalize UI-provided `TabColor.rgb` values into that representation before writing.
    const raw = color?.rgb;
    let normalized: string | null = null;
    if (raw != null) {
      const trimmed = String(raw).trim();
      if (trimmed) {
        const withoutHash = trimmed.startsWith("#") ? trimmed.slice(1) : trimmed;
        if (/^[0-9A-Fa-f]{8}$/.test(withoutHash)) {
          normalized = withoutHash.toUpperCase();
        } else if (/^[0-9A-Fa-f]{6}$/.test(withoutHash)) {
          normalized = `FF${withoutHash.toUpperCase()}`;
        } else {
          throw new Error(`Invalid tabColor (expected 6 or 8-digit hex): ${trimmed}`);
        }
      }
    }

    const before = this.getById(id)?.tabColor?.rgb ?? null;
    const normalizedColor = normalized ? ({ rgb: normalized } satisfies TabColor) : undefined;
    super.setTabColor(id, normalizedColor);
    const after = this.getById(id)?.tabColor?.rgb ?? null;
    if (after === before) return;

    this.session.transactLocal(() => {
      const idx = findCollabSheetIndexById(this.session, id);
      if (idx < 0) return;
      const entry: any = this.session.sheets.get(idx);
      if (!entry || typeof entry.set !== "function") return;

      if (!normalized) {
        if (typeof entry.delete === "function") entry.delete("tabColor");
      } else {
        entry.set("tabColor", normalized);
      }

      this.refreshKeyFromSession();
    });
  }
}
