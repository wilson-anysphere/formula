import * as Y from "yjs";
import { getSheetNameValidationErrorMessage } from "@formula/workbook-backend";
import {
  cloneYjsValue,
  getArrayRoot,
  getDocTypeConstructors,
  getMapRoot,
  getYArray,
  getYMap,
  getYText,
  yjsValueToJson,
} from "@formula/collab-yjs-utils";

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

function isRecord(value: unknown): value is Record<string, any> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function normalizeFrozenCount(value: unknown): number {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

// Defensive cap: drawing ids can be authored via remote/shared state (sheet view state). Keep
// validation strict so workbook schema normalization doesn't deep-copy pathological ids (e.g.
// multi-megabyte strings) when merging duplicate sheet entries.
const MAX_DRAWING_ID_STRING_CHARS = 4096;

function sanitizeDrawingsJson(value: unknown): any[] | null {
  if (!Array.isArray(value)) return null;
  const out: any[] = [];
  for (const entry of value) {
    if (!isRecord(entry)) continue;
    const rawId = (entry as any).id;
    let normalizedId: string | number;
    if (typeof rawId === "string") {
      if (rawId.length > MAX_DRAWING_ID_STRING_CHARS) continue;
      const trimmed = rawId.trim();
      if (!trimmed) continue;
      normalizedId = trimmed;
    } else if (typeof rawId === "number") {
      if (!Number.isSafeInteger(rawId)) continue;
      normalizedId = rawId;
    } else {
      continue;
    }
    out.push({ ...entry, id: normalizedId });
  }
  return out;
}

function sanitizeDrawingsValue(value: unknown): any[] | null {
  return sanitizeDrawingsJson(yjsValueToJson(value));
}

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
  const YArrayCtor = sheets.constructor as unknown as { new (): Y.Array<any> };
  const { Text: YTextCtor } = getDocTypeConstructors(doc as any);
  const cloneCtors = { Map: YMapCtor, Array: YArrayCtor, Text: YTextCtor as unknown as { new (): Y.Text } };

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

            // Preserve sheet view/layout metadata (frozen panes, axis sizes, drawings, merged ranges)
            // across deterministic duplicate-sheet pruning. This avoids losing shared sheet-level
            // UI metadata if it was written to (or migrated into) a losing entry.
            //
            // Note: this is best-effort and intentionally biased toward "prefer non-default" values
            // so placeholder sheets (typically `{ frozenRows: 0, frozenCols: 0 }`) don't wipe richer
            // layout state.
            const viewKeysToMerge = [
              "view",
              // Legacy/experimental top-level view keys.
              "frozenRows",
              "frozenCols",
              "backgroundImageId",
              "background_image_id",
              "colWidths",
              "rowHeights",
              "mergedRanges",
              "mergedCells",
              "merged_cells",
              "drawings",
            ];

            for (const key of viewKeysToMerge) {
              const winnerVal = winner.get(key);
              const entryVal = entry.get(key);

                if (winnerVal === undefined) {
                  if (entryVal !== undefined) {
                    if (key === "drawings") {
                      const sanitized = sanitizeDrawingsValue(entryVal);
                      winner.set(key, sanitized ?? cloneYjsValue(entryVal, cloneCtors));
                    } else if (key === "view") {
                      const viewMap = getYMap(entryVal);
                      if (viewMap) {
                        const cloned = cloneYjsValue(entryVal, cloneCtors);
                        const clonedMap = getYMap(cloned);
                        if (clonedMap) {
                          const drawings = clonedMap.get("drawings");
                          const sanitized = sanitizeDrawingsValue(drawings);
                          if (sanitized) clonedMap.set("drawings", sanitized);
                        }
                        winner.set(key, cloned);
                      } else {
                        const json = yjsValueToJson(entryVal);
                        if (isRecord(json) && Object.prototype.hasOwnProperty.call(json, "drawings")) {
                          const sanitized = sanitizeDrawingsJson((json as any).drawings);
                          winner.set(
                            key,
                            sanitized ? { ...json, drawings: sanitized } : json,
                          );
                        } else {
                          winner.set(key, json);
                        }
                      }
                    } else {
                      winner.set(key, cloneYjsValue(entryVal, cloneCtors));
                    }
                  }
                  continue;
                }

              // For legacy numeric view keys, prefer non-zero values over 0.
              if (key === "frozenRows" || key === "frozenCols") {
                const winnerNum = normalizeFrozenCount(yjsValueToJson(winnerVal));
                const entryNum = normalizeFrozenCount(yjsValueToJson(entryVal));
                if (winnerNum === 0 && entryNum > 0) {
                  winner.set(key, entryNum);
                }
              }

              if (key === "colWidths" || key === "rowHeights") {
                const winnerJson = yjsValueToJson(winnerVal);
                const entryJson = yjsValueToJson(entryVal);
                const winnerCount = isRecord(winnerJson) ? Object.keys(winnerJson).length : 0;
                const entryCount = isRecord(entryJson) ? Object.keys(entryJson).length : 0;
                if (winnerCount === 0 && entryCount > 0) {
                  winner.set(key, cloneYjsValue(entryVal, cloneCtors));
                }
              }

              // For legacy list keys, prefer non-empty over empty/undefined.
              if (key === "drawings" || key === "mergedRanges" || key === "mergedCells" || key === "merged_cells") {
                if (key === "drawings") {
                  const winnerArr = sanitizeDrawingsValue(winnerVal) ?? [];
                  const entryArr = sanitizeDrawingsValue(entryVal) ?? [];
                  if (winnerArr.length === 0 && entryArr.length > 0) {
                    winner.set(key, entryArr);
                  }
                  continue;
                }

                const winnerArr = Array.isArray(yjsValueToJson(winnerVal)) ? yjsValueToJson(winnerVal) : [];
                const entryArr = Array.isArray(yjsValueToJson(entryVal)) ? yjsValueToJson(entryVal) : [];
                if (winnerArr.length === 0 && entryArr.length > 0) {
                  winner.set(key, cloneYjsValue(entryVal, cloneCtors));
                }
              }

              if (key === "backgroundImageId" || key === "background_image_id") {
                const winnerStr = coerceString(yjsValueToJson(winnerVal))?.trim() ?? "";
                const entryStr = coerceString(yjsValueToJson(entryVal))?.trim() ?? "";
                if (!winnerStr && entryStr) {
                  winner.set(key, entryStr);
                }
              }
            }

            // If both entries have a `view` object, merge it field-by-field so we don't lose
            // shared metadata like drawings/mergedRanges when one entry is missing them.
            const winnerViewRaw = winner.get("view");
            const entryViewRaw = entry.get("view");
            if (winnerViewRaw !== undefined && entryViewRaw !== undefined) {
              const winnerViewMap = getYMap(winnerViewRaw);
              const entryViewMap = getYMap(entryViewRaw);
              if (winnerViewMap && entryViewMap) {
                const keys = Array.from(entryViewMap.keys()).sort();
                for (const k of keys) {
                  const wv = winnerViewMap.get(k);
                  const ev = entryViewMap.get(k);

                  if (wv === undefined) {
                    if (k === "drawings") {
                      const sanitized = sanitizeDrawingsValue(ev);
                      winnerViewMap.set(k, sanitized ?? cloneYjsValue(ev, cloneCtors));
                    } else {
                      winnerViewMap.set(k, cloneYjsValue(ev, cloneCtors));
                    }
                    continue;
                  }

                  if (k === "frozenRows" || k === "frozenCols") {
                    const wNum = normalizeFrozenCount(yjsValueToJson(wv));
                    const eNum = normalizeFrozenCount(yjsValueToJson(ev));
                    if (wNum === 0 && eNum > 0) winnerViewMap.set(k, eNum);
                    continue;
                  }

                  if (k === "backgroundImageId" || k === "background_image_id") {
                    const wStr = coerceString(yjsValueToJson(wv))?.trim() ?? "";
                    const eStr = coerceString(yjsValueToJson(ev))?.trim() ?? "";
                    if (!wStr && eStr) winnerViewMap.set(k, eStr);
                    continue;
                  }

                  if (k === "colWidths" || k === "rowHeights") {
                    const wJson = yjsValueToJson(wv);
                    const eJson = yjsValueToJson(ev);
                    const wCount = isRecord(wJson) ? Object.keys(wJson).length : 0;
                    const eCount = isRecord(eJson) ? Object.keys(eJson).length : 0;
                    if (wCount === 0 && eCount > 0) {
                      winnerViewMap.set(k, cloneYjsValue(ev, cloneCtors));
                    }
                    continue;
                  }

                  if (k === "drawings" || k === "mergedRanges" || k === "mergedCells" || k === "merged_cells") {
                    if (k === "drawings") {
                      const wArr = sanitizeDrawingsValue(wv) ?? [];
                      const eArr = sanitizeDrawingsValue(ev) ?? [];
                      if (wArr.length === 0 && eArr.length > 0) {
                        winnerViewMap.set(k, eArr);
                      }
                      continue;
                    }

                    const wArr = Array.isArray(yjsValueToJson(wv)) ? yjsValueToJson(wv) : [];
                    const eArr = Array.isArray(yjsValueToJson(ev)) ? yjsValueToJson(ev) : [];
                    if (wArr.length === 0 && eArr.length > 0) {
                      winnerViewMap.set(k, cloneYjsValue(ev, cloneCtors));
                    }
                    continue;
                  }
                }
              } else {
                // Fall back to JSON merge for non-map view encodings (plain objects).
                const wJson = yjsValueToJson(winnerViewRaw);
                const eJson = yjsValueToJson(entryViewRaw);
                if (isRecord(wJson) && isRecord(eJson)) {
                  /** @type {Record<string, any>} */
                  const merged = { ...wJson };
                  for (const [k, ev] of Object.entries(eJson)) {
                    const wv = merged[k];
                    if (wv === undefined) {
                      if (k === "drawings") {
                        merged[k] = sanitizeDrawingsJson(ev) ?? structuredClone(ev);
                      } else {
                        merged[k] = structuredClone(ev);
                      }
                      continue;
                    }

                    if (k === "frozenRows" || k === "frozenCols") {
                      const wNum = normalizeFrozenCount(wv);
                      const eNum = normalizeFrozenCount(ev);
                      if (wNum === 0 && eNum > 0) merged[k] = eNum;
                      continue;
                    }

                    if (k === "backgroundImageId" || k === "background_image_id") {
                      const wStr = coerceString(wv)?.trim() ?? "";
                      const eStr = coerceString(ev)?.trim() ?? "";
                      if (!wStr && eStr) merged[k] = eStr;
                      continue;
                    }

                    if (k === "colWidths" || k === "rowHeights") {
                      const wObj = isRecord(wv) ? wv : {};
                      const eObj = isRecord(ev) ? ev : {};
                      if (Object.keys(wObj).length === 0 && Object.keys(eObj).length > 0) {
                        merged[k] = structuredClone(ev);
                      }
                      continue;
                    }

                    if (k === "drawings") {
                      const wArr = sanitizeDrawingsJson(wv) ?? [];
                      const eArr = sanitizeDrawingsJson(ev) ?? [];
                      if (wArr.length === 0 && eArr.length > 0) merged[k] = eArr;
                      continue;
                    }

                    if (k === "mergedRanges" || k === "mergedCells" || k === "merged_cells") {
                      const wArr = Array.isArray(wv) ? wv : [];
                      const eArr = Array.isArray(ev) ? ev : [];
                      if (wArr.length === 0 && eArr.length > 0) merged[k] = structuredClone(ev);
                      continue;
                    }
                  }
                  winner.set("view", merged);
                }
              }
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
  if (text) return yjsValueToJson(text);
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
    const { Text } = getDocTypeConstructors(opts.doc as any);
    this.YMapCtor = cells.constructor as unknown as { new (): Y.Map<unknown> };
    this.YArrayCtor = this.sheets.constructor as unknown as { new (): Y.Array<any> };
    this.YTextCtor = Text as unknown as { new (): Y.Text };
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

      const sheetClone = cloneYjsValue(sheet, { Map: this.YMapCtor, Array: this.YArrayCtor, Text: this.YTextCtor });
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
