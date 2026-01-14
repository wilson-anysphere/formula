import {
  EXCEL_MAX_SHEET_NAME_LEN,
  INVALID_SHEET_NAME_CHARACTERS,
  getSheetNameValidationErrorMessage,
} from "@formula/workbook-backend";

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

// Keep in sync with `apps/desktop/src/workbook/workbook.ts`.
export type TabColor = {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
};

function normalizeTabColor(value: TabColor | undefined): TabColor | undefined {
  if (!value) return undefined;
  const out = { ...value } as TabColor;

  if (typeof out.rgb === "string") {
    const rgb = out.rgb.trim();
    if (!rgb) {
      delete (out as any).rgb;
    } else {
      out.rgb = rgb.toUpperCase();
    }
  }

  // If all fields are empty/undefined, treat as no color.
  if (
    out.rgb == null &&
    out.theme == null &&
    out.indexed == null &&
    out.tint == null &&
    out.auto == null
  ) {
    return undefined;
  }

  return out;
}

function tabColorEquals(a: TabColor | undefined, b: TabColor | undefined): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return (
    (a.rgb ?? null) === (b.rgb ?? null) &&
    (a.theme ?? null) === (b.theme ?? null) &&
    (a.indexed ?? null) === (b.indexed ?? null) &&
    (a.tint ?? null) === (b.tint ?? null) &&
    (a.auto ?? null) === (b.auto ?? null)
  );
}

export type SheetMeta = {
  /**
   * Stable identifier used by the DocumentController + backend.
   * (Not necessarily the same as the user-facing sheet name.)
   */
  id: string;
  /**
   * User-facing sheet name shown on tabs and in the sheet switcher.
   */
  name: string;
  /**
   * Excel-style visibility.
   */
  visibility: SheetVisibility;
  /**
   * Excel-style tab color, round-trippable through XLSX.
   */
  tabColor?: TabColor;
};

export const SHEET_NAME_MAX_LENGTH = EXCEL_MAX_SHEET_NAME_LEN;
export const SHEET_NAME_FORBIDDEN_CHARS = INVALID_SHEET_NAME_CHARACTERS;
export const SHEET_NAME_FORBIDDEN_CHARS_REGEX = /[:\\\/\?\*\[\]]/;

function normalizeSheetNameForCaseInsensitiveCompare(name: string): string {
  // Match backend workbook uniqueness semantics:
  // - Unicode NFKC normalization
  // - Unicode uppercasing
  try {
    return name.normalize("NFKC").toUpperCase();
  } catch {
    return name.toUpperCase();
  }
}

/**
 * Trim the proposed sheet name (Excel does this) and return the normalized value.
 */
export function normalizeSheetName(name: string): string {
  return String(name ?? "").trim();
}

/**
 * Validate a proposed sheet name using Excel-like constraints.
 *
 * Returns the normalized (trimmed) name when valid; otherwise throws.
 */
export function validateSheetName(
  name: string,
  opts: {
    /**
     * Existing workbook sheets for uniqueness validation.
     */
    sheets: ReadonlyArray<Pick<SheetMeta, "id" | "name">>;
    /**
     * Sheet id to ignore during uniqueness validation (rename of current sheet).
     */
    ignoreId?: string | null;
  },
): string {
  const normalized = normalizeSheetName(name);
  const existingNames = opts.sheets
    .filter((sheet) => !(opts.ignoreId && sheet.id === opts.ignoreId))
    .map((sheet) => sheet.name);

  const err = getSheetNameValidationErrorMessage(normalized, { existingNames });
  if (err) throw new Error(err);

  return normalized;
}

/**
 * Generate the next default sheet name (`Sheet1`, `Sheet2`, …), choosing the first
 * available number (case-insensitive compare).
 */
export function generateDefaultSheetName(sheets: ReadonlyArray<Pick<SheetMeta, "name">>): string {
  const existing = new Set(sheets.map((s) => normalizeSheetNameForCaseInsensitiveCompare(s.name)));
  for (let n = 1; ; n += 1) {
    const candidate = `Sheet${n}`;
    if (!existing.has(normalizeSheetNameForCaseInsensitiveCompare(candidate))) return candidate;
  }
}

function cloneSheetMeta(meta: SheetMeta): SheetMeta {
  return { ...meta, tabColor: meta.tabColor ? { ...meta.tabColor } : undefined };
}

export type WorkbookSheetStoreListener = () => void;

/**
 * A simple, synchronous workbook sheet metadata store.
 *
 * This store is intentionally UI-framework agnostic so it can be shared between
 * `main.ts` DOM wiring and React components.
 */
export class WorkbookSheetStore {
  private sheets: SheetMeta[];
  private readonly listeners = new Set<WorkbookSheetStoreListener>();

  constructor(initialSheets: SheetMeta[] = []) {
    // Validate invariants up-front so downstream consumers can rely on them.
    const byId = new Set<string>();
    for (const sheet of initialSheets) {
      const id = String(sheet.id ?? "").trim();
      if (!id) throw new Error("Sheet id cannot be empty");
      if (byId.has(id)) throw new Error(`Duplicate sheet id: ${id}`);
      byId.add(id);
    }

    const normalizedSheets: SheetMeta[] = [];
    for (const sheet of initialSheets) {
      const normalizedName = validateSheetName(sheet.name, { sheets: normalizedSheets, ignoreId: null });
      const visibility: SheetVisibility =
        sheet.visibility === "visible" || sheet.visibility === "hidden" || sheet.visibility === "veryHidden"
          ? sheet.visibility
          : "visible";
      normalizedSheets.push({
        id: String(sheet.id).trim(),
        name: normalizedName,
        visibility,
        tabColor: normalizeTabColor(sheet.tabColor),
      });
    }

    this.sheets = normalizedSheets;
  }

  subscribe(listener: WorkbookSheetStoreListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of [...this.listeners]) {
      try {
        listener();
      } catch {
        // ignore listener errors
      }
    }
  }

  listAll(): SheetMeta[] {
    return this.sheets.map(cloneSheetMeta);
  }

  listVisible(): SheetMeta[] {
    return this.sheets.filter((s) => s.visibility === "visible").map(cloneSheetMeta);
  }

  getById(id: string): SheetMeta | undefined {
    const sheet = this.sheets.find((s) => s.id === id);
    return sheet ? cloneSheetMeta(sheet) : undefined;
  }

  getName(id: string): string | undefined {
    return this.sheets.find((s) => s.id === id)?.name;
  }

  resolveIdByName(name: string): string | undefined {
    const normalized = normalizeSheetName(name);
    if (!normalized) return undefined;
    const targetCi = normalizeSheetNameForCaseInsensitiveCompare(normalized);
    return this.sheets.find((s) => normalizeSheetNameForCaseInsensitiveCompare(s.name) === targetCi)?.id;
  }

  addAfter(activeId: string, input: { id: string; name?: string }): SheetMeta {
    const id = String(input.id ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    if (this.sheets.some((s) => s.id === id)) throw new Error(`Duplicate sheet id: ${id}`);

    const name =
      input.name !== undefined
        ? validateSheetName(input.name, { sheets: this.sheets, ignoreId: null })
        : generateDefaultSheetName(this.sheets);

    const sheet: SheetMeta = { id, name, visibility: "visible" };

    const activeIdx = this.sheets.findIndex((s) => s.id === activeId);
    const insertIdx = activeIdx === -1 ? this.sheets.length : activeIdx + 1;
    const next = this.sheets.slice();
    next.splice(insertIdx, 0, sheet);
    this.sheets = next;
    this.emit();
    return cloneSheetMeta(sheet);
  }

  rename(id: string, newName: string): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");

    const normalized = validateSheetName(newName, { sheets: this.sheets, ignoreId: id });
    const current = this.sheets[idx]!;
    if (current.name === normalized) return;

    const next = this.sheets.slice();
    next[idx] = { ...current, name: normalized };
    this.sheets = next;
    this.emit();
  }

  move(id: string, toIndex: number): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    if (!Number.isInteger(toIndex)) throw new Error("Invalid index");
    if (toIndex < 0 || toIndex >= this.sheets.length) throw new Error("Invalid index");
    if (idx === toIndex) return;

    const next = this.sheets.slice();
    const [sheet] = next.splice(idx, 1);
    if (!sheet) throw new Error("Sheet not found");
    next.splice(toIndex, 0, sheet);
    this.sheets = next;
    this.emit();
  }

  remove(id: string): void {
    if (this.sheets.length <= 1) throw new Error("Cannot delete the last sheet");
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    const sheet = this.sheets[idx]!;
    // Mirror Excel behavior: prevent deleting the last visible sheet, even if hidden sheets remain.
    if (sheet.visibility === "visible") {
      const visibleCount = this.sheets.reduce((count, s) => count + (s.visibility === "visible" ? 1 : 0), 0);
      if (visibleCount <= 1) throw new Error("Cannot delete the last visible sheet");
    }
    const next = this.sheets.slice();
    next.splice(idx, 1);
    this.sheets = next;
    this.emit();
  }

  hide(id: string): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    const sheet = this.sheets[idx]!;
    if (sheet.visibility !== "visible") return;

    const visibleCount = this.sheets.reduce((count, s) => count + (s.visibility === "visible" ? 1 : 0), 0);
    if (visibleCount <= 1) throw new Error("Cannot hide the last visible sheet");

    const next = this.sheets.slice();
    next[idx] = { ...sheet, visibility: "hidden" };
    this.sheets = next;
    this.emit();
  }

  unhide(id: string): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    const sheet = this.sheets[idx]!;
    if (sheet.visibility === "visible") return;

    const next = this.sheets.slice();
    next[idx] = { ...sheet, visibility: "visible" };
    this.sheets = next;
    this.emit();
  }

  /**
   * Set a sheet's Excel visibility value.
   *
   * Unlike `hide()` / `unhide()`, this supports transitioning between "hidden" and "veryHidden"
   * while still enforcing the "cannot hide the last visible sheet" invariant.
   */
  setVisibility(id: string, visibility: SheetVisibility, opts?: { allowHideLastVisible?: boolean }): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    const sheet = this.sheets[idx]!;
    const nextVisibility: SheetVisibility =
      visibility === "visible" || visibility === "hidden" || visibility === "veryHidden" ? visibility : "visible";
    if (sheet.visibility === nextVisibility) return;

    // Mirror Excel behavior: prevent hiding the last visible sheet.
    if (sheet.visibility === "visible" && nextVisibility !== "visible") {
      const visibleCount = this.sheets.reduce((count, s) => count + (s.visibility === "visible" ? 1 : 0), 0);
      if (visibleCount <= 1 && !opts?.allowHideLastVisible) throw new Error("Cannot hide the last visible sheet");
    }

    const next = this.sheets.slice();
    next[idx] = { ...sheet, visibility: nextVisibility };
    this.sheets = next;
    this.emit();
  }

  setTabColor(id: string, color: TabColor | undefined): void {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error("Sheet not found");
    const sheet = this.sheets[idx]!;

    const nextColor = normalizeTabColor(color);
    if (tabColorEquals(sheet.tabColor, nextColor)) return;

    const next = this.sheets.slice();
    next[idx] = { ...sheet, tabColor: nextColor };
    this.sheets = next;
    this.emit();
  }

  /**
   * Replace the entire ordered sheet list.
   *
   * This is useful when switching workbooks (e.g. after opening a file) where the backend
   * provides a canonical sheet ordering + display names.
   */
  replaceAll(nextSheets: ReadonlyArray<{ id: string; name: string; visibility?: SheetVisibility; tabColor?: TabColor }>): void {
    const byId = new Set<string>();
    for (const sheet of nextSheets) {
      const id = String(sheet?.id ?? "").trim();
      if (!id) throw new Error("Sheet id cannot be empty");
      if (byId.has(id)) throw new Error(`Duplicate sheet id: ${id}`);
      byId.add(id);
    }

    const normalizedSheets: SheetMeta[] = [];
    for (const sheet of nextSheets) {
      const normalizedName = validateSheetName(sheet.name, { sheets: normalizedSheets, ignoreId: null });
      const visibility: SheetVisibility =
        sheet.visibility === "visible" || sheet.visibility === "hidden" || sheet.visibility === "veryHidden"
          ? sheet.visibility
          : "visible";
      normalizedSheets.push({
        id: String(sheet.id).trim(),
        name: normalizedName,
        visibility,
        tabColor: normalizeTabColor(sheet.tabColor),
      });
    }

    this.sheets = normalizedSheets;
    this.emit();
  }
}
