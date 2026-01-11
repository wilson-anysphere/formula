import { rewriteSheetNamesInFormula } from "./formulaRewrite";
import { fromA1, toA1 } from "@formula/spreadsheet-frontend/a1";

export type SheetVisibility = "visible" | "hidden" | "veryHidden";

export type TabColor = {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
};

export type CellValue = number | string | boolean | null;

export type Cell = {
  value: CellValue;
  formula?: string;
};

export type Sheet = {
  id: string;
  name: string;
  visibility: SheetVisibility;
  tabColor?: TabColor;
  cells: Map<string, Cell>;
};

export interface WorkbookPersistence {
  createSheet(sheet: Sheet): Promise<void> | void;
  updateSheet(sheet: Sheet): Promise<void> | void;
  deleteSheet(sheetId: string): Promise<void> | void;
  reorderSheets(sheetIdsInOrder: string[]): Promise<void> | void;
}

export class Workbook {
  readonly sheets: Sheet[] = [];

  constructor(private readonly persistence?: WorkbookPersistence) {}

  addSheet(name?: string): Sheet {
    const sheet: Sheet = {
      id: crypto.randomUUID(),
      name: name ?? this.nextDefaultSheetName(),
      visibility: "visible",
      cells: new Map(),
    };
    this.assertUniqueSheetName(sheet.name, null);

    this.sheets.push(sheet);
    this.persistence?.createSheet(sheet);
    return sheet;
  }

  renameSheet(sheetId: string, newName: string): void {
    const sheet = this.getSheet(sheetId);
    const normalized = newName.trim();
    if (!normalized) throw new Error("Sheet name cannot be empty");

    const oldName = sheet.name;
    this.assertUniqueSheetName(normalized, sheetId);

    for (const s of this.sheets) {
      for (const cell of s.cells.values()) {
        if (!cell.formula) continue;
        cell.formula = rewriteSheetNamesInFormula(cell.formula, oldName, normalized);
      }
    }

    sheet.name = normalized;
    this.persistence?.updateSheet(sheet);
  }

  reorderSheet(sheetId: string, newIndex: number): void {
    const idx = this.sheets.findIndex((s) => s.id === sheetId);
    if (idx === -1) throw new Error("Sheet not found");
    if (newIndex < 0 || newIndex >= this.sheets.length) throw new Error("Invalid index");
    if (idx === newIndex) return;

    const [sheet] = this.sheets.splice(idx, 1);
    this.sheets.splice(newIndex, 0, sheet);
    this.persistence?.reorderSheets(this.sheets.map((s) => s.id));
  }

  deleteSheet(sheetId: string): void {
    if (this.sheets.length <= 1) throw new Error("Cannot delete the last sheet");
    const idx = this.sheets.findIndex((s) => s.id === sheetId);
    if (idx === -1) throw new Error("Sheet not found");
    this.sheets.splice(idx, 1);
    this.persistence?.deleteSheet(sheetId);
  }

  setSheetVisibility(sheetId: string, visibility: SheetVisibility): void {
    const sheet = this.getSheet(sheetId);
    sheet.visibility = visibility;
    this.persistence?.updateSheet(sheet);
  }

  setSheetTabColor(sheetId: string, tabColor?: TabColor): void {
    const sheet = this.getSheet(sheetId);
    sheet.tabColor = tabColor;
    this.persistence?.updateSheet(sheet);
  }

  getSheet(sheetId: string): Sheet {
    const sheet = this.sheets.find((s) => s.id === sheetId);
    if (!sheet) throw new Error("Sheet not found");
    return sheet;
  }

  getCell(sheetId: string, a1: string): Cell {
    const sheet = this.getSheet(sheetId);
    return sheet.cells.get(a1) ?? { value: null };
  }

  setCell(sheetId: string, a1: string, cell: Cell): void {
    const sheet = this.getSheet(sheetId);
    sheet.cells.set(a1, cell);
  }

  evaluateNumberCell(sheetId: string, a1: string): number | null {
    const sheet = this.getSheet(sheetId);
    const cell = sheet.cells.get(a1);
    if (!cell) return null;
    if (!cell.formula) return typeof cell.value === "number" ? cell.value : null;
    return evalFormulaNumber(this, sheetId, cell.formula);
  }

  private nextDefaultSheetName(): string {
    for (let n = 1; ; n += 1) {
      const candidate = `Sheet${n}`;
      if (!this.sheets.some((s) => s.name.toLowerCase() === candidate.toLowerCase())) {
        return candidate;
      }
    }
  }

  private assertUniqueSheetName(name: string, ignoreId: string | null): void {
    const nameCi = name.toLowerCase();
    for (const sheet of this.sheets) {
      if (ignoreId && sheet.id === ignoreId) continue;
      if (sheet.name.toLowerCase() === nameCi) {
        throw new Error("Duplicate sheet name");
      }
    }
  }
}

function evalFormulaNumber(workbook: Workbook, currentSheetId: string, formula: string): number | null {
  const expr = formula.startsWith("=") ? formula.slice(1) : formula;
  let sum = 0;

  for (const raw of expr.split("+")) {
    const part = raw.trim();
    if (!part) continue;

    const value =
      workbookSheetRef(workbook, part) ??
      localRef(workbook, currentSheetId, part) ??
      (Number.isFinite(Number(part)) ? Number(part) : null);

    if (value === null) return null;
    sum += value;
  }

  return sum;
}

function workbookSheetRef(workbook: Workbook, part: string): number | null {
  const bang = part.indexOf("!");
  if (bang === -1) return null;

  const sheetPart = part.slice(0, bang);
  const cellPart = part.slice(bang + 1);

  const sheetName = sheetPart.startsWith("'") && sheetPart.endsWith("'")
    ? sheetPart.slice(1, -1).replace(/''/g, "'")
    : sheetPart;

  const sheet = workbook.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
  if (!sheet) return null;

  const parsed = parseA1(cellPart.replace(/\$/g, ""));
  if (!parsed) return null;

  const cell = sheet.cells.get(parsed.a1);
  return cell && typeof cell.value === "number" ? cell.value : null;
}

function localRef(workbook: Workbook, currentSheetId: string, part: string): number | null {
  const parsed = parseA1(part.replace(/\$/g, ""));
  if (!parsed) return null;
  const cell = workbook.getCell(currentSheetId, parsed.a1);
  return typeof cell.value === "number" ? cell.value : null;
}

function parseA1(a1: string): { row: number; col: number; a1: string } | null {
  try {
    const { row0, col0 } = fromA1(a1);
    return { row: row0, col: col0, a1: toA1(row0, col0) };
  } catch {
    return null;
  }
}
