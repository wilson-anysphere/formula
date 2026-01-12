import { parseA1Range, splitSheetQualifier } from "../../../../packages/search/index.js";

export type ParsedSheetQualifiedRange = {
  sheetName: string | null;
  /**
   * The unqualified A1 reference portion (e.g. `A1:B2`).
   */
  ref: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

export function parseSheetQualifiedA1Range(input: string): ParsedSheetQualifiedRange {
  const { sheetName, ref } = splitSheetQualifier(input);
  const range = parseA1Range(ref);
  return { sheetName, ref, ...range };
}

export type SheetNameLookup = {
  getName(sheetId: string): string | undefined;
};

export function getSheetDisplayName(sheetId: string, sheetStore: SheetNameLookup): string {
  return sheetStore.getName(sheetId) ?? sheetId;
}

export function buildSheetNameToIdMap(sheetIds: string[], sheetStore: SheetNameLookup): Map<string, string> {
  const out = new Map<string, string>();
  for (const sheetId of sheetIds) {
    const name = getSheetDisplayName(sheetId, sheetStore);
    const existing = out.get(name);
    if (existing && existing !== sheetId) {
      throw new Error(`Duplicate sheet name: ${name}`);
    }
    out.set(name, sheetId);
  }
  return out;
}

export function resolveSheetIdByName(args: {
  sheetName: string;
  sheetIds: string[];
  sheetStore: SheetNameLookup;
}): string {
  const sheetName = String(args.sheetName ?? "").trim();
  if (!sheetName) {
    throw new Error("Sheet name must be a non-empty string");
  }
  const map = buildSheetNameToIdMap(args.sheetIds, args.sheetStore);
  const id = map.get(sheetName);
  if (!id) {
    throw new Error(`Unknown sheet: ${sheetName}`);
  }
  return id;
}
