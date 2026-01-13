import type { DataRegionSchema, SheetSchema, TableSchema } from "./schema.js";

export function scoreRegionForQuery(
  region:
    | { type: "table" | "dataRegion"; index: number }
    | TableSchema
    | DataRegionSchema
    | null
    | undefined,
  schema: SheetSchema | null | undefined,
  query: string,
): number;

export function pickBestRegionForQuery(
  sheetSchema: SheetSchema | null | undefined,
  query: string,
): { type: "table" | "dataRegion"; index: number; range: string } | null;

