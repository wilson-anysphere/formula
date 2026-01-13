import type { DataRegionSchema, SheetSchema, TableSchema } from "./schema.js";

export type RegionType = "table" | "dataRegion";

export type RegionRef = { type: RegionType; index: number };

export function scoreRegionForQuery(
  region: RegionRef | TableSchema | DataRegionSchema | null | undefined,
  schema: SheetSchema | null | undefined,
  query: string,
): number;

export function pickBestRegionForQuery(
  sheetSchema: SheetSchema | null | undefined,
  query: string,
): { type: RegionType; index: number; range: string } | null;
