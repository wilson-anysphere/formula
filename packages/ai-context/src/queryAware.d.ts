import type { DataRegionSchema, SheetSchema, TableSchema } from "./schema.js";

export type RegionType = "table" | "dataRegion";
export type RegionRef = { type: RegionType; index: number };

/**
 * Score a table / data region for a given query.
 *
 * Higher is better. `0` means no match.
 */
export function scoreRegionForQuery(
  region: RegionRef | TableSchema | DataRegionSchema | null | undefined,
  schema: SheetSchema | null | undefined,
  query: string,
): number;

/**
 * Pick the best matching table / data region for a query.
 *
 * Returns null when no candidate receives a positive score.
 */
export function pickBestRegionForQuery(
  sheetSchema: SheetSchema | null | undefined,
  query: string,
): { type: RegionType; index: number; range: string } | null;

