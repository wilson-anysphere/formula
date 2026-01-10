import type { CellAddress } from "./address.ts";

export type AuditingMode = "precedents" | "dependents" | "both";

export interface AuditingEngine {
  precedents(cell: CellAddress, opts?: { transitive?: boolean }): Iterable<CellAddress>;
  dependents(cell: CellAddress, opts?: { transitive?: boolean }): Iterable<CellAddress>;
}

export interface AuditingOverlays {
  precedents: Set<CellAddress>;
  dependents: Set<CellAddress>;
}

export function computeAuditingOverlays(
  engine: AuditingEngine,
  cell: CellAddress,
  mode: AuditingMode = "both",
  opts: { transitive?: boolean } = {},
): AuditingOverlays {
  const precedents = new Set<CellAddress>();
  const dependents = new Set<CellAddress>();

  if (mode === "precedents" || mode === "both") {
    for (const p of engine.precedents(cell, opts)) precedents.add(p);
  }
  if (mode === "dependents" || mode === "both") {
    for (const d of engine.dependents(cell, opts)) dependents.add(d);
  }

  // Never highlight the selected cell as its own precedent/dependent.
  precedents.delete(cell);
  dependents.delete(cell);

  return { precedents, dependents };
}

