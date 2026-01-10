import { expandRange } from "../../grid/auditing-overlays/address.ts";
import type { CellAddress } from "../../grid/auditing-overlays/address.ts";
import type { DebugStep } from "./types.ts";

export function highlightsForStep(step: DebugStep): Set<CellAddress> {
  const out = new Set<CellAddress>();
  const ref = step.reference;
  if (!ref) return out;
  if (ref.type === "cell") {
    out.add(ref.cell);
  } else if (ref.type === "range") {
    for (const c of expandRange(ref.range)) out.add(c);
  }
  return out;
}

