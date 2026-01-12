export type CellStyle = Record<string, any>;

export function applyStylePatch(base: CellStyle, patch: CellStyle | null): CellStyle;

export class StyleTable {
  intern(style: CellStyle | null | undefined): number;
  get(styleId: number): CellStyle;
  get size(): number;
}

