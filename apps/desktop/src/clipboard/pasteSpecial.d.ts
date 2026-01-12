export type PasteSpecialMode = "all" | "values" | "formulas" | "formats";

export type PasteSpecialMenuItem = {
  mode: PasteSpecialMode;
  label: string;
};

export function getPasteSpecialMenuItems(): PasteSpecialMenuItem[];

