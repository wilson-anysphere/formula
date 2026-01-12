export type FormulaRange = {
  sheet_id: string;
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
};

export class DocumentControllerBridge {
  constructor(doc: any, options?: { activeSheetId?: string });

  doc: any;
  activeSheetId: string;
  sheetIds: Set<string>;
  selection: FormulaRange;

  get_active_sheet_id(): string;
  get_sheet_id(params: { name: string }): string | null;
  create_sheet(params: { name: string; index?: number | null }): string;
  get_sheet_name(params: { sheet_id: string }): string;
  rename_sheet(params: { sheet_id: string; name: string }): null;

  get_selection(): FormulaRange;
  set_selection(params: { selection: FormulaRange }): null;

  get_range_values(params: { range: FormulaRange }): any[][];
  set_cell_value(params: { range: FormulaRange; value: any }): null;
  get_cell_formula(params: { range: FormulaRange }): string | null;
  set_cell_formula(params: { range: FormulaRange; formula: string }): null;
  set_range_values(params: { range: FormulaRange; values: any }): null;
  clear_range(params: { range: FormulaRange }): null;
  set_range_format(params: { range: FormulaRange; format: any }): null;
  get_range_format(params: { range: FormulaRange }): any;
}
