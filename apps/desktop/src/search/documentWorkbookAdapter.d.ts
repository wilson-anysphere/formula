export class DocumentWorkbookAdapter {
  constructor(params: { document: any; sheetNameResolver?: any });

  document: any;
  sheetNameResolver: any;
  schemaVersion: number;
  names: Map<string, any>;
  tables: Map<string, any>;

  get sheets(): any[];
  getSheet(sheetName: string): any;

  defineName(name: string, ref: any): void;
  getName(name: string): any | null;

  addTable(table: any): void;
  getTable(name: string): any | null;

  clearSchema(): void;
  dispose(): void;
}
