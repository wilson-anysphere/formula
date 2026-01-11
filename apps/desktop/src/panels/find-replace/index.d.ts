export class FindReplaceController {
  constructor(params: any);
  query: string;
  replacement: string;
  scope: string;
  lookIn: string;
  valueMode: string;
  matchCase: boolean;
  matchEntireCell: boolean;
  useWildcards: boolean;
  searchOrder: string;

  findNext(): Promise<any>;
  findAll(): Promise<any[]>;
  replaceNext(): Promise<any>;
  replaceAll(): Promise<any>;
}

export function registerFindReplaceShortcuts(params: any): any;

