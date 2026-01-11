declare module "../../../packages/collab/permissions/index.js" {
  export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

  export type CellAddress = {
    sheetId?: string;
    sheetName?: string;
    row: number;
    col: number;
  };

  export function normalizeRestriction(restriction: unknown): unknown;

  export function getCellPermissions(args: {
    role: DocumentRole;
    restrictions?: unknown[] | null;
    userId: string | null;
    cell: CellAddress;
  }): { canRead: boolean; canEdit: boolean };
}

