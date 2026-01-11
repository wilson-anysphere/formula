export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

export const DOCUMENT_ROLES: ReadonlyArray<DocumentRole>;

export function assertValidRole(role: string): asserts role is DocumentRole;

export const ROLE_CAPABILITIES: Readonly<
  Record<DocumentRole, { read: boolean; edit: boolean; comment: boolean; share: boolean }>
>;

export function roleCanRead(role: DocumentRole): boolean;
export function roleCanEdit(role: DocumentRole): boolean;
export function roleCanComment(role: DocumentRole): boolean;
export function roleCanShare(role: DocumentRole): boolean;

export type RangeLike = {
  sheetId?: string;
  sheetName?: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export type NormalizedRange = {
  sheetId: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export function normalizeRange(range: RangeLike): NormalizedRange;

export type CellLike = {
  sheetId?: string;
  sheetName?: string;
  row: number;
  col: number;
};

export function cellInRange(cell: { sheetId: string; row: number; col: number }, range: NormalizedRange): boolean;

export type RestrictionLike =
  | (RangeLike & {
      id?: string;
      createdAt?: string | number | Date;
      readAllowlist?: string[];
      editAllowlist?: string[];
    })
  | {
      id?: string;
      createdAt?: string | number | Date;
      range: RangeLike;
      readAllowlist?: string[];
      editAllowlist?: string[];
    };

export type NormalizedRestriction = {
  id: string;
  range: NormalizedRange;
  readAllowlist?: string[];
  editAllowlist?: string[];
  createdAt: Date;
};

export function normalizeRestriction(restriction: RestrictionLike): NormalizedRestriction;

export function getCellPermissions(args: {
  role: DocumentRole;
  restrictions?: unknown[] | null;
  userId: string | null;
  cell: CellLike;
}): { canRead: boolean; canEdit: boolean };

export function maskCellValue(value: unknown): string;

export function maskValueIfUnreadable(args: { value: unknown; canRead: boolean }): unknown;

export function maskCellUpdatesForUser(args: {
  role: DocumentRole;
  restrictions?: unknown[] | null;
  userId: string | null;
  updates: Array<{ cell: CellLike; value: unknown; [key: string]: unknown }>;
}): Array<{ cell: CellLike; value: unknown; [key: string]: unknown }>;

