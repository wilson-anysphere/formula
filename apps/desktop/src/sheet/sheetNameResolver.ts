export type SheetNameResolver = {
  /**
   * Resolve a stable sheet id to the user-facing display name.
   *
   * Return null when unknown.
   */
  getSheetNameById(id: string): string | null;
  /**
   * Resolve a user-facing display name to a stable sheet id.
   *
   * Implementations must treat lookup as case-insensitive.
   * Return null when unknown.
   */
  getSheetIdByName(name: string): string | null;
};

