export type CollaboratorListEntry = {
  /**
   * Stable identifier for diffing (e.g. `${id}:${clientId}`).
   */
  key: string;
  /**
   * Display name (already sanitized for UI).
   */
  name: string;
  /**
   * CSS color (matches cursor highlight color).
   */
  color: string;
  /**
   * Optional sheet name label to display (e.g. when the user is on a different sheet).
   */
  sheetName?: string | null;
};

export class CollaboratorsListUiController {
  constructor(opts: { container: HTMLElement; maxVisible?: number | null });
  destroy(): void;
  setCollaborators(collaborators: CollaboratorListEntry[]): void;
}

