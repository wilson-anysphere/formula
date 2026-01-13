import type { CollabSession } from "@formula/collab-session";

export const READ_ONLY_SHEET_MUTATION_MESSAGE = "Read-only: you don’t have permission to modify sheets";

export type WorkbookMutationPermission = {
  allowed: boolean;
  reason?: string;
};

/**
 * Centralized permission guard for workbook-level mutations (sheet metadata/structure).
 *
 * In collab mode, viewer/commenter roles (or any session that reports `isReadOnly()`)
 * must not attempt to mutate workbook metadata locally, otherwise the UI can diverge
 * from the authoritative remote state.
 */
export function getWorkbookMutationPermission(session: CollabSession | null): WorkbookMutationPermission {
  if (!session) return { allowed: true };

  // Prefer the public CollabSession APIs (avoid peeking at private fields).
  if (!session.isReadOnly()) return { allowed: true };

  return { allowed: false, reason: READ_ONLY_SHEET_MUTATION_MESSAGE };
}
