import type { CollabSession } from "@formula/collab-session";
import * as Y from "yjs";

import { getWorkbookMutationPermission, READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards";
import type { SheetVisibility } from "./workbookSheetStore";
import { findCollabSheetIndexById } from "./collabWorkbookSheetStore";

export type InsertCollabSheetResult =
  | { inserted: true; index: number }
  | { inserted: false; reason: string };

/**
 * Attempt to insert a new sheet metadata entry into the collab session's authoritative
 * `session.sheets` array.
 *
 * This helper enforces read-only session permissions so viewer/commenter roles cannot
 * mutate workbook structure (preventing local UI divergence).
 */
export function tryInsertCollabSheet(params: {
  session: CollabSession;
  sheetId: string;
  name: string;
  visibility?: SheetVisibility;
  insertAfterSheetId?: string | null;
}): InsertCollabSheetResult {
  const permission = getWorkbookMutationPermission(params.session);
  if (!permission.allowed) {
    return { inserted: false, reason: permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE };
  }

  const visibility: SheetVisibility = params.visibility ?? "visible";
  const afterId = String(params.insertAfterSheetId ?? "").trim();
  const sheetId = String(params.sheetId ?? "").trim();
  const name = String(params.name ?? "");

  let insertIndex = params.session.sheets.length;
  if (afterId) {
    const idx = findCollabSheetIndexById(params.session, afterId);
    if (idx >= 0) insertIndex = idx + 1;
  }

  params.session.transactLocal(() => {
    const sheet = new Y.Map<unknown>();
    sheet.set("id", sheetId);
    sheet.set("name", name);
    sheet.set("visibility", visibility);
    params.session.sheets.insert(insertIndex, [sheet as any]);
  });

  return { inserted: true, index: insertIndex };
}

