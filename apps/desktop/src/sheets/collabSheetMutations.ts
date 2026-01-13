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

  const sheetsArray: any = (params.session as any)?.sheets ?? null;
  const isYjsSheetsArray = Boolean(sheetsArray && typeof sheetsArray === "object" && sheetsArray.doc);

  params.session.transactLocal(() => {
    if (isYjsSheetsArray) {
      const sheet = new Y.Map<unknown>();
      // Attach first, then populate fields (Yjs types can warn when accessed before attachment).
      params.session.sheets.insert(insertIndex, [sheet as any]);
      sheet.set("id", sheetId);
      sheet.set("name", name);
      sheet.set("visibility", visibility);
    } else {
      params.session.sheets.insert(insertIndex, [{ id: sheetId, name, visibility } as any]);
    }
  });

  return { inserted: true, index: insertIndex };
}
