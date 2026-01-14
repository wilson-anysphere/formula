import type { CollabSession } from "@formula/collab-session";

import { getWorkbookMutationPermission, READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards";
import { rewriteDocumentFormulasForSheetDelete } from "./sheetFormulaRewrite";
import { listSheetsFromCollabSession } from "./collabWorkbookSheetStore";
import { tryInsertCollabSheet } from "./collabSheetMutations";
import { generateDefaultSheetName, type WorkbookSheetStore } from "./workbookSheetStore";

export type ToastFn = (message: string, kind?: any, options?: any) => void;
export type ConfirmFn = (message: string) => Promise<boolean>;

export type SheetCommandsAppLike = {
  getCurrentSheetId(): string;
  activateSheet(sheetId: string): void;
  getDocument(): any;
  getCollabSession?(): CollabSession | null;
};

export function createAddSheetCommand(params: {
  app: SheetCommandsAppLike;
  getWorkbookSheetStore: () => WorkbookSheetStore;
  restoreFocusAfterSheetNavigation: () => void;
  showToast: ToastFn;
}): () => Promise<void> {
  const { app, getWorkbookSheetStore, restoreFocusAfterSheetNavigation, showToast } = params;

  let addSheetInFlight = false;

  return async function addSheet(): Promise<void> {
    if (addSheetInFlight) return;
    addSheetInFlight = true;
    try {
      const store = getWorkbookSheetStore();
      const activeId = app.getCurrentSheetId();
      const allSheets = store.listAll();
      const desiredName = generateDefaultSheetName(allSheets);
      const doc = app.getDocument();

      const collabSession = app.getCollabSession?.() ?? null;
      if (collabSession) {
        // In collab mode, the Yjs `session.sheets` array is the authoritative sheet list.
        // Create the new sheet by updating that metadata so it propagates to other clients.
        const existing = listSheetsFromCollabSession(collabSession);
        const existingIds = new Set(existing.map((sheet) => sheet.id));

        const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
        const generateId = () => {
          const uuid = typeof randomUuid === "function" ? randomUuid.call((globalThis as any).crypto) : null;
          return uuid ? `sheet_${uuid}` : `sheet_${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
        };

        let id = generateId();
        for (let i = 0; i < 10 && existingIds.has(id); i += 1) {
          id = generateId();
        }
        while (existingIds.has(id)) {
          id = `${id}_${Math.random().toString(16).slice(2)}`;
        }

        const inserted = tryInsertCollabSheet({
          session: collabSession,
          sheetId: id,
          name: desiredName,
          visibility: "visible",
          insertAfterSheetId: activeId,
        });
        if (!inserted.inserted) {
          showToast(inserted.reason, "error");
          return;
        }

        // DocumentController creates sheets lazily; touching any cell ensures the sheet exists.
        doc.getCell(id, { row: 0, col: 0 });
        app.activateSheet(id);
        return;
      }

      // In local (non-collab) mode, the UI sheet store is the authoritative sheet list.
      // Mutate it first so sheet operations remain undoable in the DocumentController.
      // The workbook sync bridge will persist the structural change to the native backend.
      const existingIdCi = new Set(allSheets.map((s) => s.id.trim().toLowerCase()));
      const baseId = desiredName;
      let newSheetId = baseId;
      let counter = 1;
      while (existingIdCi.has(newSheetId.toLowerCase())) {
        counter += 1;
        newSheetId = `${baseId}-${counter}`;
      }

      store.addAfter(activeId, { id: newSheetId, name: desiredName });

      // Best-effort: ensure the sheet is materialized (DocumentController can create sheets lazily).
      try {
        doc.getCell(newSheetId, { row: 0, col: 0 });
      } catch {
        // ignore
      }
      app.activateSheet(newSheetId);
    } catch (err) {
      showToast(`Failed to add sheet: ${String((err as any)?.message ?? err)}`, "error");
    } finally {
      // Ribbon dropdown menu items restore focus to the trigger button after dispatching the command.
      // Defer grid focus so it wins over that built-in focus restoration (Excel-like).
      try {
        if (typeof queueMicrotask === "function") queueMicrotask(() => restoreFocusAfterSheetNavigation());
        else restoreFocusAfterSheetNavigation();
      } catch {
        // ignore focus errors
      }
      addSheetInFlight = false;
    }
  };
}

export function createDeleteActiveSheetCommand(params: {
  app: SheetCommandsAppLike;
  getWorkbookSheetStore: () => WorkbookSheetStore;
  restoreFocusAfterSheetNavigation: () => void;
  showToast: ToastFn;
  confirm: ConfirmFn;
}): () => Promise<void> {
  const { app, getWorkbookSheetStore, restoreFocusAfterSheetNavigation, showToast, confirm } = params;

  let deleteInFlight = false;

  return async function deleteActiveSheet(): Promise<void> {
    if (deleteInFlight) return;
    deleteInFlight = true;
    try {
      const store = getWorkbookSheetStore();
      const activeId = app.getCurrentSheetId();
      const sheet = store.getById(activeId);
      if (!sheet) {
        showToast("Failed to delete sheet: active sheet not found.", "error");
        return;
      }

      const collabSession = app.getCollabSession?.() ?? null;
      if (collabSession) {
        const permission = getWorkbookMutationPermission(collabSession);
        if (!permission.allowed) {
          showToast(permission.reason ?? READ_ONLY_SHEET_MUTATION_MESSAGE, "error");
          return;
        }
      }

      let ok = false;
      try {
        ok = await confirm(`Delete sheet "${sheet.name}"?`);
      } catch {
        ok = false;
      }
      if (!ok) return;

      const deletedName = sheet.name;
      const sheetOrder = store.listAll().map((s) => s.name);
      const nextActiveId = (() => {
        // Mirror Excel: when deleting the active sheet, activate the next visible sheet to the right
        // if possible; otherwise fall back to the previous visible sheet.
        const allSheets = store.listAll();
        const visibleSheets = allSheets.filter((s) => s.visibility === "visible");
        const idx = visibleSheets.findIndex((s) => s.id === activeId);
        if (idx === -1) return null;
        return visibleSheets[idx + 1]?.id ?? visibleSheets[idx - 1]?.id ?? null;
      })();

      try {
        // In local mode, this routes the sheet delete through the existing sheet-store -> DocumentController
        // subscription so the delete is undoable.
        store.remove(activeId);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        showToast(message, "error");
        return;
      }
      // Defensive: collab sheet stores can no-op deletes when the session is read-only. Ensure the
      // sheet is actually gone before proceeding with formula rewrites.
      if (store.getById(activeId)) {
        showToast(READ_ONLY_SHEET_MUTATION_MESSAGE, "error");
        return;
      }

      // Rewrite formulas referencing the deleted sheet name (Excel-like behavior).
      // Important: do this synchronously right after `store.remove(...)` so it lands in the same
      // DocumentController batch opened by the store subscription.
      try {
        rewriteDocumentFormulasForSheetDelete(app.getDocument() as any, deletedName, sheetOrder);
      } catch (err) {
        showToast(`Failed to update formulas after delete: ${String((err as any)?.message ?? err)}`, "error");
      }

      // Switch away from the deleted sheet.
      //
      // Note: other parts of the desktop UI (e.g. `main.ts` sheet-store subscription) may
      // proactively activate a fallback sheet before the deletion is fully processed to avoid
      // DocumentController listeners re-creating the deleted sheet. We still apply the Excel-like
      // "next visible sheet to the right, otherwise left" policy here to ensure the final active
      // sheet selection is deterministic across all delete entrypoints (ribbon, tabs, dialogs).
      const fallback = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
      const next = nextActiveId ?? fallback;
      if (next && next !== activeId && app.getCurrentSheetId() !== next) {
        app.activateSheet(next);
      }
    } finally {
      restoreFocusAfterSheetNavigation();
      deleteInFlight = false;
    }
  };
}
