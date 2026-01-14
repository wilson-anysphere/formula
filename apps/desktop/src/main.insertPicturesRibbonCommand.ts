import type { SpreadsheetApp } from "./app/spreadsheetApp";
import { showToast } from "./extensions/ui.js";
import { pickLocalImageFiles } from "./drawings/pickLocalImageFiles.js";
import { showCollabEditRejectedToast } from "./collab/editRejectionToast";

export type InsertPicturesRibbonCommandApp = Pick<SpreadsheetApp, "insertPicturesFromFiles" | "focus">;

/**
 * Handles the Excel-style Ribbon "Insert → Pictures" commands.
 *
 * This is factored out of `main.ts` so it can be unit tested without importing the full app entrypoint.
 */
export async function handleInsertPicturesRibbonCommand(commandId: string, app: InsertPicturesRibbonCommandApp): Promise<boolean> {
  if (commandId === "insert.illustrations.pictures.thisDevice" || commandId === "insert.illustrations.pictures") {
    try {
      // Match SpreadsheetApp guards: don't open a native/file picker while the user is actively editing.
      // Ribbon buttons should be disabled while editing, but keep this as a defensive check so
      // command palette / programmatic execution doesn't unexpectedly steal focus.
      //
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const appAny = app as any;
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      const primaryEditing = typeof appAny?.isEditing === "function" && appAny.isEditing() === true;
      if (primaryEditing || globalEditing === true) {
        return true;
      }

      // Picture insertion mutates the workbook (sheet drawings + embedded image bytes). Block it in
      // collab read-only sessions (viewer/commenter) so the local UI doesn't diverge from the shared
      // document state.
      //
      // Guard early so we don't open a file picker only to reject the insertion after selection.
      if (typeof appAny?.isReadOnly === "function" && appAny.isReadOnly() === true) {
        showCollabEditRejectedToast([{ rejectionKind: "insertPictures", rejectionReason: "permission" }]);
        try {
          app.focus();
        } catch {
          // ignore
        }
        return true;
      }
    } catch {
      // ignore (best-effort guard only)
    }

    try {
      const files = await pickLocalImageFiles({ multiple: true });
      if (files.length > 0) {
        await app.insertPicturesFromFiles(files);
      }
      app.focus();
    } catch (err) {
      console.error("Failed to insert picture:", err);
      try {
        showToast(`Failed to insert picture: ${String((err as any)?.message ?? err)}`, "error");
      } catch {
        // `showToast` requires a #toast-root; ignore in headless contexts/tests.
      }
      app.focus();
    }
    return true;
  }

  if (commandId === "insert.illustrations.pictures.stockImages") {
    try {
      showToast("Stock Images not implemented yet");
    } catch {
      // ignore
    }
    app.focus();
    return true;
  }

  if (commandId === "insert.illustrations.pictures.onlinePictures" || commandId === "insert.illustrations.onlinePictures") {
    try {
      showToast("Online Pictures not implemented yet");
    } catch {
      // ignore
    }
    app.focus();
    return true;
  }

  return false;
}
