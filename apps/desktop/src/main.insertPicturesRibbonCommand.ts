import type { SpreadsheetApp } from "./app/spreadsheetApp";
import { showToast } from "./extensions/ui.js";
import { pickLocalImageFiles } from "./drawings/pickLocalImageFiles.js";

export type InsertPicturesRibbonCommandApp = Pick<SpreadsheetApp, "insertPicturesFromFiles" | "focus">;

/**
 * Handles the Excel-style Ribbon "Insert → Pictures" commands.
 *
 * This is factored out of `main.ts` so it can be unit tested without importing the full app entrypoint.
 */
export async function handleInsertPicturesRibbonCommand(commandId: string, app: InsertPicturesRibbonCommandApp): Promise<boolean> {
  if (commandId === "insert.illustrations.pictures.thisDevice" || commandId === "insert.illustrations.pictures") {
    try {
      const files = await pickLocalImageFiles({ multiple: true });
      if (files.length > 0) {
        await app.insertPicturesFromFiles(files);
      }
      app.focus();
    } catch (err) {
      console.error("Failed to insert picture:", err);
      showToast(`Failed to insert picture: ${String((err as any)?.message ?? err)}`, "error");
      app.focus();
    }
    return true;
  }

  if (commandId === "insert.illustrations.pictures.stockImages") {
    showToast("Stock Images not implemented yet");
    app.focus();
    return true;
  }

  if (commandId === "insert.illustrations.pictures.onlinePictures" || commandId === "insert.illustrations.onlinePictures") {
    showToast("Online Pictures not implemented yet");
    app.focus();
    return true;
  }

  return false;
}

