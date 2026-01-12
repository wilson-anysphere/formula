import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { createCollabSession } from "@formula/collab-session";
import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { DocumentController } from "../document/documentController.js";
import { bindDocumentControllerWithCollabUndo } from "./documentControllerCollabUndo";

async function flushBinderWork(): Promise<void> {
  // The Yjs↔DocumentController binder serializes work through promise chains.
  // Awaiting a couple ticks ensures both the DocumentController→Yjs write chain
  // and the Yjs→DocumentController apply chain have a chance to run.
  await new Promise<void>((resolve) => setImmediate(resolve));
  await new Promise<void>((resolve) => setImmediate(resolve));
}

describe("collaboration-safe undo/redo (desktop)", () => {
  it("undo/redo updates the DocumentController via the binder", async () => {
    const session = createCollabSession({ doc: new Y.Doc() });
    const document = new DocumentController();

    const { binder, undoService } = await bindDocumentControllerWithCollabUndo({
      session,
      documentController: document,
      defaultSheetId: "Sheet1",
    });

    document.setCellValue("Sheet1", { row: 0, col: 0 }, "local");
    await flushBinderWork();
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("local");

    undoService.undo();
    await flushBinderWork();
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(null);

    undoService.redo();
    await flushBinderWork();
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("local");

    binder.destroy();
  });

  it("local edit → remote overwrite → local undo does not overwrite remote value", async () => {
    const session = createCollabSession({ doc: new Y.Doc() });
    const document = new DocumentController();

    const { binder, undoService } = await bindDocumentControllerWithCollabUndo({
      session,
      documentController: document,
      defaultSheetId: "Sheet1",
    });

    // Two local edits so we can verify undo does *not* skip past an un-undoable change.
    document.setCellValue("Sheet1", { row: 0, col: 0 }, "local-a1");
    await flushBinderWork();
    undoService.stopCapturing();

    document.setCellValue("Sheet1", { row: 0, col: 1 }, "local-b1");
    await flushBinderWork();
    undoService.stopCapturing();

    // Remote overwrite of the *most recent* local edit (B1).
    session.doc.transact(() => {
      const ycell = new Y.Map();
      ycell.set("value", "remote-b1");
      session.cells.set("Sheet1:0:1", ycell);
    }, REMOTE_ORIGIN);
    await flushBinderWork();

    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("local-a1");
    expect(document.getCell("Sheet1", { row: 0, col: 1 }).value).toBe("remote-b1");

    // Undo should NOT:
    // - overwrite the remote value for B1
    // - skip past the un-undoable B1 edit and undo A1 instead
    undoService.undo();
    await flushBinderWork();
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("local-a1");
    expect(document.getCell("Sheet1", { row: 0, col: 1 }).value).toBe("remote-b1");

    binder.destroy();
  });
});
