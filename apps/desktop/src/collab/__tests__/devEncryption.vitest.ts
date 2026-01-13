import { beforeAll, describe, expect, it } from "vitest";

import * as Y from "yjs";
import { createCollabSession, makeCellKey } from "@formula/collab-session";

import { resolveDevCollabEncryptionFromSearch } from "../devEncryption";
import { DocumentController } from "../../document/documentController.js";
import { bindDocumentControllerWithCollabUndo } from "../documentControllerCollabUndo";

beforeAll(async () => {
  // `@formula/collab-encryption` relies on WebCrypto being present (Node 18+ provides this).
  // Provide a fallback for older Node runtimes used by some test environments.
  if (!globalThis.crypto?.subtle || !globalThis.crypto?.getRandomValues) {
    const { webcrypto } = await import("node:crypto");
    Object.defineProperty(globalThis, "crypto", { value: webcrypto });
  }
});

async function flushBinderWork(): Promise<void> {
  // The Yjs↔DocumentController binder serializes work through promise chains.
  // Awaiting a couple ticks ensures both the DocumentController→Yjs write chain
  // and the Yjs→DocumentController apply chain have a chance to run.
  await new Promise<void>((resolve) => setImmediate(resolve));
  await new Promise<void>((resolve) => setImmediate(resolve));
}

describe("dev collab encryption toggle", () => {
  it("encrypts cells in the configured range and masks reads without the key", async () => {
    const docId = "doc-dev-encryption";
    const doc = new Y.Doc({ guid: docId });

    const encryption = resolveDevCollabEncryptionFromSearch({
      search: "?collabEncrypt=1&collabEncryptRange=Sheet1!A1:A2",
      docId,
      defaultSheetId: "Sheet1",
    });
    expect(encryption).not.toBeNull();

    const sessionWithKey = createCollabSession({ docId, doc, encryption: encryption! });
    const a1 = makeCellKey({ sheetId: "Sheet1", row: 0, col: 0 });
    const a2 = makeCellKey({ sheetId: "Sheet1", row: 1, col: 0 });
    const b1 = makeCellKey({ sheetId: "Sheet1", row: 0, col: 1 });

    await sessionWithKey.setCellValue(a1, "secret");
    await sessionWithKey.setCellFormula(a2, "=SUM(1,2)");
    await sessionWithKey.setCellValue(b1, "public");

    const yA1 = sessionWithKey.cells.get(a1) as any;
    expect(yA1?.get("enc")).toBeTruthy();
    expect(yA1?.get("value")).toBeUndefined();
    expect(yA1?.get("formula")).toBeUndefined();

    const yB1 = sessionWithKey.cells.get(b1) as any;
    expect(yB1?.get("enc")).toBeUndefined();
    expect(yB1?.get("value")).toBe("public");

    const yA2 = sessionWithKey.cells.get(a2) as any;
    expect(yA2?.get("enc")).toBeTruthy();
    expect(yA2?.get("value")).toBeUndefined();
    expect(yA2?.get("formula")).toBeUndefined();

    const sessionWithoutKey = createCollabSession({ docId, doc });
    const masked = await sessionWithoutKey.getCell(a1);
    expect(masked).not.toBeNull();
    expect(masked?.value).toBe("###");
    expect(masked?.formula).toBeNull();
    expect(masked?.encrypted).toBe(true);

    const maskedFormula = await sessionWithoutKey.getCell(a2);
    expect(maskedFormula).not.toBeNull();
    expect(maskedFormula?.value).toBe("###");
    expect(maskedFormula?.formula).toBeNull();
    expect(maskedFormula?.encrypted).toBe(true);

    const outsideRange = await sessionWithoutKey.getCell(b1);
    expect(outsideRange).not.toBeNull();
    expect(outsideRange?.value).toBe("public");
    expect(outsideRange?.formula).toBeNull();

    // The dev helper should remain able to *decrypt* already-encrypted cells even if the
    // demo encryption range is later changed (writes are range-restricted via shouldEncryptCell).
    const encryptionDifferentRange = resolveDevCollabEncryptionFromSearch({
      search: "?collabEncrypt=1&collabEncryptRange=Sheet1!B1:B1",
      docId,
      defaultSheetId: "Sheet1",
    });
    expect(encryptionDifferentRange).not.toBeNull();

    const sessionDifferentRange = createCollabSession({ docId, doc, encryption: encryptionDifferentRange! });
    const decrypted = await sessionDifferentRange.getCell(a1);
    expect(decrypted).not.toBeNull();
    expect(decrypted?.value).toBe("secret");

    const decryptedFormula = await sessionDifferentRange.getCell(a2);
    expect(decryptedFormula).not.toBeNull();
    expect(decryptedFormula?.formula).toBe("=SUM(1,2)");
  });

  it("can resolve a sheet display name to a stable sheet id when provided a resolver", async () => {
    const docId = "doc-dev-encryption-sheetname";
    const doc = new Y.Doc({ guid: docId });

    const encryption = resolveDevCollabEncryptionFromSearch({
      search: "?collabEncrypt=1&collabEncryptRange=Display!A1:A1",
      docId,
      defaultSheetId: "s1",
      resolveSheetIdByName: (name) => (name === "Display" ? "s1" : null),
    });
    expect(encryption).not.toBeNull();

    const sessionWithKey = createCollabSession({
      docId,
      doc,
      encryption: encryption!,
      schema: { defaultSheetId: "s1", defaultSheetName: "Display" },
    });

    const a1 = makeCellKey({ sheetId: "s1", row: 0, col: 0 });
    await sessionWithKey.setCellValue(a1, "secret");

    const yA1 = sessionWithKey.cells.get(a1) as any;
    expect(yA1?.get("enc")).toBeTruthy();
    expect(yA1?.get("value")).toBeUndefined();
    expect(yA1?.get("formula")).toBeUndefined();

    const sessionWithoutKey = createCollabSession({
      docId,
      doc,
      schema: { defaultSheetId: "s1", defaultSheetName: "Display" },
    });
    const masked = await sessionWithoutKey.getCell(a1);
    expect(masked).not.toBeNull();
    expect(masked?.value).toBe("###");
    expect(masked?.encrypted).toBe(true);
  });

  it("encrypts DocumentController-driven edits via the binder", async () => {
    const docId = "doc-dev-encryption-binder";
    const doc = new Y.Doc({ guid: docId });

    const encryption = resolveDevCollabEncryptionFromSearch({
      search: "?collabEncrypt=1&collabEncryptRange=Sheet1!A1:A1",
      docId,
      defaultSheetId: "Sheet1",
    });
    expect(encryption).not.toBeNull();

    const sessionWithKey = createCollabSession({ docId, doc, encryption: encryption! });
    const documentWithKey = new DocumentController();
    const { binder } = await bindDocumentControllerWithCollabUndo({
      session: sessionWithKey,
      documentController: documentWithKey,
      defaultSheetId: "Sheet1",
    });

    // Binder writes into Yjs can be async (encryption via WebCrypto). Wait for a Yjs update so
    // assertions don't race the encryption/write chain.
    const whenYjsUpdated = new Promise<void>((resolve) => {
      const onUpdate = () => {
        // @ts-expect-error - Yjs typings are looser in some environments.
        doc.off("update", onUpdate);
        resolve();
      };
      // @ts-expect-error - Yjs typings are looser in some environments.
      doc.on("update", onUpdate);
    });

    documentWithKey.setCellValue("Sheet1", { row: 0, col: 0 }, "secret");
    await whenYjsUpdated;
    await flushBinderWork();

    const cellKey = makeCellKey({ sheetId: "Sheet1", row: 0, col: 0 });
    const yCell = sessionWithKey.cells.get(cellKey) as any;
    expect(yCell?.get("enc")).toBeTruthy();
    expect(yCell?.get("value")).toBeUndefined();
    expect(yCell?.get("formula")).toBeUndefined();

    binder.destroy();

    const sessionWithoutKey = createCollabSession({ docId, doc });
    const documentWithoutKey = new DocumentController();
    const { binder: binder2 } = await bindDocumentControllerWithCollabUndo({
      session: sessionWithoutKey,
      documentController: documentWithoutKey,
      defaultSheetId: "Sheet1",
    });
    await flushBinderWork();
    expect(documentWithoutKey.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("###");

    // Attempting a local edit without the key should be rejected by the binder-installed
    // canEditCell guard (so we never write plaintext into an encrypted cell).
    const encBefore = yCell?.get("enc");
    documentWithoutKey.setCellValue("Sheet1", { row: 0, col: 0 }, "hacked");
    await flushBinderWork();
    expect(documentWithoutKey.getCell("Sheet1", { row: 0, col: 0 }).value).toBe("###");
    const yCellAfter = sessionWithKey.cells.get(cellKey) as any;
    expect(yCellAfter?.get("enc")).toEqual(encBefore);
    expect(yCellAfter?.get("value")).toBeUndefined();
    expect(yCellAfter?.get("formula")).toBeUndefined();
    binder2.destroy();
  });
});
