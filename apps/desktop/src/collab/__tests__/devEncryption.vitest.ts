import { beforeAll, describe, expect, it } from "vitest";

import * as Y from "yjs";
import { createCollabSession, makeCellKey } from "@formula/collab-session";

import { resolveDevCollabEncryptionFromSearch } from "../devEncryption";

beforeAll(async () => {
  // `@formula/collab-encryption` relies on WebCrypto being present (Node 18+ provides this).
  // Provide a fallback for older Node runtimes used by some test environments.
  if (!globalThis.crypto?.subtle || !globalThis.crypto?.getRandomValues) {
    const { webcrypto } = await import("node:crypto");
    Object.defineProperty(globalThis, "crypto", { value: webcrypto });
  }
});

describe("dev collab encryption toggle", () => {
  it("encrypts cells in the configured range and masks reads without the key", async () => {
    const docId = "doc-dev-encryption";
    const doc = new Y.Doc();

    const encryption = resolveDevCollabEncryptionFromSearch({
      search: "?collabEncrypt=1&collabEncryptRange=Sheet1!A1:A1",
      docId,
      defaultSheetId: "Sheet1",
    });
    expect(encryption).not.toBeNull();

    const sessionWithKey = createCollabSession({ docId, doc, encryption: encryption! });
    const a1 = makeCellKey({ sheetId: "Sheet1", row: 0, col: 0 });
    const b1 = makeCellKey({ sheetId: "Sheet1", row: 0, col: 1 });

    await sessionWithKey.setCellValue(a1, "secret");
    await sessionWithKey.setCellValue(b1, "public");

    const yA1 = sessionWithKey.cells.get(a1) as any;
    expect(yA1?.get("enc")).toBeTruthy();
    expect(yA1?.get("value")).toBeUndefined();
    expect(yA1?.get("formula")).toBeUndefined();

    const yB1 = sessionWithKey.cells.get(b1) as any;
    expect(yB1?.get("enc")).toBeUndefined();
    expect(yB1?.get("value")).toBe("public");

    const sessionWithoutKey = createCollabSession({ docId, doc });
    const masked = await sessionWithoutKey.getCell(a1);
    expect(masked).not.toBeNull();
    expect(masked?.value).toBe("###");
    expect(masked?.formula).toBeNull();
    expect(masked?.encrypted).toBe(true);

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
  });
});
