import { describe, expect, test } from "vitest";
import * as Y from "yjs";

import { createCollabSession, makeCellKey, type CollabSession } from "@formula/collab-session";
import { bytesToBase64 } from "@formula/collab-encryption";
import { createEncryptedRangeManagerForSession, createEncryptionPolicyFromDoc } from "@formula/collab-encrypted-ranges";

import { CollabEncryptionKeyStore } from "../../encryptionKeyStore";

describe("encrypted ranges (integration)", () => {
  test("encrypt range -> set cell -> writes enc payload; missing-key sessions see masked + cannot edit", async () => {
    const docId = "doc-e2e";
    const ydoc = new Y.Doc({ guid: docId });
    const policy = createEncryptionPolicyFromDoc(ydoc);

    const keyStore1 = new CollabEncryptionKeyStore({ invoke: null });
    const keyStore2 = new CollabEncryptionKeyStore({ invoke: null });

    const session1: CollabSession = createCollabSession({
      docId,
      doc: ydoc,
      encryption: {
        shouldEncryptCell: (cell) => policy.shouldEncryptCell(cell),
        keyForCell: (cell) => {
          const keyId = policy.keyIdForCell(cell);
          if (!keyId) return null;
          return keyStore1.getCachedKey(docId, keyId);
        },
      },
    });
    const session2: CollabSession = createCollabSession({
      docId,
      doc: ydoc,
      encryption: {
        shouldEncryptCell: (cell) => policy.shouldEncryptCell(cell),
        keyForCell: (cell) => {
          const keyId = policy.keyIdForCell(cell);
          if (!keyId) return null;
          return keyStore2.getCachedKey(docId, keyId);
        },
      },
    });

    const manager1 = createEncryptedRangeManagerForSession(session1);

    const keyId = "k1";
    const keyBytes = new Uint8Array(32);
    keyBytes.fill(42);
    await keyStore1.set(docId, keyId, bytesToBase64(keyBytes));

    manager1.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId });

    const cell = { sheetId: "Sheet1", row: 0, col: 0 };
    await session1.setCellValue(makeCellKey(cell), "secret");

    const raw = session1.cells.get(makeCellKey(cell));
    expect(raw && typeof (raw as any).get === "function").toBe(true);
    expect((raw as any).get("enc")).toBeTruthy();
    expect((raw as any).get("value")).toBeUndefined();
    expect((raw as any).get("formula")).toBeUndefined();

    // Session with the key can decrypt.
    await expect(session1.getCell(makeCellKey(cell))).resolves.toMatchObject({ value: "secret", formula: null });

    // Session without the key sees masked content and cannot edit.
    await expect(session2.getCell(makeCellKey(cell))).resolves.toMatchObject({ value: "###", encrypted: true });
    expect(session2.canEditCell(cell)).toBe(false);
  });
});
