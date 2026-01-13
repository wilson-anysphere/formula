import { describe, expect, test } from "vitest";
import * as Y from "yjs";

import { createCollabSession, makeCellKey, type CollabSession } from "@formula/collab-session";
import { bytesToBase64 } from "@formula/collab-encryption";

import { CollabEncryptionKeyStore } from "../../encryptionKeyStore";
import { EncryptedRangeManager } from "../encryptedRangeManager";

function encryptedKeyIdForCell(metadata: any, cell: { sheetId: string; row: number; col: number }): string | null {
  if (!metadata) return null;
  const raw = typeof metadata.get === "function" ? metadata.get("encryptedRanges") : (metadata as any)?.encryptedRanges;
  const entries = Array.isArray(raw) ? raw : typeof raw?.toArray === "function" ? raw.toArray() : [];
  for (let i = entries.length - 1; i >= 0; i -= 1) {
    const entry = entries[i];
    if (!entry || typeof entry !== "object") continue;
    const sheetId = String((entry as any).sheetId ?? "").trim();
    if (!sheetId || sheetId !== cell.sheetId) continue;
    const keyId = String((entry as any).keyId ?? "").trim();
    if (!keyId) continue;
    const startRow = Number((entry as any).startRow);
    const startCol = Number((entry as any).startCol);
    const endRow = Number((entry as any).endRow);
    const endCol = Number((entry as any).endCol);
    if (!Number.isInteger(startRow) || !Number.isInteger(startCol) || !Number.isInteger(endRow) || !Number.isInteger(endCol)) continue;
    if (cell.row < Math.min(startRow, endRow) || cell.row > Math.max(startRow, endRow)) continue;
    if (cell.col < Math.min(startCol, endCol) || cell.col > Math.max(startCol, endCol)) continue;
    return keyId;
  }
  return null;
}

describe("encrypted ranges (integration)", () => {
  test("encrypt range -> set cell -> writes enc payload; missing-key sessions see masked + cannot edit", async () => {
    const docId = "doc-e2e";
    const ydoc = new Y.Doc({ guid: docId });

    const keyStore1 = new CollabEncryptionKeyStore({ invoke: null });
    const keyStore2 = new CollabEncryptionKeyStore({ invoke: null });

    let metadata1: any = null;
    let metadata2: any = null;

    const session1: CollabSession = createCollabSession({
      docId,
      doc: ydoc,
      encryption: {
        shouldEncryptCell: (cell) => encryptedKeyIdForCell(metadata1, cell) != null,
        keyForCell: (cell) => {
          const keyId = encryptedKeyIdForCell(metadata1, cell);
          if (!keyId) return null;
          return keyStore1.getCachedKey(docId, keyId);
        },
      },
    });
    const session2: CollabSession = createCollabSession({
      docId,
      doc: ydoc,
      encryption: {
        shouldEncryptCell: (cell) => encryptedKeyIdForCell(metadata2, cell) != null,
        keyForCell: (cell) => {
          const keyId = encryptedKeyIdForCell(metadata2, cell);
          if (!keyId) return null;
          return keyStore2.getCachedKey(docId, keyId);
        },
      },
    });

    metadata1 = session1.metadata;
    metadata2 = session2.metadata;

    const manager1 = new EncryptedRangeManager({ session: session1 });

    const keyId = "k1";
    const keyBytes = new Uint8Array(32);
    keyBytes.fill(42);
    await keyStore1.set(docId, keyId, bytesToBase64(keyBytes));

    manager1.addEncryptedRange({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId });

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
