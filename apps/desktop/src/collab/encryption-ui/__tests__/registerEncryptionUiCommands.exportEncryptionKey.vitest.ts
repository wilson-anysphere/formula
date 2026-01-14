import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../../extensions/ui.js", () => ({
  showToast: vi.fn(),
  showInputBox: vi.fn(),
  showQuickPick: vi.fn(),
}));

import * as Y from "yjs";

import { CommandRegistry } from "../../../extensions/commandRegistry.js";
import { showInputBox, showToast } from "../../../extensions/ui.js";
import { registerEncryptionUiCommands } from "../registerEncryptionUiCommands.js";
import { parseEncryptionKeyExportString } from "../keyExportFormat";

describe("registerEncryptionUiCommands", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("exportEncryptionKey prefers keyId from existing encrypted cell payload over policy metadata", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    // Store a policy range with a *different* key id than the encrypted payload.
    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    const range = new Y.Map<any>();
    range.set("id", "r1");
    range.set("sheetId", "Sheet1");
    range.set("startRow", 0);
    range.set("startCol", 0);
    range.set("endRow", 0);
    range.set("endCol", 0);
    range.set("keyId", "k-new");
    ranges.push([range]);
    metadata.set("encryptedRanges", ranges);

    // Store an encrypted cell payload under the canonical key with `keyId: k-old`.
    const cells = doc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 1,
      alg: "AES-256-GCM",
      keyId: "k-old",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    const keyBytes = new Uint8Array(32).fill(9);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "k-old" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getEncryptedRangeManager: () => ({ list: () => [] }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("done");

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    // Should have requested the key for k-old, not k-new.
    expect(keyStore.getCachedKey).toHaveBeenCalledWith(docId, "k-old");

    const inputOptions = vi.mocked(showInputBox).mock.calls[0]?.[0] as any;
    expect(inputOptions?.value).toMatch(/^formula-enc:\/\//);
    const parsed = parseEncryptionKeyExportString(String(inputOptions.value));
    expect(parsed.docId).toBe(docId);
    expect(parsed.keyId).toBe("k-old");
    expect(parsed.keyBytes).toEqual(keyBytes);

    // Should not have shown a "not inside encrypted range" warning.
    expect(showToast).not.toHaveBeenCalledWith(expect.stringMatching(/not inside an encrypted range/i), "warning");
  });
});
