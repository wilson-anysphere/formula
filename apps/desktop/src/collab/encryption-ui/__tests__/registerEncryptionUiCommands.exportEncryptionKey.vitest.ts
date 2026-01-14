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

  it("exportEncryptionKey does not fall back to policy when the encrypted payload key is unavailable locally", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

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

    // Cell ciphertext references a different key id ("k-old"), but we don't have that key locally.
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

    const keyBytes = new Uint8Array(32).fill(5);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "k-new" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    // Should attempt to use the payload key id (k-old), and warn if it's missing.
    expect(keyStore.getCachedKey).toHaveBeenCalledWith(docId, "k-old");
    expect(keyStore.getCachedKey).not.toHaveBeenCalledWith(docId, "k-new");
    expect(showInputBox).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/key id: k-old/i), "warning");
  });

  it("exportEncryptionKey does not fall back to policy when an enc payload exists but its keyId cannot be determined", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    const range = new Y.Map<any>();
    range.set("id", "r1");
    range.set("sheetId", "Sheet1");
    range.set("startRow", 0);
    range.set("startCol", 0);
    range.set("endRow", 0);
    range.set("endCol", 0);
    range.set("keyId", "policy-key");
    ranges.push([range]);
    metadata.set("encryptedRanges", ranges);

    // Cell contains an `enc` payload but the key id is malformed/unsupported.
    const cells = doc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 2,
      alg: "AES-256-GCM",
      keyId: { bogus: true },
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    const keyBytes = new Uint8Array(32).fill(3);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "policy-key" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    expect(showInputBox).not.toHaveBeenCalled();
    expect(keyStore.getCachedKey).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/cannot determine the key id/i), "error");
  });

  it("exportEncryptionKey falls back to policy metadata when the active cell is not yet encrypted", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    const range = new Y.Map<any>();
    range.set("id", "r1");
    range.set("sheetId", "Sheet1");
    range.set("startRow", 0);
    range.set("startCol", 0);
    range.set("endRow", 0);
    range.set("endCol", 0);
    range.set("keyId", "k1");
    ranges.push([range]);
    metadata.set("encryptedRanges", ranges);

    const cells = doc.getMap("cells");
    // No enc payload present yet.

    const keyBytes = new Uint8Array(32).fill(1);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "k1" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("done");

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    expect(keyStore.getCachedKey).toHaveBeenCalledWith(docId, "k1");

    const inputOptions = vi.mocked(showInputBox).mock.calls[0]?.[0] as any;
    const parsed = parseEncryptionKeyExportString(String(inputOptions.value));
    expect(parsed.keyId).toBe("k1");
    expect(parsed.keyBytes).toEqual(keyBytes);
  });

  it("exportEncryptionKey surfaces an error when the encrypted range metadata schema is unsupported (fail-closed policy)", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    // Unsupported/unknown encryptedRanges schema (newer client wrote something we can't parse).
    const metadata = doc.getMap("metadata");
    metadata.set("encryptedRanges", { foo: "bar" } as any);

    const cells = doc.getMap("cells");
    // No enc payload present yet.

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    expect(showInputBox).not.toHaveBeenCalled();
    expect(keyStore.getCachedKey).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/unsupported format/i), "error");
    expect(showToast).not.toHaveBeenCalledWith(expect.stringMatching(/not inside an encrypted range/i), "warning");
  });

  it("exportEncryptionKey can still export when encryptedRanges schema is unsupported but the cell has an enc payload", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    const metadata = doc.getMap("metadata");
    metadata.set("encryptedRanges", { foo: "bar" } as any);

    const cells = doc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 1,
      alg: "AES-256-GCM",
      keyId: "k1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    const keyBytes = new Uint8Array(32).fill(7);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "k1" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("done");

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    expect(keyStore.getCachedKey).toHaveBeenCalledWith(docId, "k1");
    expect(showToast).not.toHaveBeenCalledWith(expect.stringMatching(/unsupported format/i), "error");
  });

  it("exportEncryptionKey can still identify the key id when enc is stored as a nested Y.Map (unsupported payload schema)", async () => {
    const commandRegistry = new CommandRegistry();

    const docId = "doc-1";
    const doc = new Y.Doc({ guid: docId });

    const metadata = doc.getMap("metadata");
    metadata.set("encryptedRanges", { foo: "bar" } as any);

    const cells = doc.getMap("cells");
    const cellMap = new Y.Map<any>();
    const encMap = new Y.Map<any>();
    encMap.set("v", 2);
    encMap.set("alg", "AES-256-GCM");
    encMap.set("keyId", "k1");
    encMap.set("ivBase64", "AA==");
    encMap.set("tagBase64", "AA==");
    encMap.set("ciphertextBase64", "AA==");
    cellMap.set("enc", encMap);
    cells.set("Sheet1:0:0", cellMap);

    const keyBytes = new Uint8Array(32).fill(7);
    const keyStore = {
      getCachedKey: vi.fn((_docId: string, keyId: string) => (keyId === "k1" ? { keyId, keyBytes } : null)),
      get: vi.fn(async () => null),
    };

    const app: any = {
      getCollabSession: () => ({ doc, cells }),
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("done");

    await commandRegistry.executeCommand("collab.exportEncryptionKey");

    expect(keyStore.getCachedKey).toHaveBeenCalledWith(docId, "k1");
    expect(showToast).not.toHaveBeenCalledWith(expect.stringMatching(/cannot determine the key id/i), "error");
  });
});
