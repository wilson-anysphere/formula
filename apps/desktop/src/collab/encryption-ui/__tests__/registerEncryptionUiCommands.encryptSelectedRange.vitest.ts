import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../../extensions/ui.js", () => ({
  showToast: vi.fn(),
  showInputBox: vi.fn(),
  showQuickPick: vi.fn(),
}));

import { base64ToBytes } from "@formula/collab-encryption";

import { CommandRegistry } from "../../../extensions/commandRegistry.js";
import { showInputBox, showQuickPick, showToast } from "../../../extensions/ui.js";
import { registerEncryptionUiCommands } from "../registerEncryptionUiCommands.js";
import { parseEncryptionKeyExportString } from "../keyExportFormat";

describe("registerEncryptionUiCommands", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("encryptSelectedRange reuses existing key bytes instead of overwriting keyId", async () => {
    const commandRegistry = new CommandRegistry();

    const existingKeyBytes = new Uint8Array(32).fill(7);
    const keyStore = {
      getCachedKey: vi.fn(() => ({ keyId: "k1", keyBytes: existingKeyBytes })),
      get: vi.fn(async () => null),
      set: vi.fn(async () => ({ keyId: "k1" })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => [],
      update: vi.fn(),
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");
    vi.mocked(showQuickPick).mockResolvedValue("encrypt");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    // Should not overwrite the existing key id.
    expect(keyStore.set).not.toHaveBeenCalled();

    // Confirm prompt should reflect reuse.
    expect(showQuickPick).toHaveBeenCalledWith(
      [expect.objectContaining({ description: expect.stringMatching(/reuse existing key/) })],
      expect.anything(),
    );

    expect(manager.add).toHaveBeenCalledWith(expect.objectContaining({ sheetId: "Sheet1", keyId: "k1" }));

    const toastMessage = vi.mocked(showToast).mock.calls.find((c) => typeof c[0] === "string")?.[0] as string;
    expect(toastMessage).toMatch(/Encrypted Sheet1!/);
    const exportString = toastMessage.split("\n")[1] ?? "";
    const parsed = parseEncryptionKeyExportString(exportString);
    expect(parsed.docId).toBe("doc-1");
    expect(parsed.keyId).toBe("k1");
    expect(parsed.keyBytes).toEqual(existingKeyBytes);
  });

  it("encryptSelectedRange generates and stores a new key when keyId is missing", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
      set: vi.fn(async (_docId: string, keyId: string) => ({ keyId })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => [],
      update: vi.fn(),
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");
    vi.mocked(showQuickPick).mockResolvedValue("encrypt");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    expect(keyStore.get).toHaveBeenCalled();
    expect(keyStore.set).toHaveBeenCalledTimes(1);

    // Confirm prompt should reflect new key generation.
    expect(showQuickPick).toHaveBeenCalledWith(
      [expect.objectContaining({ description: expect.stringMatching(/new key/) })],
      expect.anything(),
    );

    const base64Arg = vi.mocked(keyStore.set).mock.calls[0]?.[2] as string;
    const storedBytes = base64ToBytes(base64Arg);

    const toastMessage = vi.mocked(showToast).mock.calls.find((c) => typeof c[0] === "string")?.[0] as string;
    const exportString = toastMessage.split("\n")[1] ?? "";
    const parsed = parseEncryptionKeyExportString(exportString);
    expect(parsed.keyId).toBe("k1");
    expect(parsed.keyBytes).toEqual(storedBytes);
  });

  it("encryptSelectedRange warns when it cannot verify whether a key id already exists", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => {
        throw new Error("backend unavailable");
      }),
      set: vi.fn(async (_docId: string, keyId: string) => ({ keyId })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => [],
      update: vi.fn(),
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");
    vi.mocked(showQuickPick).mockResolvedValue("encrypt");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    const items = vi.mocked(showQuickPick).mock.calls[0]?.[0] as any[];
    expect(items?.[0]?.description ?? "").toMatch(/could not verify existing key/i);

    expect(keyStore.set).toHaveBeenCalledTimes(1);
    expect(manager.add).toHaveBeenCalledWith(expect.objectContaining({ keyId: "k1" }));
  });

  it("encryptSelectedRange refuses to generate new key bytes for a keyId already used in policy when key bytes are missing", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
      set: vi.fn(async (_docId: string, keyId: string) => ({ keyId })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => [
        {
          id: "existing",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 0,
          endCol: 0,
          keyId: "k1",
        },
      ],
      update: vi.fn(),
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    expect(showQuickPick).not.toHaveBeenCalled();
    expect(keyStore.set).not.toHaveBeenCalled();
    expect(manager.add).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/already used by an encrypted range/i), "warning");
  });

  it("encryptSelectedRange refuses to generate new key bytes for a keyId already used by encrypted cells when key bytes are missing", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
      set: vi.fn(async (_docId: string, keyId: string) => ({ keyId })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => [],
      update: vi.fn(),
      remove: vi.fn(),
    };

    // Simulate an existing encrypted payload for the top-left cell in the selection.
    const cells = new Map<string, any>();
    const cellMap = new Map<string, any>();
    cellMap.set("enc", {
      v: 1,
      alg: "AES-256-GCM",
      keyId: "k1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        cells,
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    expect(showQuickPick).not.toHaveBeenCalled();
    expect(keyStore.set).not.toHaveBeenCalled();
    expect(manager.add).not.toHaveBeenCalled();
    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/already used by encrypted cells/i), "warning");
  });

  it("encryptSelectedRange fails early when encrypted range metadata is unreadable (avoids orphaned keys)", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
      set: vi.fn(async (_docId: string, keyId: string) => ({ keyId })),
    };

    const manager = {
      add: vi.fn(() => "r1"),
      list: () => {
        throw new Error("Unsupported metadata.encryptedRanges schema");
      },
      update: vi.fn(),
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({
        doc: { guid: "doc-1" },
        getRole: () => "editor",
        getPermissions: () => ({ userId: "u1" }),
      }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getCollabEncryptionKeyStore: () => keyStore,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    vi.mocked(showInputBox).mockResolvedValue("k1");

    await commandRegistry.executeCommand("collab.encryptSelectedRange");

    expect(showQuickPick).not.toHaveBeenCalled();
    expect(keyStore.getCachedKey).not.toHaveBeenCalled();
    expect(keyStore.get).not.toHaveBeenCalled();
    expect(keyStore.set).not.toHaveBeenCalled();
    expect(manager.add).not.toHaveBeenCalled();

    expect(showToast).toHaveBeenCalledWith(
      expect.stringMatching(/Encrypted range metadata is in an unsupported format/i),
      "error",
    );
  });
});
