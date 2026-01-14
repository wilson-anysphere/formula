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
});

