import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../../extensions/ui.js", () => ({
  showToast: vi.fn(),
  showInputBox: vi.fn(),
  showQuickPick: vi.fn(),
}));

import { CommandRegistry } from "../../../extensions/commandRegistry.js";
import { showInputBox, showQuickPick, showToast } from "../../../extensions/ui.js";
import { registerEncryptionUiCommands } from "../registerEncryptionUiCommands.js";
import { serializeEncryptionKeyExportString } from "../keyExportFormat";

describe("registerEncryptionUiCommands", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("importEncryptionKey stores a new key and rehydrates the binder", async () => {
    const commandRegistry = new CommandRegistry();

    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => null),
      set: vi.fn(async () => ({ keyId: "k1" })),
    };

    const app: any = {
      getCollabSession: () => ({ doc: { guid: "doc-1" } }),
      getCollabEncryptionKeyStore: () => keyStore,
      rehydrateCollabBinder: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    const keyBytes = new Uint8Array(32).fill(1);
    const exportString = serializeEncryptionKeyExportString({ docId: "doc-1", keyId: "k1", keyBytes });
    vi.mocked(showInputBox).mockResolvedValue(exportString);

    await commandRegistry.executeCommand("collab.importEncryptionKey");

    expect(keyStore.set).toHaveBeenCalledTimes(1);
    expect(app.rehydrateCollabBinder).toHaveBeenCalledTimes(1);
    expect(showToast).toHaveBeenCalledWith('Imported encryption key "k1".', "info");
  });

  it("importEncryptionKey does not overwrite an existing identical key", async () => {
    const commandRegistry = new CommandRegistry();

    const keyBytes = new Uint8Array(32).fill(2);
    const keyStore = {
      getCachedKey: vi.fn(() => ({ keyId: "k1", keyBytes })),
      get: vi.fn(async () => null),
      set: vi.fn(async () => ({ keyId: "k1" })),
    };

    const app: any = {
      getCollabSession: () => ({ doc: { guid: "doc-1" } }),
      getCollabEncryptionKeyStore: () => keyStore,
      rehydrateCollabBinder: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    const exportString = serializeEncryptionKeyExportString({ docId: "doc-1", keyId: "k1", keyBytes });
    vi.mocked(showInputBox).mockResolvedValue(exportString);

    await commandRegistry.executeCommand("collab.importEncryptionKey");

    expect(keyStore.set).not.toHaveBeenCalled();
    expect(app.rehydrateCollabBinder).toHaveBeenCalledTimes(1);
    expect(showToast).toHaveBeenCalledWith('Encryption key "k1" is already imported.', "info");
  });

  it("importEncryptionKey prompts before overwriting a conflicting key", async () => {
    const commandRegistry = new CommandRegistry();

    const existingBytes = new Uint8Array(32).fill(3);
    const importedBytes = new Uint8Array(32).fill(4);
    const keyStore = {
      getCachedKey: vi.fn(() => ({ keyId: "k1", keyBytes: existingBytes })),
      get: vi.fn(async () => null),
      set: vi.fn(async () => ({ keyId: "k1" })),
    };

    const app: any = {
      getCollabSession: () => ({ doc: { guid: "doc-1" } }),
      getCollabEncryptionKeyStore: () => keyStore,
      rehydrateCollabBinder: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    const exportString = serializeEncryptionKeyExportString({ docId: "doc-1", keyId: "k1", keyBytes: importedBytes });
    vi.mocked(showInputBox).mockResolvedValue(exportString);
    vi.mocked(showQuickPick).mockResolvedValue("cancel");

    await commandRegistry.executeCommand("collab.importEncryptionKey");

    expect(showQuickPick).toHaveBeenCalledTimes(1);
    expect(keyStore.set).not.toHaveBeenCalled();
    expect(app.rehydrateCollabBinder).not.toHaveBeenCalled();

    vi.clearAllMocks();
    vi.mocked(showInputBox).mockResolvedValue(exportString);
    vi.mocked(showQuickPick).mockResolvedValue("overwrite");

    await commandRegistry.executeCommand("collab.importEncryptionKey");

    expect(showQuickPick).toHaveBeenCalledTimes(1);
    expect(keyStore.set).toHaveBeenCalledTimes(1);
    expect(app.rehydrateCollabBinder).toHaveBeenCalledTimes(1);
    expect(showToast).toHaveBeenCalledWith('Overwrote encryption key "k1".', "warning");
  });

  it("importEncryptionKey prompts when it cannot verify whether a key id already exists", async () => {
    const commandRegistry = new CommandRegistry();

    const keyBytes = new Uint8Array(32).fill(5);
    const keyStore = {
      getCachedKey: vi.fn(() => null),
      get: vi.fn(async () => {
        throw new Error("backend unavailable");
      }),
      set: vi.fn(async () => ({ keyId: "k1" })),
    };

    const app: any = {
      getCollabSession: () => ({ doc: { guid: "doc-1" } }),
      getCollabEncryptionKeyStore: () => keyStore,
      rehydrateCollabBinder: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    const exportString = serializeEncryptionKeyExportString({ docId: "doc-1", keyId: "k1", keyBytes });
    vi.mocked(showInputBox).mockResolvedValue(exportString);

    vi.mocked(showQuickPick).mockResolvedValue("cancel");
    await commandRegistry.executeCommand("collab.importEncryptionKey");
    expect(keyStore.set).not.toHaveBeenCalled();
    expect(app.rehydrateCollabBinder).not.toHaveBeenCalled();

    vi.clearAllMocks();
    registerEncryptionUiCommands({ commandRegistry, app });
    vi.mocked(showInputBox).mockResolvedValue(exportString);
    vi.mocked(showQuickPick).mockResolvedValue("import");

    await commandRegistry.executeCommand("collab.importEncryptionKey");
    expect(keyStore.set).toHaveBeenCalledTimes(1);
    expect(app.rehydrateCollabBinder).toHaveBeenCalledTimes(1);
    expect(showToast).toHaveBeenCalledWith('Imported encryption key "k1".', "info");
  });
});
