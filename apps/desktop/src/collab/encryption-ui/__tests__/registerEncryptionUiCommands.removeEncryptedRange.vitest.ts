import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../../extensions/ui.js", () => ({
  showToast: vi.fn(),
  showInputBox: vi.fn(),
  showQuickPick: vi.fn(),
}));

import { CommandRegistry } from "../../../extensions/commandRegistry.js";
import { showQuickPick, showToast } from "../../../extensions/ui.js";
import { registerEncryptionUiCommands } from "../registerEncryptionUiCommands.js";

describe("registerEncryptionUiCommands", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("removeEncryptedRange removes single intersecting encrypted range without prompt", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 2,
          keyId: "k1",
        },
      ],
      remove: vi.fn(),
    };

    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getSheetDisplayNameById: (id: string) => id,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.removeEncryptedRange");

    expect(showQuickPick).not.toHaveBeenCalled();
    expect(manager.remove).toHaveBeenCalledWith("r1");
    expect(showToast).toHaveBeenCalled();
  });

  it("removeEncryptedRange prompts when multiple encrypted ranges overlap selection", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 5,
          endCol: 5,
          keyId: "k1",
        },
        {
          id: "r2",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 5,
          endCol: 5,
          keyId: "k2",
        },
      ],
      remove: vi.fn(),
    };

    vi.mocked(showQuickPick).mockResolvedValue("r2");

    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getSheetDisplayNameById: (id: string) => id,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.removeEncryptedRange");

    expect(showQuickPick).toHaveBeenCalledTimes(1);
    expect(manager.remove).toHaveBeenCalledWith("r2");
  });

  it("removeEncryptedRange surfaces remove errors via toast", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 0,
          endCol: 0,
          keyId: "k1",
        },
      ],
      remove: vi.fn(() => {
        throw new Error("boom");
      }),
    };

    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
      getCurrentSheetId: () => "Sheet1",
      getCurrentSheetDisplayName: () => "Sheet1",
      getSheetDisplayNameById: (id: string) => id,
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.removeEncryptedRange");

    expect(showToast).toHaveBeenCalledWith(expect.stringMatching(/Failed to remove encrypted range/), "error");
  });
});
