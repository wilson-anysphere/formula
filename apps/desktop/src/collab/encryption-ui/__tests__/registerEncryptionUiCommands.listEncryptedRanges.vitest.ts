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

  it("listEncryptedRanges lists ranges and selects the chosen one", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          sheetId: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 1,
          endCol: 1,
          keyId: "k1",
        },
      ],
    };

    vi.mocked(showQuickPick).mockResolvedValue(manager.list()[0] as any);

    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      getSheetDisplayNameById: (id: string) => id,
      selectRange: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.listEncryptedRanges");

    expect(showQuickPick).toHaveBeenCalledTimes(1);
    expect(app.selectRange).toHaveBeenCalledWith(
      {
        sheetId: "Sheet1",
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      },
      { scrollIntoView: true, focus: true },
    );
  });

  it("listEncryptedRanges shows a toast when there are no ranges", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = { list: () => [] };
    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      getSheetDisplayNameById: (id: string) => id,
      selectRange: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.listEncryptedRanges");

    expect(showToast).toHaveBeenCalledWith("No encrypted ranges in this workbook.", "info");
    expect(showQuickPick).not.toHaveBeenCalled();
  });
});

