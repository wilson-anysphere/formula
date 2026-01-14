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

  it("listEncryptedRanges resolves legacy sheet-name entries to stable sheet ids when selecting", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          // Legacy schema stored the sheet name instead of stable id.
          sheetId: "Summary",
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
      // Only stable sheet ids should resolve to a different display name. For the legacy
      // `sheetId: "Summary"` entry, treat it as unknown id.
      getSheetDisplayNameById: (id: string) => (id === "sheet-1" ? "Summary" : id),
      // Resolve the display name to the stable id for navigation.
      getSheetIdByName: (name: string) => (name === "Summary" ? "sheet-1" : null),
      selectRange: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.listEncryptedRanges");

    expect(app.selectRange).toHaveBeenCalledWith(
      {
        sheetId: "sheet-1",
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      },
      { scrollIntoView: true, focus: true },
    );
  });

  it("listEncryptedRanges does not treat a resolvable stable sheet id as a sheet name (avoids id/name ambiguity)", async () => {
    const commandRegistry = new CommandRegistry();

    const manager = {
      list: () => [
        {
          id: "r1",
          // This is a stable sheet id for a different sheet, but it happens to equal another
          // sheet's display name ("Summary"). We must not resolve it by name.
          sheetId: "Summary",
          startRow: 0,
          startCol: 0,
          endRow: 0,
          endCol: 0,
          keyId: "k1",
        },
      ],
    };

    vi.mocked(showQuickPick).mockResolvedValue(manager.list()[0] as any);

    const app: any = {
      getCollabSession: () => ({ getRole: () => "editor" }),
      getEncryptedRangeManager: () => manager,
      // "Summary" is a real stable sheet id for a sheet whose display name is "Data".
      getSheetDisplayNameById: (id: string) => (id === "Summary" ? "Data" : id === "sheet-1" ? "Summary" : id),
      // `Data` resolves to the stable id `Summary`. Note: this should not be called with the raw
      // sheet id (`Summary`) as a *name*.
      getSheetIdByName: vi.fn((name: string) => (name === "Data" ? "Summary" : name === "Summary" ? "sheet-1" : null)),
      selectRange: vi.fn(),
    };

    registerEncryptionUiCommands({ commandRegistry, app });

    await commandRegistry.executeCommand("collab.listEncryptedRanges");

    expect(app.getSheetIdByName).not.toHaveBeenCalledWith("Summary");
    expect(app.selectRange).toHaveBeenCalledWith(
      {
        sheetId: "Summary",
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      },
      { scrollIntoView: true, focus: true },
    );
  });
});
