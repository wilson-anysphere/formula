import { describe, expect, test } from "vitest";

import { LayoutController } from "../layoutController.js";
import { LayoutWorkspaceManager, MemoryStorage } from "../layoutPersistence.js";

describe("LayoutController persistence", () => {
  test("setSplitPaneScroll can be applied without persisting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const workspaceManager = new LayoutWorkspaceManager({ storage });
    const workbookId = "workbook-1";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = (workspaceManager as any).workbookKey(workbookId);
    const before = storage.getItem(key);

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    controller.setSplitPaneScroll("secondary", { scrollX: 1, scrollY: 2 }, { persist: false });

    expect(storage.getItem(key)).toBe(before);
    expect(changeCount).toBe(1);

    controller.persistNow();

    const after = storage.getItem(key);
    expect(after).not.toBe(before);
    expect(after).not.toBeNull();

    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(1);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(2);

    // persistNow should not emit an additional change event.
    expect(changeCount).toBe(1);
  });
});
