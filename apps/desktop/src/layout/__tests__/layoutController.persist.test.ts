import { describe, expect, test } from "vitest";

import { LayoutController } from "../layoutController.js";
import { LayoutWorkspaceManager, MemoryStorage } from "../layoutPersistence.js";

describe("LayoutController persistence", () => {
  test("setSplitPaneScroll can be applied without persisting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-1";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    // Persist an initial value so we can assert "no storage writes" on ephemeral updates.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();

    controller.setSplitPaneScroll("secondary", { scrollX: 1, scrollY: 2 }, { persist: false });

    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(2);

    controller.persistNow();

    const after = storage.getItem(key);
    expect(after).not.toBe(persistedBefore);
    expect(after).not.toBeNull();

    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(1);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(2);

    // persistNow should not emit an additional change event.
    expect(changeCount).toBe(2);
  });
});
