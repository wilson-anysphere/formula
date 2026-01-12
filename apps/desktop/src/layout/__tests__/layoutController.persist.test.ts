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

  test("switching workspaces persists pending ephemeral updates before reloading", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-2";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    // Seed initial persisted state.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();

    // Apply an ephemeral update (in-memory only).
    controller.setSplitPaneScroll("secondary", { scrollX: 5, scrollY: 6 }, { persist: false });
    expect(storage.getItem(key)).toBe(persistedBefore);

    // Switching workspaces reloads the layout; ensure we don't lose the in-memory change.
    controller.setWorkspace("analysis");

    const after = storage.getItem(key);
    expect(after).not.toBeNull();
    expect(after).not.toBe(persistedBefore);
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(5);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(6);
  });

  test("saveWorkspace(makeActive) persists pending ephemeral updates in the previous workspace", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-3";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const defaultKey = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    // Seed initial persisted state in the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(defaultKey);
    expect(persistedBefore).not.toBeNull();

    // Apply an ephemeral update to the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 7, scrollY: 8 }, { persist: false });
    expect(storage.getItem(defaultKey)).toBe(persistedBefore);

    // Saving a new workspace and making it active should flush the pending update for the default workspace
    // (otherwise the default workspace would lose the latest scroll/zoom state when we switch away).
    controller.saveWorkspace("analysis", { makeActive: true });

    const after = storage.getItem(defaultKey);
    expect(after).not.toBeNull();
    expect(after).not.toBe(persistedBefore);
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(7);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(8);
  });
});
