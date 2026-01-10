import assert from "node:assert/strict";
import test from "node:test";

import { MemoryStorage } from "../src/layout/layoutPersistence.js";
import { WindowingSessionManager } from "../src/layout/windowingPersistence.js";
import { createDefaultWindowingState, openWorkbookWindow } from "../src/layout/windowingState.js";

test("WindowingSessionManager persists and restores multi-window session", () => {
  const storage = new MemoryStorage();
  const manager1 = new WindowingSessionManager({ storage });

  let state = createDefaultWindowingState();
  state = openWorkbookWindow(state, "workbook-a", { workspaceId: "default" });
  state = openWorkbookWindow(state, "workbook-b", { workspaceId: "analysis" });

  manager1.save(state);

  const manager2 = new WindowingSessionManager({ storage });
  const restored = manager2.load();

  assert.deepEqual(restored, state);
});

test("WindowingSessionManager returns default state when nothing persisted", () => {
  const storage = new MemoryStorage();
  const manager = new WindowingSessionManager({ storage });

  assert.deepEqual(manager.load(), createDefaultWindowingState());
});

