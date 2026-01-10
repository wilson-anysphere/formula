import assert from "node:assert/strict";
import test from "node:test";

import { MemoryStorage } from "../src/layout/layoutPersistence.js";
import { WindowingController } from "../src/layout/windowingController.js";
import { WindowingSessionManager } from "../src/layout/windowingPersistence.js";

test("WindowingController persists session state and restores on reload", () => {
  const storage = new MemoryStorage();
  const sessionManager = new WindowingSessionManager({ storage });

  const controller1 = new WindowingController({ sessionManager });
  const winA = controller1.openWorkbookWindow("workbook-a", { workspaceId: "default", focus: true });
  const winB = controller1.openWorkbookWindow("workbook-b", { workspaceId: "analysis", focus: true });

  assert.equal(controller1.state.focusedWindowId, winB);
  controller1.focusWindow(winA);
  assert.equal(controller1.state.focusedWindowId, winA);

  // New controller should see persisted state.
  const controller2 = new WindowingController({ sessionManager });
  assert.equal(controller2.state.windows.length, 2);
  assert.equal(controller2.state.focusedWindowId, winA);

  controller2.closeWindow(winA);
  assert.equal(controller2.state.windows.length, 1);

  // Persisted removal.
  const controller3 = new WindowingController({ sessionManager });
  assert.equal(controller3.state.windows.length, 1);
});

