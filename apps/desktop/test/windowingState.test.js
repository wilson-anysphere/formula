import assert from "node:assert/strict";
import test from "node:test";

import { deserializeWindowingState, serializeWindowingState } from "../src/layout/windowingSerializer.js";
import {
  closeWindow,
  createDefaultWindowingState,
  focusWindow,
  openWorkbookWindow,
  setWindowBounds,
  setWindowMaximized,
  setWindowWorkspace,
} from "../src/layout/windowingState.js";

test("windowing state round-trips through serialization", () => {
  let state = createDefaultWindowingState();
  state = openWorkbookWindow(state, "workbook-a", { workspaceId: "default", bounds: { x: 10, y: 20, width: 900, height: 700 } });
  state = openWorkbookWindow(state, "workbook-b", { workspaceId: "analysis", bounds: { x: 30, y: 40, width: 800, height: 600 } });
  state = setWindowMaximized(state, state.windows[1].id, true);
  state = focusWindow(state, state.windows[0].id);
  state = setWindowWorkspace(state, state.windows[0].id, "review");

  const serialized = serializeWindowingState(state);
  const restored = deserializeWindowingState(serialized);

  assert.deepEqual(restored, state);
});

test("windowing state handles close + focus updates", () => {
  let state = createDefaultWindowingState();
  state = openWorkbookWindow(state, "workbook-a");
  state = openWorkbookWindow(state, "workbook-b");

  const firstId = state.windows[0].id;
  const secondId = state.windows[1].id;

  state = focusWindow(state, firstId);
  assert.equal(state.focusedWindowId, firstId);

  state = closeWindow(state, firstId);
  assert.equal(state.focusedWindowId, secondId);

  state = setWindowBounds(state, secondId, { width: 300 });
  assert.equal(state.windows[0].bounds.width, 480); // clamped minimum
});
