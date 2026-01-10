import assert from "node:assert/strict";
import test from "node:test";

import { deserializeLayout, serializeLayout } from "../src/layout/layoutSerializer.js";
import {
  createDefaultLayout,
  dockPanel,
  openPanel,
  setDockSize,
  setSplitDirection,
  setSplitPaneScroll,
  setSplitPaneSheet,
} from "../src/layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "../src/panels/panelRegistry.js";

test("layout state round-trips through serialization", () => {
  let layout = createDefaultLayout({ primarySheetId: "Sheet1" });

  layout = openPanel(layout, PanelIds.AI_CHAT, { panelRegistry: PANEL_REGISTRY });
  layout = dockPanel(layout, PanelIds.AI_CHAT, "left");
  layout = setDockSize(layout, "left", 420);

  layout = setSplitDirection(layout, "vertical", 0.6);
  layout = setSplitPaneSheet(layout, "secondary", "Sheet2");
  layout = setSplitPaneScroll(layout, "primary", { scrollX: 120, scrollY: 340 });
  layout = setSplitPaneScroll(layout, "secondary", { scrollX: 0, scrollY: 9001 });

  const serialized = serializeLayout(layout, { panelRegistry: PANEL_REGISTRY, primarySheetId: "Sheet1" });
  const restored = deserializeLayout(serialized, { panelRegistry: PANEL_REGISTRY, primarySheetId: "Sheet1" });

  assert.deepEqual(restored, layout);
});

test("deserializeLayout falls back to defaults on invalid JSON", () => {
  const restored = deserializeLayout("{this is not json", { primarySheetId: "Sheet1" });
  assert.deepEqual(restored, createDefaultLayout({ primarySheetId: "Sheet1" }));
});

