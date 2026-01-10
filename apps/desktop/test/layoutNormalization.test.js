import assert from "node:assert/strict";
import test from "node:test";

import { LAYOUT_STATE_VERSION } from "../src/layout/constants.js";
import { normalizeLayout } from "../src/layout/layoutNormalization.js";
import { createDefaultLayout } from "../src/layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "../src/panels/panelRegistry.js";

test("normalizeLayout removes unknown panel ids when registry is provided", () => {
  const raw = {
    version: 1,
    docks: {
      left: { size: 200, collapsed: false, panels: ["unknownPanel", PanelIds.AI_CHAT], active: "unknownPanel" },
      right: { size: 300, collapsed: false, panels: [], active: null },
      bottom: { size: 200, collapsed: false, panels: [], active: null },
    },
    floating: {
      unknownPanel: { x: 1, y: 2, width: 3, height: 4, minimized: false },
      [PanelIds.VERSION_HISTORY]: { x: 10, y: 20, width: 300, height: 400, minimized: true },
    },
    splitView: { direction: "none", ratio: 0.5, panes: {} },
  };

  const normalized = normalizeLayout(raw, { panelRegistry: PANEL_REGISTRY, primarySheetId: "Sheet1" });

  assert.deepEqual(normalized.docks.left.panels, [PanelIds.AI_CHAT]);
  assert.equal(normalized.docks.left.active, PanelIds.AI_CHAT);
  assert.equal(normalized.floating.unknownPanel, undefined);
  assert.ok(normalized.floating[PanelIds.VERSION_HISTORY]);
});

test("normalizeLayout dedupes panels across multiple docks and floating", () => {
  const raw = {
    version: 1,
    docks: {
      left: { size: 200, collapsed: false, panels: [PanelIds.AI_CHAT], active: PanelIds.AI_CHAT },
      right: { size: 300, collapsed: false, panels: [PanelIds.AI_CHAT, PanelIds.PIVOT_BUILDER], active: PanelIds.AI_CHAT },
      bottom: { size: 200, collapsed: false, panels: [PanelIds.PIVOT_BUILDER], active: PanelIds.PIVOT_BUILDER },
    },
    floating: {
      [PanelIds.AI_CHAT]: { x: 10, y: 20, width: 300, height: 400, minimized: false },
    },
    splitView: { direction: "none", ratio: 0.5, panes: {} },
  };

  const normalized = normalizeLayout(raw, { panelRegistry: PANEL_REGISTRY, primarySheetId: "Sheet1" });

  assert.deepEqual(normalized.docks.left.panels, [PanelIds.AI_CHAT]);
  assert.deepEqual(normalized.docks.right.panels, [PanelIds.PIVOT_BUILDER]);
  assert.deepEqual(normalized.docks.bottom.panels, []);
  assert.equal(normalized.floating[PanelIds.AI_CHAT], undefined);
});

test("normalizeLayout falls back to defaults when given non-object input", () => {
  const normalized = normalizeLayout(null, { primarySheetId: "Sheet1" });
  assert.deepEqual(normalized, createDefaultLayout({ primarySheetId: "Sheet1" }));
});

test("normalizeLayout coerces version to current schema version", () => {
  const normalized = normalizeLayout({ version: 999 }, { primarySheetId: "Sheet1" });
  assert.equal(normalized.version, LAYOUT_STATE_VERSION);
});

