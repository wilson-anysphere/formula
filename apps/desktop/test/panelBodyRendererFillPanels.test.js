import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractSection(source, startMarker, endMarker) {
  const startIdx = source.indexOf(startMarker);
  assert.ok(startIdx !== -1, `Expected to find start marker: ${startMarker}`);

  const endIdx = source.indexOf(endMarker, startIdx);
  assert.ok(endIdx !== -1, `Expected to find end marker: ${endMarker}`);

  return source.slice(startIdx, endIdx);
}

test("panelBodyRenderer keeps key dock panels full-height (panel-body--fill)", () => {
  const filePath = path.join(__dirname, "..", "src", "panels", "panelBodyRenderer.tsx");
  const source = fs.readFileSync(filePath, "utf8");

  const fillPanels = [
    {
      name: "AI chat",
      start: "if (panelId === PanelIds.AI_CHAT)",
      end: "if (panelId === PanelIds.QUERY_EDITOR)",
    },
    {
      name: "Query editor",
      start: "if (panelId === PanelIds.QUERY_EDITOR)",
      end: "if (panelId === PanelIds.EXTENSIONS",
    },
    {
      name: "Extensions",
      start: "if (panelId === PanelIds.EXTENSIONS",
      end: "if (panelId === PanelIds.PIVOT_BUILDER)",
    },
    {
      name: "Pivot builder",
      start: "if (panelId === PanelIds.PIVOT_BUILDER)",
      end: "if (panelId === PanelIds.DATA_QUERIES)",
    },
    {
      name: "Data queries",
      start: "if (panelId === PanelIds.DATA_QUERIES)",
      end: "if (panelId === PanelIds.MARKETPLACE)",
    },
    {
      name: "Marketplace",
      start: "if (panelId === PanelIds.MARKETPLACE)",
      end: "if (panelId === PanelIds.PYTHON)",
    },
    {
      name: "Python",
      start: "if (panelId === PanelIds.PYTHON)",
      end: "const panelDef = options.panelRegistry?.get(panelId) as any;",
    },
    {
      name: "Extension panel bodies",
      start: "if (panelDef?.source?.kind === \"extension\"",
      end: "if (panelId === PanelIds.AI_AUDIT)",
    },
    {
      name: "AI audit",
      start: "if (panelId === PanelIds.AI_AUDIT)",
      end: "if (panelId === PanelIds.VERSION_HISTORY)",
    },
    {
      name: "Version history",
      start: "if (panelId === PanelIds.VERSION_HISTORY)",
      end: "if (panelId === PanelIds.BRANCH_MANAGER)",
    },
    {
      name: "Branch manager",
      start: "if (panelId === PanelIds.BRANCH_MANAGER)",
      end: "if (panelId === PanelIds.MACROS)",
    },
    {
      name: "VBA migrate",
      start: "if (panelId === PanelIds.VBA_MIGRATE)",
      end: "body.textContent = `Panel: ${panelId}`;",
    },
  ];

  for (const panel of fillPanels) {
    const section = extractSection(source, panel.start, panel.end);
    assert.match(section, /makeBodyFillAvailableHeight\(body\);/, `Expected ${panel.name} panel to call makeBodyFillAvailableHeight(body)`);
  }
});
