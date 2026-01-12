import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Query editor React panels avoid inline styles (use query-editor.css classes)", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "query-editor");
  const components = [
    path.join(panelDir, "QueryEditorPanelContainer.tsx"),
    path.join(panelDir, "QueryEditorPanel.tsx"),
    path.join(panelDir, "components", "StepsList.tsx"),
    path.join(panelDir, "components", "SchemaView.tsx"),
    path.join(panelDir, "components", "AddStepMenu.tsx"),
    path.join(panelDir, "components", "PreviewGrid.tsx"),
  ];

  for (const filePath of components) {
    const src = fs.readFileSync(filePath, "utf8");
    assert.equal(
      /\bstyle\s*=/.test(src),
      false,
      `${path.relative(path.join(__dirname, ".."), filePath)} should not use inline styles; use src/styles/query-editor.css instead`,
    );
  }

  const cssPath = path.join(__dirname, "..", "src", "styles", "query-editor.css");
  assert.equal(fs.existsSync(cssPath), true, "Expected apps/desktop/src/styles/query-editor.css to exist");
  const css = fs.readFileSync(cssPath, "utf8");
  for (const selector of [".query-editor", ".query-editor-container", ".query-editor-preview__table"]) {
    assert.ok(css.includes(selector), `Expected query-editor.css to define ${selector}`);
  }

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.match(
    mainSrc,
    /import\s+["'][^"']*styles\/query-editor\.css["']/,
    "apps/desktop/src/main.ts should import src/styles/query-editor.css so the query editor UI is styled in production builds",
  );
});

