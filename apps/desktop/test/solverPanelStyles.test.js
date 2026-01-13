import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Solver panel React components avoid inline styles (use solver.css classes)", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "solver");
  const cssPath = path.join(__dirname, "..", "src", "styles", "solver.css");
  const mainPath = path.join(__dirname, "..", "src", "main.ts");

  const sources = {
    panel: fs.readFileSync(path.join(panelDir, "SolverPanel.tsx"), "utf8"),
    dialog: fs.readFileSync(path.join(panelDir, "SolverDialog.tsx"), "utf8"),
    progress: fs.readFileSync(path.join(panelDir, "SolverProgress.tsx"), "utf8"),
    summary: fs.readFileSync(path.join(panelDir, "SolverResultSummary.tsx"), "utf8"),
  };

  for (const [name, content] of Object.entries(sources)) {
    assert.equal(
      /\bstyle\s*=/.test(content),
      false,
      `${name} should not use inline styles; use src/styles/solver.css classes instead`,
    );
  }

  const requiredClasses = [
    "solver-panel",
    "solver-panel__header",
    "solver-panel__configure-button",
    "solver-panel__error",
    "solver-dialog",
    "solver-dialog__title",
    "solver-dialog__section",
    "solver-dialog__field",
    "solver-dialog__label",
    "solver-dialog__input",
    "solver-dialog__select",
    "solver-dialog__row",
    "solver-dialog__variable-row",
    "solver-dialog__constraint-row",
    "solver-dialog__footer",
    "solver-progress",
    "solver-results",
    "solver-results__actions",
    "solver__button",
    "solver__button--primary",
  ];

  for (const className of requiredClasses) {
    assert.ok(
      Object.values(sources).some((src) => src.includes(className)),
      `Expected solver components to reference the ${className} CSS class`,
    );
  }

  assert.equal(fs.existsSync(cssPath), true, "Expected apps/desktop/src/styles/solver.css to exist");
  const css = fs.readFileSync(cssPath, "utf8");
  for (const className of requiredClasses) {
    assert.ok(css.includes(`.${className}`), `Expected solver.css to define .${className}`);
  }

  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.match(
    mainSrc,
    /import\s+["'][^"']*styles\/solver\.css["']/,
    "apps/desktop/src/main.ts should import src/styles/solver.css so the Solver panel UI is styled in production builds",
  );
});

