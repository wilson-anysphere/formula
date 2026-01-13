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

test("desktop main.ts avoids static inline style assignments", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const source = fs.readFileSync(mainPath, "utf8");

  for (const legacyClass of ["dock-panel__mount", "panel-mount--fill-column", "dock-panel__body--fill"]) {
    assert.equal(
      source.includes(legacyClass),
      false,
      `main.ts should not reference legacy CSS class ${legacyClass}; use panel-body__container/panel-body--fill instead`,
    );
  }

  // Hidden color inputs are defined in the builtin command registration module so they can be
  // invoked from ribbon/command palette/keybindings through the same path.
  const builtinCommandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtinCommands = fs.readFileSync(builtinCommandsPath, "utf8");

  const hiddenColorInputSection = extractSection(
    builtinCommands,
    "const createHiddenColorInput",
    "let fontColorPicker",
  );
  assert.equal(
    /\.style\./.test(hiddenColorInputSection),
    false,
    "createHiddenColorInput should not set inline styles; use a CSS class instead",
  );
  assert.match(
    hiddenColorInputSection,
    /shell-hidden-input/,
    "createHiddenColorInput should apply the shell-hidden-input CSS class",
  );

  const scriptEditorMountSection = extractSection(
    source,
    "if (panelId === PanelIds.SCRIPT_EDITOR)",
    "if (panelId === PanelIds.PYTHON)",
  );
  assert.equal(
    /\.style\./.test(scriptEditorMountSection),
    false,
    "Script Editor panel mount should not set inline styles; use a CSS class instead",
  );
  assert.match(
    scriptEditorMountSection,
    /panel-body__container/,
    "Script Editor panel mount container should apply the panel-body__container CSS class",
  );
  for (const legacyClass of ["dock-panel__mount", "panel-mount--fill-column"]) {
    assert.equal(
      scriptEditorMountSection.includes(legacyClass),
      false,
      `Script Editor panel mount container should not apply legacy class ${legacyClass}; use panel-body__container instead`,
    );
  }

  const pythonMountSection = extractSection(
    source,
    "if (panelId === PanelIds.PYTHON)",
    "const panelDef = panelRegistry.get(panelId) as any;",
  );
  assert.equal(
    /\.style\./.test(pythonMountSection),
    false,
    "Python panel mount should not set inline styles; use a CSS class instead",
  );
  assert.match(
    pythonMountSection,
    /panel-body__container/,
    "Python panel mount container should apply the panel-body__container CSS class",
  );
  for (const legacyClass of ["dock-panel__mount", "panel-mount--fill-column"]) {
    assert.equal(
      pythonMountSection.includes(legacyClass),
      false,
      `Python panel mount container should not apply legacy class ${legacyClass}; use panel-body__container instead`,
    );
  }
});
