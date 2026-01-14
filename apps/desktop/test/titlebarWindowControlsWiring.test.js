import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Desktop main.ts wires titlebar window controls to Tauri window operations", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const source = stripComments(fs.readFileSync(mainPath, "utf8"));

  // Ensure the Titlebar is mounted with window control callbacks.
  assert.match(source, /\bwindowControls:\s*titlebarWindowControls\b/);

  // Ensure the callbacks exist and invoke the expected helpers.
  // Keep this loose enough to tolerate formatting changes (braces vs expression bodies, optional
  // `void`, `.catch(...)`, etc) while still catching the "buttons render but do nothing" class of
  // regressions.
  assert.match(source, /\bonClose:\s*\(\)\s*=>[\s\S]{0,200}\bhideTauriWindow\(/);
  assert.match(source, /\bonMinimize:\s*\(\)\s*=>[\s\S]{0,200}\bminimizeTauriWindow\(/);
  assert.match(source, /\bonToggleMaximize:\s*\(\)\s*=>[\s\S]{0,200}\btoggleTauriWindowMaximize\(/);
});
