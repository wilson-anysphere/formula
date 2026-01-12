import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Desktop main.ts wires titlebar window controls to Tauri window operations", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const source = fs.readFileSync(mainPath, "utf8");

  // Ensure the Titlebar is mounted with window control callbacks.
  assert.match(source, /\bwindowControls:\s*titlebarWindowControls\b/);

  // Ensure the callbacks exist and invoke the expected helpers.
  // (Keep this as simple text matching to avoid overly brittle AST parsing.)
  assert.match(source, /\bonClose:\s*\(\)\s*=>\s*\{\s*\n?\s*void\s+hideTauriWindow\(\)/);
  assert.match(source, /\bonMinimize:\s*\(\)\s*=>\s*\{\s*\n?\s*void\s+minimizeTauriWindow\(\)/);
  assert.match(source, /\bonToggleMaximize:\s*\(\)\s*=>\s*\{\s*\n?\s*void\s+toggleTauriWindowMaximize\(\)/);
});

