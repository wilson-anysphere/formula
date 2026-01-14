import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowPath = path.join(repoRoot, ".github", "workflows", "windows-arm64-smoke.yml");

async function readWorkflow() {
  return await readFile(workflowPath, "utf8");
}

test("windows-arm64-smoke workflow validates built bundles via validate-windows-bundles.ps1", async () => {
  const text = await readWorkflow();
  assert.match(text, /validate-windows-bundles\.ps1/);
  assert.match(text, /-RequireExe\b/);
  assert.match(text, /-RequireMsi\b/);
  assert.match(
    text,
    /pwsh\s+-NoProfile\s+-ExecutionPolicy\s+Bypass\s+-File\s+(?:\.\/)?scripts\/validate-windows-bundles\.ps1/,
  );
});
