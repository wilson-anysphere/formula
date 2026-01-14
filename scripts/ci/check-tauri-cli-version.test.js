import assert from "node:assert/strict";
import test from "node:test";
import { findTauriActionScriptIssues } from "./check-tauri-cli-version.mjs";

test("detects tauri-action steps missing tauriScript (dash uses form)", () => {
  const yaml = [
    "steps:",
    "  - uses: tauri-apps/tauri-action@deadbeef",
    "    with:",
    "      projectPath: apps/desktop",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 1);
  assert.equal(issues[0].line, 2);
  assert.match(issues[0].message, /tauriScript: cargo tauri/);
});

test("detects tauri-action steps missing tauriScript (name + uses form)", () => {
  const yaml = [
    "steps:",
    "  - name: Build desktop bundles",
    "    uses: tauri-apps/tauri-action@deadbeef",
    "    with:",
    "      projectPath: apps/desktop",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 1);
  // The issue should point at the step start (`- name: ...`), not the `uses:` line.
  assert.equal(issues[0].line, 2);
  assert.match(issues[0].message, /tauriScript: cargo tauri/);
});

test("reports tauriScript when it is set but not cargo tauri", () => {
  const yaml = [
    "steps:",
    "  - name: Build desktop bundles",
    "    uses: tauri-apps/tauri-action@deadbeef",
    "    with:",
    "      projectPath: apps/desktop",
    "      tauriScript: pnpm tauri",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 1);
  // tauriScript is on line 6.
  assert.equal(issues[0].line, 6);
  assert.match(issues[0].message, /pnpm tauri/);
});

test("allows tauri-action steps when tauriScript is cargo tauri", () => {
  const yaml = [
    "steps:",
    "  - name: Build desktop bundles",
    "    uses: tauri-apps/tauri-action@deadbeef",
    "    with:",
    "      projectPath: apps/desktop",
    "      tauriScript: cargo tauri",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 0);
});

