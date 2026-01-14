import assert from "node:assert/strict";
import test from "node:test";
import { extractPinnedCliVersionsFromWorkflow, findTauriActionScriptIssues } from "./check-tauri-cli-version.mjs";

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

test("ignores tauri-action YAML-ish strings inside block scalars", () => {
  const yaml = [
    "steps:",
    "  - name: Run script",
    "    run: |",
    "      # This is script content and should not be interpreted as workflow YAML.",
    "      - uses: tauri-apps/tauri-action@deadbeef",
    "      tauriScript: pnpm tauri",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 0);
});

test("does not treat tauriScript occurrences inside non-run block scalars as valid", () => {
  const yaml = [
    "steps:",
    "  - uses: tauri-apps/tauri-action@deadbeef",
    "    with:",
    "      projectPath: apps/desktop",
    "      args: |",
    "        tauriScript: cargo tauri",
  ].join("\n");

  const issues = findTauriActionScriptIssues(yaml);
  assert.equal(issues.length, 1);
  assert.equal(issues[0].line, 2);
  assert.match(issues[0].message, /tauriScript: cargo tauri/);
});

test("ignores TAURI_CLI_VERSION strings inside YAML block scalars", () => {
  const yaml = [
    "name: CI",
    "jobs:",
    "  build:",
    "    runs-on: ubuntu-24.04",
    "    steps:",
    "      - run: |",
    "          # Script content; should not count as workflow YAML.",
    "          TAURI_CLI_VERSION: 9.9.9",
    "          echo ok",
  ].join("\n");

  assert.deepEqual(extractPinnedCliVersionsFromWorkflow(yaml), []);
});
