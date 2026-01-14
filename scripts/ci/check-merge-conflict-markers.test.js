import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const sourceScriptPath = path.join(repoRoot, "scripts", "ci", "check-merge-conflict-markers.sh");
const scriptContents = readFileSync(sourceScriptPath, "utf8");

const bashProbe = spawnSync("bash", ["--version"], { encoding: "utf8" });
const hasBash = !bashProbe.error && bashProbe.status === 0;

const gitProbe = spawnSync("git", ["--version"], { encoding: "utf8" });
const hasGit = !gitProbe.error && gitProbe.status === 0;

const canRun = hasBash && hasGit;

/**
 * Runs the merge conflict marker guard in a temporary git repo.
 * @param {Record<string, string>} files
 */
function run(files) {
  const tmpdir = mkdtempSync(path.join(os.tmpdir(), "formula-merge-conflicts-"));
  try {
    let proc = spawnSync("git", ["init"], { cwd: tmpdir, encoding: "utf8" });
    assert.equal(proc.status, 0, proc.stderr);

    // Install the script under test in the temp repo so it runs `git grep` against that repo.
    const scriptPath = path.join(tmpdir, "check-merge-conflict-markers.sh");
    writeFileSync(scriptPath, scriptContents, "utf8");

    for (const [name, content] of Object.entries(files)) {
      const filePath = path.join(tmpdir, name);
      writeFileSync(filePath, `${content}\n`, "utf8");
      proc = spawnSync("git", ["add", name], { cwd: tmpdir, encoding: "utf8" });
      assert.equal(proc.status, 0, proc.stderr);
    }

    return spawnSync("bash", [scriptPath], { cwd: tmpdir, encoding: "utf8" });
  } finally {
    rmSync(tmpdir, { recursive: true, force: true });
  }
}

test("passes when no conflict markers are present", { skip: !canRun }, () => {
  const proc = run({
    "hello.txt": "hello world",
  });
  assert.equal(proc.status, 0, proc.stderr);
});

test("fails when standard conflict markers are present", { skip: !canRun }, () => {
  const proc = run({
    "conflict.txt": `
<<<<<<< ours
hello
=======
world
>>>>>>> theirs
`,
  });
  assert.notEqual(proc.status, 0);
  assert.match(proc.stderr, /Merge conflict markers detected/i);
});

test("does not match long separator lines (avoid false positives)", { skip: !canRun }, () => {
  const proc = run({
    "doc.md": `
This is a doc.
==========
Not a conflict marker because it's 10 '=' characters.
`,
  });
  assert.equal(proc.status, 0, proc.stderr);
});

