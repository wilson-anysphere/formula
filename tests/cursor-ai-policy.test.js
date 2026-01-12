import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { checkCursorAiPolicy } from "../scripts/check-cursor-ai-policy.mjs";

const SCRIPT = fileURLToPath(new URL("../scripts/check-cursor-ai-policy.mjs", import.meta.url));
const HAS_GIT = spawnSync("git", ["--version"], { encoding: "utf8" }).status === 0;

async function writeFixtureFile(root, relativePath, contents) {
  const fullPath = path.join(root, relativePath);
  await fs.mkdir(path.dirname(fullPath), { recursive: true });
  await fs.writeFile(fullPath, contents, "utf8");
}

function formatViolations(violations) {
  return violations
    .map((v) => {
      const loc = v.line ? `:${v.line}:${v.column}` : "";
      return `- ${v.file}${loc} [${v.ruleId}] ${v.message}`;
    })
    .join("\n");
}

async function runPolicyApi(rootDir, { maxViolations } = {}) {
  return await checkCursorAiPolicy({ rootDir, maxViolations });
}

function runPolicyCli(rootDir) {
  return spawnSync(process.execPath, [SCRIPT, "--root", rootDir], {
    encoding: "utf8",
  });
}

test("cursor AI policy guard passes on a clean fixture (CLI smoke)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-pass-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'export const answer = 42;\n');
    await writeFixtureFile(tmpRoot, "apps/example/src/main.ts", "export function main() { return 1; }\n");

    const proc = runPolicyCli(tmpRoot);
    assert.equal(proc.status, 0, proc.stderr || proc.stdout);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden provider strings are present (CLI smoke)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'import OpenAI from "openai";\n');

    const proc = runPolicyCli(tmpRoot);
    assert.notEqual(proc.status, 0);
    assert.match(`${proc.stdout}\n${proc.stderr}`, /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden provider strings are present in file paths", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-path-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/openai/src/index.js", "export const x = 1;\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations[0]?.ruleId, "path-openai");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in non-test source files (even without imports)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-source-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'const provider = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations.length, 1);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden provider strings are present in root config files", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-root-config-fail-"));
  try {
    // Root-level config files (package.json, Cargo.toml, etc) are scanned because
    // adding forbidden dependencies there should fail fast.
    await writeFixtureFile(tmpRoot, "package.json", '{ "name": "example", "dependencies": { "openai": "0.0.0" } }\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .env* files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-env-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".env.local", "OPENAI_API_KEY=test\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test(
  "cursor AI policy guard ignores untracked .env* files when scanning a git repo (tracked-files mode)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-untracked-env-pass-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", "export const answer = 42;\n");
      // Untracked .env files are common in developer checkouts; the policy guard should
      // only scan tracked files when running inside a git repo.
      await writeFixtureFile(tmpRoot, ".env.local", "OPENAI_API_KEY=test\n");

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/index.js"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot);
      assert.equal(result.ok, true, formatViolations(result.violations));
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard still scans tracked .env* files in a git repo (tracked-files mode)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-tracked-env-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", "export const answer = 42;\n");
      await writeFixtureFile(tmpRoot, ".env.local", "OPENAI_API_KEY=test\n");

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/index.js", ".env.local"], {
        cwd: tmpRoot,
        encoding: "utf8",
      });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.match(formatViolations(result.violations), /openai/i);
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard scans all git-tracked files (even outside the default scan roots)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-extra-dir-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "extra/notes.js", 'const provider = "OpenAI";\n');

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "extra/notes.js"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.match(formatViolations(result.violations), /openai/i);
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard ignores forbidden strings in excluded docs/ paths (git-tracked mode)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-docs-excluded-pass-"));
    try {
      await writeFixtureFile(tmpRoot, "docs/notes.md", "OpenAI\n");

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "docs/notes.md"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot);
      assert.equal(result.ok, true, formatViolations(result.violations));
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test("cursor AI policy guard ignores forbidden strings in allowlisted AGENTS.md", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-agents-allowlist-pass-"));
  try {
    await writeFixtureFile(tmpRoot, "AGENTS.md", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot);
    assert.equal(result.ok, true, formatViolations(result.violations));
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test(
  "cursor AI policy guard does not allowlist AGENTS.md by basename (only exact root path)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-agents-basename-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "notes/AGENTS.md", "OpenAI\n");

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "notes/AGENTS.md"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.match(formatViolations(result.violations), /openai/i);
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard detects violations in large git repos (git grep path)",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-grep-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", 'const provider = "OpenAI";\n');

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/index.js"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      // Add enough extra tracked entries to cross the `GIT_GREP_MIN_FILES` threshold so the policy guard uses
      // its git-grep fast path (as it does in the real repo). Use `git update-index --index-info` so we don't
      // have to physically create hundreds of files on disk.
      const extraFileCount = 199;
      const blobProc = spawnSync("git", ["hash-object", "-w", "--stdin"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: "export const x = 1;\n",
      });
      assert.equal(blobProc.status, 0, blobProc.stderr);
      const blobSha = String(blobProc.stdout || "").trim();
      assert.ok(blobSha.length > 0);

      let indexInfo = "";
      for (let i = 0; i < extraFileCount; i += 1) {
        indexInfo += `100644 ${blobSha} 0\tpackages/example/src/many/file-${i}.js\n`;
      }
      const indexInfoProc = spawnSync("git", ["update-index", "--index-info"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: indexInfo,
      });
      assert.equal(indexInfoProc.status, 0, indexInfoProc.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.match(formatViolations(result.violations), /openai/i);
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard handles ':' in filenames when using git grep output parsing",
  { skip: !HAS_GIT || process.platform === "win32" || process.platform === "darwin" },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-grep-colon-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/weird:name.js", 'const provider = "OpenAI";\n');

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/weird:name.js"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      // Add enough extra tracked entries to cross the `GIT_GREP_MIN_FILES` threshold so the policy guard uses
      // its git-grep fast path.
      const extraFileCount = 199;
      const blobProc = spawnSync("git", ["hash-object", "-w", "--stdin"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: "export const x = 1;\n",
      });
      assert.equal(blobProc.status, 0, blobProc.stderr);
      const blobSha = String(blobProc.stdout || "").trim();
      assert.ok(blobSha.length > 0);

      let indexInfo = "";
      for (let i = 0; i < extraFileCount; i += 1) {
        indexInfo += `100644 ${blobSha} 0\tpackages/example/src/many/file-${i}.js\n`;
      }
      const indexInfoProc = spawnSync("git", ["update-index", "--index-info"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: indexInfo,
      });
      assert.equal(indexInfoProc.status, 0, indexInfoProc.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.equal(result.violations[0]?.file, "packages/example/src/weird:name.js");
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test("cursor AI policy guard scans Dockerfiles for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-dockerfile-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/Dockerfile", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans Dockerfile.* variants (e.g. Dockerfile.dev)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-dockerfile-variant-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/Dockerfile.dev", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans Makefile.* variants (e.g. Makefile.dev)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-makefile-variant-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/Makefile.dev", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans extensionless dotfiles (e.g. .gitignore)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-dotfile-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".gitignore", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .gitkeep files", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-gitkeep-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/.gitkeep", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test(
  "cursor AI policy guard rejects symlinked files (avoid scan bypass)",
  { skip: process.platform === "win32" },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-symlink-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", "export const answer = 42;\n");

      const linkPath = path.join(tmpRoot, "packages", "example", "src", "link");
      // Use a dangling symlink: the target does not need to exist for lstat() to detect it.
      await fs.symlink("does-not-exist", linkPath);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.equal(result.violations.length, 1);
      assert.equal(result.violations[0]?.ruleId, "symlink");
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard rejects tracked symlinks in a git repo (tracked-files mode)",
  { skip: !HAS_GIT || process.platform === "win32" },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-symlink-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", "export const answer = 42;\n");

      const linkPath = path.join(tmpRoot, "packages", "example", "src", "link");
      await fs.symlink("does-not-exist", linkPath);

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/index.js", "packages/example/src/link"], {
        cwd: tmpRoot,
        encoding: "utf8",
      });
      assert.equal(add.status, 0, add.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.equal(result.violations.length, 1);
      assert.equal(result.violations[0]?.ruleId, "symlink");
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard rejects git submodules (gitlink entries) in a git repo",
  { skip: !HAS_GIT },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-submodule-fail-"));
    try {
      await writeFixtureFile(tmpRoot, "packages/example/src/index.js", "export const answer = 42;\n");

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const add = spawnSync("git", ["add", "packages/example/src/index.js"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(add.status, 0, add.stderr);

      // Create a tracked gitlink entry without needing a real submodule checkout.
      // `160000` is the git mode for submodules (gitlinks).
      const gitlinkSha = "0123456789012345678901234567890123456789";
      const addGitlink = spawnSync(
        "git",
        ["update-index", "--add", "--cacheinfo", "160000", gitlinkSha, "zz-submodule"],
        { cwd: tmpRoot, encoding: "utf8" },
      );
      assert.equal(addGitlink.status, 0, addGitlink.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
      assert.equal(result.ok, false);
      assert.equal(result.violations[0]?.ruleId, "git-submodule");
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test(
  "cursor AI policy guard rejects tracked symlinks in a large git repo (git grep path)",
  { skip: !HAS_GIT || process.platform === "win32" },
  async () => {
    const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-git-large-symlink-fail-"));
    try {
      const linkPath = path.join(tmpRoot, "packages", "example", "src", "link");
      await fs.mkdir(path.dirname(linkPath), { recursive: true });
      await fs.symlink("does-not-exist", linkPath);

      const init = spawnSync("git", ["init"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(init.status, 0, init.stderr);
      const addSymlink = spawnSync("git", ["add", "packages/example/src/link"], { cwd: tmpRoot, encoding: "utf8" });
      assert.equal(addSymlink.status, 0, addSymlink.stderr);

      // Add enough extra tracked entries to cross the `GIT_GREP_MIN_FILES` threshold so the policy guard uses
      // its git-grep fast path (as it does in the real repo). Use `git update-index --index-info` so we don't
      // have to physically create hundreds of files on disk.
      const extraFileCount = 198;
      const blobProc = spawnSync("git", ["hash-object", "-w", "--stdin"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: "export const x = 1;\n",
      });
      assert.equal(blobProc.status, 0, blobProc.stderr);
      const blobSha = String(blobProc.stdout || "").trim();
      assert.ok(blobSha.length > 0);

      let indexInfo = "";
      for (let i = 0; i < extraFileCount; i += 1) {
        indexInfo += `100644 ${blobSha} 0\tpackages/example/src/many/file-${i}.js\n`;
      }
      const indexInfoProc = spawnSync("git", ["update-index", "--index-info"], {
        cwd: tmpRoot,
        encoding: "utf8",
        input: indexInfo,
      });
      assert.equal(indexInfoProc.status, 0, indexInfoProc.stderr);

      // Also add a submodule gitlink entry to ensure the large-repo git-grep path still
      // rejects submodules without relying on filesystem traversal.
      const gitlinkSha = "0123456789012345678901234567890123456789";
      const addGitlink = spawnSync(
        "git",
        ["update-index", "--add", "--cacheinfo", "160000", gitlinkSha, "zz-submodule"],
        { cwd: tmpRoot, encoding: "utf8" },
      );
      assert.equal(addGitlink.status, 0, addGitlink.stderr);

      const result = await runPolicyApi(tmpRoot, { maxViolations: 2 });
      assert.equal(result.ok, false);
      assert.equal(result.violations.length, 2);
      const ruleIds = new Set(result.violations.map((v) => v.ruleId));
      assert.deepEqual(ruleIds, new Set(["symlink", "git-submodule"]));
    } finally {
      await fs.rm(tmpRoot, { recursive: true, force: true });
    }
  },
);

test("cursor AI policy guard scans markdown readmes for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-readme-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/README.md", "This should not mention OpenAI.\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans html files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-html-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "apps/example/index.html", "<!-- OpenAI -->\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans css files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-css-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "apps/example/src/styles.css", "/* OpenAI */\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans snapshot files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-snap-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/__tests__/__snapshots__/thing.snap", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans SQL files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-sql-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/migrations/0001_init.sql", "-- OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .txt files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-txt-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/NOTICE.txt", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans PowerShell scripts for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-ps1-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "tools/example/run.ps1", "Write-Host OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans XML files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-xml-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/tests/fixtures/example.xml", "<!-- OpenAI -->\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans TSV files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-tsv-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/src/data.tsv", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans jsonl files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-jsonl-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/tests/cases.jsonl", '{ "provider": "OpenAI" }\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .bas files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-bas-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/tests/fixtures/macro.bas", "' OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .m files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-m-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/test/golden/query.m", "// OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans wgsl files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-wgsl-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/shader.wgsl", "// OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans plist files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-plist-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "apps/example/entitlements.plist", "<!-- OpenAI -->\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .b64 files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-b64-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "tools/example/fixture.xlsx.b64", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .base64 files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-base64-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "crates/example/tests/fixtures/fixture.xlsx.base64", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .pem files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-pem-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/localhost-cert.pem", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .key files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-key-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/localhost.key", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .crt files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-crt-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/localhost.crt", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans ini files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-ini-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/config.ini", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans conf files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-conf-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/config.conf", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans properties files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-properties-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/config.properties", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans Kotlin script files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-kts-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/build.gradle.kts", "// OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans gradle files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-gradle-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "services/example/build.gradle", "// OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans Windows batch scripts for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-bat-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/run.bat", "rem OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans Windows cmd scripts for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-cmd-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/run.cmd", "rem OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans PowerShell module files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-psm1-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/example.psm1", "# OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans PowerShell module manifest files for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-psd1-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/example.psd1", "# OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans extensionless config files named `config`", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-config-basename-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/config", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans lockfiles (Cargo.lock) for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-cargo-lock-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "Cargo.lock", '[[package]]\nname = "openai"\nversion = "0.0.0"\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans pnpm-lock.yaml for provider strings", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-pnpm-lock-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "pnpm-lock.yaml", "openai: 0.0.0\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans scripts/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-scripts-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "scripts/example.mjs", 'import OpenAI from "openai";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans shared/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-shared-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "shared/example.js", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans extensions/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-extensions-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "extensions/example/dist/extension.js", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans python/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-python-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "python/example.py", 'provider = "OpenAI"\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .github/workflows by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-github-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".github/workflows/ci.yml", 'name: "OpenAI"\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .cargo/ config by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-cargo-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".cargo/config.toml", 'openai = "1"\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans patches/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-patches-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "patches/example.patch", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans fixtures/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-fixtures-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "fixtures/example.json", '{ "provider": "OpenAI" }\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans security/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-security-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "security/allowlist/node-audit.txt", "OpenAI\n");

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .vscode/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-vscode-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".vscode/settings.json", '{ "x": "OpenAI" }\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans .devcontainer/ directory by default", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-devcontainer-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, ".devcontainer/devcontainer.json", '{ "x": "OpenAI" }\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /openai/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when forbidden strings appear in unrelated unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.test.js", 'const x = "anthropic";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.match(formatViolations(result.violations), /anthropic/i);
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in unrelated unit tests (even without imports)", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-test-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.test.js", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations[0]?.ruleId, "openai-in-test");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans repo-level tests/ directory for unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-tests-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "tests/unrelated.test.js", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations[0]?.ruleId, "openai-in-test");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard scans repo-level test/ directory for unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-dir-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "test/unrelated.test.js", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations[0]?.ruleId, "openai-in-test");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard fails when OpenAI appears in *.vitest.* unit tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-openai-vitest-fail-"));
  try {
    await writeFixtureFile(tmpRoot, "packages/example/src/something.vitest.ts", 'const x = "OpenAI";\n');

    const result = await runPolicyApi(tmpRoot, { maxViolations: 1 });
    assert.equal(result.ok, false);
    assert.equal(result.violations[0]?.ruleId, "openai-in-test");
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});

test("cursor AI policy guard allows forbidden strings in the guard's own tests", async () => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "cursor-ai-policy-test-pass-"));
  try {
    // This file name intentionally matches the allowlist rule for policy guard tests.
    await writeFixtureFile(
      tmpRoot,
      "packages/example/src/cursor-ai-policy.test.js",
      'const fixtures = ["openai", "anthropic", "ollama", "formula:openaiApiKey"];\n',
    );

    const result = await runPolicyApi(tmpRoot);
    assert.equal(result.ok, true, formatViolations(result.violations));
  } finally {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  }
});
