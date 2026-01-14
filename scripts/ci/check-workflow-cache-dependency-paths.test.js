import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const workflowsDir = path.join(repoRoot, ".github", "workflows");

function countLeadingSpaces(text) {
  let i = 0;
  while (i < text.length && text[i] === " ") i++;
  return i;
}

function stripSurroundingQuotes(value) {
  const trimmed = value.trim();
  if (trimmed.length < 2) return trimmed;

  const first = trimmed[0];
  const last = trimmed[trimmed.length - 1];
  if ((first === "'" || first === '"') && last === first) {
    return trimmed.slice(1, -1);
  }

  return trimmed;
}

function findBadCacheDependencyPathGlobs(workflowName, workflowText) {
  // Guard against accidentally using `**/pnpm-lock.yaml` for setup-node cache discovery. The
  // actions glob implementation scans the entire repository, which can get slow once `target/` or
  // other build outputs exist.
  //
  // GitHub Actions YAML frequently uses block scalars (`run: |`) which can contain arbitrary text
  // that looks like YAML keys. We need to ignore those blocks, while still supporting the
  // *intended* multi-line form of `cache-dependency-path: |` (one path per line).
  const badGlob = "**/pnpm-lock.yaml";
  const bad = [];
  const lines = workflowText.split(/\r?\n/);

  /** @type {null | { key: string, contentIndent: number | null }} */
  let blockScalar = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (blockScalar) {
      // Blank/whitespace-only lines are always part of the scalar.
      if (line.trim() === "") continue;

      const indent = countLeadingSpaces(line);
      if (blockScalar.contentIndent === null) {
        blockScalar.contentIndent = indent;
      } else if (indent < blockScalar.contentIndent) {
        // Exited the scalar; re-process this line in normal mode.
        blockScalar = null;
        i--;
        continue;
      }

      if (blockScalar.key === "cache-dependency-path") {
        const value = stripSurroundingQuotes(line.trim());
        if (value === badGlob) {
          bad.push(`${workflowName}:${i + 1}:${line.trim()}`);
        }
      }

      continue;
    }

    const blockScalarStart = line.match(
      /^\s*(?:-\s*)?(?<key>[A-Za-z0-9_-]+):\s*(?<indicator>[|>][0-9+-]*)\s*(?:#.*)?$/,
    );
    if (blockScalarStart?.groups) {
      blockScalar = { key: blockScalarStart.groups.key, contentIndent: null };
      continue;
    }

    const cacheDependencyPath = line.match(
      /^\s*(?:-\s*)?cache-dependency-path:\s*(?<value>.*?)(?:\s+#.*)?$/,
    );
    if (cacheDependencyPath?.groups) {
      const value = stripSurroundingQuotes(cacheDependencyPath.groups.value);
      if (value === badGlob) {
        bad.push(`${workflowName}:${i + 1}:${line.trim()}`);
      }
    }
  }

  return bad;
}

test("cache-dependency-path scanner ignores YAML block scalars for unrelated keys", () => {
  const text = `
jobs:
  test:
    steps:
      - run: |
          cache-dependency-path: **/pnpm-lock.yaml
        name: Example
`;

  assert.deepEqual(findBadCacheDependencyPathGlobs("example.yml", text), []);
});

test("cache-dependency-path scanner ignores YAML block scalars with chomping/indent indicators in any order", () => {
  const text = `
jobs:
  test:
    steps:
      - run: |-2
          cache-dependency-path: **/pnpm-lock.yaml
        name: Example
`;

  assert.deepEqual(findBadCacheDependencyPathGlobs("example.yml", text), []);
});

test("cache-dependency-path scanner detects bad globs inside cache-dependency-path block scalars", () => {
  const text = `
steps:
  - uses: actions/setup-node@v4
    with:
      cache-dependency-path: |
        pnpm-lock.yaml
        **/pnpm-lock.yaml
`;

  assert.deepEqual(findBadCacheDependencyPathGlobs("example.yml", text), [
    "example.yml:7:**/pnpm-lock.yaml",
  ]);
});

test("cache-dependency-path scanner detects bad globs inside cache-dependency-path block scalars using |-2", () => {
  const text = `
steps:
  - uses: actions/setup-node@v4
    with:
      cache-dependency-path: |-2
        pnpm-lock.yaml
        **/pnpm-lock.yaml
`;

  assert.deepEqual(findBadCacheDependencyPathGlobs("example.yml", text), [
    "example.yml:7:**/pnpm-lock.yaml",
  ]);
});

test("workflows avoid recursive pnpm-lock cache-dependency-path globs (perf guardrail)", () => {
  const bad = [];
  const entries = readdirSync(workflowsDir, { withFileTypes: true })
    .filter((ent) => ent.isFile() && (ent.name.endsWith(".yml") || ent.name.endsWith(".yaml")))
    .map((ent) => ent.name)
    .sort();

  for (const name of entries) {
    const filePath = path.join(workflowsDir, name);
    bad.push(...findBadCacheDependencyPathGlobs(name, readFileSync(filePath, "utf8")));
  }

  assert.deepEqual(
    bad,
    [],
    `Found recursive pnpm lock cache-dependency-path globs (use "pnpm-lock.yaml" instead):\n${bad.join("\n")}`,
  );
});
