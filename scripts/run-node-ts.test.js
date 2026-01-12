import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

function hasTypeScriptDependency() {
  try {
    require.resolve("typescript", { paths: [repoRoot] });
    return true;
  } catch {
    return false;
  }
}

function getBuiltInTypeScriptSupport() {
  const flagProbe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  if (flagProbe.status === 0) {
    return { enabled: true, args: ["--experimental-strip-types"] };
  }

  const tmpFile = path.join(os.tmpdir(), `formula-strip-types-probe.${process.pid}.${Date.now()}.ts`);
  try {
    writeFileSync(
      tmpFile,
      [
        "export const x: number = 1;",
        "if (x !== 1) throw new Error('strip-types probe failed');",
        "",
      ].join("\n"),
      "utf8",
    );
    const fileUrl = pathToFileURL(tmpFile).href;
    const nativeProbe = spawnSync(process.execPath, ["--input-type=module", "-e", `import ${JSON.stringify(fileUrl)};`], {
      stdio: "ignore",
    });
    if (nativeProbe.status === 0) {
      return { enabled: true, args: [] };
    }
  } catch {
    // ignore
  } finally {
    rmSync(tmpFile, { force: true });
  }

  return { enabled: false, args: [] };
}

const builtInTypeScript = getBuiltInTypeScriptSupport();
const canExecuteTypeScript = builtInTypeScript.enabled || hasTypeScriptDependency();

test("run-node-ts can execute a TS entrypoint", { skip: !canExecuteTypeScript }, () => {
  const child = spawnSync(
    process.execPath,
    ["scripts/run-node-ts.mjs", "scripts/__fixtures__/run-node-ts/entry.ts"],
    { cwd: repoRoot, encoding: "utf8" },
  );

  assert.equal(
    child.status,
    0,
    `run-node-ts exited with ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.equal(child.stdout.trim(), "42,42,42");
});

test("run-node-ts strips pnpm `--` delimiters before forwarding argv", { skip: !canExecuteTypeScript }, () => {
  const child = spawnSync(
    process.execPath,
    [
      "scripts/run-node-ts.mjs",
      "scripts/__fixtures__/run-node-ts/argv-check.ts",
      "--",
      "--some-arg",
    ],
    { cwd: repoRoot, encoding: "utf8" },
  );

  assert.equal(
    child.status,
    0,
    `run-node-ts exited with ${child.status}\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
  );
  assert.equal(child.stdout.trim(), "ok");
});
