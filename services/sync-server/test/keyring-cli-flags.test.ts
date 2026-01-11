import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { KeyRing } from "../../../packages/security/crypto/keyring.js";

async function runKeyringCli(opts: {
  args: string[];
  stdin?: string;
}): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  const serviceDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const nodeWithTsx = path.join(serviceDir, "scripts", "node-with-tsx.mjs");
  const entry = path.join(serviceDir, "src", "keyring-cli.ts");

  return await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [nodeWithTsx, entry, ...opts.args], {
      cwd: serviceDir,
      stdio: ["pipe", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (d) => {
      stdout += d;
    });
    child.stderr.on("data", (d) => {
      stderr += d;
    });

    child.on("error", reject);

    const timeout = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error("Timed out waiting for keyring-cli to exit"));
    }, 10_000);
    timeout.unref();

    child.on("exit", (code) => {
      clearTimeout(timeout);
      resolve({ stdout, stderr, exitCode: code ?? 0 });
    });

    if (opts.stdin !== undefined) {
      child.stdin.write(opts.stdin);
      child.stdin.end();
    } else {
      child.stdin.end();
    }
  });
}

test("keyring-cli generate --out writes valid JSON and defaults to 0600 perms", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "sync-server-keyring-"));
  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const outPath = path.join(dir, "keyring.json");

  const result = await runKeyringCli({
    args: ["generate", "--out", outPath],
  });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(result.stdout, "");

  const json = JSON.parse(await readFile(outPath, "utf8"));
  assert.ok(KeyRing.fromJSON(json));

  if (process.platform !== "win32") {
    const mode = (await stat(outPath)).mode & 0o777;
    assert.equal(mode, 0o600);
  }
});

test("keyring-cli validate supports stdin via --in -", async () => {
  const json = JSON.stringify(KeyRing.create().toJSON());

  const result = await runKeyringCli({
    args: ["validate", "--in", "-"],
    stdin: json,
  });

  assert.equal(result.exitCode, 0, result.stderr);
  const summary = JSON.parse(result.stdout) as {
    currentVersion: number;
    availableVersions: number[];
  };
  assert.deepEqual(summary, { currentVersion: 1, availableVersions: [1] });
});

test("keyring-cli rotate supports stdin via --in - and writes rotated JSON to stdout", async () => {
  const original = KeyRing.create().toJSON();
  const originalKey = original.keys["1"];

  const result = await runKeyringCli({
    args: ["rotate", "--in", "-"],
    stdin: JSON.stringify(original),
  });

  assert.equal(result.exitCode, 0, result.stderr);
  const rotated = JSON.parse(result.stdout) as {
    currentVersion: number;
    keys: Record<string, string>;
  };
  assert.equal(rotated.currentVersion, 2);
  assert.equal(rotated.keys["1"], originalKey);
  assert.ok(typeof rotated.keys["2"] === "string" && rotated.keys["2"].length > 0);
});

