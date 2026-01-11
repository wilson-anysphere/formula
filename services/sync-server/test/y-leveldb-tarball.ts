import { execFile } from "node:child_process";
import { createRequire } from "node:module";
import { mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

function execFileAsync(
  command: string,
  args: string[],
  opts: Parameters<typeof execFile>[2]
): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(command, args, opts, (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}

export async function loadYLeveldbFromTarball(t: {
  after: (cb: () => void | Promise<void>) => void;
}): Promise<{
  LeveldbPersistence: new (location: string, opts?: any) => any;
  keyEncoding: any;
}> {
  const testDir = path.dirname(fileURLToPath(import.meta.url));
  const repoRoot = path.resolve(testDir, "../../..");
  const tarballPath = path.join(repoRoot, "y-leveldb-0.1.2.tgz");

  const extractDir = await mkdtemp(path.join(tmpdir(), "y-leveldb-"));
  t.after(async () => {
    await rm(extractDir, { recursive: true, force: true });
  });

  await execFileAsync("tar", ["-xzf", tarballPath, "-C", extractDir], {
    stdio: "ignore",
  });

  const pkgRoot = path.join(extractDir, "package");

  // y-leveldb unconditionally imports `level`. Provide a pure JS `level` module
  // that re-exports `level-mem` so tests don't need native LevelDB bindings.
  const require = createRequire(import.meta.url);
  const levelMemEntry = require.resolve("level-mem");
  const levelMemUrl = pathToFileURL(levelMemEntry).href;

  const linkDependency = async (linkName: string, pkgJsonPath: string) => {
    const targetDir = path.dirname(pkgJsonPath);
    const linkPath = path.join(pkgRoot, "node_modules", linkName);
    try {
      await symlink(
        targetDir,
        linkPath,
        process.platform === "win32" ? "junction" : "dir"
      );
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "EEXIST") throw err;
    }
  };

  // Ensure y-leveldb can resolve its runtime deps when extracted into a temp dir.
  await mkdir(path.join(pkgRoot, "node_modules"), { recursive: true });
  await linkDependency("yjs", require.resolve("yjs/package.json"));
  const ywsRequire = createRequire(require.resolve("y-websocket/package.json"));
  await linkDependency("lib0", ywsRequire.resolve("lib0/package.json"));

  const levelStubDir = path.join(pkgRoot, "node_modules", "level");
  await mkdir(levelStubDir, { recursive: true });
  await writeFile(
    path.join(levelStubDir, "package.json"),
    JSON.stringify(
      {
        name: "level",
        version: "0.0.0-test",
        type: "module",
        main: "./index.js",
      },
      null,
      2
    )
  );
  await writeFile(
    path.join(levelStubDir, "index.js"),
    `import levelMem from ${JSON.stringify(levelMemUrl)};\nexport default levelMem;\n`
  );

  const yLeveldbUrl = pathToFileURL(path.join(pkgRoot, "src", "y-leveldb.js")).href;
  const mod = (await import(yLeveldbUrl)) as any;

  return {
    LeveldbPersistence: mod.LeveldbPersistence,
    keyEncoding: mod.keyEncoding,
  };
}

