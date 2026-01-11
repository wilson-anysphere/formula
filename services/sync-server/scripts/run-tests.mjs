import { spawn } from "node:child_process";
import { readdir } from "node:fs/promises";
import path from "node:path";

async function runTestFile(filePath) {
  const nodeWithTsx = path.join(process.cwd(), "scripts", "node-with-tsx.mjs");

  await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [nodeWithTsx, "--test", filePath], {
      stdio: "inherit",
    });

    child.on("error", reject);
    child.on("exit", (code, signal) => {
      if (signal) {
        reject(new Error(`Test process terminated by signal ${signal}`));
        return;
      }
      if (code === 0) resolve();
      else reject(new Error(`Test process exited with code ${code ?? 0}`));
    });
  });
}

const testDir = path.join(process.cwd(), "test");
const entries = await readdir(testDir, { withFileTypes: true });
const files = entries
  .filter((entry) => entry.isFile() && entry.name.endsWith(".test.ts"))
  .map((entry) => path.join("test", entry.name))
  .sort();

for (const file of files) {
  await runTestFile(file);
}
