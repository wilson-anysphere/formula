import { createHash } from "node:crypto";
import { createWriteStream } from "node:fs";
import { createReadStream } from "node:fs";
import { mkdir, rename, stat } from "node:fs/promises";
import path from "node:path";
import { pipeline } from "node:stream/promises";
import { fileURLToPath } from "node:url";

const PYODIDE_VERSION = "0.25.1";
const PYODIDE_CDN_BASE_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pyodideDir = path.resolve(__dirname, "../public/pyodide", `v${PYODIDE_VERSION}`, "full");

const requiredFiles = {
  "pyodide.js": "b9cb64a73cc4127eef7cdc75f0cd8307db9e90e93b66b1b6a789319511e1937c",
  "pyodide.asm.js": "512042dfdd406971c6fc920b6932e1a8eb5dd2ab3521aa89a020980e4a08bd4b",
  "pyodide.asm.wasm": "aa920641c032c3db42eb1fb018eec611dbef96f0fa4dbdfa6fe3cb1b335aed3c",
  "python_stdlib.zip": "52866039fa3097e549649b9a62ffae8a1125f01ace7b2d077f34e3cbaff8d0ca",
  "pyodide-lock.json": "6526dae570ab7db75019fe2c7ccc6b7b82765c56417a498a7b57e1aaebec39f5",
};

async function fileExists(filePath) {
  try {
    const info = await stat(filePath);
    return info.isFile() && info.size > 0;
  } catch {
    return false;
  }
}

async function sha256File(filePath) {
  const hash = createHash("sha256");
  await pipeline(createReadStream(filePath), hash);
  return hash.digest("hex");
}

async function hasExpectedHash(filePath, expectedSha256) {
  if (!(await fileExists(filePath))) return false;
  const digest = await sha256File(filePath);
  return digest === expectedSha256;
}

async function downloadFile(url, destPath) {
  const res = await fetch(url);
  if (!res.ok || !res.body) {
    throw new Error(`Failed to download ${url} (${res.status} ${res.statusText})`);
  }

  await mkdir(path.dirname(destPath), { recursive: true });

  // Download to a temp file, then move into place to avoid leaving corrupt
  // partial files if the process is interrupted.
  const tmpPath = `${destPath}.tmp-${process.pid}-${Date.now()}`;
  await pipeline(res.body, createWriteStream(tmpPath));
  await rename(tmpPath, destPath);
}

async function main() {
  await mkdir(pyodideDir, { recursive: true });

  const missingFiles = [];
  for (const [fileName, expectedSha256] of Object.entries(requiredFiles)) {
    const destPath = path.join(pyodideDir, fileName);
    if (!(await hasExpectedHash(destPath, expectedSha256))) {
      missingFiles.push(fileName);
    }
  }

  if (missingFiles.length === 0) return;

  console.log(`Downloading Pyodide v${PYODIDE_VERSION} assets to ${path.relative(process.cwd(), pyodideDir)}`);

  const downloads = missingFiles.map(async (fileName) => {
    const destPath = path.join(pyodideDir, fileName);
    const url = `${PYODIDE_CDN_BASE_URL}${fileName}`;
    console.log(`- ${fileName}`);
    await downloadFile(url, destPath);
    const expectedSha256 = requiredFiles[fileName];
    if (!(await hasExpectedHash(destPath, expectedSha256))) {
      throw new Error(`Downloaded ${fileName} but sha256 did not match expected ${expectedSha256}`);
    }
  });

  await Promise.all(downloads);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
