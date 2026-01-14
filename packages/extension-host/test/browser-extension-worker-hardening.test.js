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

async function writeFiles(rootDir, files) {
  for (const [relPath, contents] of Object.entries(files)) {
    const fullPath = path.join(rootDir, relPath);
    await fs.mkdir(path.dirname(fullPath), { recursive: true });
    await fs.writeFile(fullPath, contents, "utf8");
  }
}

function withTimeout(promise, timeoutMs, message) {
  /** @type {ReturnType<typeof setTimeout> | null} */
  let timeout = null;
  const timeoutPromise = new Promise((_, reject) => {
    timeout = setTimeout(() => reject(new Error(message)), timeoutMs);
    timeout.unref?.();
  });
  return Promise.race([promise, timeoutPromise]).finally(() => {
    if (timeout) clearTimeout(timeout);
  });
}

async function activateExtensionWorker({
  mainUrl,
  extensionPath,
  sandbox,
  apiHandler
}) {
  const extensionWorkerUrl = pathToFileURL(
    path.resolve(__dirname, "../src/browser/extension-worker.mjs")
  ).href;

  const wrapperDir = await createTempDir("formula-ext-worker-wrapper-");
  const loaderPath = path.join(wrapperDir, "loader.mjs");
  const extensionApiUrl = pathToFileURL(
    path.resolve(__dirname, "../../extension-api/index.mjs")
  ).href;
  await fs.writeFile(
    loaderPath,
    `export async function resolve(specifier, context, nextResolve) {
  if (specifier === "@formula/extension-api" || specifier === "formula") {
    return { url: ${JSON.stringify(extensionApiUrl)}, shortCircuit: true };
  }
  return nextResolve(specifier, context);
}
`,
    "utf8"
  );

  const wrapperPath = path.join(wrapperDir, "wrapper.mjs");
  await fs.writeFile(
    wrapperPath,
    `import { parentPort } from "node:worker_threads";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

globalThis.self = globalThis;

// Stub browser-only networking primitives so the extension worker can hard-disable them
// even when running in Node's worker_threads test environment.
if (typeof globalThis.XMLHttpRequest !== "function") {
  globalThis.XMLHttpRequest = function XMLHttpRequest() {};
}
if (typeof globalThis.EventSource !== "function") {
  globalThis.EventSource = function EventSource() {};
}
if (typeof globalThis.WebTransport !== "function") {
  globalThis.WebTransport = function WebTransport() {};
}
if (typeof globalThis.RTCPeerConnection !== "function") {
  globalThis.RTCPeerConnection = function RTCPeerConnection() {};
}
if (!globalThis.navigator || typeof globalThis.navigator !== "object") {
  globalThis.navigator = {};
}
if (typeof globalThis.navigator.sendBeacon !== "function") {
  globalThis.navigator.sendBeacon = () => true;
}
if (typeof globalThis.Worker !== "function") {
  globalThis.Worker = function Worker() {};
}
if (typeof globalThis.SharedWorker !== "function") {
  globalThis.SharedWorker = function SharedWorker() {};
}
if (typeof globalThis.importScripts !== "function") {
  globalThis.importScripts = function importScripts() {};
}

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

let aliasBigBytes = null;

globalThis.fetch = async (input, init) => {
  const url = typeof input === "string" ? input : input?.url ?? String(input);
  if (url.endsWith("/redirect-outside.mjs")) {
    return {
      ok: true,
      status: 200,
      redirected: true,
      url: "https://evil.invalid/evil.mjs",
      async arrayBuffer() {
        return new ArrayBuffer(0);
      }
    };
  }
  if (url.endsWith("/whitespace-url.mjs")) {
    const data = await readFile(fileURLToPath(url));
    return {
      ok: true,
      status: 200,
      redirected: false,
      url: "  " + url + "  ",
      async arrayBuffer() {
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
      }
    };
  }
  if (/\\/alias-big\\d+\\.mjs$/.test(url)) {
    const targetUrl = url.replace(/\\/alias-big\\d+\\.mjs$/, "/target-big.mjs");
    if (!aliasBigBytes) {
      const bigComment = "a".repeat(240 * 1024);
      aliasBigBytes = new TextEncoder().encode(\`/*\${bigComment}*/\\nexport const value = 1;\\n\`);
    }
    return {
      ok: true,
      status: 200,
      redirected: true,
      url: targetUrl,
      async arrayBuffer() {
        return aliasBigBytes.buffer.slice(aliasBigBytes.byteOffset, aliasBigBytes.byteOffset + aliasBigBytes.byteLength);
      }
    };
  }
  if (/\\/alias\\d+\\.mjs$/.test(url)) {
    const targetUrl = url.replace(/\\/alias\\d+\\.mjs$/, "/target.mjs");
    const bytes = new TextEncoder().encode("export const value = 1;\\n");
    return {
      ok: true,
      status: 200,
      redirected: true,
      url: targetUrl,
      async arrayBuffer() {
        return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      }
    };
  }
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

  const allowedFlags =
    process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
      ? process.allowedNodeEnvironmentFlags
      : new Set();
  const loaderUrl = pathToFileURL(loaderPath).href;
  const supportsRegister = typeof require("node:module")?.register === "function";
  /** @type {string[]} */
  const execArgv = [];

  if (supportsRegister && allowedFlags.has("--import")) {
    const registerScript = `import { register } from "node:module"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    execArgv.push("--import", dataUrl);
  } else if (allowedFlags.has("--loader")) {
    execArgv.push("--loader", loaderUrl);
  } else {
    execArgv.push("--experimental-loader", loaderUrl);
  }
  if (allowedFlags.has("--disable-warning")) {
    execArgv.unshift("--disable-warning=ExperimentalWarning");
  } else if (allowedFlags.has("--no-warnings")) {
    execArgv.unshift("--no-warnings");
  }

  const worker = new Worker(pathToFileURL(wrapperPath), {
    type: "module",
    execArgv
  });
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
      10000,
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

test("extension-worker: rejects static data: URL imports", async () => {
  const dir = await createTempDir("formula-ext-worker-data-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "data:text/javascript,export default 1";\nexport async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /URL\/protocol imports are not allowed/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects bare module specifiers", async () => {
  const dir = await createTempDir("formula-ext-worker-bare-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "lodash";\nexport async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /Disallowed import specifier 'lodash'.*only relative imports/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects absolute-path imports", async () => {
  const dir = await createTempDir("formula-ext-worker-absolute-path-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "/evil.mjs";\nexport async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /absolute imports are not allowed/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects disallowed imports in dependency modules", async () => {
  const dir = await createTempDir("formula-ext-worker-dep-disallowed-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "./dep.mjs";\nexport async function activate() {}\n`,
      "dep.mjs": `import "https://evil.invalid/evil.mjs";\nexport const x = 1;\n`
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

test("extension-worker: rejects dynamic import() in dependency modules", async () => {
  const dir = await createTempDir("formula-ext-worker-dep-dynamic-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "./dep.mjs";\nexport async function activate() {}\n`,
      "dep.mjs": `export async function run() {\n  return import("https://evil.invalid/evil.mjs");\n}\n`
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

test("extension-worker: rejects relative imports that escape the extension base URL", async () => {
  const dir = await createTempDir("formula-ext-worker-escape-relative-import-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "../outside.mjs";\nexport async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /resolved outside the extension base URL/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects module graphs that exceed maxModules", async () => {
  const dir = await createTempDir("formula-ext-worker-maxmodules-");
  try {
    /** @type {Record<string, string>} */
    const files = {
      "main.mjs": `import "./mod1.mjs";\nexport async function activate() {}\n`
    };

    // Total modules = main + 200 deps = 201. Limit is 200.
    for (let i = 1; i <= 200; i++) {
      const next = i === 200 ? "" : `import "./mod${i + 1}.mjs";\n`;
      files[`mod${i}.mjs`] = `${next}export const value = ${i};\n`;
    }

    await writeFiles(dir, files);
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /exceeded limit of 200 modules/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects modules that exceed maxModuleBytes", async () => {
  const dir = await createTempDir("formula-ext-worker-maxmodulebytes-");
  try {
    const bigComment = "a".repeat(270 * 1024);
    await writeFiles(dir, {
      "main.mjs": `import "./big.mjs";\nexport async function activate() {}\n`,
      "big.mjs": `/*${bigComment}*/\nexport const x = 1;\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /module too large/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects module graphs that exceed maxTotalBytes", async () => {
  const dir = await createTempDir("formula-ext-worker-maxtotalbytes-");
  try {
    /** @type {Record<string, string>} */
    const files = {
      "main.mjs": `import "./mod1.mjs";\nexport async function activate() {}\n`
    };
    const bigComment = "a".repeat(240 * 1024);

    // main + 22 deps ≈ 5.2MB+ (per-module below 256KB).
    for (let i = 1; i <= 22; i++) {
      const next = i === 22 ? "" : `import "./mod${i + 1}.mjs";\n`;
      files[`mod${i}.mjs`] = `/*${bigComment}*/\n${next}export const value = ${i};\n`;
    }

    await writeFiles(dir, files);
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /graph too large/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: maxTotalBytes counts redirected duplicates", async () => {
  const dir = await createTempDir("formula-ext-worker-maxtotalbytes-redirected-");
  try {
    /** @type {Record<string, string>} */
    const files = {};
    let imports = "";
    for (let i = 1; i <= 30; i++) {
      imports += `import "./alias-big${i}.mjs";\n`;
      files[`alias-big${i}.mjs`] = `export const value = ${i};\n`;
    }
    files["main.mjs"] = `${imports}export async function activate() {}\n`;

    await writeFiles(dir, files);
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /graph too large/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: does not treat obj.import(...) as dynamic import", async () => {
  const dir = await createTempDir("formula-ext-worker-import-prop-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const obj = { import: (value) => value };\n  obj.import("https://evil.invalid/evil.mjs");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await activateExtensionWorker({ mainUrl, extensionPath });
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects modules that redirect outside the extension base URL", async () => {
  const dir = await createTempDir("formula-ext-worker-redirect-outside-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "./redirect-outside.mjs";\nexport async function activate() {}\n`,
      "redirect-outside.mjs": `export const x = 1;\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /redirected outside the extension base URL/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: trims response.url before validating extension root prefix", async () => {
  const dir = await createTempDir("formula-ext-worker-redirect-trim-");
  try {
    await writeFiles(dir, {
      "main.mjs": `import "./whitespace-url.mjs";\nexport async function activate() {}\n`,
      "whitespace-url.mjs": `export const x = 1;\n`,
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await activateExtensionWorker({ mainUrl, extensionPath });
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

test("extension-worker: eval works when disableEval is disabled", async () => {
  const dir = await createTempDir("formula-ext-worker-eval-enabled-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const value = eval("1 + 1");\n  if (value !== 2) throw new Error("bad");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await activateExtensionWorker({ mainUrl, extensionPath, sandbox: { disableEval: false } });
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: strictImports=false allows dynamic import of relative modules", async () => {
  const dir = await createTempDir("formula-ext-worker-dynamic-import-relative-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const mod = await import("./dep.mjs");\n  if (mod.value !== 123) throw new Error("bad");\n}\n`,
      "dep.mjs": `export const value = 123;\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await activateExtensionWorker({ mainUrl, extensionPath, sandbox: { strictImports: false } });
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

test("extension-worker: Function() throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-function-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  Function("return 1");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /Function is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: AsyncFunction constructor throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-async-function-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;\n  AsyncFunction("return 1");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /AsyncFunction is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: GeneratorFunction constructor throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-generator-function-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const GeneratorFunction = Object.getPrototypeOf(function*(){}).constructor;\n  GeneratorFunction("return 1");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /GeneratorFunction is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: AsyncGeneratorFunction constructor throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-async-generator-function-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  const AsyncGeneratorFunction = Object.getPrototypeOf(async function*(){}).constructor;\n  AsyncGeneratorFunction("return 1");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /AsyncGeneratorFunction is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: setTimeout(string) throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-timeout-string-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  setTimeout("1 + 1", 0);\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /setTimeout with a string callback is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: setInterval(string) throws when disableEval is enabled", async () => {
  const dir = await createTempDir("formula-ext-worker-interval-string-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  setInterval("1 + 1", 0);\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /setInterval with a string callback is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: XMLHttpRequest is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-xhr-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new XMLHttpRequest();\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /XMLHttpRequest is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: Worker is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-worker-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new Worker("data:text/javascript,export default 1", { type: "module" });\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /Worker is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: SharedWorker is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-sharedworker-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new SharedWorker("data:text/javascript,export default 1", { type: "module" });\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /SharedWorker is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: importScripts is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-importscripts-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  importScripts("https://example.invalid/");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /importScripts is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: EventSource is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-eventsource-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new EventSource("https://example.invalid/");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /EventSource is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: navigator.sendBeacon is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-sendbeacon-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  navigator.sendBeacon("https://example.invalid/");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /navigator\.sendBeacon is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: WebTransport is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-webtransport-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new WebTransport("https://example.invalid/");\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /WebTransport is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: RTCPeerConnection is blocked when present", async () => {
  const dir = await createTempDir("formula-ext-worker-rtcpeer-");
  try {
    await writeFiles(dir, {
      "main.mjs": `export async function activate() {\n  new RTCPeerConnection();\n}\n`
    });
    const mainUrl = pathToFileURL(path.join(dir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${dir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /RTCPeerConnection is not allowed in extensions/i
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("extension-worker: rejects entrypoints outside extensionPath base URL", async () => {
  const rootDir = await createTempDir("formula-ext-worker-root-");
  const outsideDir = await createTempDir("formula-ext-worker-outside-");
  try {
    await writeFiles(outsideDir, {
      "main.mjs": `export async function activate() {}\n`
    });
    const mainUrl = pathToFileURL(path.join(outsideDir, "main.mjs")).href;
    const extensionPath = pathToFileURL(`${rootDir}${path.sep}`).href;

    await assert.rejects(
      () => activateExtensionWorker({ mainUrl, extensionPath }),
      /entrypoint must resolve inside the extension base URL/i
    );
  } finally {
    await fs.rm(rootDir, { recursive: true, force: true });
    await fs.rm(outsideDir, { recursive: true, force: true });
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
