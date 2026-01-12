const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const { spawn } = require("node:child_process");

const { ExtensionHost } = require("../src");

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function writeJson(filePath, value) {
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), "utf8");
}

async function createRequireTestExtension(rootDir) {
  const extDir = path.join(rootDir, "require-test-ext");
  const distDir = path.join(extDir, "dist");
  await fs.mkdir(distDir, { recursive: true });

  await writeJson(path.join(extDir, "package.json"), {
    name: "require-test",
    publisher: "formula",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: ["onCommand:sandboxTest.require"],
    contributes: {
      commands: [
        {
          command: "sandboxTest.require",
          title: "Sandbox require test"
        }
      ]
    },
    permissions: ["ui.commands"]
  });

  await fs.writeFile(
    path.join(distDir, "extension.js"),
    `const formula = require("formula");

async function activate(context) {
  context.subscriptions.push(
    await formula.commands.registerCommand("sandboxTest.require", (moduleName) => {
      require(String(moduleName));
      return "ok";
    })
  );
}

module.exports = { activate };
`,
    "utf8"
  );

  return extDir;
}

test("sandbox: blocks disallowed Node builtin modules (including subpaths)", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-builtin-"));
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker thread startup can be slow when node:test runs files in parallel.
    // Keep this generous so the test exercises sandboxed require() behavior rather than flaking.
    activationTimeoutMs: 20_000
  });

  t.after(async () => {
    await host.dispose();
    await fs.rm(dir, { recursive: true, force: true });
  });

  const extPath = await createRequireTestExtension(dir);
  await host.loadExtension(extPath);

  const denied = [
    { request: "fs", normalized: "fs" },
    { request: "node:fs", normalized: "fs" },
    { request: "fs/promises", normalized: "fs/promises" },
    { request: "node:fs/promises", normalized: "fs/promises" },
    { request: "process", normalized: "process" },
    { request: "node:process", normalized: "process" },
    // node:-only builtins (still should be blocked)
    { request: "node:test", normalized: "test" },
    { request: "node:test/reporters", normalized: "test/reporters" },
    { request: "http", normalized: "http" },
    { request: "node:http", normalized: "http" },
    { request: "https", normalized: "https" },
    { request: "node:https", normalized: "https" },
    { request: "http2", normalized: "http2" },
    { request: "node:http2", normalized: "http2" },
    { request: "_http_agent", normalized: "_http_agent" },
    { request: "node:_http_agent", normalized: "_http_agent" },
    { request: "_http_common", normalized: "_http_common" },
    { request: "node:_http_common", normalized: "_http_common" },
    { request: "_http_incoming", normalized: "_http_incoming" },
    { request: "node:_http_incoming", normalized: "_http_incoming" },
    { request: "_http_outgoing", normalized: "_http_outgoing" },
    { request: "node:_http_outgoing", normalized: "_http_outgoing" },
    { request: "net", normalized: "net" },
    { request: "node:net", normalized: "net" },
    { request: "tls", normalized: "tls" },
    { request: "node:tls", normalized: "tls" },
    { request: "_tls_common", normalized: "_tls_common" },
    { request: "node:_tls_common", normalized: "_tls_common" },
    { request: "_tls_wrap", normalized: "_tls_wrap" },
    { request: "node:_tls_wrap", normalized: "_tls_wrap" },
    { request: "dgram", normalized: "dgram" },
    { request: "node:dgram", normalized: "dgram" },
    { request: "dns", normalized: "dns" },
    { request: "node:dns", normalized: "dns" },
    { request: "dns/promises", normalized: "dns/promises" },
    { request: "node:dns/promises", normalized: "dns/promises" },
    { request: "child_process", normalized: "child_process" },
    { request: "node:child_process", normalized: "child_process" },
    { request: "worker_threads", normalized: "worker_threads" },
    { request: "node:worker_threads", normalized: "worker_threads" },
    { request: "cluster", normalized: "cluster" },
    { request: "node:cluster", normalized: "cluster" },
    { request: "module", normalized: "module" },
    { request: "node:module", normalized: "module" },
    { request: "vm", normalized: "vm" },
    { request: "node:vm", normalized: "vm" },
    { request: "inspector", normalized: "inspector" },
    { request: "node:inspector", normalized: "inspector" },
    { request: "inspector/promises", normalized: "inspector/promises" },
    { request: "node:inspector/promises", normalized: "inspector/promises" },
    { request: "_http_client", normalized: "_http_client" },
    { request: "node:_http_client", normalized: "_http_client" },
    { request: "_http_server", normalized: "_http_server" },
    { request: "node:_http_server", normalized: "_http_server" }
  ];

  // Newer Node versions ship additional node:-only builtins. Keep the sandbox locked down
  // while allowing the test suite to run across versions.
  const hostBuiltins = require("node:module").builtinModules;
  if (hostBuiltins.includes("node:sqlite")) {
    denied.push({ request: "node:sqlite", normalized: "sqlite" });
  }
  if (hostBuiltins.includes("node:sea")) {
    denied.push({ request: "node:sea", normalized: "sea" });
  }

  for (const { request, normalized } of denied) {
    await assert.rejects(
      () => host.executeCommand("sandboxTest.require", request),
      new RegExp(`Access to Node builtin module '${escapeRegExp(normalized)}'`)
    );
  }
});

test("sandbox: blocks require() through symlinks (even with --preserve-symlinks)", async (t) => {
  const tmpRoot = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-symlink-"));

  // Gate if symlinks are not supported (e.g. some Windows CI setups).
  const outside = path.join(tmpRoot, "outside.js");
  const probeLink = path.join(tmpRoot, "probe-link.js");
  const outsideDir = path.join(tmpRoot, "outside-dir");
  const probeDirLink = path.join(tmpRoot, "probe-dir");
  await fs.writeFile(outside, "module.exports = 123;\n", "utf8");
  await fs.mkdir(outsideDir);
  try {
    await fs.symlink(outside, probeLink, process.platform === "win32" ? "file" : undefined);
    await fs.unlink(probeLink);
    await fs.symlink(outsideDir, probeDirLink, process.platform === "win32" ? "junction" : undefined);
    try {
      await fs.unlink(probeDirLink);
    } catch {
      await fs.rmdir(probeDirLink);
    }
  } catch (error) {
    await fs.rm(tmpRoot, { recursive: true, force: true });
    const code = error && typeof error === "object" ? error.code : null;
    if (code === "EPERM" || code === "EACCES" || code === "ENOSYS") {
      t.skip(`Symlinks not supported in this environment (${code})`);
      return;
    }
    throw error;
  }

  t.after(async () => {
    await fs.rm(tmpRoot, { recursive: true, force: true });
  });

  const extensionHostEntry = require.resolve("../src");
  const scriptPath = path.join(tmpRoot, "symlink-sandbox-runner.js");

  await fs.writeFile(
    scriptPath,
    `const { ExtensionHost } = require(${JSON.stringify(extensionHostEntry)});
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");

async function writeJson(filePath, value) {
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), "utf8");
}

async function main() {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-symlink-child-"));
  try {
    const outsidePath = path.join(root, "outside.js");
    await fs.writeFile(outsidePath, "module.exports = { ok: true };\\n", "utf8");
    const outsideDirPath = path.join(root, "outside-dir");
    await fs.mkdir(outsideDirPath, { recursive: true });
    await fs.writeFile(path.join(outsideDirPath, "index.js"), "module.exports = { ok: true };\\n", "utf8");

    const extDir = path.join(root, "symlink-ext");
    const distDir = path.join(extDir, "dist");
    await fs.mkdir(distDir, { recursive: true });

    // Symlink inside the extension pointing to code outside.
    await fs.symlink(outsidePath, path.join(extDir, "linked.js"), process.platform === "win32" ? "file" : undefined);
    await fs.symlink(
      outsideDirPath,
      path.join(extDir, "linked-dir"),
      process.platform === "win32" ? "junction" : undefined
    );

    await writeJson(path.join(extDir, "package.json"), {
      name: "symlink-test",
      publisher: "formula",
      version: "1.0.0",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onCommand:symlinkTest.escapeFile", "onCommand:symlinkTest.escapeDir"],
      contributes: {
        commands: [
          {
            command: "symlinkTest.escapeFile",
            title: "Symlink escape file"
          },
          {
            command: "symlinkTest.escapeDir",
            title: "Symlink escape dir"
          }
        ]
      },
      permissions: ["ui.commands"]
    });

    await fs.writeFile(
      path.join(distDir, "extension.js"),
      \`const formula = require("formula");
async function activate(context) {
  context.subscriptions.push(await formula.commands.registerCommand("symlinkTest.escapeFile", () => {
    require("../linked.js");
    return "ok";
  }));
  context.subscriptions.push(await formula.commands.registerCommand("symlinkTest.escapeDir", () => {
    require("../linked-dir");
    return "ok";
  }));
}
module.exports = { activate };
\`,
      "utf8"
    );

    const host = new ExtensionHost({
      engineVersion: "1.0.0",
      permissionsStoragePath: path.join(root, "permissions.json"),
      extensionStoragePath: path.join(root, "storage.json"),
      permissionPrompt: async () => true,
      // Worker thread startup can be slow under heavy CI load; avoid flaking on activation.
      activationTimeoutMs: 20_000
    });

    try {
      await host.loadExtension(extDir);
      const failures = [];

      for (const cmd of ["symlinkTest.escapeFile", "symlinkTest.escapeDir"]) {
        try {
          await host.executeCommand(cmd);
          console.error("Expected " + cmd + " to be blocked, but it succeeded");
          process.exitCode = 1;
          return;
        } catch (error) {
          const msg = String(error?.message ?? error);
          if (!/outside their extension folder/.test(msg)) {
            console.error("Unexpected error for " + cmd + ":", msg);
            failures.push(cmd);
          }
        }
      }

      process.exitCode = failures.length === 0 ? 0 : 1;
    } finally {
      await host.dispose().catch(() => {});
    }
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
`,
    "utf8"
  );

  const result = await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, ["--preserve-symlinks", "--no-warnings", scriptPath], {
      stdio: ["ignore", "pipe", "pipe"]
    });

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });

    child.on("error", reject);
    child.on("exit", (code, signal) => {
      resolve({ code, signal, stdout, stderr });
    });
  });

  if (result.signal) {
    assert.fail(`child exited with signal ${result.signal}\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`);
  }

  assert.equal(
    result.code,
    0,
    `expected child to exit 0\nstdout:\n${result.stdout}\nstderr:\n${result.stderr}`
  );
});
