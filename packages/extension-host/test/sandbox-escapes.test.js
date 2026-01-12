const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const net = require("node:net");

const { ExtensionHost } = require("../src");

async function writeExtension(dir, { name, publisher, main, activateBody }) {
  await fs.mkdir(dir, { recursive: true });

  const manifest = {
    name,
    displayName: name,
    version: "1.0.0",
    publisher,
    main,
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${publisher}.${name}.activate`],
    contributes: {
      commands: [
        {
          command: `${publisher}.${name}.activate`,
          title: "Activate"
        }
      ]
    },
    permissions: []
  };

  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(
    path.join(dir, main),
    `module.exports.activate = async () => {\n${activateBody}\n};\n`,
    "utf8"
  );

  return manifest.contributes.commands[0].command;
}

test("sandbox: blocks dynamic import('node:fs/promises')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-import-fs-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-import-fs",
    publisher: "formula",
    main: "extension.js",
    activateBody: "await import('node:fs/promises');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep this security test focused on
    // correctness (dynamic import is rejected) rather than the default 5s activation SLA.
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(
    () => host.executeCommand(commandId),
    /Dynamic import is not allowed in extensions.*node:fs\/promises/
  );
});

test("sandbox: blocks process.binding('fs')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-binding-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-binding",
    publisher: "formula",
    main: "extension.js",
    activateBody: "process.binding('fs');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /process\.binding\(\) is not allowed/);
});

test("sandbox: blocks dynamic import('node:http2')", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-import-http2-"));
  const extDir = path.join(dir, "ext");

  const commandId = await writeExtension(extDir, {
    name: "sandbox-import-http2",
    publisher: "formula",
    main: "extension.js",
    activateBody: "await import('node:http2');"
  });

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(
    () => host.executeCommand(commandId),
    /Dynamic import is not allowed in extensions.*node:http2/
  );
});

test("sandbox: WebSocket events cannot be used as a vm escape hatch", async (t) => {
  const server = net.createServer((socket) => {
    socket.destroy();
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  const port = typeof address === "object" && address ? address.port : null;
  if (!port) throw new Error("Failed to allocate test port");

  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-ws-escape-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
  });

  const commandId = "formula-test.wsEscape.attempt";
  const manifest = {
    name: "ws-escape",
    displayName: "WebSocket Escape Attempt",
    version: "1.0.0",
    publisher: "formula-test",
    main: "./extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${commandId}`],
    contributes: { commands: [{ command: commandId, title: "WebSocket Escape Attempt" }] },
    permissions: ["ui.commands", "network"]
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(
    path.join(extDir, "extension.js"),
    `
      const formula = require("@formula/extension-api");

       exports.activate = async (context) => {
         context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async (port) => {
           return await new Promise((resolve) => {
             const ws = new WebSocket(\`ws://127.0.0.1:\${Number(port)}/\`);
             const on = ws.addEventListener ? (name, fn) => ws.addEventListener(name, fn) : (name, fn) => (ws[\`on\${name}\`] = fn);
             const handler = (evt) => {
               try {
                 const proc = evt?.constructor?.constructor?.("return process")();
                 resolve({ escaped: true, pid: proc?.pid ?? null });
               } catch (error) {
                 resolve({ escaped: false, error: String(error?.message ?? error) });
               }
             };
             on("close", handler);
             on("error", handler);
             setTimeout(() => resolve({ escaped: false, error: "timeout" }), 2000);
           });
         }));
       };
     `,
    "utf8"
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId, port);
  assert.equal(result.escaped, false);
  assert.match(String(result.error ?? ""), /Code generation from strings disallowed|not allowed|disallowed/);
});

test("sandbox: blocks Error.prepareStackTrace CallSite escape", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-stack-escape-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "formula-test.stackEscape.attempt";
  const manifest = {
    name: "stack-escape",
    displayName: "Stack Escape Attempt",
    version: "1.0.0",
    publisher: "formula-test",
    main: "./extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${commandId}`],
    contributes: { commands: [{ command: commandId, title: "Stack Escape Attempt" }] },
    permissions: ["ui.commands"]
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(
    path.join(extDir, "extension.js"),
    `
      const formula = require("@formula/extension-api");

      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
          try {
            Error.prepareStackTrace = (err, stack) => stack;
            const stack = new Error().stack;
            const callSite = Array.isArray(stack) ? stack.find((cs) => cs && typeof cs.getFunction === "function") : null;
            const fn = callSite ? callSite.getFunction() : null;
            if (typeof fn === "function") {
              // If this ever succeeds, the sandbox is broken.
              const proc = fn.constructor("return process")();
              return { escaped: true, pid: proc?.pid ?? null };
            }
            return { escaped: false, error: "no-host-function" };
          } catch (error) {
            return { escaped: false, error: String(error?.message ?? error) };
          }
        }));
      };
    `,
    "utf8"
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId);
  assert.equal(result.escaped, false);
  assert.match(String(result.error ?? ""), /Error.prepareStackTrace is not allowed in extensions/);
});

test("sandbox: blocks arguments.callee.caller escape", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-caller-escape-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "formula-test.callerEscape.attempt";
  const manifest = {
    name: "caller-escape",
    displayName: "Caller Escape Attempt",
    version: "1.0.0",
    publisher: "formula-test",
    main: "./extension.js",
    engines: { formula: "^1.0.0" },
    activationEvents: [`onCommand:${commandId}`],
    contributes: { commands: [{ command: commandId, title: "Caller Escape Attempt" }] },
    permissions: ["ui.commands"]
  };

  await fs.writeFile(path.join(extDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(
    path.join(extDir, "extension.js"),
    `
      const formula = require("@formula/extension-api");

      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
          try {
            const caller = (function () {
              // If modules run in sloppy mode, this returns a host function (full escape).
              return arguments.callee.caller;
            })();

            if (typeof caller === "function") {
              const proc = caller.constructor("return process")();
              return { escaped: true, pid: proc?.pid ?? null };
            }

            return { escaped: false, error: String(caller) };
          } catch (error) {
            return { escaped: false, error: String(error?.message ?? error) };
          }
        }));
      };
    `,
    "utf8"
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId);
  assert.equal(result.escaped, false);
  assert.match(String(result.error ?? ""), /caller|callee|strict/i);
});

test("sandbox: does not treat obj.import(...) as dynamic import", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sandbox-import-prop-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir, { recursive: true });

  await fs.writeFile(
    path.join(extDir, "package.json"),
    JSON.stringify(
      {
        name: "sandbox-import-prop",
        displayName: "sandbox-import-prop",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onStartupFinished"],
        contributes: { commands: [] },
        permissions: []
      },
      null,
      2
    ),
    "utf8"
  );

  await fs.writeFile(
    path.join(extDir, "extension.js"),
    `
      module.exports.activate = async () => {
        const obj = { import: (value) => String(value ?? "ok") };
        if (obj.import("ok") !== "ok") {
          throw new Error("import property call failed");
        }
      };
    `,
    "utf8"
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: 20_000,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);
  await host.startup();
});
