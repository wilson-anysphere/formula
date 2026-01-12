const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { PermissionManager, PermissionError } = require("../src/permission-manager");

test("permission gating: prompts once and persists grants", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-"));
  const storePath = path.join(dir, "permissions.json");

  const promptCalls = [];

  const pm1 = new PermissionManager({
    storagePath: storePath,
    prompt: async ({ permissions }) => {
      promptCalls.push([...permissions]);
      return true;
    }
  });

  await pm1.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );

  assert.equal(promptCalls.length, 1);
  assert.deepEqual(promptCalls[0], ["cells.write"]);

  const pm2 = new PermissionManager({
    storagePath: storePath,
    prompt: async () => {
      throw new Error("prompt should not be called after persistence");
    }
  });

  await pm2.ensurePermissions(
    {
      extensionId: "pub.ext",
      displayName: "Ext",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );
});

test("permission gating: rejects permission not declared in manifest", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-"));
  const storePath = path.join(dir, "permissions.json");

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  await assert.rejects(
    () =>
      pm.ensurePermissions(
        {
          extensionId: "pub.ext",
          displayName: "Ext",
          declaredPermissions: []
        },
        ["cells.write"]
      ),
    PermissionError
  );
});

test("permission storage: migrates legacy string-array grants to v2 permission records", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-migrate-"));
  const storePath = path.join(dir, "permissions.json");

  await fs.writeFile(
    storePath,
    JSON.stringify(
      {
        "pub.ext": ["cells.write", "network"]
      },
      null,
      2
    ),
    "utf8"
  );

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  const granted = await pm.getGrantedPermissions("pub.ext");
  assert.deepEqual(granted, {
    "cells.write": true,
    network: { mode: "full" }
  });

  const raw = JSON.parse(await fs.readFile(storePath, "utf8"));
  assert.deepEqual(raw, {
    "pub.ext": {
      "cells.write": true,
      network: { mode: "full" }
    }
  });
});

test("permission storage: resetPermissions clears a single extension and forces re-prompt", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-reset-"));
  const storePath = path.join(dir, "permissions.json");
  const extensionId = "pub.ext";

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId,
      displayName: "Ext",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );

  assert.deepEqual(await pm.getGrantedPermissions(extensionId), { "cells.write": true });

  await pm.resetPermissions(extensionId);
  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {});

  let promptCalls = 0;
  const pm2 = new PermissionManager({
    storagePath: storePath,
    prompt: async ({ permissions }) => {
      promptCalls += 1;
      assert.deepEqual(permissions, ["cells.write"]);
      return true;
    }
  });

  await pm2.ensurePermissions(
    {
      extensionId,
      displayName: "Ext",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );

  assert.equal(promptCalls, 1);
});

test("permission gating: accepts object-form declared permissions", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-object-declared-"));
  const storePath = path.join(dir, "permissions.json");
  const extensionId = "pub.obj";

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId,
      displayName: "Obj",
      declaredPermissions: [{ network: { mode: "allowlist", hosts: ["example.com"] } }, { clipboard: true }]
    },
    ["network", "clipboard"]
  );

  assert.deepEqual(await pm.getGrantedPermissions(extensionId), {
    network: { mode: "full" },
    clipboard: true
  });
});

test("permission storage: resetAllPermissions clears all extensions and forces re-prompt", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-reset-all-"));
  const storePath = path.join(dir, "permissions.json");

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  await pm.ensurePermissions(
    {
      extensionId: "pub.one",
      displayName: "One",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );

  await pm.ensurePermissions(
    {
      extensionId: "pub.two",
      displayName: "Two",
      declaredPermissions: ["clipboard"]
    },
    ["clipboard"]
  );

  assert.deepEqual(await pm.getGrantedPermissions("pub.one"), { "cells.write": true });
  assert.deepEqual(await pm.getGrantedPermissions("pub.two"), { clipboard: true });

  await pm.resetAllPermissions();
  assert.deepEqual(await pm.getGrantedPermissions("pub.one"), {});
  assert.deepEqual(await pm.getGrantedPermissions("pub.two"), {});

  const promptCalls = [];
  const pm2 = new PermissionManager({
    storagePath: storePath,
    prompt: async ({ extensionId, permissions }) => {
      promptCalls.push({ extensionId, permissions: [...permissions] });
      return true;
    }
  });

  await pm2.ensurePermissions(
    {
      extensionId: "pub.one",
      displayName: "One",
      declaredPermissions: ["cells.write"]
    },
    ["cells.write"]
  );

  await pm2.ensurePermissions(
    {
      extensionId: "pub.two",
      displayName: "Two",
      declaredPermissions: ["clipboard"]
    },
    ["clipboard"]
  );

  assert.deepEqual(promptCalls, [
    { extensionId: "pub.one", permissions: ["cells.write"] },
    { extensionId: "pub.two", permissions: ["clipboard"] }
  ]);
});

test("permission storage: does not rewrite v2 permission records when no migration is needed", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-no-migrate-"));
  const storePath = path.join(dir, "permissions.json");

  // Use non-standard formatting so we can detect unexpected migration writes.
  const initial = JSON.stringify(
    {
      "pub.ext": {
        "cells.write": true,
        network: { mode: "allowlist", hosts: ["api.example.com"] }
      }
    },
    null,
    4
  );

  await fs.writeFile(storePath, initial, "utf8");

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => {
      throw new Error("prompt should not be called");
    }
  });

  assert.deepEqual(await pm.getGrantedPermissions("pub.ext"), {
    "cells.write": true,
    network: { mode: "allowlist", hosts: ["api.example.com"] }
  });

  const after = await fs.readFile(storePath, "utf8");
  assert.equal(after, initial);
});

test("permission storage: getGrantedPermissions for an unknown extension does not create an empty record", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-perms-empty-record-"));
  const storePath = path.join(dir, "permissions.json");

  const pm = new PermissionManager({
    storagePath: storePath,
    prompt: async () => true
  });

  assert.deepEqual(await pm.getGrantedPermissions("pub.unknown"), {});

  await pm.ensurePermissions(
    {
      extensionId: "pub.other",
      displayName: "Other",
      declaredPermissions: ["clipboard"]
    },
    ["clipboard"]
  );

  const stored = JSON.parse(await fs.readFile(storePath, "utf8"));
  assert.deepEqual(stored, { "pub.other": { clipboard: true } });
});
