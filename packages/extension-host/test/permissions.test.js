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

