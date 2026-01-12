const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("browser BrowserExtensionHost.clearExtensionStorage awaits async storageApi.clearExtensionStore", async () => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let cleared = false;
  const storageApi = {
    getExtensionStore() {
      return {};
    },
    async clearExtensionStore() {
      await new Promise((resolve) => setTimeout(resolve, 10));
      cleared = true;
    }
  };

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    storageApi
  });

  await host.clearExtensionStorage("pub.ext");
  assert.equal(cleared, true);
});

