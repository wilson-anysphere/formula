const test = require("node:test");

const { build } = require("./build");

test("sample-hello: dist/extension.js matches src/extension.js", async () => {
  await build({ check: true });
});

