const test = require("node:test");

const { build } = require("./build");

test("sample-hello: dist entrypoints are in sync with src", async () => {
  await build({ check: true });
});
