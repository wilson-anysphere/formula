require("./src/runtime.js");

const api = globalThis[Symbol.for("formula.extensionApi.api")];
if (!api) {
  throw new Error("@formula/extension-api runtime failed to initialize");
}

module.exports = api;
