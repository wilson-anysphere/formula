// Used by `scripts/run-node-ts.test.js` to ensure the wrapper strips pnpm's `--`
// delimiter before forwarding args to the underlying TypeScript entrypoint.
if (process.argv.includes("--")) {
  console.error("Unexpected `--` argument in forwarded argv");
  process.exit(1);
}

console.log("ok");

