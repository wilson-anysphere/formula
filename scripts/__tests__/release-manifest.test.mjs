// Thin wrapper so `node --test scripts/__tests__/release-manifest.test.mjs` works,
// while the canonical suite lives in `.test.js` for `pnpm test:node` discovery
// (scripts/run-node-tests.mjs collects `*.test.js`).
import "./release-manifest.test.js";

