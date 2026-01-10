import { createRequire } from "node:module";

// y-websocket's server utilities are CommonJS and `require('yjs')`.
// Import Yjs the same way to ensure we share a single Yjs module instance.
const require = createRequire(import.meta.url);

// eslint-disable-next-line @typescript-eslint/no-var-requires
export const Y = require("yjs") as typeof import("yjs");

