import { createRequire } from "node:module";

// Keep sync-server tests on the CommonJS build of Yjs + y-websocket so we don't
// end up loading both the `import` (ESM) and `require` (CJS) entrypoints in the
// same process. Mixing them triggers Yjs' "already imported" warning and can
// break `instanceof` checks inside the CRDT / provider implementations.
const require = createRequire(import.meta.url);

// eslint-disable-next-line @typescript-eslint/no-var-requires
export const Y = require("yjs") as typeof import("yjs");

// eslint-disable-next-line @typescript-eslint/no-var-requires
const yWebsocket: typeof import("y-websocket") = require("y-websocket");

export const { WebsocketProvider } = yWebsocket;

