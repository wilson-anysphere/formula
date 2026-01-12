// Backwards-compatible alias for the Cursor tab completion client.
//
// Some consumers refer to this as a "completion client" (as opposed to other
// Cursor-backed clients in the repo). Keep this alias thin to avoid any prompt
// or model-routing logic in this package: the backend owns the harness.
export { CursorTabCompletionClient as CursorCompletionClient } from "./cursorTabCompletionClient.js";

