// Focused integration test target for the legacy DBCS / byte-count text functions.
//
// This keeps the `*B` semantics tests runnable without compiling the full `tests/functions/*`
// suite (which pulls in additional dev-dependencies like proptest/criterion).

#[path = "functions/harness.rs"]
mod harness;

#[path = "functions/text_dbcs.rs"]
mod text_dbcs;
