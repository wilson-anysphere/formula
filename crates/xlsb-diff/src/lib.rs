//! Deprecated: use [`xlsx-diff`](https://crates.io/crates/xlsx-diff).
//!
//! Historically this crate implemented a separate XLSB-specific diff engine.
//! `xlsx-diff` already supports `.xlsx`, `.xlsm`, and `.xlsb` at the Open
//! Packaging Convention (ZIP part) layer, so maintaining two implementations
//! only duplicated logic and features.
//!
//! This crate is now a thin compatibility wrapper that re-exports the `xlsx-diff`
//! API so existing internal tooling can continue to compile.

pub use xlsx_diff::*;

