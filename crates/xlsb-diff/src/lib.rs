//! Deprecated: use [`xlsx-diff`](https://crates.io/crates/xlsx-diff).
//!
//! Historically this crate implemented a separate XLSB-specific diff engine.
//! `xlsx-diff` already supports `.xlsx`, `.xlsm`, and `.xlsb` at the Open
//! Packaging Convention (ZIP part) layer, so maintaining two implementations
//! only duplicated logic and features.
//!
//! ## Encrypted XLSB support
//!
//! Password-protected workbooks that use an OLE `EncryptedPackage` wrapper are
//! supported by providing a password (via [`DiffInput`] in the library API, or
//! `--password` / `--password-file` on the CLI). The decrypted OOXML package is
//! kept in-memory and diffed like a normal ZIP-based workbook.
//!
//! This crate is now a thin compatibility wrapper that re-exports the `xlsx-diff`
//! API so existing internal tooling can continue to compile.

#![deprecated(note = "Use `xlsx-diff`; `xlsb-diff` is a deprecated compatibility wrapper.")]

pub use xlsx_diff::*;
