//! Re-exports for the BIFF function table (Ftab).
//!
//! The authoritative BIFF12 function id <-> name mapping lives in the shared
//! `formula-biff` crate so all BIFF consumers (XLSB, XLS, etc) stay consistent.

pub use formula_biff::{function_id_from_name, function_name_from_id, FTAB_USER_DEFINED};
