//! Minimal XLSX/XLSM package handling focused on macro preservation.
//!
//! The long-term project goal is a full-fidelity Excel compatibility layer.
//! For now we implement just enough OPC plumbing to:
//! - Load an XLSX/XLSM ZIP archive.
//! - Preserve unknown parts byte-for-byte.
//! - Preserve `xl/vbaProject.bin` exactly on write.
//! - Optionally parse `vbaProject.bin` to expose modules for UI display.

mod package;
pub mod comments;
pub mod outline;
pub mod pivots;
pub mod print;
pub mod shared_strings;
pub mod vba;

pub mod conditional_formatting;
pub mod styles;

pub use conditional_formatting::*;
pub use package::{XlsxError, XlsxPackage};
pub use pivots::{PivotCacheDefinitionPart, PivotCacheRecordsPart, PivotTablePart, XlsxPivots};
pub use styles::*;
