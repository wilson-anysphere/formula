//! Minimal XLSX/XLSM package handling focused on macro preservation.
//!
//! The long-term project goal is a full-fidelity Excel compatibility layer.
//! For now we implement just enough OPC plumbing to:
//! - Load an XLSX/XLSM ZIP archive.
//! - Preserve unknown parts byte-for-byte.
//! - Preserve `xl/vbaProject.bin` exactly on write.
//! - Optionally parse `vbaProject.bin` to expose modules for UI display.

mod package;
pub mod pivots;
pub mod vba;

pub use package::{XlsxError, XlsxPackage};
pub use pivots::{PivotCacheDefinitionPart, PivotCacheRecordsPart, PivotTablePart, XlsxPivots};
