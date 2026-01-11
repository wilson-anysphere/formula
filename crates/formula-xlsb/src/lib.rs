//! Minimal XLSB (Excel binary workbook) reader.
//!
//! Focus: fast-ish streaming reads of large worksheets, with preservation hooks.

pub mod biff12_varint;
pub mod format;
pub mod rgce;
pub mod workbook_context;
mod opc;
mod patch;
mod parser;
mod writer;

pub use opc::{OpenOptions, XlsbWorkbook};
pub use patch::{patch_sheet_bin, CellEdit};
pub use parser::{Cell, CellValue, Dimension, Error, Formula, SheetData, SheetMeta};

#[cfg(feature = "write")]
pub use formula_biff::{encode_rgce as encode_formula_rgce, EncodeRgceError};
