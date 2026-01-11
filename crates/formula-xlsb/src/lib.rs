//! Minimal XLSB (Excel binary workbook) reader.
//!
//! Focus: fast-ish streaming reads of large worksheets, with preservation hooks.

mod opc;
mod patch;
mod parser;
mod writer;

pub use opc::{OpenOptions, XlsbWorkbook};
pub use patch::{patch_sheet_bin, CellEdit};
pub use parser::{Cell, CellValue, Dimension, Error, Formula, SheetData, SheetMeta};
