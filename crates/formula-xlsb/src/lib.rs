//! Minimal XLSB (Excel binary workbook) reader.
//!
//! Focus: fast-ish streaming reads of large worksheets, with preservation hooks.

mod opc;
mod parser;

pub use opc::{OpenOptions, XlsbWorkbook};
pub use parser::{Cell, CellValue, Dimension, Error, Formula, SheetData, SheetMeta};
