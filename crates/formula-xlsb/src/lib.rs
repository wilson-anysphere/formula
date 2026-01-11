//! Minimal XLSB (Excel binary workbook) reader.
//!
//! Focus: fast-ish streaming reads of large worksheets, with preservation hooks.

pub mod biff12_varint;
pub mod errors;
pub mod format;
pub mod formula_text;
pub mod ftab;
pub mod rgce;
pub mod workbook_context;
mod opc;
mod patch;
mod parser;
mod shared_strings;
mod writer;

pub use opc::{OpenOptions, XlsbWorkbook};
pub use patch::{patch_sheet_bin, CellEdit};
pub use parser::{
    CalcMode, Cell, CellValue, Dimension, Error, Formula, SheetData, SheetMeta, SheetVisibility,
    WorkbookProperties,
};
pub use shared_strings::SharedString;

#[cfg(feature = "write")]
pub use formula_biff::{encode_rgce as encode_formula_rgce, EncodeRgceError};

/// Parse a worksheet `.bin` stream (BIFF12) and return all discovered cells.
///
/// This is primarily intended for tests and tools that already have the
/// worksheet bytes available (e.g. from an OPC reader).
pub fn parse_sheet_bin<R: std::io::Read>(
    sheet_bin: &mut R,
    shared_strings: &[String],
) -> Result<SheetData, Error> {
    let ctx = workbook_context::WorkbookContext::default();
    parser::parse_sheet(sheet_bin, shared_strings, &ctx)
}

/// Parse a worksheet `.bin` stream (BIFF12) using the provided workbook context.
///
/// This enables decoding of formulas that reference workbook-defined names, 3D
/// references, external defined names (`PtgNameX`), and add-in/UDF calls.
pub fn parse_sheet_bin_with_context<R: std::io::Read>(
    sheet_bin: &mut R,
    shared_strings: &[String],
    ctx: &workbook_context::WorkbookContext,
) -> Result<SheetData, Error> {
    parser::parse_sheet(sheet_bin, shared_strings, ctx)
}
