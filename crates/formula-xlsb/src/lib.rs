//! Minimal XLSB (Excel binary workbook) reader.
//!
//! Focus: fast-ish streaming reads of large worksheets, with preservation hooks.

pub mod biff12_varint;
pub mod errors;
pub mod format;
pub mod formula_text;
pub mod ftab;
mod opc;
mod parser;
mod patch;
pub mod rgce;
mod shared_strings;
mod shared_strings_write;
mod strings;
mod styles;
pub mod workbook_context;
mod writer;

pub use opc::{OpenOptions, XlsbWorkbook};
pub use parser::{
    CalcMode, Cell, CellValue, DefinedName, Dimension, Error, Formula, SheetData, SheetMeta,
    SheetVisibility, WorkbookProperties,
};
pub use patch::{patch_sheet_bin, patch_sheet_bin_streaming, rgce_references_rgcb, CellEdit};
pub use shared_strings::SharedString;
pub use shared_strings_write::{SharedStringsWriter, SharedStringsWriterStreaming};
pub use strings::{OpaqueRichText, ParsedXlsbString};
pub use styles::{StyleInfo, Styles};

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
    parser::parse_sheet(sheet_bin, shared_strings, &ctx, true, true)
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
    parser::parse_sheet(sheet_bin, shared_strings, ctx, true, true)
}
