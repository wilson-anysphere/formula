//! XLSB / BIFF12 error code helpers.
//!
//! Excel stores error values in XLSB as a single byte code (used by both cell values and
//! `PtgErr` in formula token streams).
//!
//! ## Error code table
//!
//! The classic BIFF table (documented in MS-XLSB) covers the 7 legacy Excel errors plus
//! `#GETTING_DATA` (`0x2B`). Modern Excel (Microsoft 365) introduces additional error literals
//! (dynamic arrays / data types / connectivity). Excel stores those newer literals in XLSB using
//! additional u8 codes continuing the same internal numbering scheme.
//!
//! Note: these extended codes are used in the wild by Excel for both `PtgErr` and cached cell
//! values (`BrtCellBoolErr`, `BrtFmlaError`) but are not always included in published file format
//! references.
//!
//! | Code | Literal |
//! |------|---------|
//! | 0x00 | `#NULL!` |
//! | 0x07 | `#DIV/0!` |
//! | 0x0F | `#VALUE!` |
//! | 0x17 | `#REF!` |
//! | 0x1D | `#NAME?` |
//! | 0x24 | `#NUM!` |
//! | 0x2A | `#N/A` |
//! | 0x2B | `#GETTING_DATA` |
//! | 0x2C | `#SPILL!` |
//! | 0x2D | `#CALC!` |
//! | 0x2E | `#FIELD!` |
//! | 0x2F | `#CONNECT!` |
//! | 0x30 | `#BLOCKED!` |
//! | 0x31 | `#UNKNOWN!` |
//!
//! Newer versions of Excel may introduce additional codes for newer error literals
//! (e.g. `#SPILL!`). Until we have an authoritative mapping for those, callers should treat
//! unknown codes as forward-compatible and provide their own fallback representation.

use formula_biff::errors::{biff_error_code_from_literal, biff_error_literal};

/// Return the canonical Excel error literal for an XLSB error `code`, if known.
///
/// Codes are the legacy BIFF/Excel internal error ids used by XLSB records like `BrtCellBoolErr`
/// and `BrtFmlaError`, as well as the `PtgErr` formula token.
pub fn xlsb_error_literal(code: u8) -> Option<&'static str> {
    biff_error_literal(code)
}

/// Convert an Excel error literal (e.g. `#DIV/0!`) into an XLSB/BIFF12 internal error code.
///
/// Returns `None` for unknown/unsupported literals.
pub fn xlsb_error_code_from_literal(literal: &str) -> Option<u8> {
    biff_error_code_from_literal(literal)
}

/// Human-readable display string for an XLSB error `code`.
///
/// Unknown codes are displayed as `#ERR(0x??)` so the raw value isn't lost.
pub fn xlsb_error_display(code: u8) -> String {
    match xlsb_error_literal(code) {
        Some(lit) => lit.to_string(),
        None => format!("#ERR({code:#04x})"),
    }
}
