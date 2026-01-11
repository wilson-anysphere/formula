//! XLSB / BIFF12 error code helpers.
//!
//! Excel stores error values in XLSB as a single byte code (used by both cell values and
//! `PtgErr` in formula token streams). The mapping is the classic BIFF error table.
//!
//! Newer versions of Excel may introduce additional codes for newer error literals
//! (e.g. `#SPILL!`). Until we have an authoritative mapping for those, callers should treat
//! unknown codes as forward-compatible and provide their own fallback representation.

/// Return the canonical Excel error literal for an XLSB error `code`, if known.
///
/// Codes are the legacy BIFF/Excel internal error ids used by XLSB records like `BrtCellBoolErr`
/// and `BrtFmlaError`, as well as the `PtgErr` formula token.
pub fn xlsb_error_literal(code: u8) -> Option<&'static str> {
    match code {
        0x00 => Some("#NULL!"),
        0x07 => Some("#DIV/0!"),
        0x0F => Some("#VALUE!"),
        0x17 => Some("#REF!"),
        0x1D => Some("#NAME?"),
        0x24 => Some("#NUM!"),
        0x2A => Some("#N/A"),
        0x2B => Some("#GETTING_DATA"),
        _ => None,
    }
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

