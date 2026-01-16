//! BIFF / Excel internal error code helpers.
//!
//! Excel stores error values in BIFF token streams (and BIFF12/XLSB records) as a single-byte code.
//! These codes are used by:
//! - `PtgErr` in rgce token streams
//! - cached cell values in various BIFF12/XLSB records
//!
//! The table below is consistent with MS-XLSB for the classic 7 errors plus `#GETTING_DATA`, and
//! includes additional modern Excel errors observed in the wild (Microsoft 365).
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

/// Return the canonical Excel error literal for a BIFF/XLSB error `code`, if known.
pub fn biff_error_literal(code: u8) -> Option<&'static str> {
  match code {
    0x00 => Some("#NULL!"),
    0x07 => Some("#DIV/0!"),
    0x0F => Some("#VALUE!"),
    0x17 => Some("#REF!"),
    0x1D => Some("#NAME?"),
    0x24 => Some("#NUM!"),
    0x2A => Some("#N/A"),
    0x2B => Some("#GETTING_DATA"),
    0x2C => Some("#SPILL!"),
    0x2D => Some("#CALC!"),
    0x2E => Some("#FIELD!"),
    0x2F => Some("#CONNECT!"),
    0x30 => Some("#BLOCKED!"),
    0x31 => Some("#UNKNOWN!"),
    _ => None,
  }
}

/// Convert an Excel error literal (e.g. `#DIV/0!`) into a BIFF/XLSB internal error code.
///
/// Returns `None` for unknown/unsupported literals.
pub fn biff_error_code_from_literal(literal: &str) -> Option<u8> {
  let lit = literal.trim();
  if lit.eq_ignore_ascii_case("#NULL!") {
    return Some(0x00);
  }
  if lit.eq_ignore_ascii_case("#DIV/0!") {
    return Some(0x07);
  }
  if lit.eq_ignore_ascii_case("#VALUE!") {
    return Some(0x0F);
  }
  if lit.eq_ignore_ascii_case("#REF!") {
    return Some(0x17);
  }
  if lit.eq_ignore_ascii_case("#NAME?") {
    return Some(0x1D);
  }
  if lit.eq_ignore_ascii_case("#NUM!") {
    return Some(0x24);
  }
  if lit.eq_ignore_ascii_case("#N/A") || lit.eq_ignore_ascii_case("#N/A!") {
    return Some(0x2A);
  }
  if lit.eq_ignore_ascii_case("#GETTING_DATA") {
    return Some(0x2B);
  }
  if lit.eq_ignore_ascii_case("#SPILL!") {
    return Some(0x2C);
  }
  if lit.eq_ignore_ascii_case("#CALC!") {
    return Some(0x2D);
  }
  if lit.eq_ignore_ascii_case("#FIELD!") {
    return Some(0x2E);
  }
  if lit.eq_ignore_ascii_case("#CONNECT!") {
    return Some(0x2F);
  }
  if lit.eq_ignore_ascii_case("#BLOCKED!") {
    return Some(0x30);
  }
  if lit.eq_ignore_ascii_case("#UNKNOWN!") {
    return Some(0x31);
  }
  None
}

#[cfg(test)]
mod tests {
  use super::*;

  #[test]
  fn biff_error_literal_roundtrips_with_code_from_literal() {
    for (code, lit) in [
      (0x00, "#NULL!"),
      (0x07, "#DIV/0!"),
      (0x0F, "#VALUE!"),
      (0x17, "#REF!"),
      (0x1D, "#NAME?"),
      (0x24, "#NUM!"),
      (0x2A, "#N/A"),
      (0x2B, "#GETTING_DATA"),
      (0x2C, "#SPILL!"),
      (0x2D, "#CALC!"),
      (0x2E, "#FIELD!"),
      (0x2F, "#CONNECT!"),
      (0x30, "#BLOCKED!"),
      (0x31, "#UNKNOWN!"),
    ] {
      assert_eq!(biff_error_literal(code), Some(lit));
      assert_eq!(biff_error_code_from_literal(lit), Some(code));
    }
  }

  #[test]
  fn na_exclamation_alias_maps_to_na_code() {
    assert_eq!(biff_error_code_from_literal("#N/A!"), Some(0x2A));
    assert_eq!(biff_error_code_from_literal("  #n/a!  "), Some(0x2A));
  }
}

