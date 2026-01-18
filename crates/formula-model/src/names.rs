use core::fmt;

use serde::{Deserialize, Serialize};

use crate::WorksheetId;

/// Identifier for a defined name.
pub type DefinedNameId = u32;

/// Maximum length of a defined name in characters (Excel-compatible).
pub const EXCEL_DEFINED_NAME_MAX_LEN: usize = 255;

/// Built-in defined name for a sheet's print area.
pub const XLNM_PRINT_AREA: &str = "_xlnm.Print_Area";
/// Built-in defined name for a sheet's print titles.
pub const XLNM_PRINT_TITLES: &str = "_xlnm.Print_Titles";
/// Built-in defined name for a sheet's autofilter database range.
pub const XLNM_FILTER_DATABASE: &str = "_xlnm._FilterDatabase";

fn is_false(v: &bool) -> bool {
    !*v
}

/// Scope of a defined name (workbook-scoped or worksheet-scoped).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(tag = "type", content = "worksheet_id", rename_all = "snake_case")]
pub enum DefinedNameScope {
    Workbook,
    Sheet(WorksheetId),
}

/// A workbook- or sheet-scoped defined name (named range / constant / formula).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct DefinedName {
    /// Stable identifier for this name.
    pub id: DefinedNameId,
    /// User-visible defined name.
    pub name: String,
    /// Workbook or sheet scope for the name.
    pub scope: DefinedNameScope,
    /// Definition formula, stored **without** leading `=`.
    pub refers_to: String,
    /// Optional comment from the source file or UI.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment: Option<String>,
    /// Hidden names are not shown in Excel's Name Manager UI.
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    /// XLSX `localSheetId` value, preserved for round-trip fidelity.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xlsx_local_sheet_id: Option<u32>,
}

/// Excel-compatible validation errors for defined names.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DefinedNameValidationError {
    Empty,
    TooLong { len: usize, max: usize },
    InvalidStartCharacter(char),
    InvalidCharacter { ch: char, index: usize },
    LooksLikeCellReference,
}

impl fmt::Display for DefinedNameValidationError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DefinedNameValidationError::Empty => f.write_str("defined name cannot be empty"),
            DefinedNameValidationError::TooLong { len, max } => {
                write!(f, "defined name is too long ({len} > {max})")
            }
            DefinedNameValidationError::InvalidStartCharacter(ch) => {
                write!(
                    f,
                    "invalid first character '{ch}' (must start with a letter, '_' or '\\\\')"
                )
            }
            DefinedNameValidationError::InvalidCharacter { ch, index } => {
                write!(f, "invalid character '{ch}' at index {index}")
            }
            DefinedNameValidationError::LooksLikeCellReference => {
                f.write_str("defined name cannot look like a cell reference (e.g. A1 or R1C1)")
            }
        }
    }
}

impl std::error::Error for DefinedNameValidationError {}

/// Errors raised by workbook defined name operations.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum DefinedNameError {
    SheetNotFound(WorksheetId),
    DefinedNameNotFound(DefinedNameId),
    DuplicateName,
    InvalidName(DefinedNameValidationError),
}

impl fmt::Display for DefinedNameError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DefinedNameError::SheetNotFound(id) => write!(f, "sheet not found: {id}"),
            DefinedNameError::DefinedNameNotFound(id) => write!(f, "defined name not found: {id}"),
            DefinedNameError::DuplicateName => f.write_str("defined name already exists in scope"),
            DefinedNameError::InvalidName(err) => write!(f, "{err}"),
        }
    }
}

impl std::error::Error for DefinedNameError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            DefinedNameError::InvalidName(err) => Some(err),
            _ => None,
        }
    }
}

fn looks_like_a1_cell_reference(name: &str) -> bool {
    let bytes = name.as_bytes();
    if bytes.is_empty() {
        return false;
    }

    let mut i = 0;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }

    // Excel columns are 1-3 letters.
    if i == 0 || i > 3 {
        return false;
    }

    let digit_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    // Must end with digits and contain at least one digit.
    digit_start != i && i == bytes.len()
}

fn looks_like_r1c1_cell_reference(name: &str) -> bool {
    if name.eq_ignore_ascii_case("r") || name.eq_ignore_ascii_case("c") {
        return true;
    }

    let bytes = name.as_bytes();
    if bytes.first().copied().map(|b| b.to_ascii_uppercase()) != Some(b'R') {
        return false;
    }

    let mut i = 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    if i >= bytes.len() || bytes[i].to_ascii_uppercase() != b'C' {
        return false;
    }

    i += 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    i == bytes.len()
}

fn looks_like_cell_reference(name: &str) -> bool {
    looks_like_a1_cell_reference(name) || looks_like_r1c1_cell_reference(name)
}

/// Validate an Excel-compatible defined name.
///
/// Rules implemented (based on Excel and Microsoft documentation):
/// - must not be empty
/// - must be <= [`EXCEL_DEFINED_NAME_MAX_LEN`]
/// - must start with a letter, `_`, or `\`
/// - remaining characters may be letters, digits, `_`, or `.`
/// - must not match an A1 or R1C1-style cell reference (e.g. `A1`, `R1C1`)
pub fn validate_defined_name(name: &str) -> Result<(), DefinedNameValidationError> {
    let name = name.trim();
    if name.is_empty() {
        return Err(DefinedNameValidationError::Empty);
    }

    let len = name.chars().count();
    if len > EXCEL_DEFINED_NAME_MAX_LEN {
        return Err(DefinedNameValidationError::TooLong {
            len,
            max: EXCEL_DEFINED_NAME_MAX_LEN,
        });
    }

    if looks_like_cell_reference(name) {
        return Err(DefinedNameValidationError::LooksLikeCellReference);
    }

    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        debug_assert!(false, "name was checked non-empty but chars() yielded none");
        return Err(DefinedNameValidationError::Empty);
    };
    if !(first.is_alphabetic() || first == '_' || first == '\\') {
        return Err(DefinedNameValidationError::InvalidStartCharacter(first));
    }

    for (index, ch) in name.chars().enumerate().skip(1) {
        if !(ch.is_alphabetic() || ch.is_ascii_digit() || ch == '_' || ch == '.') {
            return Err(DefinedNameValidationError::InvalidCharacter { ch, index });
        }
    }

    Ok(())
}
