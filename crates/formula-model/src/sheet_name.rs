use core::fmt;

/// Maximum worksheet name length enforced by Excel.
///
/// Excel stores sheet names as UTF-16 and enforces the 31-character limit in terms of UTF-16 code
/// units. That means characters outside the BMP (e.g. many emoji) count as 2.
pub const EXCEL_MAX_SHEET_NAME_LEN: usize = 31;

/// Errors returned when validating worksheet names.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetNameError {
    /// The name is empty or consists only of whitespace.
    EmptyName,
    /// The name exceeds [`EXCEL_MAX_SHEET_NAME_LEN`].
    TooLong,
    /// The name contains a character that Excel forbids in worksheet names.
    InvalidCharacter(char),
    /// Excel forbids worksheet names that begin or end with `'`.
    LeadingOrTrailingApostrophe,
    /// The name conflicts with an existing sheet (case-insensitive).
    DuplicateName,
}

impl fmt::Display for SheetNameError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SheetNameError::EmptyName => f.write_str("sheet name cannot be blank"),
            SheetNameError::TooLong => write!(
                f,
                "sheet name cannot exceed {EXCEL_MAX_SHEET_NAME_LEN} characters"
            ),
            SheetNameError::InvalidCharacter(ch) => {
                write!(f, "sheet name contains invalid character `{ch}`")
            }
            SheetNameError::LeadingOrTrailingApostrophe => {
                f.write_str("sheet name cannot begin or end with an apostrophe")
            }
            SheetNameError::DuplicateName => f.write_str("sheet name already exists"),
        }
    }
}

impl std::error::Error for SheetNameError {}

/// Validate a worksheet name using Excel-compatible rules.
pub fn validate_sheet_name(name: &str) -> Result<(), SheetNameError> {
    if name.trim().is_empty() {
        return Err(SheetNameError::EmptyName);
    }

    if name.encode_utf16().count() > EXCEL_MAX_SHEET_NAME_LEN {
        return Err(SheetNameError::TooLong);
    }

    if let Some(ch) = name
        .chars()
        .find(|ch| matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(SheetNameError::InvalidCharacter(ch));
    }

    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(SheetNameError::LeadingOrTrailingApostrophe);
    }

    Ok(())
}
