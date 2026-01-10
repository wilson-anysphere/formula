use std::fmt;

/// Spreadsheet-compatible error codes used by Excel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelError {
    /// `#NUM!`
    Num,
    /// `#VALUE!`
    Value,
    /// `#DIV/0!`
    Div0,
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ExcelError::Num => write!(f, "#NUM!"),
            ExcelError::Value => write!(f, "#VALUE!"),
            ExcelError::Div0 => write!(f, "#DIV/0!"),
        }
    }
}

impl std::error::Error for ExcelError {}

pub type ExcelResult<T> = Result<T, ExcelError>;
