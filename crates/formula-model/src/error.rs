use core::fmt;
use core::str::FromStr;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize, Serializer};

/// Excel-style error values.
///
/// Includes the classic Excel 7 errors plus newer dynamic array / data errors
/// referenced in the project docs (e.g. `#SPILL!`, `#CALC!`).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub enum ErrorValue {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    GettingData,
    Spill,
    Calc,
    Field,
    Connect,
    Blocked,
    Unknown,
}

impl ErrorValue {
    /// Excel's canonical spelling for the error (including punctuation).
    pub const fn as_str(self) -> &'static str {
        match self {
            ErrorValue::Null => "#NULL!",
            ErrorValue::Div0 => "#DIV/0!",
            ErrorValue::Value => "#VALUE!",
            ErrorValue::Ref => "#REF!",
            ErrorValue::Name => "#NAME?",
            ErrorValue::Num => "#NUM!",
            ErrorValue::NA => "#N/A",
            ErrorValue::GettingData => "#GETTING_DATA",
            ErrorValue::Spill => "#SPILL!",
            ErrorValue::Calc => "#CALC!",
            ErrorValue::Field => "#FIELD!",
            ErrorValue::Connect => "#CONNECT!",
            ErrorValue::Blocked => "#BLOCKED!",
            ErrorValue::Unknown => "#UNKNOWN!",
        }
    }

    /// Alias for [`ErrorValue::as_str`].
    ///
    /// Historically some callers used `as_code()` for the canonical error text; keep this helper
    /// to avoid breaking downstream code while `ErrorValue` is the shared canonical error type.
    pub const fn as_code(self) -> &'static str {
        self.as_str()
    }

    /// Numeric error code used by Excel in various internal representations.
    ///
    /// Values are based on the mapping documented in `docs/01-formula-engine.md`.
    pub const fn code(self) -> u8 {
        match self {
            ErrorValue::Null => 1,
            ErrorValue::Div0 => 2,
            ErrorValue::Value => 3,
            ErrorValue::Ref => 4,
            ErrorValue::Name => 5,
            ErrorValue::Num => 6,
            ErrorValue::NA => 7,
            ErrorValue::GettingData => 8,
            ErrorValue::Spill => 9,
            ErrorValue::Calc => 10,
            ErrorValue::Field => 11,
            ErrorValue::Connect => 12,
            ErrorValue::Blocked => 13,
            ErrorValue::Unknown => 14,
        }
    }

    /// Convert from an Excel error code.
    pub const fn from_code(code: u8) -> Option<Self> {
        match code {
            1 => Some(ErrorValue::Null),
            2 => Some(ErrorValue::Div0),
            3 => Some(ErrorValue::Value),
            4 => Some(ErrorValue::Ref),
            5 => Some(ErrorValue::Name),
            6 => Some(ErrorValue::Num),
            7 => Some(ErrorValue::NA),
            8 => Some(ErrorValue::GettingData),
            9 => Some(ErrorValue::Spill),
            10 => Some(ErrorValue::Calc),
            11 => Some(ErrorValue::Field),
            12 => Some(ErrorValue::Connect),
            13 => Some(ErrorValue::Blocked),
            14 => Some(ErrorValue::Unknown),
            _ => None,
        }
    }
}

impl fmt::Display for ErrorValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for ErrorValue {
    type Err = ParseErrorValueError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let s = s.trim();
        if s.eq_ignore_ascii_case("#NULL!") {
            return Ok(ErrorValue::Null);
        }
        if s.eq_ignore_ascii_case("#DIV/0!") {
            return Ok(ErrorValue::Div0);
        }
        if s.eq_ignore_ascii_case("#VALUE!") {
            return Ok(ErrorValue::Value);
        }
        if s.eq_ignore_ascii_case("#REF!") {
            return Ok(ErrorValue::Ref);
        }
        if s.eq_ignore_ascii_case("#NAME?") {
            return Ok(ErrorValue::Name);
        }
        if s.eq_ignore_ascii_case("#NUM!") {
            return Ok(ErrorValue::Num);
        }
        if s.eq_ignore_ascii_case("#N/A") || s.eq_ignore_ascii_case("#N/A!") {
            return Ok(ErrorValue::NA);
        }
        if s.eq_ignore_ascii_case("#GETTING_DATA") {
            return Ok(ErrorValue::GettingData);
        }
        if s.eq_ignore_ascii_case("#SPILL!") {
            return Ok(ErrorValue::Spill);
        }
        if s.eq_ignore_ascii_case("#CALC!") {
            return Ok(ErrorValue::Calc);
        }
        if s.eq_ignore_ascii_case("#FIELD!") {
            return Ok(ErrorValue::Field);
        }
        if s.eq_ignore_ascii_case("#CONNECT!") {
            return Ok(ErrorValue::Connect);
        }
        if s.eq_ignore_ascii_case("#BLOCKED!") {
            return Ok(ErrorValue::Blocked);
        }
        if s.eq_ignore_ascii_case("#UNKNOWN!") {
            return Ok(ErrorValue::Unknown);
        }
        Err(ParseErrorValueError)
    }
}

/// Failed to parse an [`ErrorValue`] from a string.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub struct ParseErrorValueError;

impl fmt::Display for ParseErrorValueError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("invalid Excel error value")
    }
}

impl std::error::Error for ParseErrorValueError {}

impl Serialize for ErrorValue {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_str(self.as_str())
    }
}

impl<'de> Deserialize<'de> for ErrorValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let s = String::deserialize(deserializer)?;
        s.parse::<ErrorValue>()
            .map_err(|_| D::Error::custom(format!("unknown Excel error: {s}")))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn error_value_from_str_is_case_insensitive_and_trims() {
        assert_eq!("#ref!".parse::<ErrorValue>().unwrap(), ErrorValue::Ref);
        assert_eq!("  #n/a  ".parse::<ErrorValue>().unwrap(), ErrorValue::NA);
        assert_eq!(
            "#getting_data".parse::<ErrorValue>().unwrap(),
            ErrorValue::GettingData
        );
        assert!("nope".parse::<ErrorValue>().is_err());
    }

    #[test]
    fn error_string_roundtrip() {
        for (err, s) in [
            (ErrorValue::Null, "#NULL!"),
            (ErrorValue::Div0, "#DIV/0!"),
            (ErrorValue::Value, "#VALUE!"),
            (ErrorValue::Ref, "#REF!"),
            (ErrorValue::Name, "#NAME?"),
            (ErrorValue::Num, "#NUM!"),
            (ErrorValue::NA, "#N/A"),
            (ErrorValue::GettingData, "#GETTING_DATA"),
            (ErrorValue::Spill, "#SPILL!"),
            (ErrorValue::Calc, "#CALC!"),
            (ErrorValue::Field, "#FIELD!"),
            (ErrorValue::Connect, "#CONNECT!"),
            (ErrorValue::Blocked, "#BLOCKED!"),
            (ErrorValue::Unknown, "#UNKNOWN!"),
        ] {
            assert_eq!(err.as_str(), s);
            assert_eq!(err.to_string(), s);
            assert_eq!(s.parse::<ErrorValue>().unwrap(), err);
        }
    }

    #[test]
    fn na_exclamation_alias_parses_as_na() {
        assert_eq!("#N/A!".parse::<ErrorValue>().unwrap(), ErrorValue::NA);
        assert_eq!("  #n/a!  ".parse::<ErrorValue>().unwrap(), ErrorValue::NA);
    }

    #[test]
    fn error_codes_roundtrip() {
        for err in [
            ErrorValue::Null,
            ErrorValue::Div0,
            ErrorValue::Value,
            ErrorValue::Ref,
            ErrorValue::Name,
            ErrorValue::Num,
            ErrorValue::NA,
            ErrorValue::GettingData,
            ErrorValue::Spill,
            ErrorValue::Calc,
            ErrorValue::Field,
            ErrorValue::Connect,
            ErrorValue::Blocked,
            ErrorValue::Unknown,
        ] {
            assert_eq!(ErrorValue::from_code(err.code()), Some(err));
        }
        assert_eq!(ErrorValue::from_code(0), None);
        assert_eq!(ErrorValue::from_code(255), None);
    }
}
