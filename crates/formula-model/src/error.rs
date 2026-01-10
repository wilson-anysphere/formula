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
}

impl fmt::Display for ErrorValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for ErrorValue {
    type Err = ParseErrorValueError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let normalized = s.trim().to_ascii_uppercase();
        match normalized.as_str() {
            "#NULL!" => Ok(ErrorValue::Null),
            "#DIV/0!" => Ok(ErrorValue::Div0),
            "#VALUE!" => Ok(ErrorValue::Value),
            "#REF!" => Ok(ErrorValue::Ref),
            "#NAME?" => Ok(ErrorValue::Name),
            "#NUM!" => Ok(ErrorValue::Num),
            "#N/A" => Ok(ErrorValue::NA),
            "#GETTING_DATA" => Ok(ErrorValue::GettingData),
            "#SPILL!" => Ok(ErrorValue::Spill),
            "#CALC!" => Ok(ErrorValue::Calc),
            "#FIELD!" => Ok(ErrorValue::Field),
            "#CONNECT!" => Ok(ErrorValue::Connect),
            "#BLOCKED!" => Ok(ErrorValue::Blocked),
            "#UNKNOWN!" => Ok(ErrorValue::Unknown),
            _ => Err(ParseErrorValueError),
        }
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
}
