use chrono::{DateTime, Utc};

use crate::coercion::datetime::parse_value_text;
use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::functions::wildcard::WildcardPattern;
use crate::simd::{CmpOp, NumericCriteria};
use crate::value::{parse_number, NumberLocale};
use crate::{ErrorKind, Value};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CriteriaOp {
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

#[derive(Debug, Clone, PartialEq)]
enum CriteriaRhs {
    Blank,
    Number(f64),
    Bool(bool),
    Error(ErrorKind),
    Text(TextCriteria),
}

#[derive(Debug, Clone, PartialEq)]
struct TextCriteria {
    /// Case-folded pattern used for `>`/`<` comparisons (wildcards treated as literals).
    literal_folded: String,
    wildcard: WildcardPattern,
}

impl TextCriteria {
    fn new(raw: &str) -> Self {
        let folded = fold_case(raw);
        let wildcard = WildcardPattern::new(&folded);
        let literal_folded = wildcard.literal_pattern();

        Self {
            literal_folded,
            wildcard,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Criteria {
    op: CriteriaOp,
    rhs: CriteriaRhs,
    number_locale: NumberLocale,
}

impl Criteria {
    /// Parse an Excel criteria value using Excel's default 1900 date system (with Lotus compat).
    pub fn parse(input: &Value) -> Result<Self, ErrorKind> {
        Self::parse_with_date_system(input, ExcelDateSystem::EXCEL_1900)
    }

    /// Parse an Excel criteria value, resolving date/time strings using the supplied workbook
    /// date system.
    pub fn parse_with_date_system(
        input: &Value,
        system: ExcelDateSystem,
    ) -> Result<Self, ErrorKind> {
        Self::parse_with_date_system_and_locale(
            input,
            system,
            ValueLocaleConfig::en_us(),
            Utc::now(),
        )
    }

    /// Parse an Excel criteria value, resolving numeric/date/time strings using the provided value
    /// locale (decimal/thousands separators, date order) and workbook date system.
    ///
    /// Excel interprets criteria strings using the workbook locale. For example, `">1,5"` in the
    /// `de-DE` locale should be treated as `>1.5`.
    pub fn parse_with_date_system_and_locale(
        input: &Value,
        system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
        now_utc: DateTime<Utc>,
    ) -> Result<Self, ErrorKind> {
        let separators = value_locale.separators;
        let number_locale =
            NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));

        match input {
            Value::Number(n) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Number(*n),
                number_locale,
            }),
            Value::Bool(b) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Bool(*b),
                number_locale,
            }),
            Value::Error(e) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Error(*e),
                number_locale,
            }),
            Value::Blank => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Blank,
                number_locale,
            }),
            Value::Text(s) => {
                parse_criteria_string(s, system, value_locale, now_utc, number_locale)
            }
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    /// If this criteria can be represented as a simple numeric comparator, return the SIMD
    /// kernel representation.
    pub fn as_numeric_criteria(&self) -> Option<NumericCriteria> {
        let op = match self.op {
            CriteriaOp::Eq => CmpOp::Eq,
            CriteriaOp::Ne => CmpOp::Ne,
            CriteriaOp::Lt => CmpOp::Lt,
            CriteriaOp::Lte => CmpOp::Le,
            CriteriaOp::Gt => CmpOp::Gt,
            CriteriaOp::Gte => CmpOp::Ge,
        };

        match &self.rhs {
            CriteriaRhs::Number(n) => Some(NumericCriteria::new(op, *n)),
            CriteriaRhs::Bool(b) => Some(NumericCriteria::new(op, if *b { 1.0 } else { 0.0 })),
            _ => None,
        }
    }

    pub fn matches(&self, value: &Value) -> bool {
        // Criteria functions never propagate errors from candidate cells. Errors only match
        // error criteria.
        match value {
            Value::Error(_) | Value::Spill { .. } => {
                return matches_error_criteria(&self.op, &self.rhs, value);
            }
            _ => {}
        }

        match &self.rhs {
            CriteriaRhs::Blank => match self.op {
                CriteriaOp::Eq => is_blank_value(value),
                CriteriaOp::Ne => !is_blank_value(value),
                _ => false,
            },
            CriteriaRhs::Error(_) => matches_error_criteria(&self.op, &self.rhs, value),
            CriteriaRhs::Bool(b) => matches_numeric_criteria(
                self.op,
                if *b { 1.0 } else { 0.0 },
                value,
                self.number_locale,
            ),
            CriteriaRhs::Number(n) => {
                matches_numeric_criteria(self.op, *n, value, self.number_locale)
            }
            CriteriaRhs::Text(pattern) => matches_text_criteria(self.op, pattern, value),
        }
    }
}

fn matches_error_criteria(op: &CriteriaOp, rhs: &CriteriaRhs, value: &Value) -> bool {
    let CriteriaRhs::Error(err) = rhs else {
        return false;
    };

    let candidate = match value {
        Value::Error(e) => Some(*e),
        Value::Spill { .. } => Some(ErrorKind::Spill),
        _ => None,
    };
    let is_target_error = candidate == Some(*err);
    match op {
        CriteriaOp::Eq => is_target_error,
        CriteriaOp::Ne => !is_target_error,
        _ => false,
    }
}

fn matches_numeric_criteria(op: CriteriaOp, rhs: f64, value: &Value, locale: NumberLocale) -> bool {
    let Some(value_num) = coerce_to_number(value, locale) else {
        return false;
    };

    match op {
        CriteriaOp::Eq => value_num == rhs,
        CriteriaOp::Ne => value_num != rhs,
        CriteriaOp::Lt => value_num < rhs,
        CriteriaOp::Lte => value_num <= rhs,
        CriteriaOp::Gt => value_num > rhs,
        CriteriaOp::Gte => value_num >= rhs,
    }
}

fn matches_text_criteria(op: CriteriaOp, pattern: &TextCriteria, value: &Value) -> bool {
    let Some(value_text) = coerce_to_text(value) else {
        // Excel criteria functions treat blanks as "not text" for text-pattern matching (e.g.
        // COUNTIF(range,"*") does not count truly empty cells). For `<>` text criteria, non-text
        // values still satisfy the predicate because they are not equal to the text pattern.
        return matches!(op, CriteriaOp::Ne);
    };
    let value_folded = fold_case(&value_text);

    match op {
        CriteriaOp::Eq => {
            if !pattern.wildcard.has_wildcards() {
                return value_folded == pattern.literal_folded;
            }
            pattern.wildcard.matches_folded(&value_folded)
        }
        CriteriaOp::Ne => {
            if !pattern.wildcard.has_wildcards() {
                return value_folded != pattern.literal_folded;
            }
            !pattern.wildcard.matches_folded(&value_folded)
        }
        CriteriaOp::Lt => value_folded < pattern.literal_folded,
        CriteriaOp::Lte => value_folded <= pattern.literal_folded,
        CriteriaOp::Gt => value_folded > pattern.literal_folded,
        CriteriaOp::Gte => value_folded >= pattern.literal_folded,
    }
}

fn parse_criteria_string(
    raw: &str,
    system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    number_locale: NumberLocale,
) -> Result<Criteria, ErrorKind> {
    let (op, rhs_str) = split_op(raw);
    let rhs_trimmed = rhs_str.trim();

    // Blank criteria are driven by the explicit RHS, not by whitespace.
    if rhs_str.is_empty() {
        return match op {
            CriteriaOp::Eq | CriteriaOp::Ne => Ok(Criteria {
                op,
                rhs: CriteriaRhs::Blank,
                number_locale,
            }),
            _ => Err(ErrorKind::Value),
        };
    }

    if let Some(err) = parse_error_kind(rhs_trimmed) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Error(err),
            number_locale,
        });
    }

    if rhs_trimmed.eq_ignore_ascii_case("TRUE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(true),
            number_locale,
        });
    }
    if rhs_trimmed.eq_ignore_ascii_case("FALSE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(false),
            number_locale,
        });
    }

    if let Ok(serial) = parse_value_text(rhs_str, value_locale, now_utc, system) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Number(serial),
            number_locale,
        });
    }

    Ok(Criteria {
        op,
        rhs: CriteriaRhs::Text(TextCriteria::new(rhs_str)),
        number_locale,
    })
}

fn split_op(raw: &str) -> (CriteriaOp, &str) {
    let raw = raw;
    for (prefix, op) in [
        ("<>", CriteriaOp::Ne),
        ("<=", CriteriaOp::Lte),
        (">=", CriteriaOp::Gte),
        ("<", CriteriaOp::Lt),
        (">", CriteriaOp::Gt),
        ("=", CriteriaOp::Eq),
    ] {
        if let Some(rest) = raw.strip_prefix(prefix) {
            return (op, rest);
        }
    }
    (CriteriaOp::Eq, raw)
}

fn is_blank_value(value: &Value) -> bool {
    match value {
        Value::Blank => true,
        Value::Text(s) if s.is_empty() => true,
        _ => false,
    }
}

fn coerce_to_number(value: &Value, locale: NumberLocale) -> Option<f64> {
    match value {
        Value::Number(n) => Some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        Value::Text(s) => parse_number(s, locale).ok(),
        Value::Array(arr) => coerce_to_number(&arr.top_left(), locale),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

fn coerce_to_text(value: &Value) -> Option<String> {
    match value {
        Value::Blank => None,
        Value::Number(_) | Value::Text(_) | Value::Bool(_) => value.coerce_to_string().ok(),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
        Value::Array(arr) => coerce_to_text(&arr.top_left()),
    }
}

fn parse_error_kind(raw: &str) -> Option<ErrorKind> {
    let normalized = raw.trim().to_ascii_uppercase();
    match normalized.as_str() {
        "#NULL!" => Some(ErrorKind::Null),
        "#DIV/0!" => Some(ErrorKind::Div0),
        "#VALUE!" => Some(ErrorKind::Value),
        "#REF!" => Some(ErrorKind::Ref),
        "#NAME?" => Some(ErrorKind::Name),
        "#NUM!" => Some(ErrorKind::Num),
        "#N/A" => Some(ErrorKind::NA),
        "#SPILL!" => Some(ErrorKind::Spill),
        "#CALC!" => Some(ErrorKind::Calc),
        _ => None,
    }
}

fn fold_case(input: &str) -> String {
    input.chars().flat_map(|c| c.to_uppercase()).collect()
}
