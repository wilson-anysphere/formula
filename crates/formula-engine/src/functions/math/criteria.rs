use chrono::{DateTime, Utc};

use crate::coercion::datetime::parse_value_text;
use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::functions::wildcard::WildcardPattern;
use crate::simd::{CmpOp, NumericCriteria};
use crate::value::format_number_general_with_options;
use crate::value::{casefold, parse_number, NumberLocale};
use crate::{ErrorKind, LocaleConfig, Value};

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
        let folded = casefold(raw);
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
    value_locale: ValueLocaleConfig,
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
        // `LocaleConfig::en_us()` disables thousands separators during lexing because `,` collides
        // with argument separators. Criteria values are already strings, so we can safely enable
        // grouping separators and reuse `LocaleConfig::parse_number`.
        let separators = value_locale.separators;
        let mut criteria_locale = LocaleConfig::en_us();
        criteria_locale.decimal_separator = separators.decimal_sep;
        criteria_locale.thousands_separator = Some(separators.thousands_sep);

        Self::parse_with_date_system_and_locales(
            input,
            system,
            value_locale,
            now_utc,
            criteria_locale,
        )
    }

    /// Parse an Excel criteria value using separate value (text -> number/date parsing) and
    /// formula locales.
    ///
    /// This is primarily used to parse numbers that appear inside string literals, such as
    /// criteria arguments entered in the workbook locale (e.g. `">1,5"` in `de-DE`).
    pub fn parse_with_date_system_and_locales(
        input: &Value,
        system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
        now_utc: DateTime<Utc>,
        locale: LocaleConfig,
    ) -> Result<Self, ErrorKind> {
        let separators = value_locale.separators;
        let number_locale =
            NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));

        match input {
            Value::Number(n) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Number(*n),
                value_locale,
                number_locale,
            }),
            Value::Bool(b) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Bool(*b),
                value_locale,
                number_locale,
            }),
            Value::Error(e) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Error(*e),
                value_locale,
                number_locale,
            }),
            Value::Blank => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Blank,
                value_locale,
                number_locale,
            }),
            Value::Text(s) => {
                parse_criteria_string(s, system, value_locale, now_utc, number_locale, &locale)
            }
            Value::Entity(entity) => parse_criteria_string(
                entity.display.as_str(),
                system,
                value_locale,
                now_utc,
                number_locale,
                &locale,
            ),
            Value::Record(record) => parse_criteria_string(
                record.display.as_str(),
                system,
                value_locale,
                now_utc,
                number_locale,
                &locale,
            ),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    /// Parse an Excel criteria value, resolving numeric/date/time strings using the provided value
    /// locale (decimal/thousands separators, date order) and workbook date system.
    ///
    /// This is equivalent to [`Criteria::parse_with_date_system_and_locale`] but uses a more
    /// ergonomic argument order for callers that already have a [`ValueLocaleConfig`] and
    /// deterministic `now_utc`.
    pub fn parse_with_locale_config(
        input: &Value,
        cfg: ValueLocaleConfig,
        now_utc: DateTime<Utc>,
        system: ExcelDateSystem,
    ) -> Result<Self, ErrorKind> {
        Self::parse_with_date_system_and_locale(input, system, cfg, now_utc)
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
            Value::Error(_) | Value::Lambda(_) | Value::Spill { .. } => {
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
            CriteriaRhs::Text(pattern) => {
                matches_text_criteria(self.op, pattern, value, self.value_locale)
            }
        }
    }
}

fn matches_error_criteria(op: &CriteriaOp, rhs: &CriteriaRhs, value: &Value) -> bool {
    let CriteriaRhs::Error(err) = rhs else {
        return false;
    };

    let candidate = match value {
        Value::Error(e) => Some(*e),
        Value::Lambda(_) => Some(ErrorKind::Value),
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

fn matches_text_criteria(
    op: CriteriaOp,
    pattern: &TextCriteria,
    value: &Value,
    value_locale: ValueLocaleConfig,
) -> bool {
    let Some(value_text) = coerce_to_text(value, value_locale) else {
        // Excel criteria functions treat blanks as "not text" for text-pattern matching (e.g.
        // COUNTIF(range,"*") does not count truly empty cells). For `<>` text criteria, non-text
        // values still satisfy the predicate because they are not equal to the text pattern.
        return matches!(op, CriteriaOp::Ne);
    };
    let value_folded = casefold(&value_text);

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
    locale: &LocaleConfig,
) -> Result<Criteria, ErrorKind> {
    let (op, rhs_str) = split_op(raw);
    let rhs_trimmed = rhs_str.trim();

    // Blank criteria are driven by the explicit RHS, not by whitespace.
    if rhs_str.is_empty() {
        return match op {
            CriteriaOp::Eq | CriteriaOp::Ne => Ok(Criteria {
                op,
                rhs: CriteriaRhs::Blank,
                value_locale,
                number_locale,
            }),
            _ => Err(ErrorKind::Value),
        };
    }

    // Excel criteria strings allow quoting the RHS as a text literal, e.g. `="foo"` or `=""`.
    // Treat quoted RHS values as text (bypassing number/date parsing). This matters for blank
    // criteria: `=""` should match blank cells (equivalent to `""` / `"="`).
    if let Some(text_literal) = parse_criteria_rhs_text_literal(rhs_trimmed) {
        if text_literal.is_empty() {
            return match op {
                CriteriaOp::Eq | CriteriaOp::Ne => Ok(Criteria {
                    op,
                    rhs: CriteriaRhs::Blank,
                    value_locale,
                    number_locale,
                }),
                _ => Err(ErrorKind::Value),
            };
        }

        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Text(TextCriteria::new(&text_literal)),
            value_locale,
            number_locale,
        });
    }

    if let Some(err) = parse_error_kind(rhs_trimmed) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Error(err),
            value_locale,
            number_locale,
        });
    }

    if rhs_trimmed.eq_ignore_ascii_case("TRUE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(true),
            value_locale,
            number_locale,
        });
    }
    if rhs_trimmed.eq_ignore_ascii_case("FALSE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(false),
            value_locale,
            number_locale,
        });
    }

    // Prefer locale-aware numeric parsing for values that appear inside string literals (e.g.
    // `">1,5"` in `de-DE`). This accepts both the locale decimal separator and canonical `.`.
    if let Some(num) = locale.parse_number(rhs_str) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Number(num),
            value_locale,
            number_locale,
        });
    }

    // Fall back to VALUE-like parsing for date/time strings (and a few other numeric formats).
    if let Ok(serial) = parse_value_text(rhs_str, value_locale, now_utc, system) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Number(serial),
            value_locale,
            number_locale,
        });
    }

    Ok(Criteria {
        op,
        rhs: CriteriaRhs::Text(TextCriteria::new(rhs_str)),
        value_locale,
        number_locale,
    })
}

/// Parse a quoted text literal in a criteria RHS.
///
/// In Excel, criteria strings can include a quoted RHS (e.g. `="foo"` or `<>""`). The quoting
/// uses Excel string escaping rules where `""` represents a literal `"`.
///
/// Returns `None` when `raw` is not a quoted literal or contains invalid escaping.
fn parse_criteria_rhs_text_literal(raw: &str) -> Option<String> {
    if raw.len() < 2 {
        return None;
    }
    let mut chars = raw.chars();
    if chars.next()? != '"' {
        return None;
    }
    if raw.chars().last()? != '"' {
        return None;
    }

    let inner = &raw[1..raw.len() - 1];
    let mut out = String::new();
    let mut it = inner.chars().peekable();
    while let Some(ch) = it.next() {
        if ch == '"' {
            if matches!(it.peek(), Some('"')) {
                it.next();
                out.push('"');
            } else {
                // Unescaped quote inside the quoted literal.
                return None;
            }
        } else {
            out.push(ch);
        }
    }
    Some(out)
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
        Value::Entity(e) if e.display.is_empty() => true,
        Value::Record(r) if r.display.is_empty() => true,
        _ => false,
    }
}

fn coerce_to_number(value: &Value, locale: NumberLocale) -> Option<f64> {
    match value {
        Value::Number(n) => Some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        Value::Text(s) => parse_number(s, locale).ok(),
        Value::Entity(entity) => parse_number(entity.display.as_str(), locale).ok(),
        Value::Record(record) => parse_number(record.display.as_str(), locale).ok(),
        Value::Array(arr) => coerce_to_number(&arr.top_left(), locale),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

fn coerce_to_text(value: &Value, value_locale: ValueLocaleConfig) -> Option<String> {
    match value {
        Value::Blank => None,
        Value::Text(_) | Value::Entity(_) | Value::Bool(_) => value.coerce_to_string().ok(),
        Value::Record(record) => {
            if let Some(display_field) = record.display_field.as_deref() {
                if let Some(value) = record.get_field_case_insensitive(display_field) {
                    return coerce_to_text(&value, value_locale);
                }
            }
            Some(record.display.clone())
        }
        Value::Number(n) => {
            // Criteria text matching always treats numbers as numbers under the "General" format;
            // the date system is irrelevant.
            Some(format_number_general_with_options(
                *n,
                value_locale.separators,
                ExcelDateSystem::EXCEL_1900,
            ))
        }
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
        Value::Array(arr) => coerce_to_text(&arr.top_left(), value_locale),
    }
}

fn parse_error_kind(raw: &str) -> Option<ErrorKind> {
    ErrorKind::from_code(raw)
}
