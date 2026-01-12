use std::collections::HashMap;
use std::cmp::Ordering;
use std::fmt;
use std::sync::Arc;

use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::{FunctionContext, Reference};
use crate::locale::ValueLocaleConfig;
use formula_model::CellRef;
use formula_format::{DateSystem, FormatOptions, Value as FmtValue};

mod number_parse;

pub use number_parse::NumberLocale;
pub(crate) use number_parse::parse_number;
use crate::date::ExcelDateSystem;

pub(crate) fn cmp_ascii_case_insensitive(a: &str, b: &str) -> Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

pub(crate) fn cmp_case_insensitive(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    // Compare using Unicode-aware uppercasing so matches behave like Excel (e.g. ß -> SS).
    // This intentionally uses the same `char::to_uppercase` logic as criteria matching and
    // lookup semantics.
    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

pub(crate) fn casefold(s: &str) -> String {
    if s.is_ascii() {
        return s.to_ascii_uppercase();
    }

    s.chars().flat_map(|c| c.to_uppercase()).collect()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ErrorKind {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    Spill,
    Calc,
}

impl ErrorKind {
    pub fn as_code(self) -> &'static str {
        match self {
            ErrorKind::Null => "#NULL!",
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
            ErrorKind::Spill => "#SPILL!",
            ErrorKind::Calc => "#CALC!",
        }
    }
}

impl fmt::Display for ErrorKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_code())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Array {
    pub rows: usize,
    pub cols: usize,
    /// Row-major order values (length = rows * cols).
    pub values: Vec<Value>,
}

impl Array {
    #[must_use]
    pub fn new(rows: usize, cols: usize, values: Vec<Value>) -> Self {
        debug_assert_eq!(rows.saturating_mul(cols), values.len());
        Self { rows, cols, values }
    }

    #[must_use]
    pub fn get(&self, row: usize, col: usize) -> Option<&Value> {
        if row >= self.rows || col >= self.cols {
            return None;
        }
        self.values.get(row * self.cols + col)
    }

    #[must_use]
    pub fn top_left(&self) -> Value {
        self.get(0, 0).cloned().unwrap_or(Value::Blank)
    }

    pub fn iter(&self) -> impl Iterator<Item = &Value> {
        self.values.iter()
    }
}

#[derive(Clone, PartialEq)]
pub struct Lambda {
    pub params: Arc<[String]>,
    pub body: Arc<CompiledExpr>,
    pub env: Arc<HashMap<String, Value>>,
}

impl fmt::Debug for Lambda {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        struct SortedEnv<'a>(&'a HashMap<String, Value>);

        impl fmt::Debug for SortedEnv<'_> {
            fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
                let mut entries: Vec<_> = self.0.iter().collect();
                entries.sort_by(|(a, _), (b, _)| a.cmp(b));
                let mut map = f.debug_map();
                for (k, v) in entries {
                    map.entry(&k, v);
                }
                map.finish()
            }
        }

        f.debug_struct("Lambda")
            .field("params", &self.params)
            .field("body", &self.body)
            .field("env", &SortedEnv(&self.env))
            .finish()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
    Error(ErrorKind),
    /// Reference value returned by functions like OFFSET/INDIRECT.
    ///
    /// This variant is not intended to be stored in grid cells directly; the evaluator will
    /// typically dereference it into a scalar or dynamic array result depending on context.
    Reference(Reference),
    /// Multi-area reference value (union) returned by reference-producing functions.
    ReferenceUnion(Vec<Reference>),
    /// Dynamic array result.
    Array(Array),
    Lambda(Lambda),
    /// Marker for a cell that belongs to a spilled range.
    ///
    /// The engine generally resolves spill markers to the concrete spilled value
    /// when reading cell values; this variant is primarily used internally to
    /// track spill occupancy.
    Spill {
        origin: CellRef,
    },
}

impl Value {
    pub fn is_error(&self) -> bool {
        matches!(self, Value::Error(_))
    }

    pub fn coerce_to_number_with_ctx(&self, ctx: &dyn FunctionContext) -> Result<f64, ErrorKind> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Ok(0.0),
            Value::Text(s) => coerce_text_to_number(s, ctx.value_locale(), ctx.now_utc(), ctx.date_system()),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_number(&self) -> Result<f64, ErrorKind> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Ok(0.0),
            Value::Text(s) => coerce_text_to_number(
                s,
                ValueLocaleConfig::en_us(),
                chrono::Utc::now(),
                ExcelDateSystem::EXCEL_1900,
            ),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_number_with_locale(&self, locale: NumberLocale) -> Result<f64, ErrorKind> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Ok(0.0),
            Value::Text(s) => parse_number(s, locale).map_err(map_excel_error),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_i64_with_ctx(&self, ctx: &dyn FunctionContext) -> Result<i64, ErrorKind> {
        let n = self.coerce_to_number_with_ctx(ctx)?;
        Ok(n.trunc() as i64)
    }

    pub fn coerce_to_i64(&self) -> Result<i64, ErrorKind> {
        let n = self.coerce_to_number()?;
        Ok(n.trunc() as i64)
    }

    pub fn coerce_to_bool_with_ctx(&self, ctx: &dyn FunctionContext) -> Result<bool, ErrorKind> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Blank => Ok(false),
            Value::Text(s) => {
                let t = s.trim();
                if t.is_empty() {
                    return Ok(false);
                }
                if t.eq_ignore_ascii_case("TRUE") {
                    return Ok(true);
                }
                if t.eq_ignore_ascii_case("FALSE") {
                    return Ok(false);
                }
                Ok(self.coerce_to_number_with_ctx(ctx)? != 0.0)
            }
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_bool(&self) -> Result<bool, ErrorKind> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Blank => Ok(false),
            Value::Text(s) => {
                let t = s.trim();
                if t.is_empty() {
                    return Ok(false);
                }
                if t.eq_ignore_ascii_case("TRUE") {
                    return Ok(true);
                }
                if t.eq_ignore_ascii_case("FALSE") {
                    return Ok(false);
                }
                Ok(coerce_text_to_number(
                    s,
                    ValueLocaleConfig::en_us(),
                    chrono::Utc::now(),
                    ExcelDateSystem::EXCEL_1900,
                )? != 0.0)
            }
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_bool_with_locale(&self, locale: NumberLocale) -> Result<bool, ErrorKind> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Blank => Ok(false),
            Value::Text(s) => {
                let t = s.trim();
                if t.is_empty() {
                    return Ok(false);
                }
                if t.eq_ignore_ascii_case("TRUE") {
                    return Ok(true);
                }
                if t.eq_ignore_ascii_case("FALSE") {
                    return Ok(false);
                }
                let n = parse_number(t, locale).map_err(map_excel_error)?;
                Ok(n != 0.0)
            }
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_string(&self) -> Result<String, ErrorKind> {
        match self {
            Value::Text(s) => Ok(s.clone()),
            Value::Number(n) => Ok(format_number_general(*n)),
            Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            Value::Blank => Ok(String::new()),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_string_with_ctx(&self, ctx: &dyn FunctionContext) -> Result<String, ErrorKind> {
        match self {
            Value::Text(s) => Ok(s.clone()),
            Value::Number(n) => {
                let options = FormatOptions {
                    locale: ctx.value_locale().separators,
                    date_system: match ctx.date_system() {
                        // `formula-format` always uses the Lotus 1-2-3 leap-year bug behavior
                        // for the 1900 date system (Excel compatibility).
                        ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
                        ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
                    },
                };
                Ok(formula_format::format_value(FmtValue::Number(*n), None, &options).text)
            }
            Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            Value::Blank => Ok(String::new()),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }
}

fn map_excel_error(error: ExcelError) -> ErrorKind {
    match error {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn coerce_text_to_number(
    text: &str,
    cfg: ValueLocaleConfig,
    now_utc: chrono::DateTime<chrono::Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Ok(0.0);
    }

    crate::coercion::datetime::parse_value_text(trimmed, cfg, now_utc, system).map_err(map_excel_error)
}

fn format_number_general(n: f64) -> String {
    if !n.is_finite() {
        return n.to_string();
    }

    // Excel does not surface a negative sign for zero under "General".
    if n == 0.0 {
        return "0".to_string();
    }

    let abs = n.abs();
    // Approximate Excel's General switching thresholds:
    // - Scientific when abs >= 1e11
    // - Scientific when abs < 1e-9
    let scientific = abs >= 1e11 || abs < 1e-9;
    if scientific {
        format_number_general_scientific(n)
    } else {
        format_number_general_decimal(n)
    }
}

fn format_number_general_decimal(n: f64) -> String {
    let abs = n.abs();
    let digits_before_decimal = if abs >= 1.0 {
        // `log10` can be slightly off for some values; use it only for determining a reasonable
        // fixed precision and rely on formatting+trimming for the final representation.
        abs.log10().floor() as i32 + 1
    } else {
        0
    };

    let precision = (15 - digits_before_decimal).clamp(0, 15) as usize;
    let mut out = format!("{:.*}", precision, n);
    trim_trailing_decimal_zeros(&mut out);
    out
}

fn format_number_general_scientific(n: f64) -> String {
    let abs = n.abs();
    // Compute exponent such that 1 <= mantissa < 10.
    let mut exp = abs.log10().floor() as i32;
    let mut mantissa = abs / 10_f64.powi(exp);
    if mantissa < 1.0 {
        exp -= 1;
        mantissa *= 10.0;
    } else if mantissa >= 10.0 {
        exp += 1;
        mantissa /= 10.0;
    }

    // Excel uses 15 significant digits.
    let mut mantissa = (mantissa * 1e14).round() / 1e14;
    if mantissa >= 10.0 {
        mantissa /= 10.0;
        exp += 1;
    }

    let mut mantissa_str = format!("{:.14}", mantissa);
    trim_trailing_decimal_zeros(&mut mantissa_str);
    if n.is_sign_negative() {
        mantissa_str.insert(0, '-');
    }

    let exp_sign = if exp >= 0 { '+' } else { '-' };
    let exp_abs = exp.unsigned_abs();
    let exp_str = if exp_abs < 10 {
        format!("0{exp_abs}")
    } else {
        exp_abs.to_string()
    };

    format!("{mantissa_str}E{exp_sign}{exp_str}")
}

fn trim_trailing_decimal_zeros(s: &mut String) {
    if let Some(dot) = s.find('.') {
        // Strip trailing zeros after the decimal point.
        while s.ends_with('0') {
            s.pop();
        }
        // Remove the decimal point if it is now the last character.
        if s.len() == dot + 1 {
            s.pop();
        }
    }

    if s == "-0" {
        *s = "0".to_string();
    }
}

impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Value::Number(value)
    }
}

impl From<i64> for Value {
    fn from(value: i64) -> Self {
        Value::Number(value as f64)
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Value::Bool(value)
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Value::Text(value.to_string())
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Number(n) => write!(f, "{n}"),
            Value::Text(s) => f.write_str(s),
            Value::Bool(b) => write!(f, "{b}"),
            Value::Blank => f.write_str(""),
            Value::Error(e) => write!(f, "{e}"),
            Value::Reference(_) | Value::ReferenceUnion(_) => {
                f.write_str(ErrorKind::Value.as_code())
            }
            Value::Array(arr) => write!(f, "{}", arr.top_left()),
            Value::Lambda(_) => f.write_str(ErrorKind::Calc.as_code()),
            Value::Spill { .. } => f.write_str(ErrorKind::Spill.as_code()),
        }
    }
}
