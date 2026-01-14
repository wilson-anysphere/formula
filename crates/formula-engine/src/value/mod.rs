use std::cmp::Ordering;
use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::{FunctionContext, Reference};
use crate::locale::ValueLocaleConfig;
use formula_format::{FormatOptions, Value as FmtValue};
use formula_model::{CellRef, ErrorValue as ModelErrorValue};

mod formatting;
mod number_parse;

use crate::date::ExcelDateSystem;
pub(crate) use formatting::format_number_general_with_options;
pub(crate) use number_parse::parse_number;
pub use number_parse::NumberLocale;

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
    GettingData,
    Spill,
    Calc,
    Field,
    Connect,
    Blocked,
    Unknown,
}

impl ErrorKind {
    /// Excel's canonical spelling for the error (including punctuation).
    pub const fn as_code(self) -> &'static str {
        match self {
            ErrorKind::Null => "#NULL!",
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
            ErrorKind::GettingData => "#GETTING_DATA",
            ErrorKind::Spill => "#SPILL!",
            ErrorKind::Calc => "#CALC!",
            ErrorKind::Field => "#FIELD!",
            ErrorKind::Connect => "#CONNECT!",
            ErrorKind::Blocked => "#BLOCKED!",
            ErrorKind::Unknown => "#UNKNOWN!",
        }
    }

    /// Numeric error code used by Excel in various internal representations.
    ///
    /// Values are based on the mapping documented in `docs/01-formula-engine.md`.
    pub const fn code(self) -> u8 {
        match self {
            ErrorKind::Null => 1,
            ErrorKind::Div0 => 2,
            ErrorKind::Value => 3,
            ErrorKind::Ref => 4,
            ErrorKind::Name => 5,
            ErrorKind::Num => 6,
            ErrorKind::NA => 7,
            ErrorKind::GettingData => 8,
            ErrorKind::Spill => 9,
            ErrorKind::Calc => 10,
            ErrorKind::Field => 11,
            ErrorKind::Connect => 12,
            ErrorKind::Blocked => 13,
            ErrorKind::Unknown => 14,
        }
    }

    /// Parse an Excel error literal (e.g. `#DIV/0!`) into an [`ErrorKind`].
    ///
    /// Returns `None` for unknown error literals.
    pub fn from_code(raw: &str) -> Option<Self> {
        let raw = raw.trim();
        if raw.eq_ignore_ascii_case("#NULL!") {
            return Some(ErrorKind::Null);
        }
        if raw.eq_ignore_ascii_case("#DIV/0!") {
            return Some(ErrorKind::Div0);
        }
        if raw.eq_ignore_ascii_case("#VALUE!") {
            return Some(ErrorKind::Value);
        }
        if raw.eq_ignore_ascii_case("#REF!") {
            return Some(ErrorKind::Ref);
        }
        if raw.eq_ignore_ascii_case("#NAME?") {
            return Some(ErrorKind::Name);
        }
        if raw.eq_ignore_ascii_case("#NUM!") {
            return Some(ErrorKind::Num);
        }
        if raw.eq_ignore_ascii_case("#N/A") || raw.eq_ignore_ascii_case("#N/A!") {
            return Some(ErrorKind::NA);
        }
        if raw.eq_ignore_ascii_case("#GETTING_DATA") {
            return Some(ErrorKind::GettingData);
        }
        if raw.eq_ignore_ascii_case("#SPILL!") {
            return Some(ErrorKind::Spill);
        }
        if raw.eq_ignore_ascii_case("#CALC!") {
            return Some(ErrorKind::Calc);
        }
        if raw.eq_ignore_ascii_case("#FIELD!") {
            return Some(ErrorKind::Field);
        }
        if raw.eq_ignore_ascii_case("#CONNECT!") {
            return Some(ErrorKind::Connect);
        }
        if raw.eq_ignore_ascii_case("#BLOCKED!") {
            return Some(ErrorKind::Blocked);
        }
        if raw.eq_ignore_ascii_case("#UNKNOWN!") {
            return Some(ErrorKind::Unknown);
        }
        None
    }
}

impl From<ModelErrorValue> for ErrorKind {
    fn from(value: ModelErrorValue) -> Self {
        match value {
            ModelErrorValue::Null => ErrorKind::Null,
            ModelErrorValue::Div0 => ErrorKind::Div0,
            ModelErrorValue::Value => ErrorKind::Value,
            ModelErrorValue::Ref => ErrorKind::Ref,
            ModelErrorValue::Name => ErrorKind::Name,
            ModelErrorValue::Num => ErrorKind::Num,
            ModelErrorValue::NA => ErrorKind::NA,
            ModelErrorValue::GettingData => ErrorKind::GettingData,
            ModelErrorValue::Spill => ErrorKind::Spill,
            ModelErrorValue::Calc => ErrorKind::Calc,
            ModelErrorValue::Field => ErrorKind::Field,
            ModelErrorValue::Connect => ErrorKind::Connect,
            ModelErrorValue::Blocked => ErrorKind::Blocked,
            ModelErrorValue::Unknown => ErrorKind::Unknown,
        }
    }
}

impl From<ErrorKind> for ModelErrorValue {
    fn from(value: ErrorKind) -> Self {
        match value {
            ErrorKind::Null => ModelErrorValue::Null,
            ErrorKind::Div0 => ModelErrorValue::Div0,
            ErrorKind::Value => ModelErrorValue::Value,
            ErrorKind::Ref => ModelErrorValue::Ref,
            ErrorKind::Name => ModelErrorValue::Name,
            ErrorKind::Num => ModelErrorValue::Num,
            ErrorKind::NA => ModelErrorValue::NA,
            ErrorKind::GettingData => ModelErrorValue::GettingData,
            ErrorKind::Spill => ModelErrorValue::Spill,
            ErrorKind::Calc => ModelErrorValue::Calc,
            ErrorKind::Field => ModelErrorValue::Field,
            ErrorKind::Connect => ModelErrorValue::Connect,
            ErrorKind::Blocked => ModelErrorValue::Blocked,
            ErrorKind::Unknown => ModelErrorValue::Unknown,
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

#[derive(Debug, Clone, PartialEq)]
pub struct EntityValue {
    /// Display string shown in the grid UI.
    pub display: String,
    /// Optional entity type discriminator (e.g. `"stock"`, `"geography"`).
    pub entity_type: Option<String>,
    /// Optional entity identifier (e.g. `"AAPL"`).
    pub entity_id: Option<String>,
    /// Field values for `.` access (case-insensitive at lookup time).
    pub fields: HashMap<String, Value>,
}

impl EntityValue {
    #[must_use]
    pub fn new(display_value: impl Into<String>) -> Self {
        Self {
            display: display_value.into(),
            entity_type: None,
            entity_id: None,
            fields: HashMap::new(),
        }
    }

    #[must_use]
    pub fn with_fields(display: impl Into<String>, fields: HashMap<String, Value>) -> Self {
        Self {
            display: display.into(),
            entity_type: None,
            entity_id: None,
            fields,
        }
    }

    #[must_use]
    pub fn with_properties<I, K, V>(display: impl Into<String>, properties: I) -> Self
    where
        I: IntoIterator<Item = (K, V)>,
        K: Into<String>,
        V: Into<Value>,
    {
        let mut map = HashMap::new();
        for (k, v) in properties {
            map.insert(k.into(), v.into());
        }
        Self::with_fields(display, map)
    }

    #[must_use]
    pub fn property(mut self, name: impl Into<String>, value: impl Into<Value>) -> Self {
        self.fields.insert(name.into(), value.into());
        self
    }

    pub fn get_field_case_insensitive(&self, field: &str) -> Option<Value> {
        // Common fast paths:
        // - exact-key match (already case-correct)
        // - pre-folded key storage (some builders may store case-folded keys)
        if let Some(v) = self.fields.get(field) {
            return Some(v.clone());
        }
        let folded = casefold(field);
        if let Some(v) = self.fields.get(&folded) {
            return Some(v.clone());
        }

        self.fields
            .iter()
            .find(|(k, _)| cmp_case_insensitive(k, field) == Ordering::Equal)
            .map(|(_, v)| v.clone())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RecordValue {
    /// Display string shown in the grid UI.
    pub display: String,
    /// Optional name of the field that should be used for display.
    pub display_field: Option<String>,
    /// Field values for `.` access (case-insensitive at lookup time).
    pub fields: HashMap<String, Value>,
}

impl RecordValue {
    #[must_use]
    pub fn new(display_value: impl Into<String>) -> Self {
        Self {
            display: display_value.into(),
            display_field: None,
            fields: HashMap::new(),
        }
    }

    #[must_use]
    pub fn with_fields(display: impl Into<String>, fields: HashMap<String, Value>) -> Self {
        Self {
            display: display.into(),
            display_field: None,
            fields,
        }
    }

    #[must_use]
    pub fn with_fields_iter<I, K, V>(display: impl Into<String>, fields: I) -> Self
    where
        I: IntoIterator<Item = (K, V)>,
        K: Into<String>,
        V: Into<Value>,
    {
        let mut map = HashMap::new();
        for (k, v) in fields {
            map.insert(k.into(), v.into());
        }
        Self::with_fields(display, map)
    }

    #[must_use]
    pub fn field(mut self, name: impl Into<String>, value: impl Into<Value>) -> Self {
        self.fields.insert(name.into(), value.into());
        self
    }

    pub fn get_field_case_insensitive(&self, field: &str) -> Option<Value> {
        if let Some(v) = self.fields.get(field) {
            return Some(v.clone());
        }
        let folded = casefold(field);
        if let Some(v) = self.fields.get(&folded) {
            return Some(v.clone());
        }

        self.fields
            .iter()
            .find(|(k, _)| cmp_case_insensitive(k, field) == Ordering::Equal)
            .map(|(_, v)| v.clone())
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

/// Convenience alias for [`EntityValue`].
pub type Entity = EntityValue;

/// Convenience alias for [`RecordValue`].
pub type Record = RecordValue;

#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Entity(EntityValue),
    Record(RecordValue),
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
            Value::Text(s) => {
                coerce_text_to_number(s, ctx.value_locale(), ctx.now_utc(), ctx.date_system())
            }
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
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
            Value::Entity(v) => Ok(v.display.clone()),
            Value::Record(v) => {
                if let Some(display_field) = v.display_field.as_deref() {
                    if let Some(value) = v.get_field_case_insensitive(display_field) {
                        return value.coerce_to_string();
                    }
                }
                Ok(v.display.clone())
            }
            Value::Number(n) => Ok(format_number_general_with_options(
                *n,
                ValueLocaleConfig::en_us().separators,
                ExcelDateSystem::EXCEL_1900,
            )),
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

    pub fn coerce_to_string_with_ctx(
        &self,
        ctx: &dyn FunctionContext,
    ) -> Result<String, ErrorKind> {
        match self {
            Value::Text(s) => Ok(s.clone()),
            Value::Entity(v) => Ok(v.display.clone()),
            Value::Record(v) => {
                if let Some(display_field) = v.display_field.as_deref() {
                    if let Some(value) = v.get_field_case_insensitive(display_field) {
                        return value.coerce_to_string_with_ctx(ctx);
                    }
                }
                Ok(v.display.clone())
            }
            Value::Number(n) => Ok(format_number_general_with_options(
                *n,
                ctx.value_locale().separators,
                ctx.date_system(),
            )),
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

    crate::coercion::datetime::parse_value_text(trimmed, cfg, now_utc, system)
        .map_err(map_excel_error)
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
            Value::Number(n) => {
                // `Display` intentionally uses Excel-like "General" formatting (en-US fallback)
                // so any accidental `.to_string()` or `format!("{value}")` behaves like Excel.
                let options = FormatOptions::default();
                let fmt_value = FmtValue::Number(*n);
                let formatted = formula_format::format_value(fmt_value, None, &options).text;
                f.write_str(&formatted)
            }
            Value::Text(s) => f.write_str(s),
            Value::Bool(b) => f.write_str(if *b { "TRUE" } else { "FALSE" }),
            Value::Entity(entity) => f.write_str(&entity.display),
            Value::Record(record) => {
                if let Some(display_field) = record.display_field.as_deref() {
                    if let Some(value) = record.get_field_case_insensitive(display_field) {
                        return write!(f, "{value}");
                    }
                }
                f.write_str(&record.display)
            }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn display_number_uses_excel_like_general_formatting() {
        assert_eq!(Value::Number(-0.0).to_string(), "0");
    }

    #[test]
    fn display_bool_uses_excel_spelling() {
        assert_eq!(Value::Bool(true).to_string(), "TRUE");
        assert_eq!(Value::Bool(false).to_string(), "FALSE");
    }

    #[test]
    fn display_record_honors_display_field() {
        let mut record = RecordValue::with_fields_iter("Fallback", [("name", "Apple")]);
        record.display_field = Some("Name".to_string());
        assert_eq!(Value::Record(record).to_string(), "Apple");
    }
}
