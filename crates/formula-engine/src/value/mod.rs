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

pub(crate) fn try_vec_with_capacity<T>(len: usize) -> Result<Vec<T>, ErrorKind> {
    let mut out = Vec::new();
    out.try_reserve_exact(len).map_err(|_| ErrorKind::Num)?;
    Ok(out)
}

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
    // Avoid the `ToUppercase` iterator for ASCII characters, even in mixed Unicode strings.
    let mut a_iter = FoldedUppercaseChars::new(a);
    let mut b_iter = FoldedUppercaseChars::new(b);
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

struct FoldedUppercaseChars<'a> {
    chars: std::str::Chars<'a>,
    pending: Option<std::char::ToUppercase>,
    ascii_needs_uppercasing: bool,
}

impl<'a> FoldedUppercaseChars<'a> {
    fn new(s: &'a str) -> Self {
        let ascii_needs_uppercasing = s.as_bytes().iter().any(|b| b.is_ascii_lowercase());
        Self {
            chars: s.chars(),
            pending: None,
            ascii_needs_uppercasing,
        }
    }
}

impl Iterator for FoldedUppercaseChars<'_> {
    type Item = char;

    fn next(&mut self) -> Option<char> {
        loop {
            if let Some(pending) = &mut self.pending {
                if let Some(ch) = pending.next() {
                    return Some(ch);
                }
                self.pending = None;
            }

            let ch = self.chars.next()?;
            if ch.is_ascii() {
                if self.ascii_needs_uppercasing {
                    return Some(ch.to_ascii_uppercase());
                }
                return Some(ch);
            }
            self.pending = Some(ch.to_uppercase());
        }
    }
}

#[inline]
pub(crate) fn eq_case_insensitive(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        a.eq_ignore_ascii_case(b)
    } else {
        cmp_case_insensitive(a, b) == Ordering::Equal
    }
}

fn try_fold_to_uppercase_string(s: &str) -> Option<String> {
    if s.is_empty() {
        return Some(String::new());
    }
    if s.is_ascii() {
        let bytes = s.as_bytes();
        if !bytes.iter().any(|b| b.is_ascii_lowercase()) {
            let mut out = String::new();
            if out.try_reserve_exact(s.len()).is_err() {
                return None;
            }
            out.push_str(s);
            return Some(out);
        }
        let mut out = String::new();
        if out.try_reserve_exact(s.len()).is_err() {
            return None;
        }
        for &b in bytes {
            out.push((b as char).to_ascii_uppercase());
        }
        return Some(out);
    }

    let mut out = String::new();
    if out.try_reserve_exact(s.len()).is_err() {
        return None;
    }
    for ch in s.chars() {
        if ch.is_ascii() {
            out.push(ch.to_ascii_uppercase());
        } else {
            out.extend(ch.to_uppercase());
        }
    }
    Some(out)
}


fn fold_to_lowercase_string(s: &str) -> String {
    if s.is_ascii() {
        let bytes = s.as_bytes();
        if !bytes.iter().any(|b| b.is_ascii_uppercase()) {
            let mut out = String::new();
            if out.try_reserve_exact(s.len()).is_err() {
                debug_assert!(false, "allocation failed (ascii lowercase copy)");
                return String::new();
            }
            out.push_str(s);
            return out;
        }
        let mut out = String::new();
        if out.try_reserve_exact(s.len()).is_err() {
            debug_assert!(false, "allocation failed (ascii lowercase)");
            return String::new();
        }
        for &b in bytes {
            out.push((b as char).to_ascii_lowercase());
        }
        return out;
    }

    let mut out = String::new();
    if out.try_reserve_exact(s.len()).is_err() {
        debug_assert!(false, "allocation failed (unicode lowercase)");
        return String::new();
    }
    for ch in s.chars() {
        if ch.is_ascii() {
            out.push(ch.to_ascii_lowercase());
        } else {
            out.extend(ch.to_lowercase());
        }
    }
    out
}

/// Fold a string using Unicode-aware uppercasing into a `Vec<char>`.
///
/// This is used by case-insensitive matching in places that need a character stream (rather than a
/// `String`) because Unicode uppercasing can expand a single character into multiple characters
/// (e.g. `ß` → `SS`).
pub(crate) fn fold_to_uppercase_chars(s: &str) -> Vec<char> {
    if s.is_ascii() {
        let needs_uppercasing = s.as_bytes().iter().any(|b| b.is_ascii_lowercase());
        let mut out: Vec<char> = Vec::new();
        if out.try_reserve_exact(s.len()).is_err() {
            debug_assert!(false, "allocation failed (fold_to_uppercase_chars ascii)");
            return Vec::new();
        }
        if needs_uppercasing {
            out.extend(s.as_bytes().iter().map(|&b| (b as char).to_ascii_uppercase()));
        } else {
            out.extend(s.chars());
        }
        return out;
    }

    let mut out: Vec<char> = Vec::new();
    if out.try_reserve_exact(s.len()).is_err() {
        debug_assert!(false, "allocation failed (fold_to_uppercase_chars unicode)");
        return Vec::new();
    }
    for ch in s.chars() {
        if ch.is_ascii() {
            out.push(ch.to_ascii_uppercase());
        } else {
            out.extend(ch.to_uppercase());
        }
    }
    out
}

pub(crate) fn fold_str_to_uppercase_with_starts(
    s: &str,
    ascii_needs_uppercasing: bool,
) -> (Vec<char>, Vec<usize>) {
    let mut folded: Vec<char> = Vec::new();
    if folded.try_reserve_exact(s.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (fold_str_to_uppercase_with_starts folded)"
        );
        return (Vec::new(), Vec::new());
    }
    let mut starts: Vec<usize> = Vec::new();
    if starts.try_reserve_exact(s.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (fold_str_to_uppercase_with_starts starts)"
        );
        return (Vec::new(), Vec::new());
    }
    for ch in s.chars() {
        starts.push(folded.len());
        if ch.is_ascii() {
            if ascii_needs_uppercasing {
                folded.push(ch.to_ascii_uppercase());
            } else {
                folded.push(ch);
            }
        } else {
            folded.extend(ch.to_uppercase());
        }
    }
    (folded, starts)
}

pub(crate) fn casefold_owned(mut s: String) -> String {
    let mut ascii_needs_uppercasing = false;
    for &b in s.as_bytes() {
        if b >= 0x80 {
            return match try_fold_to_uppercase_string(&s) {
                Some(v) => v,
                None => {
                    debug_assert!(
                        false,
                        "allocation failed (casefold_owned unicode, len={})",
                        s.len()
                    );
                    s
                }
            };
        }
        ascii_needs_uppercasing |= b.is_ascii_lowercase();
    }
    if ascii_needs_uppercasing {
        s.make_ascii_uppercase();
    }
    s
}

pub(crate) fn lowercase_owned(mut s: String) -> String {
    let mut ascii_needs_lowercasing = false;
    for &b in s.as_bytes() {
        if b >= 0x80 {
            let folded = fold_to_lowercase_string(&s);
            if folded.is_empty() && !s.is_empty() {
                debug_assert!(
                    false,
                    "allocation failed (lowercase_owned unicode, len={})",
                    s.len()
                );
                return s;
            }
            return folded;
        }
        ascii_needs_lowercasing |= b.is_ascii_uppercase();
    }
    if ascii_needs_lowercasing {
        s.make_ascii_lowercase();
    }
    s
}

#[inline]
pub(crate) fn with_ascii_uppercased_key<R>(s: &str, f: impl FnOnce(&str) -> R) -> R {
    // Equivalent to `s.to_ascii_uppercase()`, but avoids allocating for common short strings by
    // uppercasing into a small stack buffer.
    //
    // ASCII uppercasing preserves UTF-8 validity because it only mutates bytes in the `a-z`
    // range (all non-ASCII UTF-8 bytes are >= 0x80).
    let bytes = s.as_bytes();
    if !bytes.iter().any(|b| b.is_ascii_lowercase()) {
        return f(s);
    }

    let mut buf = [0u8; 64];
    if bytes.len() <= buf.len() {
        for (dst, src) in buf[..bytes.len()].iter_mut().zip(bytes) {
            *dst = src.to_ascii_uppercase();
        }
        let upper = match std::str::from_utf8(&buf[..bytes.len()]) {
            Ok(s) => s,
            Err(_) => {
                debug_assert!(false, "ASCII uppercasing should preserve UTF-8");
                return f(s);
            }
        };
        return f(upper);
    }

    let mut upper = String::new();
    if upper.try_reserve_exact(bytes.len()).is_err() {
        debug_assert!(false, "allocation failed (ascii uppercase key)");
        return f(s);
    }
    for &b in bytes {
        upper.push((b as char).to_ascii_uppercase());
    }
    f(&upper)
}

pub(crate) fn with_casefolded_key<R>(s: &str, f: impl FnOnce(&str) -> R) -> R {
    if s.is_ascii() {
        return with_ascii_uppercased_key(s, f);
    }

    match try_fold_to_uppercase_string(s) {
        Some(folded) => f(&folded),
        None => {
            debug_assert!(
                false,
                "allocation failed (with_casefolded_key, len={})",
                s.len()
            );
            f(s)
        }
    }
}

pub(crate) fn try_casefold(s: &str) -> Result<String, ErrorKind> {
    if s.is_empty() {
        return Ok(String::new());
    }
    try_fold_to_uppercase_string(s).ok_or(ErrorKind::Num)
}

#[inline]
pub(crate) fn casefolded_key_arc(s: &str) -> Arc<str> {
    with_casefolded_key(s, |folded| Arc::from(folded))
}

#[inline]
pub(crate) fn casefolded_key_arc_if(s: &str, pred: impl FnOnce(&str) -> bool) -> Option<Arc<str>> {
    with_casefolded_key(s, |folded| pred(folded).then(|| Arc::from(folded)))
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
        if let Some(v) = with_casefolded_key(field, |folded| self.fields.get(folded).cloned()) {
            return Some(v);
        }

        self.fields
            .iter()
            .find(|(k, _)| eq_case_insensitive(k, field))
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
        if let Some(v) = with_casefolded_key(field, |folded| self.fields.get(folded).cloned()) {
            return Some(v);
        }

        self.fields
            .iter()
            .find(|(k, _)| eq_case_insensitive(k, field))
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
                let mut entries: Vec<(&String, &Value)> = Vec::new();
                if entries.try_reserve_exact(self.0.len()).is_err() {
                    debug_assert!(false, "allocation failed (Lambda debug env)");
                    let mut map = f.debug_map();
                    for (k, v) in self.0 {
                        map.entry(&k, v);
                    }
                    return map.finish();
                }
                for (k, v) in self.0.iter() {
                    entries.push((k, v));
                }
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

    #[test]
    fn with_ascii_uppercased_key_matches_to_ascii_uppercase() {
        let long = "a".repeat(100);
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup", long.as_str()] {
            let out = with_ascii_uppercased_key(s, |upper| upper.to_string());
            assert_eq!(out, s.to_ascii_uppercase(), "input={s:?}");
        }
    }

    #[test]
    fn with_casefolded_key_matches_casefold() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            let mut out: Option<String> = None;
            with_casefolded_key(s, |key| {
                out = Some(key.to_string());
            });
            let expected = try_casefold(s).unwrap();
            assert_eq!(out.as_deref(), Some(expected.as_str()), "input={s:?}");
        }
    }

    #[test]
    fn cmp_case_insensitive_handles_mixed_ascii_and_unicode() {
        assert_eq!(cmp_case_insensitive("aö", "AÖ"), Ordering::Equal);
        assert_eq!(cmp_case_insensitive("straße", "STRASSE"), Ordering::Equal);
    }

    #[test]
    fn casefolded_key_arc_if_matches_casefold_when_predicate_passes() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            let key = casefolded_key_arc_if(s, |_| true).unwrap();
            let expected = try_casefold(s).unwrap();
            assert_eq!(key.as_ref(), expected.as_str(), "input={s:?}");
        }
    }

    #[test]
    fn casefolded_key_arc_if_returns_none_when_predicate_fails() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            assert!(casefolded_key_arc_if(s, |_| false).is_none(), "input={s:?}");
        }
    }

    #[test]
    fn casefold_owned_matches_casefold() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            let expected = try_casefold(s).unwrap();
            assert_eq!(casefold_owned(s.to_string()), expected, "input={s:?}");
        }
    }

    #[test]
    fn lowercase_owned_matches_to_lowercase() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            assert_eq!(
                lowercase_owned(s.to_string()),
                s.to_lowercase(),
                "input={s:?}"
            );
        }
    }

    #[test]
    fn fold_to_uppercase_chars_matches_unicode_to_uppercase_expansion() {
        for s in ["foo", "FOO", "Straße", "ß", "_xlfn.xlookup"] {
            let folded: Vec<char> = fold_to_uppercase_chars(s);
            let mut expected: Vec<char> = Vec::new();
            if expected.try_reserve_exact(s.len()).is_err() {
                panic!("allocation failed (upper casefold expected, input={s:?})");
            }
            for c in s.chars().flat_map(|c| c.to_uppercase()) {
                expected.push(c);
            }
            assert_eq!(folded, expected, "input={s:?}");
        }
    }
}
