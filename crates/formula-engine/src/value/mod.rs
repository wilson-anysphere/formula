use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use crate::functions::Reference;
use formula_model::CellRef;

use crate::eval::CompiledExpr;

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

    pub fn coerce_to_number(&self) -> Result<f64, ErrorKind> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Ok(0.0),
            Value::Text(s) => parse_number_from_text(s).ok_or(ErrorKind::Value),
            Value::Error(e) => Err(*e),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub fn coerce_to_i64(&self) -> Result<i64, ErrorKind> {
        let n = self.coerce_to_number()?;
        Ok(n.trunc() as i64)
    }

    pub fn coerce_to_bool(&self) -> Result<bool, ErrorKind> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Blank => Ok(false),
            Value::Text(s) => {
                let t = s.trim();
                if t.eq_ignore_ascii_case("TRUE") {
                    return Ok(true);
                }
                if t.eq_ignore_ascii_case("FALSE") {
                    return Ok(false);
                }
                if let Some(n) = parse_number_from_text(t) {
                    return Ok(n != 0.0);
                }
                Err(ErrorKind::Value)
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
}

fn parse_number_from_text(s: &str) -> Option<f64> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.parse::<f64>().ok()
}

fn format_number_general(n: f64) -> String {
    if n == 0.0 {
        return "0".to_string();
    }
    if n.fract() == 0.0 {
        return format!("{:.0}", n);
    }
    let s = n.to_string();
    if s == "-0" || s == "-0.0" {
        "0".to_string()
    } else {
        s
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
            Value::Lambda(_) => f.write_str("<LAMBDA>"),
            Value::Spill { .. } => f.write_str(ErrorKind::Spill.as_code()),
        }
    }
}
