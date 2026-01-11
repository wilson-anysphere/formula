use ordered_float::OrderedFloat;
use std::fmt;
use std::sync::Arc;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum Value {
    Blank,
    Number(OrderedFloat<f64>),
    Text(Arc<str>),
    Boolean(bool),
}

impl Value {
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Value::Number(n) => Some(n.0),
            _ => None,
        }
    }

    pub fn is_blank(&self) -> bool {
        matches!(self, Value::Blank)
    }

    pub fn truthy(&self) -> Result<bool, &'static str> {
        match self {
            Value::Boolean(b) => Ok(*b),
            Value::Number(n) => Ok(n.0 != 0.0),
            Value::Blank => Ok(false),
            Value::Text(_) => Err("cannot interpret text as boolean"),
        }
    }
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Blank => write!(f, "BLANK"),
            Value::Number(n) => write!(f, "{n}"),
            Value::Text(s) => write!(f, "{s:?}"),
            Value::Boolean(b) => write!(f, "{b}"),
        }
    }
}

impl From<f64> for Value {
    fn from(value: f64) -> Self {
        Value::Number(OrderedFloat(value))
    }
}

impl From<i64> for Value {
    fn from(value: i64) -> Self {
        Value::Number(OrderedFloat(value as f64))
    }
}

impl From<i32> for Value {
    fn from(value: i32) -> Self {
        Value::Number(OrderedFloat(value as f64))
    }
}

impl From<bool> for Value {
    fn from(value: bool) -> Self {
        Value::Boolean(value)
    }
}

impl From<&str> for Value {
    fn from(value: &str) -> Self {
        Value::Text(Arc::<str>::from(value))
    }
}

impl From<String> for Value {
    fn from(value: String) -> Self {
        Value::Text(Arc::<str>::from(value))
    }
}

impl From<Arc<str>> for Value {
    fn from(value: Arc<str>) -> Self {
        Value::Text(value)
    }
}
