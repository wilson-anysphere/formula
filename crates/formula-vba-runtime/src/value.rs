use std::fmt;
use std::rc::Rc;

use crate::object_model::VbaObjectRef;

/// A minimal VBA Variant-like value.
#[derive(Clone)]
pub enum VbaValue {
    Empty,
    Null,
    Boolean(bool),
    Double(f64),
    String(String),
    Object(VbaObjectRef),
    Array(Rc<Vec<VbaValue>>),
}

impl Default for VbaValue {
    fn default() -> Self {
        Self::Empty
    }
}

impl fmt::Debug for VbaValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => write!(f, "Empty"),
            Self::Null => write!(f, "Null"),
            Self::Boolean(v) => write!(f, "Boolean({v})"),
            Self::Double(v) => write!(f, "Double({v})"),
            Self::String(v) => write!(f, "String({v:?})"),
            Self::Object(_) => write!(f, "Object(..)"),
            Self::Array(v) => write!(f, "Array(len={})", v.len()),
        }
    }
}

impl PartialEq for VbaValue {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Empty, Self::Empty) => true,
            (Self::Null, Self::Null) => true,
            (Self::Boolean(a), Self::Boolean(b)) => a == b,
            (Self::Double(a), Self::Double(b)) => a == b,
            (Self::String(a), Self::String(b)) => a == b,
            _ => false,
        }
    }
}

impl VbaValue {
    pub fn is_truthy(&self) -> bool {
        match self {
            Self::Boolean(v) => *v,
            Self::Double(v) => *v != 0.0,
            Self::String(v) => !v.is_empty(),
            Self::Empty | Self::Null => false,
            Self::Object(_) => true,
            Self::Array(v) => !v.is_empty(),
        }
    }

    pub fn to_f64(&self) -> Option<f64> {
        match self {
            Self::Double(v) => Some(*v),
            Self::Boolean(v) => Some(if *v { -1.0 } else { 0.0 }),
            Self::String(s) => s.parse::<f64>().ok(),
            _ => None,
        }
    }

    pub fn to_string_lossy(&self) -> String {
        match self {
            Self::Empty => "".to_string(),
            Self::Null => "Null".to_string(),
            Self::Boolean(v) => {
                if *v {
                    "True".to_string()
                } else {
                    "False".to_string()
                }
            }
            Self::Double(v) => {
                // Excel/VBA have a complicated formatting model; keep it simple.
                let mut s = format!("{v}");
                if s.ends_with(".0") {
                    s.truncate(s.len() - 2);
                }
                s
            }
            Self::String(v) => v.clone(),
            Self::Object(_) => "<Object>".to_string(),
            Self::Array(v) => format!("<Array {}>", v.len()),
        }
    }

    pub fn as_object(&self) -> Option<VbaObjectRef> {
        match self {
            Self::Object(o) => Some(o.clone()),
            _ => None,
        }
    }
}

impl From<&str> for VbaValue {
    fn from(value: &str) -> Self {
        Self::String(value.to_string())
    }
}

impl From<String> for VbaValue {
    fn from(value: String) -> Self {
        Self::String(value)
    }
}

impl From<bool> for VbaValue {
    fn from(value: bool) -> Self {
        Self::Boolean(value)
    }
}

impl From<f64> for VbaValue {
    fn from(value: f64) -> Self {
        Self::Double(value)
    }
}

impl From<i32> for VbaValue {
    fn from(value: i32) -> Self {
        Self::Double(value as f64)
    }
}
