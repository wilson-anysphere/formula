use std::cell::RefCell;
use std::fmt;
use std::rc::Rc;

use crate::object_model::VbaObjectRef;

#[derive(Clone, Debug)]
pub struct VbaArray {
    pub lower_bound: i32,
    pub values: Vec<VbaValue>,
}

impl VbaArray {
    pub fn new(lower_bound: i32, values: Vec<VbaValue>) -> Self {
        Self { lower_bound, values }
    }

    pub fn len(&self) -> usize {
        self.values.len()
    }

    pub fn upper_bound(&self) -> i32 {
        self.lower_bound + self.values.len().saturating_sub(1) as i32
    }

    pub fn get(&self, index: i32) -> Option<&VbaValue> {
        let offset = index.checked_sub(self.lower_bound)? as isize;
        if offset < 0 {
            return None;
        }
        self.values.get(offset as usize)
    }

    pub fn get_mut(&mut self, index: i32) -> Option<&mut VbaValue> {
        let offset = index.checked_sub(self.lower_bound)? as isize;
        if offset < 0 {
            return None;
        }
        self.values.get_mut(offset as usize)
    }
}

pub type VbaArrayRef = Rc<RefCell<VbaArray>>;

/// A minimal VBA Variant-like value.
#[derive(Clone, Default)]
pub enum VbaValue {
    #[default]
    Empty,
    Null,
    Boolean(bool),
    Double(f64),
    Date(f64),
    String(String),
    Object(VbaObjectRef),
    Array(VbaArrayRef),
}

impl fmt::Debug for VbaValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Empty => write!(f, "Empty"),
            Self::Null => write!(f, "Null"),
            Self::Boolean(v) => write!(f, "Boolean({v})"),
            Self::Double(v) => write!(f, "Double({v})"),
            Self::Date(v) => write!(f, "Date({v})"),
            Self::String(v) => write!(f, "String({v:?})"),
            Self::Object(_) => write!(f, "Object(..)"),
            Self::Array(v) => write!(f, "Array(len={})", v.borrow().len()),
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
            (Self::Date(a), Self::Date(b)) => a == b,
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
            Self::Date(v) => *v != 0.0,
            Self::String(v) => !v.is_empty(),
            Self::Empty | Self::Null => false,
            Self::Object(_) => true,
            Self::Array(v) => v.borrow().len() != 0,
        }
    }

    pub fn to_f64(&self) -> Option<f64> {
        match self {
            Self::Double(v) => Some(*v),
            Self::Date(v) => Some(*v),
            Self::Boolean(v) => Some(if *v { -1.0 } else { 0.0 }),
            Self::String(s) => s.parse::<f64>().ok(),
            Self::Empty => Some(0.0),
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
            Self::Date(v) => {
                // Format as an OLE Automation date serial for now.
                let mut s = format!("{v}");
                if s.ends_with(".0") {
                    s.truncate(s.len() - 2);
                }
                s
            }
            Self::String(v) => v.clone(),
            Self::Object(_) => "<Object>".to_string(),
            Self::Array(v) => format!("<Array {}>", v.borrow().len()),
        }
    }

    pub fn as_object(&self) -> Option<VbaObjectRef> {
        match self {
            Self::Object(o) => Some(o.clone()),
            _ => None,
        }
    }

    pub fn as_array(&self) -> Option<VbaArrayRef> {
        match self {
            Self::Array(arr) => Some(arr.clone()),
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
