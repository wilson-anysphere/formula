use crate::{ErrorKind, Value};
use crate::functions::wildcard::wildcard_match;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CriteriaOp {
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum CriteriaRhs {
    Blank,
    Number(f64),
    Bool(bool),
    Error(ErrorKind),
    Text(String),
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Criteria {
    pub(crate) op: CriteriaOp,
    pub(crate) rhs: CriteriaRhs,
}

impl Criteria {
    pub(crate) fn parse(input: &Value) -> Result<Self, ErrorKind> {
        match input {
            Value::Number(n) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Number(*n),
            }),
            Value::Bool(b) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Bool(*b),
            }),
            Value::Error(e) => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Error(*e),
            }),
            Value::Blank => Ok(Criteria {
                op: CriteriaOp::Eq,
                rhs: CriteriaRhs::Blank,
            }),
            Value::Text(s) => parse_criteria_string(s),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    pub(crate) fn matches(&self, value: &Value) -> bool {
        match &self.rhs {
            CriteriaRhs::Blank => match self.op {
                CriteriaOp::Eq => is_blank_value(value),
                CriteriaOp::Ne => !is_blank_value(value),
                _ => false,
            },
            CriteriaRhs::Error(err) => match self.op {
                CriteriaOp::Eq => matches!(value, Value::Error(e) if e == err),
                CriteriaOp::Ne => !matches!(value, Value::Error(e) if e == err),
                _ => false,
            },
            CriteriaRhs::Bool(b) => {
                let Some(value_bool) = coerce_to_bool(value) else {
                    return false;
                };
                match self.op {
                    CriteriaOp::Eq => value_bool == *b,
                    CriteriaOp::Ne => value_bool != *b,
                    _ => false,
                }
            }
            CriteriaRhs::Number(n) => {
                let Some(value_num) = coerce_to_number(value) else {
                    return false;
                };
                match self.op {
                    CriteriaOp::Eq => value_num == *n,
                    CriteriaOp::Ne => value_num != *n,
                    CriteriaOp::Lt => value_num < *n,
                    CriteriaOp::Lte => value_num <= *n,
                    CriteriaOp::Gt => value_num > *n,
                    CriteriaOp::Gte => value_num >= *n,
                }
            }
            CriteriaRhs::Text(pattern) => {
                let value_text = coerce_to_text(value);
                match self.op {
                    CriteriaOp::Eq => wildcard_match(pattern, &value_text),
                    CriteriaOp::Ne => !wildcard_match(pattern, &value_text),
                    CriteriaOp::Lt => {
                        value_text.to_ascii_uppercase() < pattern.to_ascii_uppercase()
                    }
                    CriteriaOp::Lte => {
                        value_text.to_ascii_uppercase() <= pattern.to_ascii_uppercase()
                    }
                    CriteriaOp::Gt => {
                        value_text.to_ascii_uppercase() > pattern.to_ascii_uppercase()
                    }
                    CriteriaOp::Gte => {
                        value_text.to_ascii_uppercase() >= pattern.to_ascii_uppercase()
                    }
                }
            }
        }
    }
}

fn parse_criteria_string(raw: &str) -> Result<Criteria, ErrorKind> {
    let (op, rhs_str) = split_op(raw);
    let rhs_str = rhs_str.trim();

    if rhs_str.is_empty() {
        return match op {
            CriteriaOp::Eq | CriteriaOp::Ne => Ok(Criteria {
                op,
                rhs: CriteriaRhs::Blank,
            }),
            _ => Err(ErrorKind::Value),
        };
    }

    if let Some(err) = parse_error_kind(rhs_str) {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Error(err),
        });
    }

    if rhs_str.eq_ignore_ascii_case("TRUE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(true),
        });
    }
    if rhs_str.eq_ignore_ascii_case("FALSE") {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Bool(false),
        });
    }

    if let Ok(num) = rhs_str.parse::<f64>() {
        return Ok(Criteria {
            op,
            rhs: CriteriaRhs::Number(num),
        });
    }

    Ok(Criteria {
        op,
        rhs: CriteriaRhs::Text(rhs_str.to_string()),
    })
}

fn split_op(raw: &str) -> (CriteriaOp, &str) {
    let raw = raw.trim_start();
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

fn coerce_to_number(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => Some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        Value::Text(s) => s.trim().parse::<f64>().ok(),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

fn coerce_to_bool(value: &Value) -> Option<bool> {
    match value {
        Value::Bool(b) => Some(*b),
        Value::Number(n) => Some(*n != 0.0),
        Value::Text(s) => {
            if s.eq_ignore_ascii_case("TRUE") {
                Some(true)
            } else if s.eq_ignore_ascii_case("FALSE") {
                Some(false)
            } else {
                None
            }
        }
        Value::Blank => Some(false),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

fn coerce_to_text(value: &Value) -> String {
    match value {
        Value::Blank => String::new(),
        Value::Number(n) => n.to_string(),
        Value::Text(s) => s.clone(),
        Value::Bool(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        Value::Error(e) => e.to_string(),
        Value::Reference(_) | Value::ReferenceUnion(_) => ErrorKind::Value.to_string(),
        Value::Array(arr) => arr.top_left().to_string(),
        Value::Lambda(_) => "<LAMBDA>".to_string(),
        Value::Spill { .. } => ErrorKind::Spill.to_string(),
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

// Wildcard matching is shared with lookup functions (XLOOKUP/XMATCH).
