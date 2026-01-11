use crate::{ErrorKind, Value};

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Token {
    Star,
    QMark,
    Literal(char),
}

pub(crate) fn wildcard_match(pattern: &str, text: &str) -> bool {
    let pattern = pattern.to_ascii_uppercase();
    let text = text.to_ascii_uppercase();
    let tokens = tokenize_pattern(&pattern);
    wildcard_match_tokens(&tokens, &text.chars().collect::<Vec<_>>())
}

fn tokenize_pattern(pattern: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '~' => {
                if let Some(next) = chars.next() {
                    tokens.push(Token::Literal(next));
                } else {
                    tokens.push(Token::Literal('~'));
                }
            }
            '*' => tokens.push(Token::Star),
            '?' => tokens.push(Token::QMark),
            other => tokens.push(Token::Literal(other)),
        }
    }
    tokens
}

fn wildcard_match_tokens(pattern: &[Token], text: &[char]) -> bool {
    let mut pi = 0usize;
    let mut ti = 0usize;
    let mut star: Option<usize> = None;
    let mut star_text = 0usize;

    while ti < text.len() {
        if pi < pattern.len() {
            match pattern[pi] {
                Token::Literal(c) if c == text[ti] => {
                    pi += 1;
                    ti += 1;
                    continue;
                }
                Token::QMark => {
                    pi += 1;
                    ti += 1;
                    continue;
                }
                Token::Star => {
                    star = Some(pi);
                    pi += 1;
                    star_text = ti;
                    continue;
                }
                _ => {}
            }
        }

        if let Some(star_pos) = star {
            pi = star_pos + 1;
            star_text += 1;
            ti = star_text;
        } else {
            return false;
        }
    }

    while pi < pattern.len() && pattern[pi] == Token::Star {
        pi += 1;
    }

    pi == pattern.len()
}
