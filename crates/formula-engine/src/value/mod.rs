use std::fmt;

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
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
    Error(ErrorKind),
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
        }
    }

    pub fn coerce_to_string(&self) -> Result<String, ErrorKind> {
        match self {
            Value::Text(s) => Ok(s.clone()),
            Value::Number(n) => Ok(format_number_general(*n)),
            Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            Value::Blank => Ok(String::new()),
            Value::Error(e) => Err(*e),
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
        }
    }
}
