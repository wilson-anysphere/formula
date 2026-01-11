use crate::{ErrorKind, Value};

/// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
pub fn textjoin(
    delimiter: &str,
    ignore_empty: bool,
    values: &[Value],
) -> Result<String, ErrorKind> {
    let mut out = String::new();
    let mut first = true;

    for value in values {
        let piece = match value {
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
            Value::Error(e) => return Err(*e),
            Value::Reference(_) | Value::ReferenceUnion(_) => return Err(ErrorKind::Value),
            Value::Array(arr) => arr.top_left().to_string(),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            Value::Spill { .. } => return Err(ErrorKind::Spill),
        };

        if ignore_empty && piece.is_empty() {
            continue;
        }

        if !first {
            out.push_str(delimiter);
        }
        first = false;
        out.push_str(&piece);
    }

    Ok(out)
}
