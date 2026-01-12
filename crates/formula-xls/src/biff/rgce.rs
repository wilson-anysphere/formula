//! BIFF8 `rgce` (formula token stream) decode helpers.
//!
//! This module is currently used for decoding the `rgce` stream stored in BIFF8 `NAME` records
//! (defined names / named ranges).
//!
//! The decoder is intentionally best-effort: it aims to produce readable, parseable Excel formula
//! text for common patterns (refs, operators, and function calls) while preserving warnings for
//! unsupported/unknown constructs.

#![allow(dead_code)]

use std::borrow::Cow;

use crate::biff::strings;

/// Result of best-effort BIFF8 `rgce` decoding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DecodedRgce {
    /// Best-effort decoded Excel formula text (without the leading `=`).
    pub(crate) text: Option<String>,
    /// Any non-fatal decode warnings.
    pub(crate) warnings: Vec<String>,
}

/// Decode a BIFF8 `rgce` token stream used in a defined-name (`NAME`) record.
///
/// The returned text does **not** include a leading `=`.
pub(crate) fn decode_defined_name_rgce(rgce: &[u8], codepage: u16) -> DecodedRgce {
    let mut warnings = Vec::new();
    let text = decode_defined_name_rgce_impl(rgce, codepage, &mut warnings).ok();
    DecodedRgce { text, warnings }
}

#[derive(Clone, Debug)]
struct ExprFragment {
    text: String,
    precedence: u8,
    /// `true` if this fragment contains the union operator (`,`).
    contains_union: bool,
    is_missing: bool,
}

impl ExprFragment {
    fn new(text: String) -> Self {
        Self {
            text,
            precedence: 100,
            contains_union: false,
            is_missing: false,
        }
    }

    fn missing() -> Self {
        Self {
            text: String::new(),
            precedence: 100,
            contains_union: false,
            is_missing: true,
        }
    }
}

// Precedence values match `formula-engine` (`Expr::precedence`).
fn binary_precedence(ptg: u8) -> Option<u8> {
    match ptg {
        0x11 => Some(82),                                    // range (:)
        0x0F => Some(81),                                    // intersect ( )
        0x10 => Some(80),                                    // union (,)
        0x07 => Some(50),                                    // power (^)
        0x05 | 0x06 => Some(40),                             // mul/div
        0x03 | 0x04 => Some(30),                             // add/sub
        0x08 => Some(20),                                    // concat (&)
        0x09 | 0x0A | 0x0B | 0x0C | 0x0D | 0x0E => Some(10), // comparisons
        _ => None,
    }
}

fn op_str(ptg: u8) -> Option<&'static str> {
    match ptg {
        0x03 => Some("+"),
        0x04 => Some("-"),
        0x05 => Some("*"),
        0x06 => Some("/"),
        0x07 => Some("^"),
        0x08 => Some("&"),
        0x09 => Some("<"),
        0x0A => Some("<="),
        0x0B => Some("="),
        0x0C => Some(">"),
        0x0D => Some(">="),
        0x0E => Some("<>"),
        0x0F => Some(" "),
        0x10 => Some(","),
        0x11 => Some(":"),
        _ => None,
    }
}

fn format_function_call(name: &str, args: Vec<ExprFragment>) -> ExprFragment {
    let contains_union = args.iter().any(|a| a.contains_union);

    let mut text = String::new();
    text.push_str(name);
    text.push('(');
    for (i, arg) in args.into_iter().enumerate() {
        if i > 0 {
            text.push(',');
        }
        if arg.is_missing {
            continue;
        }
        // The union operator uses `,`, which is also the function argument separator. To keep the
        // decoded formula parseable by our formula parser, wrap any argument containing union in
        // parentheses (Excel's canonical form, e.g. `SUM((A1,B1))`).
        if arg.contains_union {
            // Avoid double-parenthesizing an explicit `PtgParen` arg.
            if arg.precedence == 100 && arg.text.starts_with('(') && arg.text.ends_with(')') {
                text.push_str(&arg.text);
            } else {
                text.push('(');
                text.push_str(&arg.text);
                text.push(')');
            }
        } else {
            text.push_str(&arg.text);
        }
    }
    text.push(')');

    ExprFragment {
        text,
        precedence: 100,
        contains_union,
        is_missing: false,
    }
}

fn escape_excel_string(value: &str) -> String {
    // Excel escapes `"` inside a string literal by doubling it.
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        if ch == '"' {
            out.push('"');
            out.push('"');
        } else {
            out.push(ch);
        }
    }
    out
}

fn decode_defined_name_rgce_impl(
    rgce: &[u8],
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Result<String, String> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    // RPN stack.
    let mut stack: Vec<ExprFragment> = Vec::new();

    let mut input = rgce;
    while !input.is_empty() {
        let ptg = input[0];
        input = &input[1..];

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    return Err(format!("unsupported ptg token 0x{ptg:02X}"));
                };
                let prec = binary_precedence(ptg).expect("precedence for binary ops");

                let right = stack.pop().ok_or_else(|| "rgce stack underflow".to_string())?;
                let left = stack.pop().ok_or_else(|| "rgce stack underflow".to_string())?;

                let left_s = if left.precedence < prec && !left.is_missing {
                    format!("({})", left.text)
                } else {
                    left.text
                };
                let right_s = if right.precedence < prec && !right.is_missing {
                    format!("({})", right.text)
                } else {
                    right.text
                };

                let mut text = String::with_capacity(left_s.len() + op.len() + right_s.len());
                text.push_str(&left_s);
                text.push_str(op);
                text.push_str(&right_s);

                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union: left.contains_union || right.contains_union || ptg == 0x10,
                    is_missing: false,
                });
            }
            // Unary plus/minus.
            0x12 | 0x13 => {
                let prec = 70;
                let op = if ptg == 0x12 { "+" } else { "-" };
                let expr = stack.pop().ok_or_else(|| "rgce stack underflow".to_string())?;
                let inner = if expr.precedence < prec && !expr.is_missing {
                    format!("({})", expr.text)
                } else {
                    expr.text
                };
                stack.push(ExprFragment {
                    text: format!("{op}{inner}"),
                    precedence: prec,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Percent postfix.
            0x14 => {
                let prec = 60;
                let expr = stack.pop().ok_or_else(|| "rgce stack underflow".to_string())?;
                let inner = if expr.precedence < prec && !expr.is_missing {
                    format!("({})", expr.text)
                } else {
                    expr.text
                };
                stack.push(ExprFragment {
                    text: format!("{inner}%"),
                    precedence: prec,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Explicit parentheses.
            0x15 => {
                let expr = stack.pop().ok_or_else(|| "rgce stack underflow".to_string())?;
                stack.push(ExprFragment {
                    text: format!("({})", expr.text),
                    precedence: 100,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Missing argument.
            0x16 => stack.push(ExprFragment::missing()),
            // String literal (ShortXLUnicodeString).
            0x17 => {
                let (s, consumed) =
                    strings::parse_biff8_short_string(input, codepage).map_err(|e| e)?;
                input = input.get(consumed..).ok_or_else(|| "unexpected eof".to_string())?;
                stack.push(ExprFragment::new(format!(
                    "\"{}\"",
                    escape_excel_string(&s)
                )));
            }
            // Error literal.
            0x1C => {
                let (&err, rest) = input.split_first().ok_or_else(|| "unexpected eof".to_string())?;
                input = rest;
                let text = match err {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => "#UNKNOWN!",
                };
                if text == "#UNKNOWN!" {
                    warnings.push(format!(
                        "unknown error literal 0x{err:02X} in BIFF8 rgce stream"
                    ));
                }
                stack.push(ExprFragment::new(text.to_string()));
            }
            // Bool literal.
            0x1D => {
                let (&b, rest) = input.split_first().ok_or_else(|| "unexpected eof".to_string())?;
                input = rest;
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // Int literal.
            0x1E => {
                if input.len() < 2 {
                    return Err("unexpected eof".to_string());
                }
                let n = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];
                stack.push(ExprFragment::new(n.to_string()));
            }
            // Num literal.
            0x1F => {
                if input.len() < 8 {
                    return Err("unexpected eof".to_string());
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&input[..8]);
                input = &input[8..];
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgFunc: [iftab: u16] (fixed arg count is implicit).
            0x21 | 0x41 | 0x61 => {
                if input.len() < 2 {
                    return Err("unexpected eof".to_string());
                }
                let func_id = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];

                let (name, argc): (Cow<'static, str>, usize) =
                    match formula_biff::function_spec_from_id(func_id) {
                        Some(spec) if spec.min_args == spec.max_args => {
                            (Cow::Borrowed(spec.name), spec.min_args as usize)
                        }
                        _ => {
                            warnings.push(format!(
                                "unknown BIFF function id 0x{func_id:04X} (PtgFunc) in defined name formula"
                            ));
                            (Cow::Owned(format!("#UNKNOWN_FUNC(0x{func_id:04X})")), 0)
                        }
                    };

                if stack.len() < argc {
                    return Err("rgce stack underflow".to_string());
                }
                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().expect("len checked"));
                }
                args.reverse();
                stack.push(format_function_call(name.as_ref(), args));
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if input.len() < 3 {
                    return Err("unexpected eof".to_string());
                }
                let argc = input[0] as usize;
                let func_id = u16::from_le_bytes([input[1], input[2]]);
                input = &input[3..];

                if stack.len() < argc {
                    return Err("rgce stack underflow".to_string());
                }

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name, followed by args.
                if func_id == 0x00FF {
                    if argc == 0 {
                        return Err("rgce stack underflow".to_string());
                    }
                    let func_name = stack.pop().expect("len checked").text;
                    let mut args = Vec::with_capacity(argc.saturating_sub(1));
                    for _ in 0..argc.saturating_sub(1) {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format_function_call(&func_name, args));
                } else {
                    let name_owned;
                    let name = match formula_biff::function_id_to_name(func_id) {
                        Some(name) => name,
                        None => {
                            warnings.push(format!(
                                "unknown BIFF function id 0x{func_id:04X} (PtgFuncVar) in defined name formula"
                            ));
                            name_owned = format!("#UNKNOWN_FUNC(0x{func_id:04X})");
                            &name_owned
                        }
                    };

                    let mut args = Vec::with_capacity(argc);
                    for _ in 0..argc {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format_function_call(name, args));
                }
            }
            // PtgRef: [rw: u16][col: u16]
            0x24 | 0x44 | 0x64 => {
                if input.len() < 4 {
                    return Err("unexpected eof".to_string());
                }
                let row0 = u16::from_le_bytes([input[0], input[1]]);
                let col_field = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];

                stack.push(ExprFragment::new(format_cell_ref_from_field(
                    row0, col_field,
                )));
            }
            // PtgArea: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x25 | 0x45 | 0x65 => {
                if input.len() < 8 {
                    return Err("unexpected eof".to_string());
                }
                let row1 = u16::from_le_bytes([input[0], input[1]]);
                let row2 = u16::from_le_bytes([input[2], input[3]]);
                let col1 = u16::from_le_bytes([input[4], input[5]]);
                let col2 = u16::from_le_bytes([input[6], input[7]]);
                input = &input[8..];

                let start = format_cell_ref_from_field(row1, col1);
                let end = format_cell_ref_from_field(row2, col2);

                let is_single_cell =
                    row1 == row2 && (col1 & 0x3FFF) == (col2 & 0x3FFF) && (col1 & 0xC000) == (col2 & 0xC000);
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                if is_value_class && !is_single_cell {
                    // Preserve legacy implicit intersection semantics for value-class ranges.
                    text.push('@');
                }
                if is_single_cell {
                    text.push_str(&start);
                } else {
                    text.push_str(&start);
                    text.push(':');
                    text.push_str(&end);
                }

                let mut frag = ExprFragment::new(text);
                if is_value_class && !is_single_cell {
                    // Unary `@` has the same precedence as unary +/- in our formula parser.
                    frag.precedence = 70;
                }
                stack.push(frag);
            }
            _ => return Err(format!("unsupported ptg token 0x{ptg:02X}")),
        }
    }

    if stack.len() == 1 {
        Ok(stack.pop().expect("len checked").text)
    } else {
        Err(format!(
            "rgce decode finished with stack size {} (expected 1)",
            stack.len()
        ))
    }
}

fn push_column_label(mut col: u32, out: &mut String) {
    // Excel columns are 1-based in the alphabetic representation.
    col = col.saturating_add(1);
    let mut buf = [0u8; 8];
    let mut i = 0usize;
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        buf[i] = b'A' + rem;
        i += 1;
        col = (col - 1) / 26;
    }
    for ch in buf[..i].iter().rev() {
        out.push(*ch as char);
    }
}

fn format_cell_ref_from_field(row0: u16, col_field: u16) -> String {
    let row1 = (row0 as u32).saturating_add(1);
    let col = (col_field & 0x3FFF) as u32;
    let row_relative = (col_field & 0x4000) == 0x4000;
    let col_relative = (col_field & 0x8000) == 0x8000;

    let mut out = String::new();
    if !col_relative {
        out.push('$');
    }
    push_column_label(col, &mut out);
    if !row_relative {
        out.push('$');
    }
    out.push_str(&row1.to_string());
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decodes_sum_1_2_from_ptg_funcvar() {
        // SUM(1,2):
        //   PtgInt 1
        //   PtgInt 2
        //   PtgFuncVar argc=2 iftab=4 (SUM)
        let rgce = vec![0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x22, 0x02, 0x04, 0x00];

        let decoded = decode_defined_name_rgce(&rgce, 1252);
        assert_eq!(decoded.text.as_deref(), Some("SUM(1,2)"));
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
    }

    #[test]
    fn decodes_if_true_1_2_from_ptg_funcvar() {
        // IF(TRUE,1,2):
        //   PtgBool TRUE
        //   PtgInt 1
        //   PtgInt 2
        //   PtgFuncVar argc=3 iftab=1 (IF)
        let rgce = vec![
            0x1D, 0x01, // TRUE
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x22, 0x03, 0x01, 0x00, // IF(argc=3)
        ];

        let decoded = decode_defined_name_rgce(&rgce, 1252);
        assert_eq!(decoded.text.as_deref(), Some("IF(TRUE,1,2)"));
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
    }
}
