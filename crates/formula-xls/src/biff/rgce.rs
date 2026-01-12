//! BIFF8 `rgce` (formula token stream) decoding helpers.
//!
//! This module implements a best-effort stack-based (RPN) decoder for the subset of BIFF8 tokens
//! we need to import defined names (`NAME` records):
//! - basic operators and constants
//! - 2D references (`PtgRef`, `PtgArea`)
//! - 3D references (`PtgRef3d`, `PtgArea3d`) via `EXTERNSHEET`
//! - defined-name references (`PtgName`)
//!
//! Unsupported tokens yield a placeholder string and warnings, but never hard-fail `.xls` import.

use super::strings;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ExternSheetRef {
    pub(crate) itab_first: u16,
    pub(crate) itab_last: u16,
}

#[derive(Debug, Clone)]
pub(crate) struct DefinedNameMeta {
    pub(crate) name: String,
    /// BIFF sheet index (0-based) for local names, or `None` for workbook scope.
    pub(crate) scope_sheet: Option<usize>,
}

pub(crate) struct RgceDecodeContext<'a> {
    pub(crate) codepage: u16,
    pub(crate) sheet_names: &'a [String],
    pub(crate) externsheet: &'a [ExternSheetRef],
    pub(crate) defined_names: &'a [DefinedNameMeta],
}

#[derive(Debug, Clone)]
pub(crate) struct DecodeRgceResult {
    pub(crate) text: String,
    pub(crate) warnings: Vec<String>,
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
        // decoded formula parseable, wrap any argument containing union in parentheses (Excel's
        // canonical form, e.g. `SUM((A1,B1))`).
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

pub(crate) fn decode_biff8_rgce(rgce: &[u8], ctx: &RgceDecodeContext<'_>) -> DecodeRgceResult {
    if rgce.is_empty() {
        return DecodeRgceResult {
            text: String::new(),
            warnings: Vec::new(),
        };
    }

    let mut input = rgce;
    let mut stack: Vec<ExprFragment> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    while !input.is_empty() {
        let ptg = input[0];
        input = &input[1..];

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    warnings.push(format!("unsupported rgce token 0x{ptg:02X}"));
                    return unsupported(ptg, warnings);
                };
                let prec = binary_precedence(ptg).expect("precedence for binary ops");

                let right = match stack.pop() {
                    Some(v) => v,
                    None => {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
                    }
                };
                let left = match stack.pop() {
                    Some(v) => v,
                    None => {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
                    }
                };

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
                let expr = match stack.pop() {
                    Some(v) => v,
                    None => {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
                    }
                };
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
                let expr = match stack.pop() {
                    Some(v) => v,
                    None => {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
                    }
                };
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
                let expr = match stack.pop() {
                    Some(v) => v,
                    None => {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
                    }
                };
                stack.push(ExprFragment {
                    text: format!("({})", expr.text),
                    precedence: 100,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Missing argument.
            0x16 => {
                stack.push(ExprFragment::missing());
            }
            // String literal (ShortXLUnicodeString).
            0x17 => match strings::parse_biff8_short_string(input, ctx.codepage) {
                Ok((s, consumed)) => {
                    input = input.get(consumed..).unwrap_or_default();
                    let escaped = s.replace('"', "\"\"");
                    stack.push(ExprFragment::new(format!("\"{escaped}\"")));
                }
                Err(err) => {
                    warnings.push(format!("failed to decode PtgStr: {err}"));
                    return unsupported(ptg, warnings);
                }
            },
            // PtgAttr: [grbit: u8][wAttr: u16]
            //
            // Most PtgAttr bits are evaluation hints (or formatting metadata) that do not affect
            // the printed formula text, but some do (notably tAttrSum).
            0x19 => {
                if input.len() < 3 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let grbit = input[0];
                let w_attr = u16::from_le_bytes([input[1], input[2]]);
                input = &input[3..];

                const T_ATTR_CHOOSE: u8 = 0x04;
                const T_ATTR_SUM: u8 = 0x10;

                if grbit & T_ATTR_CHOOSE != 0 {
                    // tAttrChoose is followed by a jump table of u16 offsets; consume it.
                    let needed = (w_attr as usize).saturating_mul(2);
                    if input.len() < needed {
                        warnings.push("unexpected end of rgce stream".to_string());
                        return unsupported(ptg, warnings);
                    }
                    input = &input[needed..];
                }

                if grbit & T_ATTR_SUM != 0 {
                    let expr = match stack.pop() {
                        Some(v) => v,
                        None => {
                            warnings.push("rgce stack underflow".to_string());
                            return unsupported(ptg, warnings);
                        }
                    };
                    stack.push(format_function_call("SUM", vec![expr]));
                }
            }
            // Error literal.
            0x1C => {
                let Some((&err, rest)) = input.split_first() else {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                };
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
                    other => {
                        warnings.push(format!("unknown error literal 0x{other:02X} in rgce"));
                        "#UNKNOWN!"
                    }
                };
                stack.push(ExprFragment::new(text.to_string()));
            }
            // Bool literal.
            0x1D => {
                let Some((&b, rest)) = input.split_first() else {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                };
                input = rest;
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // Int literal.
            0x1E => {
                if input.len() < 2 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let n = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];
                stack.push(ExprFragment::new(n.to_string()));
            }
            // Num literal.
            0x1F => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&input[..8]);
                input = &input[8..];
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgFunc: [iftab: u16] (fixed arg count is implicit).
            0x21 | 0x41 | 0x61 => {
                if input.len() < 2 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let func_id = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];

                let Some(spec) = formula_biff::function_spec_from_id(func_id) else {
                    warnings.push(format!(
                        "unknown BIFF function id 0x{func_id:04X} (PtgFunc) in rgce"
                    ));
                    return unsupported(ptg, warnings);
                };

                // Only handle fixed arity here.
                if spec.min_args != spec.max_args {
                    warnings.push(format!(
                        "unsupported variable-arity BIFF function id 0x{func_id:04X} (PtgFunc) in rgce"
                    ));
                    return unsupported(ptg, warnings);
                }

                let argc = spec.min_args as usize;
                if stack.len() < argc {
                    warnings.push("rgce stack underflow".to_string());
                    return unsupported(ptg, warnings);
                }
                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().expect("len checked"));
                }
                args.reverse();
                stack.push(format_function_call(spec.name, args));
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if input.len() < 3 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let argc = input[0] as usize;
                let func_id = u16::from_le_bytes([input[1], input[2]]);
                input = &input[3..];

                if stack.len() < argc {
                    warnings.push("rgce stack underflow".to_string());
                    return unsupported(ptg, warnings);
                }

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name, followed by args.
                if func_id == formula_biff::FTAB_USER_DEFINED {
                    if argc == 0 {
                        warnings.push("rgce stack underflow".to_string());
                        return unsupported(ptg, warnings);
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
                                "unknown BIFF function id 0x{func_id:04X} (PtgFuncVar) in rgce"
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
            // PtgName (defined name reference).
            0x23 | 0x43 | 0x63 => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }

                let name_id = u32::from_le_bytes([input[0], input[1], input[2], input[3]]);
                // Skip reserved bytes.
                input = &input[6..];

                let idx = name_id.saturating_sub(1) as usize;
                let Some(meta) = ctx.defined_names.get(idx) else {
                    warnings.push(format!("PtgName references missing name index {name_id}"));
                    stack.push(ExprFragment::new(format!("#NAME_ID({name_id})")));
                    continue;
                };

                let text = match meta.scope_sheet {
                    None => meta.name.clone(),
                    Some(sheet_idx) => match ctx.sheet_names.get(sheet_idx) {
                        Some(sheet_name) => {
                            let sheet = quote_sheet_name_if_needed(sheet_name);
                            format!("{sheet}!{}", meta.name)
                        }
                        None => {
                            warnings.push(format!(
                                "PtgName references sheet-scoped name `{}` with out-of-range sheet index {sheet_idx}",
                                meta.name
                            ));
                            meta.name.clone()
                        }
                    },
                };

                stack.push(ExprFragment::new(text));
            }
            // PtgRef (2D)
            0x24 | 0x44 | 0x64 => {
                if input.len() < 4 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row = u16::from_le_bytes([input[0], input[1]]);
                let col = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];
                stack.push(ExprFragment::new(format_cell_ref(row, col)));
            }
            // PtgArea (2D)
            0x25 | 0x45 | 0x65 => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row1 = u16::from_le_bytes([input[0], input[1]]);
                let row2 = u16::from_le_bytes([input[2], input[3]]);
                let col1 = u16::from_le_bytes([input[4], input[5]]);
                let col2 = u16::from_le_bytes([input[6], input[7]]);
                input = &input[8..];
                stack.push(ExprFragment::new(format_area_ref(row1, col1, row2, col2)));
            }
            // PtgMem* tokens: no-op for printing, but consume payload to keep offsets aligned.
            //
            // Payload: [cce: u16][rgce: cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if input.len() < 2 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let cce = u16::from_le_bytes([input[0], input[1]]) as usize;
                input = &input[2..];
                if input.len() < cce {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[cce..];
            }
            // PtgRef3d
            0x3A | 0x5A | 0x7A => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row = u16::from_le_bytes([input[2], input[3]]);
                let col = u16::from_le_bytes([input[4], input[5]]);
                input = &input[6..];

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        warnings.push(err);
                        format!("#SHEET(ixti={ixti})!")
                    }
                };
                let cell = format_cell_ref(row, col);
                stack.push(ExprFragment::new(format!("{sheet_prefix}{cell}")));
            }
            // PtgArea3d
            0x3B | 0x5B | 0x7B => {
                if input.len() < 10 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row1 = u16::from_le_bytes([input[2], input[3]]);
                let row2 = u16::from_le_bytes([input[4], input[5]]);
                let col1 = u16::from_le_bytes([input[6], input[7]]);
                let col2 = u16::from_le_bytes([input[8], input[9]]);
                input = &input[10..];

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        warnings.push(err);
                        format!("#SHEET(ixti={ixti})!")
                    }
                };
                let area = format_area_ref(row1, col1, row2, col2);
                stack.push(ExprFragment::new(format!("{sheet_prefix}{area}")));
            }
            other => {
                warnings.push(format!("unsupported rgce token 0x{other:02X}"));
                return unsupported(other, warnings);
            }
        }
    }

    let text = match stack.len() {
        0 => String::new(),
        1 => stack.pop().expect("len checked").text,
        _ => {
            warnings.push(format!(
                "rgce decode ended with {} expressions on stack",
                stack.len()
            ));
            stack.pop().expect("non-empty").text
        }
    };

    DecodeRgceResult { text, warnings }
}

fn unsupported(ptg: u8, warnings: Vec<String>) -> DecodeRgceResult {
    DecodeRgceResult {
        text: format!("#UNSUPPORTED_PTG(0x{ptg:02X})"),
        warnings,
    }
}

fn format_cell_ref(row: u16, col_with_flags: u16) -> String {
    let row_rel = (col_with_flags & 0x4000) != 0;
    let col_rel = (col_with_flags & 0x8000) != 0;
    let col = col_with_flags & 0x3FFF;

    let mut out = String::new();
    if !col_rel {
        out.push('$');
    }
    push_column(col as u32, &mut out);
    if !row_rel {
        out.push('$');
    }
    out.push_str(&(row as u32 + 1).to_string());
    out
}

fn format_area_ref(row1: u16, col1: u16, row2: u16, col2: u16) -> String {
    let start = format_cell_ref(row1, col1);
    let end = format_cell_ref(row2, col2);
    if start == end {
        start
    } else {
        format!("{start}:{end}")
    }
}

fn format_sheet_ref(ixti: u16, ctx: &RgceDecodeContext<'_>) -> Result<String, String> {
    let Some(entry) = ctx.externsheet.get(ixti as usize) else {
        return Err(format!(
            "PtgRef3d/PtgArea3d references missing EXTERNSHEET entry ixti={ixti}"
        ));
    };

    let itab_first = entry.itab_first as usize;
    let itab_last = entry.itab_last as usize;

    let Some(first) = ctx.sheet_names.get(itab_first) else {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} refers to out-of-range itabFirst={itab_first} (sheet count={})",
            ctx.sheet_names.len()
        ));
    };
    let Some(last) = ctx.sheet_names.get(itab_last) else {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} refers to out-of-range itabLast={itab_last} (sheet count={})",
            ctx.sheet_names.len()
        ));
    };

    if itab_first == itab_last {
        Ok(format!("{}!", quote_sheet_name_if_needed(first)))
    } else {
        Ok(format!(
            "{}:{}!",
            quote_sheet_name_if_needed(first),
            quote_sheet_name_if_needed(last)
        ))
    }
}

fn quote_sheet_name_if_needed(name: &str) -> String {
    if is_unquoted_sheet_name(name) {
        return name.to_string();
    }
    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn is_unquoted_sheet_name(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }
    for ch in chars {
        if !(ch.is_ascii_alphanumeric() || ch == '_') {
            return false;
        }
    }
    true
}

fn push_column(col: u32, out: &mut String) {
    // Excel columns are 1-based in A1 notation. We store 0-based internally.
    let mut n = col + 1;
    let mut buf = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        buf.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    buf.reverse();
    out.push_str(&String::from_utf8(buf).expect("A1 column bytes"));
}

