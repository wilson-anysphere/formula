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

// BIFF8 supports 65,536 rows (0-based 0..=65,535).
const BIFF8_MAX_ROW0: i64 = u16::MAX as i64;
// Columns are stored in a 14-bit field in many BIFF8 structures.
const BIFF8_MAX_COL0: i64 = 0x3FFF;

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
            // PtgNameX (external name reference).
            //
            // [MS-XLS] 2.5.198.41
            // Payload: [ixti: u16][iname: u16][reserved: u16]
            //
            // We do not currently parse the supporting `SUPBOOK`/`EXTERNNAME` tables needed to
            // resolve the external name, so decode best-effort as `#REF!` (but still consume the
            // payload so we can keep decoding the remainder of the rgce stream).
            0x39 | 0x59 | 0x79 => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let iname = u16::from_le_bytes([input[2], input[3]]);
                input = &input[6..];
                warnings.push(format!(
                    "PtgNameX external name reference is not supported (ixti={ixti}, iname={iname})"
                ));
                stack.push(ExprFragment::new("#REF!".to_string()));
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
            // PtgRefErr (2D): consume payload and emit `#REF!`.
            0x2A | 0x4A | 0x6A => {
                if input.len() < 4 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[4..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr (2D): consume payload and emit `#REF!`.
            0x2B | 0x4B | 0x6B => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[8..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN (relative reference): [row_off: i16][col_off: i16]
            0x2C | 0x4C | 0x6C => {
                if input.len() < 4 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row_off = i16::from_le_bytes([input[0], input[1]]) as i64;
                let col_off = i16::from_le_bytes([input[2], input[3]]) as i64;
                input = &input[4..];

                // Defined-name formulas do not have a stable origin cell; decode relative refs
                // best-effort relative to (0,0) / A1.
                let abs_row0 = row_off;
                let abs_col0 = col_off;
                if abs_row0 < 0
                    || abs_row0 > BIFF8_MAX_ROW0
                    || abs_col0 < 0
                    || abs_col0 > BIFF8_MAX_COL0
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    stack.push(ExprFragment::new(format_cell_ref_no_dollars(
                        abs_row0 as u32,
                        abs_col0 as u32,
                    )));
                }
            }
            // PtgAreaN (relative area): [rowFirst_off: i16][rowLast_off: i16][colFirst_off: i16][colLast_off: i16]
            0x2D | 0x4D | 0x6D => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row_first_off = i16::from_le_bytes([input[0], input[1]]) as i64;
                let row_last_off = i16::from_le_bytes([input[2], input[3]]) as i64;
                let col_first_off = i16::from_le_bytes([input[4], input[5]]) as i64;
                let col_last_off = i16::from_le_bytes([input[6], input[7]]) as i64;
                input = &input[8..];

                let abs_row_first = row_first_off;
                let abs_row_last = row_last_off;
                let abs_col_first = col_first_off;
                let abs_col_last = col_last_off;

                if abs_row_first < 0
                    || abs_row_first > BIFF8_MAX_ROW0
                    || abs_row_last < 0
                    || abs_row_last > BIFF8_MAX_ROW0
                    || abs_col_first < 0
                    || abs_col_first > BIFF8_MAX_COL0
                    || abs_col_last < 0
                    || abs_col_last > BIFF8_MAX_COL0
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let start = format_cell_ref_no_dollars(abs_row_first as u32, abs_col_first as u32);
                    let end = format_cell_ref_no_dollars(abs_row_last as u32, abs_col_last as u32);
                    if start == end {
                        stack.push(ExprFragment::new(start));
                    } else {
                        stack.push(ExprFragment::new(format!("{start}:{end}")));
                    }
                }
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
            // PtgRefErr3d: consume payload and emit `#REF!`.
            //
            // Payload matches `PtgRef3d`: [ixti: u16][row: u16][col: u16]
            0x3C | 0x5C | 0x7C => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[6..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr3d: consume payload and emit `#REF!`.
            //
            // Payload matches `PtgArea3d`: [ixti: u16][row1: u16][row2: u16][col1: u16][col2: u16]
            0x3D | 0x5D | 0x7D => {
                if input.len() < 10 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[10..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN3d (relative 3D reference): [ixti: u16][row_off: i16][col_off: i16]
            0x3E | 0x5E | 0x7E => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row_off = i16::from_le_bytes([input[2], input[3]]) as i64;
                let col_off = i16::from_le_bytes([input[4], input[5]]) as i64;
                input = &input[6..];

                // Defined-name formulas do not have a stable origin cell; decode relative refs
                // best-effort relative to (0,0) / A1.
                let abs_row0 = row_off;
                let abs_col0 = col_off;
                if abs_row0 < 0
                    || abs_row0 > BIFF8_MAX_ROW0
                    || abs_col0 < 0
                    || abs_col0 > BIFF8_MAX_COL0
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                        Ok(v) => v,
                        Err(err) => {
                            warnings.push(err);
                            format!("#SHEET(ixti={ixti})!")
                        }
                    };
                    let cell = format_cell_ref_no_dollars(abs_row0 as u32, abs_col0 as u32);
                    stack.push(ExprFragment::new(format!("{sheet_prefix}{cell}")));
                }
            }
            // PtgAreaN3d (relative 3D area): [ixti: u16][rowFirst_off: i16][rowLast_off: i16][colFirst_off: i16][colLast_off: i16]
            0x3F | 0x5F | 0x7F => {
                if input.len() < 10 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row_first_off = i16::from_le_bytes([input[2], input[3]]) as i64;
                let row_last_off = i16::from_le_bytes([input[4], input[5]]) as i64;
                let col_first_off = i16::from_le_bytes([input[6], input[7]]) as i64;
                let col_last_off = i16::from_le_bytes([input[8], input[9]]) as i64;
                input = &input[10..];

                let abs_row_first = row_first_off;
                let abs_row_last = row_last_off;
                let abs_col_first = col_first_off;
                let abs_col_last = col_last_off;

                if abs_row_first < 0
                    || abs_row_first > BIFF8_MAX_ROW0
                    || abs_row_last < 0
                    || abs_row_last > BIFF8_MAX_ROW0
                    || abs_col_first < 0
                    || abs_col_first > BIFF8_MAX_COL0
                    || abs_col_last < 0
                    || abs_col_last > BIFF8_MAX_COL0
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                        Ok(v) => v,
                        Err(err) => {
                            warnings.push(err);
                            format!("#SHEET(ixti={ixti})!")
                        }
                    };

                    let start =
                        format_cell_ref_no_dollars(abs_row_first as u32, abs_col_first as u32);
                    let end = format_cell_ref_no_dollars(abs_row_last as u32, abs_col_last as u32);
                    let area = if start == end {
                        start
                    } else {
                        format!("{start}:{end}")
                    };
                    stack.push(ExprFragment::new(format!("{sheet_prefix}{area}")));
                }
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

fn format_cell_ref_no_dollars(row0: u32, col0: u32) -> String {
    let mut out = String::new();
    push_column(col0, &mut out);
    out.push_str(&(row0.saturating_add(1)).to_string());
    out
}

fn format_row_ref(row: u16, row_rel: bool) -> String {
    let mut out = String::new();
    if !row_rel {
        out.push('$');
    }
    out.push_str(&(row as u32 + 1).to_string());
    out
}

fn format_col_ref(col: u16, col_rel: bool) -> String {
    let mut out = String::new();
    if !col_rel {
        out.push('$');
    }
    push_column(col as u32, &mut out);
    out
}

fn format_area_ref(row1: u16, col1: u16, row2: u16, col2: u16) -> String {
    const BIFF8_MAX_ROW: u16 = 0xFFFF;
    const BIFF8_MAX_COL: u16 = 0x00FF;

    let col1_idx = col1 & 0x3FFF;
    let col2_idx = col2 & 0x3FFF;

    let row1_rel = (col1 & 0x4000) != 0;
    let col1_rel = (col1 & 0x8000) != 0;
    let row2_rel = (col2 & 0x4000) != 0;
    let col2_rel = (col2 & 0x8000) != 0;

    // Whole-column references (`$A:$A`, `$A:$C`).
    if row1 == 0 && row2 == BIFF8_MAX_ROW && col1_idx <= BIFF8_MAX_COL && col2_idx <= BIFF8_MAX_COL
    {
        let start = format_col_ref(col1_idx as u16, col1_rel);
        let end = format_col_ref(col2_idx as u16, col2_rel);
        // Excel canonical form includes the `:` even for single-column ranges.
        return format!("{start}:{end}");
    }

    // Whole-row references (`$1:$1`, `$1:$3`).
    if col1_idx == 0 && col2_idx == BIFF8_MAX_COL {
        let start = format_row_ref(row1, row1_rel);
        let end = format_row_ref(row2, row2_rel);
        // Excel canonical form includes the `:` even for single-row ranges.
        return format!("{start}:{end}");
    }

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
            "{}!",
            quote_sheet_range_name_if_needed(first, last)
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

fn quote_sheet_range_name_if_needed(start: &str, end: &str) -> String {
    // Excel represents 3D sheet ranges as:
    // - `Sheet1:Sheet3!A1` for simple sheet identifiers
    // - `'Sheet 1:Sheet3'!A1` when either side requires quoting.
    //
    // Note: quoting each side independently (`'Sheet 1':Sheet3!A1`) is not a valid 3D sheet range.
    if is_unquoted_sheet_name(start) && is_unquoted_sheet_name(end) {
        return format!("{start}:{end}");
    }

    let mut out = String::new();
    out.push('\'');
    out.push_str(&start.replace('\'', "''"));
    out.push(':');
    out.push_str(&end.replace('\'', "''"));
    out.push('\'');
    out
}

fn is_unquoted_sheet_name(name: &str) -> bool {
    // Be conservative: quoting is always accepted by Excel, but the unquoted form is only valid
    // for a subset of identifier-like sheet names. In particular, `TRUE` / `FALSE` and A1-style
    // cell references must be quoted to avoid being tokenized as boolean/cell literals by our
    // `formula-engine` lexer.
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return false;
    }
    if starts_like_a1_cell_ref(name) {
        return false;
    }

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

fn starts_like_a1_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut col_letters = String::new();
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_alphabetic() {
            col_letters.push(ch);
            chars.next();
        } else {
            break;
        }
    }
    if col_letters.is_empty() {
        return false;
    }

    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut row_digits = String::new();
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_digit() {
            row_digits.push(ch);
            chars.next();
        } else {
            break;
        }
    }
    if row_digits.is_empty() {
        return false;
    }

    if col_from_a1(&col_letters).is_none() {
        return false;
    }
    matches!(row_digits.parse::<u32>(), Ok(v) if v != 0)
}

fn col_from_a1(col: &str) -> Option<u32> {
    let mut value: u32 = 0;
    let mut count: usize = 0;
    for ch in col.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        count += 1;
        if count > 3 {
            return None;
        }
        value = value * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    if value == 0 || value > 16_384 {
        return None;
    }
    Some(value - 1)
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

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::{parse_formula, ParseOptions};

    fn assert_parseable(expr: &str) {
        parse_formula(&format!("={expr}"), ParseOptions::default()).expect("parse formula");
    }

    const BIFF8_MAX_ROW: u16 = 0xFFFF;
    const BIFF8_MAX_COL: u16 = 0x00FF;

    fn empty_ctx<'a>(
        sheet_names: &'a [String],
        externsheet: &'a [ExternSheetRef],
        defined_names: &'a [DefinedNameMeta],
    ) -> RgceDecodeContext<'a> {
        RgceDecodeContext {
            codepage: 1252,
            sheet_names,
            externsheet,
            defined_names,
        }
    }

    #[test]
    fn decodes_ptg_refn_to_a1() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // row_off=0 col_off=0 => A1 (base A1).
        let rgce = [0x2C, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_refn_value_class_variant() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // row_off=1 col_off=1 => B2.
        let rgce = [0x4C, 0x01, 0x00, 0x01, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_refn_oob_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // row_off=-1 => out-of-bounds.
        let rgce = [0x2C, 0xFF, 0xFF, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_refn_col_oob_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // col_off=20000 -> out of Excel column bounds (16,384 cols).
        let col_off: i16 = 20_000;
        let [c0, c1] = col_off.to_le_bytes();
        let rgce = [0x2C, 0x00, 0x00, c0, c1];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_arean_to_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // A1:B2.
        let rgce = [
            0x2D, // PtgAreaN
            0x00, 0x00, // rowFirst_off = 0
            0x01, 0x00, // rowLast_off = 1
            0x00, 0x00, // colFirst_off = 0
            0x01, 0x00, // colLast_off = 1
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_arean_array_class_variant() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // C3:D4.
        let rgce = [
            0x6D, // PtgAreaNA
            0x02, 0x00, // rowFirst_off = 2
            0x03, 0x00, // rowLast_off = 3
            0x02, 0x00, // colFirst_off = 2
            0x03, 0x00, // colLast_off = 3
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "C3:D4");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_referr_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (4 bytes).
        let rgce = [0x2A, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_areaerr_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (8 bytes).
        let rgce = [0x2B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_referr3d_to_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (6 bytes): ixti=0 row=0 col=0.
        let rgce = [0x3C, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_areaerr3d_to_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (10 bytes): ixti=0 + 8 bytes area.
        let rgce = [
            0x3D, 0x00, 0x00, // ixti
            0x00, 0x00, // row1
            0x00, 0x00, // row2
            0x00, 0x00, // col1
            0x00, 0x00, // col2
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_row_area_as_row_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Whole row 1: row1==row2, col spanning 0..255 (A..IV).
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A, absolute)
        rgce.extend_from_slice(&BIFF8_MAX_COL.to_le_bytes()); // colLast (IV, absolute)

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "$1:$1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_column_area_as_col_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Whole column A: col1==col2, row spanning 0..65535 (1..65536).
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&BIFF8_MAX_ROW.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A, absolute)
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colLast (A, absolute)

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "$A:$A");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn continues_to_render_rectangular_ranges_as_a1_areas() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // $A$1:$B$2
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&1u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst (A)
        rgce.extend_from_slice(&1u16.to_le_bytes()); // colLast (B)

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "$A$1:$B$2");
    }

    #[test]
    fn decodes_print_titles_union_as_row_and_col_ranges() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // `$1:$1,$A:$A` (union of whole-row and whole-column).
        let mut rgce = Vec::new();
        // $1:$1
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes());
        rgce.extend_from_slice(&0u16.to_le_bytes());
        rgce.extend_from_slice(&0u16.to_le_bytes());
        rgce.extend_from_slice(&BIFF8_MAX_COL.to_le_bytes());
        // $A:$A
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes());
        rgce.extend_from_slice(&BIFF8_MAX_ROW.to_le_bytes());
        rgce.extend_from_slice(&0u16.to_le_bytes());
        rgce.extend_from_slice(&0u16.to_le_bytes());
        // Union
        rgce.push(0x10); // PtgUnion

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "$1:$1,$A:$A");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_row_area3d_as_row_range() {
        // Same whole-row area but stored as PtgArea3d with an EXTERNSHEET sheet prefix.
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let mut rgce = Vec::new();
        rgce.push(0x3B); // PtgArea3d
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&BIFF8_MAX_COL.to_le_bytes()); // colLast

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!$1:$1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_column_area3d_as_col_range() {
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let mut rgce = Vec::new();
        rgce.push(0x3B); // PtgArea3d
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&BIFF8_MAX_ROW.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colLast

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!$A:$A");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_print_titles_union_area3d_as_row_and_col_ranges() {
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // `Sheet1!$1:$1,Sheet1!$A:$A` in RPN form: [area3d row][area3d col][union].
        let mut rgce = Vec::new();
        // $1:$1
        rgce.push(0x3B); // PtgArea3d
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&BIFF8_MAX_COL.to_le_bytes()); // colLast
        // $A:$A
        rgce.push(0x3B); // PtgArea3d
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&BIFF8_MAX_ROW.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colLast
        // Union
        rgce.push(0x10); // PtgUnion

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!$1:$1,Sheet1!$A:$A");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_arean3d_to_sheet_range() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Sheet1!A1:B2.
        let rgce = [
            0x3F, 0x00, 0x00, // ixti
            0x00, 0x00, // rowFirst_off = 0
            0x01, 0x00, // rowLast_off = 1
            0x00, 0x00, // colFirst_off = 0
            0x01, 0x00, // colLast_off = 1
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_refn3d_to_sheet_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // ixti=0, row_off=0, col_off=0 => Sheet1!A1.
        let rgce = [0x3E, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!A1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_to_ref_placeholder() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (6 bytes): ixti=0, iname=1, reserved=0.
        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("PtgNameX external name reference is not supported")),
            "warnings={:?}",
            decoded.warnings
        );
    }

    #[test]
    fn decodes_sum_1_2_from_ptg_funcvar() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM(1,2):
        //   PtgInt 1
        //   PtgInt 2
        //   PtgFuncVar argc=2 iftab=4 (SUM)
        let rgce = vec![0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x22, 0x02, 0x04, 0x00];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM(1,2)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_if_true_1_2_from_ptg_funcvar() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

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

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "IF(TRUE,1,2)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_abs_neg1_from_ptg_func() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // ABS(-1):
        //   PtgInt 1
        //   PtgUminus
        //   PtgFunc iftab=24 (ABS)
        let rgce = vec![
            0x1E, 0x01, 0x00, // 1
            0x13, // unary minus
            0x21, 0x18, 0x00, // ABS (iftab 24)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "ABS(-1)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_sum_sheet1_a1_2_from_ptg_ref3d_and_funcvar() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM(Sheet1!$A$1,2):
        //   PtgRef3d ixti=0 row=0 col=0 (absolute A1)
        //   PtgInt 2
        //   PtgFuncVar argc=2 iftab=4 (SUM)
        let rgce = vec![
            0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // Sheet1!$A$1 via ixti=0
            0x1E, 0x02, 0x00, // 2
            0x22, 0x02, 0x04, 0x00, // SUM(argc=2)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM(Sheet1!$A$1,2)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_quotes_sheet_names_with_spaces() {
        let sheet_names: Vec<String> = vec!["My Sheet".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgRef3d: [ptg][ixti: u16][rw: u16][col: u16] => sheet-qualified $A$1.
        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'My Sheet'!$A$1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_escapes_apostrophes_in_sheet_names() {
        let sheet_names: Vec<String> = vec!["O'Brien".to_string()];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'O''Brien'!$A$1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_renders_sheet_ranges_as_single_quoted_ident() {
        let sheet_names: Vec<String> = vec![
            "Sheet 1".to_string(),
            "Sheet 2".to_string(),
            "Sheet 3".to_string(),
        ];
        let externsheet: Vec<ExternSheetRef> = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 2,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'Sheet 1:Sheet 3'!$A$1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_area3d_sheet_range_with_quoting() {
        let sheet_names = vec!["Sheet 1".to_string(), "Sheet3".to_string()];
        let externsheet = vec![ExternSheetRef {
            itab_first: 0,
            itab_last: 1,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgArea3d (0x3B): [ixti=0][A1:B2] on sheet range Sheet 1:Sheet3.
        let mut rgce = Vec::new();
        rgce.push(0x3B);
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rowFirst0
        rgce.extend_from_slice(&1u16.to_le_bytes()); // rowLast0
        rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst=A relative
        rgce.extend_from_slice(&0xC001u16.to_le_bytes()); // colLast=B relative

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'Sheet 1:Sheet3'!A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptgname_workbook_and_sheet_scope() {
        let sheet_names = vec!["Sheet 1".to_string()];
        let externsheet: Vec<ExternSheetRef> = Vec::new();
        let defined_names = vec![
            DefinedNameMeta {
                name: "MyName".to_string(),
                scope_sheet: None,
            },
            DefinedNameMeta {
                name: "LocalName".to_string(),
                scope_sheet: Some(0),
            },
        ];
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Workbook-scoped name (id=1).
        let rgce1 = [0x23, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded1 = decode_biff8_rgce(&rgce1, &ctx);
        assert_eq!(decoded1.text, "MyName");
        assert!(decoded1.warnings.is_empty(), "warnings={:?}", decoded1.warnings);
        assert_parseable(&decoded1.text);

        // Sheet-scoped name (id=2).
        let rgce2 = [0x23, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded2 = decode_biff8_rgce(&rgce2, &ctx);
        assert_eq!(decoded2.text, "'Sheet 1'!LocalName");
        assert!(decoded2.warnings.is_empty(), "warnings={:?}", decoded2.warnings);
        assert_parseable(&decoded2.text);
    }
}
