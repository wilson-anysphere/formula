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
//!
//! ## Relative references (`PtgRefN` / `PtgAreaN`)
//!
//! BIFF8 defined-name formulas may use the relative-reference ptgs `PtgRefN` / `PtgAreaN`. These
//! encode row/column *offsets* relative to a base cell (origin) rather than absolute coordinates.
//!
//! The origin cell is not always known at decode time; for workbook-scoped defined names Excel
//! evaluates relative references relative to the cell where the name is used. When no meaningful
//! base is known, the decoder defaults to `(0,0)` (A1) but still preserves `$` absolute/relative
//! markers in the rendered A1 text.

use super::{
    externsheet::ExternSheetEntry,
    strings,
    supbook::{SupBookInfo, SupBookKind},
};

// BIFF8 supports 65,536 rows (0-based 0..=65,535).
const BIFF8_MAX_ROW0: i64 = u16::MAX as i64;
// Columns are stored in a 14-bit field in many BIFF8 structures.
const BIFF8_MAX_COL0: i64 = 0x3FFF;

/// Context needed to decode BIFF8 `rgce` streams that may reference workbook-scoped metadata such
/// as sheets (`EXTERNSHEET`) and other defined names (`NAME` table).
#[derive(Debug, Clone)]
pub(crate) struct DefinedNameMeta {
    pub(crate) name: String,
    /// BIFF sheet index (0-based) for local names, or `None` for workbook scope.
    pub(crate) scope_sheet: Option<usize>,
}

pub(crate) struct RgceDecodeContext<'a> {
    pub(crate) codepage: u16,
    pub(crate) sheet_names: &'a [String],
    pub(crate) externsheet: &'a [ExternSheetEntry],
    pub(crate) supbooks: &'a [SupBookInfo],
    pub(crate) defined_names: &'a [DefinedNameMeta],
}

#[derive(Debug, Clone)]
pub(crate) struct DecodeRgceResult {
    pub(crate) text: String,
    pub(crate) warnings: Vec<String>,
}

/// (0-indexed) cell coordinate used as the origin for relative-reference ptgs (`PtgRefN` /
/// `PtgAreaN`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellCoord {
    pub(crate) row: u32,
    pub(crate) col: u32,
}

impl CellCoord {
    pub(crate) const fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }
}

/// Legacy/compat alias used by earlier BIFF defined-name decoding codepaths.
///
/// `DecodedRgce` intentionally matches `DecodeRgceResult` so callers can access `.text` and
/// `.warnings` directly.
pub(crate) type DecodedRgce = DecodeRgceResult;

/// Decode a BIFF8 `rgce` token stream used in a defined-name (`NAME`) record.
///
/// This is a convenience wrapper around [`decode_biff8_rgce`] that builds an empty decode context.
/// The returned text does **not** include a leading `=`.
#[allow(dead_code)]
pub(crate) fn decode_defined_name_rgce(rgce: &[u8], codepage: u16) -> DecodedRgce {
    let sheet_names: &[String] = &[];
    let externsheet: &[ExternSheetEntry] = &[];
    let supbooks: &[SupBookInfo] = &[];
    let defined_names: &[DefinedNameMeta] = &[];
    let ctx = RgceDecodeContext {
        codepage,
        sheet_names,
        externsheet,
        supbooks,
        defined_names,
    };
    decode_defined_name_rgce_with_context(rgce, codepage, &ctx)
}

/// Decode a BIFF8 `rgce` token stream used in a defined-name (`NAME`) record using workbook
/// context (sheet names, `EXTERNSHEET`, and defined-name metadata).
///
/// This is a thin wrapper around [`decode_biff8_rgce`]. The `codepage` parameter overrides any
/// `ctx.codepage` value so callers can pass a shared context and vary string decoding if needed.
///
/// The returned text does **not** include a leading `=`.
pub(crate) fn decode_defined_name_rgce_with_context(
    rgce: &[u8],
    codepage: u16,
    ctx: &RgceDecodeContext<'_>,
) -> DecodedRgce {
    let ctx = RgceDecodeContext {
        codepage,
        sheet_names: ctx.sheet_names,
        externsheet: ctx.externsheet,
        supbooks: ctx.supbooks,
        defined_names: ctx.defined_names,
    };
    decode_biff8_rgce(rgce, &ctx)
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
    decode_biff8_rgce_with_base(rgce, ctx, None)
}

pub(crate) fn decode_biff8_rgce_with_base(
    rgce: &[u8],
    ctx: &RgceDecodeContext<'_>,
    base: Option<CellCoord>,
) -> DecodeRgceResult {
    let base_is_default = base.is_none();
    let base = base.unwrap_or(CellCoord::new(0, 0));
    if rgce.is_empty() {
        return DecodeRgceResult {
            text: String::new(),
            warnings: Vec::new(),
        };
    }

    let mut input = rgce;
    let mut stack: Vec<ExprFragment> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();
    let mut warned_default_base_for_relative = false;

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
            // Spill range postfix (`#`).
            0x2F => {
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
                    text: format!("{inner}#"),
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
                            // Emit a parseable identifier-like function name that preserves the
                            // raw BIFF function id.
                            name_owned = format!("_UNKNOWN_FUNC_0x{func_id:04X}");
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
                    // `#NAME_ID(...)` is not a valid Excel token; fall back to a parseable error
                    // literal and keep the id in warnings.
                    stack.push(ExprFragment::new("#NAME?".to_string()));
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
            0x39 | 0x59 | 0x79 => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let iname = u16::from_le_bytes([input[2], input[3]]);
                input = &input[6..];

                // Excel uses `PtgNameX` for external workbook defined names and for add-in/UDF
                // function calls. For UDF calls, the token sequence is typically:
                //
                //   args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
                //
                // In this case we must decode `PtgNameX` into a *function identifier* (no workbook
                // prefix or quoting), otherwise the rendered formula won't be parseable.
                let is_udf_call = matches!(input.get(0), Some(0x22 | 0x42 | 0x62))
                    && input.len() >= 4
                    && u16::from_le_bytes([input[2], input[3]]) == 0x00FF;

                match format_namex_ref(ixti, iname, is_udf_call, ctx) {
                    Ok(txt) => stack.push(ExprFragment::new(txt)),
                    Err(err) => {
                        warnings.push(err);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                    }
                }
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
                let area = format_area_ref_ptg_area(row1, col1, row2, col2, &mut warnings);
                stack.push(ExprFragment::new(area));
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
            // PtgRefErr / PtgRefErrN: [rw: u16][col: u16]
            //
            // Some documentation refers to the relative-reference error tokens as `PtgRefErrN` /
            // `PtgAreaErrN`; in BIFF8 the ptg id is the same regardless of context.
            0x2A | 0x4A | 0x6A => {
                if input.len() < 4 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[4..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr / PtgAreaErrN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                input = &input[8..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN: [rw: u16][col: u16]
            0x2C | 0x4C | 0x6C => {
                if input.len() < 4 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row_raw = u16::from_le_bytes([input[0], input[1]]);
                let col_field = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];
                if base_is_default
                    && !warned_default_base_for_relative
                    && (col_field & RELATIVE_MASK) != 0
                {
                    warnings.push(
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)"
                            .to_string(),
                    );
                    warned_default_base_for_relative = true;
                }
                stack.push(ExprFragment::new(decode_ref_n(
                    row_raw,
                    col_field,
                    base,
                    &mut warnings,
                    "PtgRefN",
                )));
            }
            // PtgAreaN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2D | 0x4D | 0x6D => {
                if input.len() < 8 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let row_first_raw = u16::from_le_bytes([input[0], input[1]]);
                let row_last_raw = u16::from_le_bytes([input[2], input[3]]);
                let col_first_field = u16::from_le_bytes([input[4], input[5]]);
                let col_last_field = u16::from_le_bytes([input[6], input[7]]);
                input = &input[8..];
                if base_is_default
                    && !warned_default_base_for_relative
                    && ((col_first_field & RELATIVE_MASK) != 0
                        || (col_last_field & RELATIVE_MASK) != 0)
                {
                    warnings.push(
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)"
                            .to_string(),
                    );
                    warned_default_base_for_relative = true;
                }
                stack.push(ExprFragment::new(decode_area_n(
                    row_first_raw,
                    col_first_field,
                    row_last_raw,
                    col_last_field,
                    base,
                    &mut warnings,
                    "PtgAreaN",
                )));
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
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
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
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
                    }
                };
                let area = format_area_ref_ptg_area(row1, col1, row2, col2, &mut warnings);
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
            // PtgRefN3d (relative/absolute 3D reference): [ixti: u16][rw: u16][col: u16]
            0x3E | 0x5E | 0x7E => {
                if input.len() < 6 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row_raw = u16::from_le_bytes([input[2], input[3]]);
                let col_field = u16::from_le_bytes([input[4], input[5]]);
                input = &input[6..];

                if base_is_default
                    && !warned_default_base_for_relative
                    && (col_field & RELATIVE_MASK) != 0
                {
                    warnings.push(
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)"
                            .to_string(),
                    );
                    warned_default_base_for_relative = true;
                }

                let cell = decode_ref_n(row_raw, col_field, base, &mut warnings, "PtgRefN3d");
                if cell == "#REF!" {
                    stack.push(ExprFragment::new(cell));
                    continue;
                }

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        warnings.push(err);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
                    }
                };
                stack.push(ExprFragment::new(format!("{sheet_prefix}{cell}")));
            }
            // PtgAreaN3d (relative/absolute 3D area):
            // [ixti: u16][rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x3F | 0x5F | 0x7F => {
                if input.len() < 10 {
                    warnings.push("unexpected end of rgce stream".to_string());
                    return unsupported(ptg, warnings);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row_first_raw = u16::from_le_bytes([input[2], input[3]]);
                let row_last_raw = u16::from_le_bytes([input[4], input[5]]);
                let col_first_field = u16::from_le_bytes([input[6], input[7]]);
                let col_last_field = u16::from_le_bytes([input[8], input[9]]);
                input = &input[10..];

                if base_is_default
                    && !warned_default_base_for_relative
                    && ((col_first_field & RELATIVE_MASK) != 0
                        || (col_last_field & RELATIVE_MASK) != 0)
                {
                    warnings.push(
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)"
                            .to_string(),
                    );
                    warned_default_base_for_relative = true;
                }

                let area = decode_area_n(
                    row_first_raw,
                    col_first_field,
                    row_last_raw,
                    col_last_field,
                    base,
                    &mut warnings,
                    "PtgAreaN3d",
                );
                if area == "#REF!" {
                    stack.push(ExprFragment::new(area));
                    continue;
                }

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        warnings.push(err);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
                    }
                };
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
    let mut warnings = warnings;
    let msg = format!("unsupported rgce token 0x{ptg:02X}");
    if !warnings.iter().any(|w| w == &msg) {
        warnings.push(msg);
    }
    DecodeRgceResult {
        // Use a parseable, stable Excel error literal so callers can round-trip the decoded
        // formula through `formula-engine`.
        //
        // We intentionally use `#UNKNOWN!` here rather than encoding custom placeholders like
        // `#UNSUPPORTED_PTG_0xNN!`. `formula-engine` maps unrecognized error literals to `#VALUE!`,
        // losing the intent that the decoded token was unknown/unsupported.
        text: "#UNKNOWN!".to_string(),
        warnings,
    }
}

const COL_INDEX_MASK: u16 = 0x3FFF;
const ROW_RELATIVE_BIT: u16 = 0x4000;
const COL_RELATIVE_BIT: u16 = 0x8000;
const RELATIVE_MASK: u16 = 0xC000;

fn decode_ref_n(
    row_raw: u16,
    col_field: u16,
    base: CellCoord,
    warnings: &mut Vec<String>,
    ptg_name: &str,
) -> String {
    let row_relative = (col_field & ROW_RELATIVE_BIT) == ROW_RELATIVE_BIT;
    let col_relative = (col_field & COL_RELATIVE_BIT) == COL_RELATIVE_BIT;
    let col_raw = col_field & COL_INDEX_MASK;

    let row_off = if row_relative { row_raw as i16 } else { 0 };
    let col_off = if col_relative {
        sign_extend_14(col_raw)
    } else {
        0
    };

    let abs_row = if row_relative {
        (base.row as i64).saturating_add(row_off as i64)
    } else {
        row_raw as i64
    };

    let abs_col = if col_relative {
        (base.col as i64).saturating_add(col_off as i64)
    } else {
        col_raw as i64
    };

    if !cell_in_bounds(abs_row, abs_col) {
        let row_desc = if row_relative {
            format!("row_off={row_off}")
        } else {
            format!("row={row_raw}")
        };
        let col_desc = if col_relative {
            format!("col_off={col_off}")
        } else {
            format!("col={col_raw}")
        };
        warnings.push(format!(
            "{ptg_name} produced out-of-bounds reference: base=({},{}), {row_desc}, {col_desc} -> #REF!",
            base.row, base.col
        ));
        return "#REF!".to_string();
    }

    let col_abs_field: u16 = ((abs_col as u16) & COL_INDEX_MASK) | (col_field & RELATIVE_MASK);
    format_cell_ref(abs_row as u16, col_abs_field)
}

fn decode_area_n(
    row1_raw: u16,
    col1_field: u16,
    row2_raw: u16,
    col2_field: u16,
    base: CellCoord,
    warnings: &mut Vec<String>,
    ptg_name: &str,
) -> String {
    let row1_relative = (col1_field & ROW_RELATIVE_BIT) == ROW_RELATIVE_BIT;
    let col1_relative = (col1_field & COL_RELATIVE_BIT) == COL_RELATIVE_BIT;
    let row2_relative = (col2_field & ROW_RELATIVE_BIT) == ROW_RELATIVE_BIT;
    let col2_relative = (col2_field & COL_RELATIVE_BIT) == COL_RELATIVE_BIT;

    let col1_raw = col1_field & COL_INDEX_MASK;
    let col2_raw = col2_field & COL_INDEX_MASK;

    let abs_row1 = if row1_relative {
        (base.row as i64).saturating_add(row1_raw as i16 as i64)
    } else {
        row1_raw as i64
    };
    let abs_row2 = if row2_relative {
        (base.row as i64).saturating_add(row2_raw as i16 as i64)
    } else {
        row2_raw as i64
    };

    let abs_col1 = if col1_relative {
        let col_off = sign_extend_14(col1_raw) as i64;
        (base.col as i64).saturating_add(col_off)
    } else {
        col1_raw as i64
    };
    let abs_col2 = if col2_relative {
        let col_off = sign_extend_14(col2_raw) as i64;
        (base.col as i64).saturating_add(col_off)
    } else {
        col2_raw as i64
    };

    if !cell_in_bounds(abs_row1, abs_col1) || !cell_in_bounds(abs_row2, abs_col2) {
        let row1_desc = if row1_relative {
            format!("row1_off={}", row1_raw as i16)
        } else {
            format!("row1={row1_raw}")
        };
        let row2_desc = if row2_relative {
            format!("row2_off={}", row2_raw as i16)
        } else {
            format!("row2={row2_raw}")
        };
        let col1_desc = if col1_relative {
            format!("col1_off={}", sign_extend_14(col1_raw))
        } else {
            format!("col1={col1_raw}")
        };
        let col2_desc = if col2_relative {
            format!("col2_off={}", sign_extend_14(col2_raw))
        } else {
            format!("col2={col2_raw}")
        };
        warnings.push(format!(
            "{ptg_name} produced out-of-bounds area: base=({},{}), {row1_desc}, {row2_desc}, {col1_desc}, {col2_desc} -> #REF!",
            base.row, base.col
        ));
        return "#REF!".to_string();
    }

    let col1_abs_field: u16 = ((abs_col1 as u16) & COL_INDEX_MASK) | (col1_field & RELATIVE_MASK);
    let col2_abs_field: u16 = ((abs_col2 as u16) & COL_INDEX_MASK) | (col2_field & RELATIVE_MASK);

    format_area_ref(
        abs_row1 as u16,
        col1_abs_field,
        abs_row2 as u16,
        col2_abs_field,
    )
}

fn cell_in_bounds(row: i64, col: i64) -> bool {
    row >= 0 && row <= BIFF8_MAX_ROW0 && col >= 0 && col <= BIFF8_MAX_COL0
}

fn sign_extend_14(v: u16) -> i16 {
    debug_assert!(v <= COL_INDEX_MASK);
    // 14-bit two's complement. If bit13 is set, treat as negative.
    if (v & 0x2000) != 0 {
        (v | 0xC000) as i16
    } else {
        v as i16
    }
}

#[cfg(test)]
fn format_cell_ref_no_dollars(row0: u32, col0: u32) -> String {
    let mut out = String::new();
    push_column(col0, &mut out);
    out.push_str(&(row0 + 1).to_string());
    out
}

fn format_cell_ref(row: u16, col_with_flags: u16) -> String {
    let row_rel = (col_with_flags & 0x4000) != 0;
    let col_rel = (col_with_flags & 0x8000) != 0;
    let col = col_with_flags & COL_INDEX_MASK;

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

fn format_area_ref_ptg_area(
    row1: u16,
    col1: u16,
    row2: u16,
    col2: u16,
    warnings: &mut Vec<String>,
) -> String {
    let col1_idx = col1 & COL_INDEX_MASK;
    let col2_idx = col2 & COL_INDEX_MASK;

    // Relative bits are stored in the upper bits of the `col` field.
    let row1_rel = (col1 & ROW_RELATIVE_BIT) != 0;
    let row2_rel = (col2 & ROW_RELATIVE_BIT) != 0;
    let col1_rel = (col1 & COL_RELATIVE_BIT) != 0;
    let col2_rel = (col2 & COL_RELATIVE_BIT) != 0;

    // Best-effort whole-row / whole-column decoding for BIFF8 areas.
    //
    // Print titles (`_xlnm.Print_Titles`) are commonly stored as full-row/full-column areas; render
    // them using Excel's canonical shorthand (`$1:$1`, `$A:$A`) rather than `A1:IV1` /
    // `A1:A65536`.
    let is_full_width = col2_idx == 0x00FF || col2_idx == 0x3FFF;
    let is_full_height = row2 == 0xFFFF;

    let is_whole_row = row1 == row2 && col1_idx == 0 && is_full_width;
    let is_whole_col = col1_idx == col2_idx && row1 == 0 && is_full_height;

    if is_whole_row && is_whole_col {
        // Degenerate/garbage: avoid choosing one shorthand over the other.
        warnings.push(format!(
            "BIFF8 area matches both whole-row and whole-column patterns (rwFirst={row1}, rwLast={row2}, colFirst=0x{col1:04X}, colLast=0x{col2:04X}); rendering as explicit A1-style area"
        ));
        return format_area_ref(row1, col1, row2, col2);
    }

    if is_whole_row {
        let row_rel = if row1_rel != row2_rel {
            warnings.push(format!(
                "BIFF8 whole-row area has mismatched row-relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
            ));
            row1_rel
        } else {
            row1_rel
        };

        let r = format_row_ref(row1, row_rel);
        // Excel includes the `:` even for a single-row span.
        return format!("{r}:{r}");
    }

    if is_whole_col {
        let col_rel = if col1_rel != col2_rel {
            warnings.push(format!(
                "BIFF8 whole-column area has mismatched col-relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
            ));
            col1_rel
        } else {
            col1_rel
        };

        let c = format_col_ref(col1_idx as u16, col_rel);
        // Excel includes the `:` even for a single-column span.
        return format!("{c}:{c}");
    }

    format_area_ref(row1, col1, row2, col2)
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
    // Some producers use `0x3FFF` as the "max column" sentinel (full 14-bit width); treat it as
    // full-width for whole-row formatting too.
    if col1_idx == 0 && (col2_idx == BIFF8_MAX_COL || col2_idx == 0x3FFF) {
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

    // Internal workbook reference.
    if entry.supbook == 0 {
        return format_internal_sheet_ref(ixti, entry.itab_first, entry.itab_last, ctx);
    }

    // Some writers may still reference the internal workbook SUPBOOK explicitly; detect it via
    // virtPath marker if present.
    if let Some(sb) = ctx.supbooks.get(entry.supbook as usize) {
        if sb.is_internal() {
            return format_internal_sheet_ref(ixti, entry.itab_first, entry.itab_last, ctx);
        }
    }

    // External workbook reference. Best-effort: if SUPBOOK metadata is missing or incomplete,
    // surface an error so the caller can fall back to a parseable placeholder (e.g. `#REF!`).
    let sb = ctx.supbooks.get(entry.supbook as usize).ok_or_else(|| {
        format!(
            "EXTERNSHEET entry ixti={ixti} references missing SUPBOOK index {} (supbook count={})",
            entry.supbook,
            ctx.supbooks.len()
        )
    })?;

    let workbook_raw = sb
        .workbook_name
        .as_deref()
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| sb.virt_path.as_str());
    if workbook_raw.is_empty() {
        return Err(format!(
            "SUPBOOK index {} referenced by EXTERNSHEET ixti={ixti} has empty workbook name",
            entry.supbook
        ));
    }

    if entry.itab_first < 0 || entry.itab_last < 0 {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} has negative sheet indices itabFirst={} itabLast={} for external workbook",
            entry.itab_first, entry.itab_last
        ));
    }

    let itab_first = entry.itab_first as usize;
    let itab_last = entry.itab_last as usize;

    let Some(sheet_first) = sb.sheet_names.get(itab_first) else {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} refers to out-of-range external itabFirst={itab_first} (SUPBOOK sheet count={})",
            sb.sheet_names.len()
        ));
    };
    let Some(sheet_last) = sb.sheet_names.get(itab_last) else {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} refers to out-of-range external itabLast={itab_last} (SUPBOOK sheet count={})",
            sb.sheet_names.len()
        ));
    };

    let workbook = format_external_workbook_name(workbook_raw);
    let start = format!("{workbook}{sheet_first}");
    if itab_first == itab_last {
        Ok(format!("{}!", quote_sheet_name_if_needed(&start)))
    } else {
        // External 3D sheet ranges only include the workbook prefix once:
        // `'[Book.xlsx]SheetA:SheetC'!A1`
        Ok(format!(
            "{}!",
            quote_sheet_range_name_if_needed(&start, sheet_last)
        ))
    }
}

fn format_namex_ref(
    ixti: u16,
    iname: u16,
    is_function: bool,
    ctx: &RgceDecodeContext<'_>,
) -> Result<String, String> {
    if iname == 0 {
        return Err(format!(
            "PtgNameX has invalid iname=0 (expected 1-based external name index, ixti={ixti})"
        ));
    }

    // In BIFF8, PtgNameX stores `ixti` (index into EXTERNSHEET), which in turn points at a SUPBOOK.
    // Some writers may instead store the SUPBOOK index directly (when EXTERNSHEET is missing). Be
    // permissive and treat missing EXTERNSHEET as a signal to interpret `ixti` as `iSupBook`.
    let (supbook_index, sheet_ref_available) = match ctx.externsheet.get(ixti as usize) {
        Some(entry) => (entry.supbook, true),
        None => (ixti, false),
    };

    let Some(sb) = ctx.supbooks.get(supbook_index as usize) else {
        return Err(format!(
            "PtgNameX references missing SUPBOOK index {supbook_index} (ixti={ixti}, supbook count={})",
            ctx.supbooks.len()
        ));
    };

    let name_idx = (iname as usize)
        .checked_sub(1)
        .ok_or_else(|| "iname underflow".to_string())?;
    let Some(extern_name) = sb.extern_names.get(name_idx) else {
        return Err(format!(
            "PtgNameX references missing EXTERNNAME iname={iname} for SUPBOOK index {supbook_index} (extern name count={})",
            sb.extern_names.len()
        ));
    };

    if is_function {
        if extern_name.is_empty() {
            return Err(format!(
                "PtgNameX function reference has empty EXTERNNAME (ixti={ixti}, iname={iname})"
            ));
        }
        return Ok(extern_name.clone());
    }

    match sb.kind {
        // Internal workbook EXTERNNAME (rare, but some producers use PtgNameX even for internal
        // sheet-scoped names). If a usable internal sheet prefix is available via EXTERNSHEET,
        // include it so the rendered formula matches Excel’s `Sheet1!Name` style.
        SupBookKind::Internal => {
            if sheet_ref_available {
                if let Ok(prefix) = format_sheet_ref(ixti, ctx) {
                    return Ok(format!("{prefix}{extern_name}"));
                }
            }
            Ok(extern_name.clone())
        }
        // Workbook- or sheet-scoped external name in another workbook.
        SupBookKind::ExternalWorkbook => {
            // If the EXTERNSHEET entry has a meaningful sheet ref (itab values), format
            // `'[Book.xlsx]Sheet1'!MyName`. Otherwise fall back to workbook-scoped `'[Book.xlsx]MyName'`.
            if sheet_ref_available {
                if let Ok(prefix) = format_sheet_ref(ixti, ctx) {
                    return Ok(format!("{prefix}{extern_name}"));
                }
            }

            let workbook_raw = sb
                .workbook_name
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| sb.virt_path.as_str());
            if workbook_raw.is_empty() {
                return Err(format!(
                    "SUPBOOK index {supbook_index} referenced by PtgNameX has empty workbook name"
                ));
            }

            let workbook = format_external_workbook_name(workbook_raw);
            // Workbook-scoped external names use the Excel form `[Book]Name`, but our formula
            // parser can mis-tokenize unquoted `[Book]Name` as a structured reference. Quote the
            // full token so it becomes a `QuotedIdent`.
            let token = format!("{workbook}{extern_name}");
            Ok(quote_sheet_name_if_needed(&token))
        }
        // Add-in/UDF/library. Best-effort: use the extern name itself.
        _ => Ok(extern_name.clone()),
    }
}

fn format_internal_sheet_ref(
    ixti: u16,
    itab_first: i16,
    itab_last: i16,
    ctx: &RgceDecodeContext<'_>,
) -> Result<String, String> {
    if itab_first < 0 || itab_last < 0 {
        return Err(format!(
            "EXTERNSHEET entry ixti={ixti} has negative sheet indices itabFirst={itab_first} itabLast={itab_last}"
        ));
    }

    let itab_first = itab_first as usize;
    let itab_last = itab_last as usize;

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

fn format_external_workbook_name(workbook: &str) -> String {
    // Excel external workbook refs wrap the workbook in brackets: `[Book1.xlsx]`.
    //
    // Some producers may already include brackets; normalize to a single set.
    let trimmed = workbook.trim();
    let inner = trimmed
        .strip_prefix('[')
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(trimmed);
    format!("[{inner}]")
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
    use std::str::FromStr;

    fn assert_parseable(expr: &str) {
        let expr = expr.trim();
        assert!(!expr.is_empty(), "decoded expression must be non-empty");
        let formula = format!("={expr}");
        let ast = parse_formula(&formula, ParseOptions::default()).unwrap_or_else(|err| {
            panic!(
                "expected decoded expression to be parseable, expr={expr:?}, err={err:?}, formula={formula:?}"
            );
        });

        // Validate that any error literals in the parsed AST are *known* Excel errors (so we don't
        // regress back to emitting custom, unsupported error-like placeholders such as `#SHEET`).
        fn assert_known_errors(expr: &formula_engine::Expr) {
            match expr {
                formula_engine::Expr::Error(e) => {
                    assert!(
                        formula_model::ErrorValue::from_str(e).is_ok(),
                        "unexpected Excel error literal: {e:?}"
                    );
                }
                formula_engine::Expr::FieldAccess(fa) => assert_known_errors(&fa.base),
                formula_engine::Expr::Array(arr) => {
                    for row in &arr.rows {
                        for el in row {
                            assert_known_errors(el);
                        }
                    }
                }
                formula_engine::Expr::FunctionCall(call) => {
                    for arg in &call.args {
                        assert_known_errors(arg);
                    }
                }
                formula_engine::Expr::Call(call) => {
                    assert_known_errors(&call.callee);
                    for arg in &call.args {
                        assert_known_errors(arg);
                    }
                }
                formula_engine::Expr::Unary(u) => assert_known_errors(&u.expr),
                formula_engine::Expr::Postfix(p) => assert_known_errors(&p.expr),
                formula_engine::Expr::Binary(b) => {
                    assert_known_errors(&b.left);
                    assert_known_errors(&b.right);
                }
                // Leaf nodes.
                formula_engine::Expr::Number(_)
                | formula_engine::Expr::String(_)
                | formula_engine::Expr::Boolean(_)
                | formula_engine::Expr::NameRef(_)
                | formula_engine::Expr::CellRef(_)
                | formula_engine::Expr::ColRef(_)
                | formula_engine::Expr::RowRef(_)
                | formula_engine::Expr::StructuredRef(_)
                | formula_engine::Expr::Missing => {}
            }
        }

        assert_known_errors(&ast.expr);
    }

    /*
    // Legacy helpers retained for reference: originally `assert_parseable` only validated A1 ranges
    // for print area / print titles style defined names.
    fn is_row_range(s: &str) -> bool {
        let Some((a, b)) = s.split_once(':') else {
            return false;
        };

        let a = a.trim().replace('$', "");
        let b = b.trim().replace('$', "");
        if a.is_empty() || b.is_empty() {
            return false;
        }
        if !a.chars().all(|c| c.is_ascii_digit()) || !b.chars().all(|c| c.is_ascii_digit()) {
            return false;
        }
        match (a.parse::<u32>(), b.parse::<u32>()) {
            (Ok(a), Ok(b)) => a > 0 && b > 0,
            _ => false,
        }
    }

    fn is_col_range(s: &str) -> bool {
        let Some((a, b)) = s.split_once(':') else {
            return false;
        };

        let a = a.trim().replace('$', "");
        let b = b.trim().replace('$', "");
        if a.is_empty() || b.is_empty() {
            return false;
        }
        if !a.chars().all(|c| c.is_ascii_alphabetic())
            || !b.chars().all(|c| c.is_ascii_alphabetic())
        {
            return false;
        }

        col_from_a1(&a).is_some() && col_from_a1(&b).is_some()
    }

    */

    fn assert_print_area_parseable(sheet_name: &str, expr: &str) {
        let mut warnings = Vec::<crate::ImportWarning>::new();
        crate::parse_print_area_refers_to(sheet_name, expr, &mut warnings)
            .expect("parse print area defined name");
    }

    fn assert_print_titles_parseable(sheet_name: &str, expr: &str) {
        let mut warnings = Vec::<crate::ImportWarning>::new();
        crate::parse_print_titles_refers_to(sheet_name, expr, &mut warnings)
            .expect("parse print titles defined name");
    }

    #[test]
    fn formats_cell_ref_no_dollars() {
        assert_eq!(format_cell_ref_no_dollars(0, 0), "A1");
        assert_eq!(format_cell_ref_no_dollars(1, 1), "B2");
        assert_eq!(format_cell_ref_no_dollars(0, 26), "AA1");
    }

    const BIFF8_MAX_ROW: u16 = 0xFFFF;
    const BIFF8_MAX_COL: u16 = 0x00FF;

    fn empty_ctx<'a>(
        sheet_names: &'a [String],
        externsheet: &'a [ExternSheetEntry],
        defined_names: &'a [DefinedNameMeta],
    ) -> RgceDecodeContext<'a> {
        RgceDecodeContext {
            codepage: 1252,
            sheet_names,
            externsheet,
            supbooks: &[],
            defined_names,
        }
    }

    fn encode_col_field(col_value_14: u16, col_relative: bool, row_relative: bool) -> u16 {
        let mut field = col_value_14 & COL_INDEX_MASK;
        if col_relative {
            field |= COL_RELATIVE_BIT;
        }
        if row_relative {
            field |= ROW_RELATIVE_BIT;
        }
        field
    }

    #[test]
    fn decodes_ptg_ref_n_default_base() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // row offset = 0, col offset = 0 (both relative) => A1.
        let row_raw = 0u16;
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x2C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_spill_range_postfix() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // A1# (spill operator applied to a relative ref).
        let row_raw = 0u16;
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x24,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
            0x2F,
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1#");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_isect_intersection_operator() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // A1 B1 (intersection operator).
        let row = 0u16;
        let col_a = encode_col_field(0, true, true);
        let col_b = encode_col_field(1, true, true);
        let rgce = [
            0x24, // PtgRef
            row.to_le_bytes()[0],
            row.to_le_bytes()[1],
            col_a.to_le_bytes()[0],
            col_a.to_le_bytes()[1],
            0x24, // PtgRef
            row.to_le_bytes()[0],
            row.to_le_bytes()[1],
            col_b.to_le_bytes()[0],
            col_b.to_le_bytes()[1],
            0x0F, // PtgIsect (intersection)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1 B1");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_range_colon_operator() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // A1:B2 encoded via two PtgRef tokens + PtgRange (0x11).
        let row_a1 = 0u16;
        let col_a1 = encode_col_field(0, true, true);
        let row_b2 = 1u16;
        let col_b2 = encode_col_field(1, true, true);
        let rgce = [
            0x24, // PtgRef
            row_a1.to_le_bytes()[0],
            row_a1.to_le_bytes()[1],
            col_a1.to_le_bytes()[0],
            col_a1.to_le_bytes()[1],
            0x24, // PtgRef
            row_b2.to_le_bytes()[0],
            row_b2.to_le_bytes()[1],
            col_b2.to_le_bytes()[0],
            col_b2.to_le_bytes()[1],
            0x11, // PtgRange (:)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_paren_union_arg_without_double_parenthesizing() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM((A1,B1)) with explicit PtgParen around the union expression.
        //
        // Without the `PtgParen` special-case in `format_function_call`, we could end up with
        // SUM(((A1,B1))) (triple parens).
        let a1_col = encode_col_field(0, true, true);
        let b1_col = encode_col_field(1, true, true);
        let rgce = vec![
            0x24, 0x00, 0x00, a1_col.to_le_bytes()[0], a1_col.to_le_bytes()[1], // A1
            0x24, 0x00, 0x00, b1_col.to_le_bytes()[0], b1_col.to_le_bytes()[1], // B1
            0x10, // union operator
            0x15, // explicit paren
            0x22, 0x01, 0x04, 0x00, // SUM(argc=1)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM((A1,B1))");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_attr_sum() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // BIFF8 can encode SUM with a PtgAttr token (tAttrSum bit).
        //
        // rgce:
        //   PtgRef A1
        //   PtgAttr grbit=tAttrSum (0x10), wAttr=0
        let a1_col = encode_col_field(0, true, true);
        let rgce = vec![
            0x24, 0x00, 0x00, a1_col.to_le_bytes()[0], a1_col.to_le_bytes()[1], // A1
            0x19, 0x10, 0x00, 0x00, // PtgAttr(tAttrSum)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM(A1)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_unsupported_tokens_as_parseable_excel_errors() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgExp (0x01) is not supported by this best-effort NAME rgce printer.
        let rgce = [0x01];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!");
        assert!(!decoded.warnings.is_empty(), "expected warnings");
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_missing_ptgname_indices_as_parseable_excel_errors() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgName: name_id=1, reserved bytes=0.
        let rgce = [0x23, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#NAME?");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("missing name index 1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_unknown_ptgfuncvar_ids_as_parseable_function_calls() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgFuncVar with argc=0, func_id=0xFFFF (unknown).
        let rgce = [0x22, 0x00, 0xFF, 0xFF];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "_UNKNOWN_FUNC_0xFFFF()");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0xFFFF")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_unknown_ptgfuncvar_ids_with_arguments_as_parseable_function_calls() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // _UNKNOWN_FUNC_0xFFFF(1,2):
        //   PtgInt 1
        //   PtgInt 2
        //   PtgFuncVar argc=2 iftab=0xFFFF (unknown)
        let rgce = [
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x22, 0x02, 0xFF, 0xFF, // funcvar(argc=2, iftab=0xFFFF)
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "_UNKNOWN_FUNC_0xFFFF(1,2)");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0xFFFF")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_user_defined_ptgfuncvar_calls_with_unresolved_namex_as_parseable_formulas() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Excel encodes user-defined / add-in / future functions as PtgFuncVar with iftab=255,
        // where the function name is the top-of-stack expression.
        //
        // If that name expression is an unresolved PtgNameX (external name reference), we still
        // want the printed formula text to be parseable (even though semantics are best-effort).
        //
        // rgce: PtgNameX(ixti=0,iname=0) ; PtgFuncVar(argc=1,iftab=255)
        let rgce = [
            0x39, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // PtgNameX payload
            0x22, 0x01, 0xFF, 0x00, // PtgFuncVar (user-defined)
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!()");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgNameX")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_3d_refs_with_missing_supbook_as_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &[],
            defined_names: &defined_names,
        };

        // PtgRef3d ixti=0 referencing an EXTERNSHEET entry with iSupBook=1, but no SUPBOOK table.
        let rgce = [0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("SUPBOOK")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_unresolvable_3d_sheet_refs_as_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgRef3d ixti=0 with missing EXTERNSHEET table.
        let rgce = [0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("ixti=0")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref_n_value_class_variant() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // row offset = 1, col offset = 1 (both relative) => B2.
        let row_raw = 1u16;
        let col_field = encode_col_field(1, true, true);
        let rgce = [
            0x4C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "B2");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref_n_with_base() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let base = CellCoord::new(10, 10); // K11
                                           // row offset = -2, col offset = +3 => N9.
        let row_raw = (-2i16) as u16;
        let col_field = encode_col_field(3, true, true);
        let rgce = [
            0x2C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "N9");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn ptg_ref_n_out_of_range_emits_ref_error() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let base = CellCoord::new(0, 0);
        // row offset = -1 => invalid.
        let row_raw = (-1i16) as u16;
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x2C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgRefN")),
            "warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("row_off=-1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn ptg_ref_n_col_out_of_range_emits_ref_error() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // base at max BIFF column index, then offset +1.
        let base = CellCoord::new(0, BIFF8_MAX_COL0 as u32);
        let row_raw = 0u16;
        let col_field = encode_col_field(1, true, true);
        let rgce = [
            0x2C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgRefN")),
            "warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("col_off=1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn ptg_ref_n_preserves_absolute_relative_flags() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let base = CellCoord::new(5, 5);
        // Row is relative (+1), column is absolute (C => 2).
        let row_raw = 1u16;
        let col_field = encode_col_field(2, false, true);
        let rgce = [
            0x2C,
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "$C7");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_area_n_default_base() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // (0,0)-(1,1) with all components relative => A1:B2.
        let row_first = 0u16;
        let row_last = 1u16;
        let col_first = encode_col_field(0, true, true);
        let col_last = encode_col_field(1, true, true);
        let rgce = [
            0x2D,
            row_first.to_le_bytes()[0],
            row_first.to_le_bytes()[1],
            row_last.to_le_bytes()[0],
            row_last.to_le_bytes()[1],
            col_first.to_le_bytes()[0],
            col_first.to_le_bytes()[1],
            col_last.to_le_bytes()[0],
            col_last.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "A1:B2");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_area_n_with_base() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let base = CellCoord::new(10, 10);
        // First corner offset (-1,-1), last corner offset (+1,+2) => J10:M12
        let row_first = (-1i16) as u16;
        let row_last = (1i16) as u16;
        let col_first_off_14 = 0x3FFF; // -1 in 14-bit two's complement
        let col_last_off_14 = 2u16;
        let col_first = encode_col_field(col_first_off_14, true, true);
        let col_last = encode_col_field(col_last_off_14, true, true);
        let rgce = [
            0x2D, // PtgAreaN
            row_first.to_le_bytes()[0],
            row_first.to_le_bytes()[1],
            row_last.to_le_bytes()[0],
            row_last.to_le_bytes()[1],
            col_first.to_le_bytes()[0],
            col_first.to_le_bytes()[1],
            col_last.to_le_bytes()[0],
            col_last.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "J10:M12");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_area_n_array_class_variant() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // C3:D4 (all components relative).
        let row_first = 2u16;
        let row_last = 3u16;
        let col_first = encode_col_field(2, true, true);
        let col_last = encode_col_field(3, true, true);
        let rgce = [
            0x6D, // PtgAreaNA
            row_first.to_le_bytes()[0],
            row_first.to_le_bytes()[1],
            row_last.to_le_bytes()[0],
            row_last.to_le_bytes()[1],
            col_first.to_le_bytes()[0],
            col_first.to_le_bytes()[1],
            col_last.to_le_bytes()[0],
            col_last.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "C3:D4");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_referr_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (4 bytes).
        let rgce = [0x2A, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_areaerr_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (8 bytes).
        let rgce = [0x2B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_referr3d_to_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (6 bytes): ixti=0 row=0 col=0.
        let rgce = [0x3C, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_areaerr3d_to_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_row_area_as_row_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_row_area_with_3fff_max_col_as_row_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Some writers use colLast=0x3FFF (full 14-bit width) for whole-row spans.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&0x3FFFu16.to_le_bytes()); // colLast

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "$1:$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_column_area_as_col_range() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn continues_to_render_rectangular_ranges_as_a1_areas() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_whole_row_area3d_as_row_range() {
        // Same whole-row area but stored as PtgArea3d with an EXTERNSHEET sheet prefix.
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_titles_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_whole_row_area3d_with_3fff_max_col_as_row_range() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
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
        rgce.extend_from_slice(&0x3FFFu16.to_le_bytes()); // colLast (tolerant max col)

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!$1:$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_titles_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_whole_column_area3d_as_col_range() {
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_titles_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_print_titles_union_area3d_as_row_and_col_ranges() {
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_titles_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_ptg_arean3d_to_sheet_range() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Sheet1!A1:B2.
        let row_first_raw = 0u16;
        let row_last_raw = 1u16;
        let col_first_field = encode_col_field(0, true, true);
        let col_last_field = encode_col_field(1, true, true);
        let rgce = [
            0x3F, // PtgAreaN3d
            0x00,
            0x00, // ixti
            row_first_raw.to_le_bytes()[0],
            row_first_raw.to_le_bytes()[1],
            row_last_raw.to_le_bytes()[0],
            row_last_raw.to_le_bytes()[1],
            col_first_field.to_le_bytes()[0],
            col_first_field.to_le_bytes()[1],
            col_last_field.to_le_bytes()[0],
            col_last_field.to_le_bytes()[1],
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!A1:B2");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_area_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_ptg_refn3d_to_sheet_ref() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // ixti=0, row offset=0, col offset=0 (both relative) => Sheet1!A1.
        let row_raw = 0u16;
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x3E, // PtgRefN3d
            0x00,
            0x00, // ixti
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!A1");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("interpreted relative to A1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_area_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_ptg_refn3d_with_base() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);
        let base = CellCoord::new(10, 10); // K11 (0-based)

        // ixti=0, row offset=-2, col offset=+3 (both relative) => Sheet1!N9.
        let row_raw = (-2i16) as u16;
        let col_field = encode_col_field(3, true, true);
        let rgce = [
            0x3E, // PtgRefN3d
            0x00,
            0x00, // ixti
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "Sheet1!N9");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_area_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn decodes_ptg_arean3d_with_base() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);
        let base = CellCoord::new(10, 10); // K11 (0-based)

        // ixti=0, rows -2..-1 and cols +3..+4 (all relative) => Sheet1!N9:O10.
        let row_first_raw = (-2i16) as u16;
        let row_last_raw = (-1i16) as u16;
        let col_first_field = encode_col_field(3, true, true);
        let col_last_field = encode_col_field(4, true, true);
        let rgce = [
            0x3F, // PtgAreaN3d
            0x00,
            0x00, // ixti
            row_first_raw.to_le_bytes()[0],
            row_first_raw.to_le_bytes()[1],
            row_last_raw.to_le_bytes()[0],
            row_last_raw.to_le_bytes()[1],
            col_first_field.to_le_bytes()[0],
            col_first_field.to_le_bytes()[1],
            col_last_field.to_le_bytes()[0],
            col_last_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "Sheet1!N9:O10");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_area_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn ptg_refn3d_preserves_absolute_relative_flags() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);
        let base = CellCoord::new(5, 5);

        // Row is relative (+1), column is absolute (C => 2).
        let row_raw = 1u16;
        let col_field = encode_col_field(2, false, true);
        let rgce = [
            0x3E, // PtgRefN3d
            0x00,
            0x00, // ixti
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "Sheet1!$C7");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_print_area_parseable("Sheet1", &decoded.text);
    }

    #[test]
    fn ptg_refn3d_col_out_of_range_emits_ref_error() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // base at max BIFF column index, then offset +1.
        let base = CellCoord::new(0, BIFF8_MAX_COL0 as u32);
        let row_raw = 0u16;
        let col_field = encode_col_field(1, true, true);
        let rgce = [
            0x3E, // PtgRefN3d
            0x00,
            0x00, // ixti
            row_raw.to_le_bytes()[0],
            row_raw.to_le_bytes()[1],
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgRefN3d")),
            "warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("col_off=1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn ptg_arean3d_col_out_of_range_emits_ref_error() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // base at max BIFF column index, then offset +1 in the second corner.
        let base = CellCoord::new(0, BIFF8_MAX_COL0 as u32);
        let row_first_raw = 0u16;
        let row_last_raw = 0u16;
        let col_first_field = encode_col_field(0, true, true);
        let col_last_field = encode_col_field(1, true, true);
        let rgce = [
            0x3F, // PtgAreaN3d
            0x00,
            0x00, // ixti
            row_first_raw.to_le_bytes()[0],
            row_first_raw.to_le_bytes()[1],
            row_last_raw.to_le_bytes()[0],
            row_last_raw.to_le_bytes()[1],
            col_first_field.to_le_bytes()[0],
            col_first_field.to_le_bytes()[1],
            col_last_field.to_le_bytes()[0],
            col_last_field.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce_with_base(&rgce, &ctx, Some(base));
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgAreaN3d")),
            "warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("col2_off=1")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref3d_with_missing_sheet_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgRef3d with an ixti that points to a sheet index we don't have.
        let rgce = [0x3A, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("ixti=0")),
            "expected warning for missing sheet, warnings={:?}",
            decoded.warnings,
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_name_missing_to_name_error() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgName with name_id=1 but empty ctx.defined_names.
        let rgce = [0x23, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#NAME?");
        assert!(
            !decoded.warnings.is_empty(),
            "expected warning for missing name, warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_unknown_ptg_to_parseable_error_literal() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0x00")),
            "expected warning to include original ptg id, warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_stack_underflow_to_name_error_literal_with_ptg_warning() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgAdd requires two operands; with an empty stack this triggers stack underflow and the
        // decoder falls back to a parseable Excel error literal.
        let rgce = [0x03];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("stack underflow")),
            "expected stack underflow warning, warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0x03")),
            "expected ptg id in warnings, warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_truncated_ptg_payload_to_name_error_literal_with_ptg_warning() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgInt expects 2 bytes of payload; provide only 1 to trigger unexpected EOF and the
        // decoder falls back to a parseable Excel error literal.
        let rgce = [0x1E, 0x01];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!");
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w.contains("unexpected end of rgce stream")),
            "expected unexpected EOF warning, warnings={:?}",
            decoded.warnings
        );
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0x1E")),
            "expected ptg id in warnings, warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_unknown_error_code_to_unknown_error_literal() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgErr with an unknown error code byte.
        let rgce = [0x1C, 0xFF];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0xFF")),
            "expected warning to include unknown error code, warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_with_missing_context_to_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Dummy payload (6 bytes): ixti=0, iname=1, reserved=0.
        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgNameX")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_external_workbook_workbook_scoped_name() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: -1,
            itab_last: -1,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0001}".to_string(),
                kind: SupBookKind::Internal,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: Vec::new(),
            },
            SupBookInfo {
                ctab: 0,
                virt_path: "Book2.xlsx".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: Some("Book2.xlsx".to_string()),
                sheet_names: vec!["Sheet1".to_string()],
                extern_names: vec!["MyName".to_string()],
            },
        ];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'[Book2.xlsx]MyName'");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_external_workbook_sheet_scoped_name() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0001}".to_string(),
                kind: SupBookKind::Internal,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: Vec::new(),
            },
            SupBookInfo {
                ctab: 0,
                virt_path: "Book2.xlsx".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: Some("Book2.xlsx".to_string()),
                sheet_names: vec!["Sheet1".to_string()],
                extern_names: vec!["MyName".to_string()],
            },
        ];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'[Book2.xlsx]Sheet1'!MyName");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_internal_sheet_scoped_name() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![SupBookInfo {
            ctab: 0,
            virt_path: "\u{0001}".to_string(),
            kind: SupBookKind::Internal,
            workbook_name: None,
            sheet_names: Vec::new(),
            extern_names: vec!["MyName".to_string()],
        }];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!MyName");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref3d_external_workbook_sheet_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0001}".to_string(),
                kind: SupBookKind::Internal,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: Vec::new(),
            },
            SupBookInfo {
                ctab: 0,
                virt_path: "Book2.xlsx".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: Some("Book2.xlsx".to_string()),
                sheet_names: vec!["Sheet1".to_string()],
                extern_names: Vec::new(),
            },
        ];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        // '[Book2.xlsx]Sheet1'!A1
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x3A, // PtgRef3d
            0x00, 0x00, // ixti=0
            0x00, 0x00, // row=0
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'[Book2.xlsx]Sheet1'!A1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_udf_function_name_for_ptg_funcvar_00ff() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: -1,
            itab_last: -1,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0001}".to_string(),
                kind: SupBookKind::Internal,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: Vec::new(),
            },
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0002}".to_string(),
                kind: SupBookKind::Other,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: vec!["MyFunc".to_string()],
            },
        ];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        // MyFunc(1) encoded as: [PtgInt 1][PtgNameX][PtgFuncVar argc=2 iftab=0x00FF]
        let rgce = vec![
            0x1E, 0x01, 0x00, // PtgInt 1
            0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, // PtgNameX (ixti=0, iname=1)
            0x22, 0x02, 0xFF, 0x00, // PtgFuncVar(argc=2, iftab=0x00FF)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "MyFunc(1)");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_sum_1_2_from_ptg_funcvar() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM(1,2):
        //   PtgInt 1
        //   PtgInt 2
        //   PtgFuncVar argc=2 iftab=4 (SUM)
        let rgce = vec![0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x22, 0x02, 0x04, 0x00];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM(1,2)");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn wraps_union_operator_when_used_as_single_function_argument() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM((A1,B1)):
        // - A1 and B1 are combined with the union operator (`,`) into a single reference operand.
        // - Excel requires parentheses around a union operand when used as a *single* function arg
        //   so the comma is not interpreted as an argument separator.
        //
        // rgce:
        //   PtgRef A1
        //   PtgRef B1
        //   PtgUnion
        //   PtgFuncVar argc=1 iftab=4 (SUM)
        let a1_col = encode_col_field(0, true, true);
        let b1_col = encode_col_field(1, true, true);
        let rgce = vec![
            0x24, 0x00, 0x00, a1_col.to_le_bytes()[0], a1_col.to_le_bytes()[1], // A1
            0x24, 0x00, 0x00, b1_col.to_le_bytes()[0], b1_col.to_le_bytes()[1], // B1
            0x10, // union operator
            0x22, 0x01, 0x04, 0x00, // SUM(argc=1)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM((A1,B1))");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_if_true_1_2_from_ptg_funcvar() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_abs_neg1_from_ptg_func() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_sum_sheet1_a1_2_from_ptg_ref3d_and_funcvar() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_quotes_sheet_names_with_spaces() {
        let sheet_names: Vec<String> = vec!["My Sheet".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgRef3d: [ptg][ixti: u16][rw: u16][col: u16] => sheet-qualified $A$1.
        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'My Sheet'!$A$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_escapes_apostrophes_in_sheet_names() {
        let sheet_names: Vec<String> = vec!["O'Brien".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'O''Brien'!$A$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn defined_name_3d_ref_renders_sheet_ranges_as_single_quoted_ident() {
        let sheet_names: Vec<String> = vec![
            "Sheet 1".to_string(),
            "Sheet 2".to_string(),
            "Sheet 3".to_string(),
        ];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 2,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'Sheet 1:Sheet 3'!$A$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_area3d_sheet_range_with_quoting() {
        let sheet_names = vec!["Sheet 1".to_string(), "Sheet3".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
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
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptgname_workbook_and_sheet_scope() {
        let sheet_names = vec!["Sheet 1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
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
        assert!(
            decoded1.warnings.is_empty(),
            "warnings={:?}",
            decoded1.warnings
        );
        assert_parseable(&decoded1.text);

        // Sheet-scoped name (id=2).
        let rgce2 = [0x23, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded2 = decode_biff8_rgce(&rgce2, &ctx);
        assert_eq!(decoded2.text, "'Sheet 1'!LocalName");
        assert!(
            decoded2.warnings.is_empty(),
            "warnings={:?}",
            decoded2.warnings
        );
        assert_parseable(&decoded2.text);
    }

    #[test]
    fn decodes_ptgref3d_with_sheet_prefix() {
        // Sheet1!A1 (PtgRef3d).
        let rgce = [
            0x3A, // PtgRef3d
            0x00, 0x00, // ixti=0
            0x00, 0x00, // row=0
            0x00, 0xC0, // col=A (relative)
        ];

        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!A1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptgarea3d_sheet_span_with_quoting() {
        // 'Sheet 1:Sheet 3'!A1:B2 (PtgArea3d with sheet span).
        let rgce = [
            0x3B, // PtgArea3d
            0x00, 0x00, // ixti=0
            0x00, 0x00, // rowFirst=0
            0x01, 0x00, // rowLast=1
            0x00, 0xC0, // colFirst=A
            0x01, 0xC0, // colLast=B
        ];

        let sheet_names = vec![
            "Sheet 1".to_string(),
            "Sheet2".to_string(),
            "Sheet 3".to_string(),
        ];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 2,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'Sheet 1:Sheet 3'!A1:B2");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptgname_workbook_and_sheet_scoped() {
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names = vec![
            DefinedNameMeta {
                name: "GlobalName".to_string(),
                scope_sheet: None,
            },
            DefinedNameMeta {
                name: "LocalName".to_string(),
                scope_sheet: Some(0),
            },
        ];
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce1 = [
            0x23, // PtgName
            0x01, 0x00, 0x00, 0x00, // nameId=1
            0x00, 0x00, // reserved
        ];
        let decoded1 = decode_biff8_rgce(&rgce1, &ctx);
        assert_eq!(decoded1.text, "GlobalName");
        assert!(
            decoded1.warnings.is_empty(),
            "warnings={:?}",
            decoded1.warnings
        );
        assert_parseable(&decoded1.text);

        let rgce2 = [
            0x23, // PtgName
            0x02, 0x00, 0x00, 0x00, // nameId=2
            0x00, 0x00, // reserved
        ];
        let decoded2 = decode_biff8_rgce(&rgce2, &ctx);
        assert_eq!(decoded2.text, "Sheet1!LocalName");
        assert!(
            decoded2.warnings.is_empty(),
            "warnings={:?}",
            decoded2.warnings
        );
        assert_parseable(&decoded2.text);
    }

    #[test]
    fn quotes_sheet_names_that_look_like_cell_refs() {
        // Sheet name "A1" must be quoted or it will lex as a Cell token instead of a sheet prefix.
        let rgce = [
            0x3A, // PtgRef3d
            0x00, 0x00, // ixti=0
            0x00, 0x00, // row=0
            0x00, 0xC0, // col=A (relative)
        ];

        let sheet_names = vec!["A1".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'A1'!A1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }
}
