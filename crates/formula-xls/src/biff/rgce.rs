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

// Cap the number of decode warnings produced for a single rgce stream. Corrupt token streams can
// otherwise generate an unbounded number of warnings (e.g. repeated PtgExp/PtgTbl tokens), which
// can lead to excessive memory usage and noisy UX.
const MAX_RGCE_WARNINGS: usize = 50;
const RGCE_WARNINGS_SUPPRESSED_MESSAGE: &str = "additional rgce decode warnings suppressed";

fn push_warning(warnings: &mut Vec<String>, msg: impl Into<String>, suppressed: &mut bool) {
    if *suppressed {
        return;
    }

    if warnings.len() < MAX_RGCE_WARNINGS {
        warnings.push(msg.into());
        return;
    }

    warnings.push(RGCE_WARNINGS_SUPPRESSED_MESSAGE.to_string());
    *suppressed = true;
}

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

fn namex_is_udf_call(remaining: &[u8]) -> bool {
    // Excel encodes user-defined / add-in / future functions as a `PtgFuncVar` token with
    // `iftab=0x00FF` (FTAB_USER_DEFINED), where the *function name* is the top-of-stack expression
    // (often a `PtgNameX`).
    //
    // In practice, there may be `PtgAttr` tokens between the name and `PtgFuncVar` (e.g. spacing
    // / evaluation attributes). To reliably render a parseable function identifier, treat `PtgNameX`
    // as a function name if the next non-attr token is `PtgFuncVar(0x00FF)`.
    //
    // [MS-XLS] 2.5.198.42 (PtgFuncVar) and 2.5.198.3 (PtgAttr)
    const PTG_ATTR: u8 = 0x19;
    const T_ATTR_CHOOSE: u8 = 0x04;

    let mut input = remaining;
    while let Some(&ptg) = input.first() {
        match ptg {
            PTG_ATTR => {
                // PtgAttr: [ptg][grbit: u8][wAttr: u16][optional jump table for tAttrChoose]
                if input.len() < 4 {
                    return false;
                }
                let grbit = input[1];
                let w_attr = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];

                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = (w_attr as usize).saturating_mul(2);
                    if input.len() < needed {
                        return false;
                    }
                    input = &input[needed..];
                }
            }
            // PtgFuncVar variants.
            0x22 | 0x42 | 0x62 => {
                if input.len() < 4 {
                    return false;
                }
                let func_id = u16::from_le_bytes([input[2], input[3]]);
                return func_id == formula_biff::FTAB_USER_DEFINED;
            }
            _ => return false,
        }
    }
    false
}

pub(crate) fn decode_biff8_rgce(rgce: &[u8], ctx: &RgceDecodeContext<'_>) -> DecodeRgceResult {
    decode_biff8_rgce_with_base(rgce, ctx, None)
}

pub(crate) fn decode_biff8_rgce_with_base(
    rgce: &[u8],
    ctx: &RgceDecodeContext<'_>,
    base: Option<CellCoord>,
) -> DecodeRgceResult {
    decode_biff8_rgce_with_base_and_rgcb_opt(rgce, None, ctx, base)
}

/// Decode a BIFF8 `rgce` token stream using the trailing `rgcb` data blocks referenced by certain
/// ptgs (notably `PtgArray`).
///
/// The returned text does **not** include a leading `=`.
pub(crate) fn decode_biff8_rgce_with_base_and_rgcb(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &RgceDecodeContext<'_>,
    base: Option<CellCoord>,
) -> DecodeRgceResult {
    decode_biff8_rgce_with_base_and_rgcb_opt(rgce, Some(rgcb), ctx, base)
}

fn decode_biff8_rgce_with_base_and_rgcb_opt(
    rgce: &[u8],
    rgcb: Option<&[u8]>,
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
    let mut rgcb_pos: usize = 0;
    let mut stack: Vec<ExprFragment> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();
    let mut warnings_suppressed = false;
    let mut warned_default_base_for_relative = false;

    while !input.is_empty() {
        let ptg = input[0];
        input = &input[1..];

        match ptg {
            // PtgExp / PtgTbl: shared/array formula tokens.
            //
            // These are not expected in NAME records, but can appear in the wild. We don't have
            // enough context to resolve them, so we render a parseable placeholder while still
            // consuming the fixed-size payload to keep the token stream aligned.
            //
            // Payload: 4 bytes (row/col).
            0x01 | 0x02 => {
                if input.len() < 4 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[4..];
                push_warning(
                    &mut warnings,
                    format!(
                        "encountered unsupported rgce token 0x{ptg:02X} (PtgExp/PtgTbl); rendering #UNKNOWN!"
                    ),
                    &mut warnings_suppressed,
                );
                stack.push(ExprFragment::new("#UNKNOWN!".to_string()));
            }
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    push_warning(
                        &mut warnings,
                        format!("unsupported rgce token 0x{ptg:02X}"),
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                };
                let prec = binary_precedence(ptg).expect("precedence for binary ops");

                let right = match stack.pop() {
                    Some(v) => v,
                    None => {
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
                    }
                };
                let left = match stack.pop() {
                    Some(v) => v,
                    None => {
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                    push_warning(
                        &mut warnings,
                        format!("failed to decode PtgStr: {err}"),
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
            },
            // PtgExtend / PtgExtendV / PtgExtendA (ptg=0x18 variants).
            //
            // Modern Excel can embed newer operand subtypes behind `PtgExtend` using an `etpg`
            // subtype byte. In particular, structured references (tables) use `etpg=0x19`
            // (PtgList).
            //
            // Some `.xls` files in the wild also include a 5-byte opaque token with this ptg value
            // (calamine treats it as a fixed-size payload and skips it). For compatibility, if the
            // subtype is not recognized we treat the token as a 5-byte no-op payload so the stream
            // stays aligned.
            0x18 | 0x38 | 0x58 | 0x78 => {
                let Some(&etpg) = input.first() else {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                };

                match etpg {
                    // Structured reference / table ref (PtgList).
                    0x19 => {
                        // Payload: [etpg: u8][12-byte PtgList payload]
                        if input.len() < 13 {
                            push_warning(
                                &mut warnings,
                                "unexpected end of rgce stream",
                                &mut warnings_suppressed,
                            );
                            return unsupported(ptg, warnings, &mut warnings_suppressed);
                        }
                        let mut payload = [0u8; 12];
                        payload.copy_from_slice(&input[1..13]);
                        input = &input[13..];

                        let decoded = decode_ptg_list_payload_best_effort(&payload);

                        // Interpret row/item flags. Accept unknown bits and continue decoding.
                        let flags16 = (decoded.flags & 0xFFFF) as u16;
                        const FLAG_ALL: u16 = 0x0001;
                        const FLAG_HEADERS: u16 = 0x0002;
                        const FLAG_DATA: u16 = 0x0004;
                        const FLAG_TOTALS: u16 = 0x0008;
                        const FLAG_THIS_ROW: u16 = 0x0010;
                        const KNOWN_FLAGS: u16 =
                            FLAG_ALL | FLAG_HEADERS | FLAG_DATA | FLAG_TOTALS | FLAG_THIS_ROW;
                        let unknown = flags16 & !KNOWN_FLAGS;
                        if unknown != 0 {
                            push_warning(
                                &mut warnings,
                                format!(
                                    "PtgList structured ref has unknown flags 0x{unknown:04X} (raw=0x{:08X})",
                                    decoded.flags
                                ),
                                &mut warnings_suppressed,
                            );
                        }

                        let item = if flags16 & FLAG_THIS_ROW != 0 {
                            Some(StructuredRefItem::ThisRow)
                        } else if flags16 & FLAG_HEADERS != 0 {
                            Some(StructuredRefItem::Headers)
                        } else if flags16 & FLAG_TOTALS != 0 {
                            Some(StructuredRefItem::Totals)
                        } else if flags16 & FLAG_ALL != 0 {
                            Some(StructuredRefItem::All)
                        } else if flags16 & FLAG_DATA != 0 {
                            Some(StructuredRefItem::Data)
                        } else {
                            None
                        };

                        let table_name = format!("Table{}", decoded.table_id);

                        let col_first = decoded.col_first;
                        let col_last = decoded.col_last;

                        let columns = if col_first == 0 && col_last == 0 {
                            StructuredColumns::All
                        } else if col_first == col_last {
                            StructuredColumns::Single(format!("Column{col_first}"))
                        } else {
                            StructuredColumns::Range {
                                start: format!("Column{col_first}"),
                                end: format!("Column{col_last}"),
                            }
                        };

                        let display_table_name = match item {
                            Some(StructuredRefItem::ThisRow) => None,
                            _ => Some(table_name.as_str()),
                        };

                        let mut out = format_structured_ref(display_table_name, item, &columns);

                        let mut prec = 100;
                        let is_value_class = ptg == 0x38;
                        if is_value_class && !structured_ref_is_single_cell(item, &columns) {
                            // Like value-class range/name tokens, Excel uses value-class list
                            // tokens to represent legacy implicit intersection.
                            out = format!("@{out}");
                            prec = 70;
                        }

                        stack.push(ExprFragment {
                            text: out,
                            precedence: prec,
                            contains_union: false,
                            is_missing: false,
                        });
                    }
                    _ => {
                        if input.len() < 5 {
                            push_warning(
                                &mut warnings,
                                "unexpected end of rgce stream",
                                &mut warnings_suppressed,
                            );
                            return unsupported(ptg, warnings, &mut warnings_suppressed);
                        }
                        input = &input[5..];
                        push_warning(
                            &mut warnings,
                            format!(
                                "skipped opaque 5-byte payload token 0x{ptg:02X} (Ptg18 variant) in rgce"
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                }
            }
            // PtgAttr: [grbit: u8][wAttr: u16]
            //
            // Most PtgAttr bits are evaluation hints (or formatting metadata) that do not affect
            // the printed formula text, but some do (notably tAttrSum).
            0x19 => {
                if input.len() < 3 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                        push_warning(
                            &mut warnings,
                            "unexpected end of rgce stream",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
                    }
                    input = &input[needed..];
                }

                if grbit & T_ATTR_SUM != 0 {
                    let expr = match stack.pop() {
                        Some(v) => v,
                        None => {
                            push_warning(
                                &mut warnings,
                                "rgce stack underflow",
                                &mut warnings_suppressed,
                            );
                            return unsupported(ptg, warnings, &mut warnings_suppressed);
                        }
                    };
                    stack.push(format_function_call("SUM", vec![expr]));
                }
            }
            // Error literal.
            0x1C => {
                let Some((&err, rest)) = input.split_first() else {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                    0x2C => "#SPILL!",
                    0x2D => "#CALC!",
                    0x2E => "#FIELD!",
                    0x2F => "#CONNECT!",
                    0x30 => "#BLOCKED!",
                    0x31 => "#UNKNOWN!",
                    other => {
                        push_warning(
                            &mut warnings,
                            format!("unknown error literal 0x{other:02X} in rgce"),
                            &mut warnings_suppressed,
                        );
                        "#UNKNOWN!"
                    }
                };
                stack.push(ExprFragment::new(text.to_string()));
            }
            // Bool literal.
            0x1D => {
                let Some((&b, rest)) = input.split_first() else {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                };
                input = rest;
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // Int literal.
            0x1E => {
                if input.len() < 2 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let n = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];
                stack.push(ExprFragment::new(n.to_string()));
            }
            // Num literal.
            0x1F => {
                if input.len() < 8 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&input[..8]);
                input = &input[8..];
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgArray: [unused: 7 bytes] + array constant values stored in rgcb.
            0x20 | 0x40 | 0x60 => {
                if input.len() < 7 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[7..];

                if let Some(rgcb) = rgcb {
                    match decode_array_constant(
                        rgcb,
                        &mut rgcb_pos,
                        &mut warnings,
                        &mut warnings_suppressed,
                    ) {
                        Some(arr) => stack.push(ExprFragment::new(arr)),
                        None => {
                            push_warning(
                                &mut warnings,
                                "failed to decode PtgArray constant from rgcb; rendering #UNKNOWN!",
                                &mut warnings_suppressed,
                            );
                            stack.push(ExprFragment::new("#UNKNOWN!".to_string()));
                        }
                    }
                } else {
                    // Defined-name (`NAME`) formulas do not have access to the trailing `rgcb`
                    // data blocks.
                    push_warning(
                        &mut warnings,
                        "PtgArray constant is not supported; rendering #UNKNOWN!",
                        &mut warnings_suppressed,
                    );
                    stack.push(ExprFragment::new("#UNKNOWN!".to_string()));
                }
            }
            // PtgFunc: [iftab: u16] (fixed arg count is implicit).
            0x21 | 0x41 | 0x61 => {
                if input.len() < 2 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let func_id = u16::from_le_bytes([input[0], input[1]]);
                input = &input[2..];

                if let Some(spec) = formula_biff::function_spec_from_id(func_id) {
                    // Only handle fixed arity here.
                    if spec.min_args != spec.max_args {
                        push_warning(
                            &mut warnings,
                            format!(
                                "unsupported variable-arity BIFF function id 0x{func_id:04X} (PtgFunc) in rgce"
                            ),
                            &mut warnings_suppressed,
                        );
                        // Best-effort: treat as a unary function call so the formula remains
                        // parseable.
                        let name = spec.name;
                        let mut args = Vec::new();
                        if let Some(arg) = stack.pop() {
                            args.push(arg);
                        }
                        stack.push(format_function_call(name, args));
                        continue;
                    }

                    let argc = spec.min_args as usize;
                    if stack.len() < argc {
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
                    }
                    let mut args = Vec::with_capacity(argc);
                    for _ in 0..argc {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format_function_call(spec.name, args));
                } else {
                    // Unknown function metadata: we can still emit a parseable call expression, but
                    // we don't know the argument count because `PtgFunc` is fixed-arity and does
                    // not store argc. Best-effort: assume a unary call if possible.
                    let name_owned;
                    let name = match formula_biff::function_id_to_name(func_id) {
                        Some(name) => name,
                        None => {
                            name_owned = format!("_UNKNOWN_FUNC_0x{func_id:04X}");
                            &name_owned
                        }
                    };
                    push_warning(
                        &mut warnings,
                        format!(
                            "unknown BIFF function id 0x{func_id:04X} (PtgFunc) in rgce; assuming unary"
                        ),
                        &mut warnings_suppressed,
                    );
                    let mut args = Vec::new();
                    if let Some(arg) = stack.pop() {
                        args.push(arg);
                    }
                    stack.push(format_function_call(name, args));
                }
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if input.len() < 3 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let argc = input[0] as usize;
                let func_id = u16::from_le_bytes([input[1], input[2]]);
                input = &input[3..];

                if stack.len() < argc {
                    push_warning(
                        &mut warnings,
                        "rgce stack underflow",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name, followed by args.
                if func_id == formula_biff::FTAB_USER_DEFINED {
                    if argc == 0 {
                        push_warning(
                            &mut warnings,
                            "rgce stack underflow",
                            &mut warnings_suppressed,
                        );
                        return unsupported(ptg, warnings, &mut warnings_suppressed);
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
                            push_warning(
                                &mut warnings,
                                format!(
                                    "unknown BIFF function id 0x{func_id:04X} (PtgFuncVar) in rgce"
                                ),
                                &mut warnings_suppressed,
                            );
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
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }

                let is_value_class = (ptg & 0x60) == 0x40;

                let name_id = u32::from_le_bytes([input[0], input[1], input[2], input[3]]);
                // Skip reserved bytes.
                input = &input[6..];

                let idx = name_id.saturating_sub(1) as usize;
                let Some(meta) = ctx.defined_names.get(idx) else {
                    push_warning(
                        &mut warnings,
                        format!("PtgName references missing name index {name_id}"),
                        &mut warnings_suppressed,
                    );
                    // `#NAME_ID(...)` is not a valid Excel token; fall back to a parseable error
                    // literal and keep the id in warnings.
                    stack.push(ExprFragment::new("#NAME?".to_string()));
                    continue;
                };
                if meta.name.is_empty() {
                    push_warning(
                        &mut warnings,
                        format!(
                            "PtgName references empty defined name at index {name_id} (0-based idx={idx})"
                        ),
                        &mut warnings_suppressed,
                    );
                    stack.push(ExprFragment::new("#NAME?".to_string()));
                    continue;
                }
                // `formula-engine`'s parser does not accept boolean literals (TRUE/FALSE) after a
                // sheet prefix (`Sheet1!TRUE`), since `TRUE` would be lexed as a boolean rather than
                // an identifier. If a BIFF file contains a sheet-scoped name with one of these
                // reserved keywords, fall back to a parseable name error literal.
                if meta.scope_sheet.is_some()
                    && (meta.name.eq_ignore_ascii_case("TRUE")
                        || meta.name.eq_ignore_ascii_case("FALSE")
                        || meta.name.starts_with('#')
                        || meta.name.starts_with('\''))
                {
                    push_warning(
                        &mut warnings,
                        format!(
                            "PtgName references sheet-scoped name `{}` (index {name_id}) that cannot be rendered parseably after a sheet prefix; using #NAME?",
                            meta.name
                        ),
                        &mut warnings_suppressed,
                    );
                    stack.push(ExprFragment::new("#NAME?".to_string()));
                    continue;
                }

                let text = match meta.scope_sheet {
                    None => meta.name.clone(),
                    Some(sheet_idx) => match ctx.sheet_names.get(sheet_idx) {
                        Some(sheet_name) => {
                            let sheet = quote_sheet_name_if_needed(sheet_name);
                            format!("{sheet}!{}", meta.name)
                        }
                        None => {
                            push_warning(
                                &mut warnings,
                                format!(
                                    "PtgName references sheet-scoped name `{}` with out-of-range sheet index {sheet_idx}",
                                    meta.name
                                ),
                                &mut warnings_suppressed,
                            );
                            meta.name.clone()
                        }
                    },
                };

                if is_value_class {
                    // Like value-class range tokens, a value-class name can require legacy implicit
                    // intersection (e.g. when the name refers to a multi-cell range). Emit an
                    // explicit `@` so the decoded text preserves scalar semantics.
                    stack.push(ExprFragment {
                        text: format!("@{text}"),
                        precedence: 70,
                        contains_union: false,
                        is_missing: false,
                    });
                } else {
                    stack.push(ExprFragment::new(text));
                }
            }
            // PtgNameX (external name reference).
            //
            // [MS-XLS] 2.5.198.41
            // Payload: [ixti: u16][iname: u16][reserved: u16]
            0x39 | 0x59 | 0x79 => {
                if input.len() < 6 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let iname = u16::from_le_bytes([input[2], input[3]]);
                input = &input[6..];

                let is_value_class = (ptg & 0x60) == 0x40;

                // Excel uses `PtgNameX` for external workbook defined names and for add-in/UDF
                // function calls. For UDF calls, the token sequence is typically:
                //
                //   args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
                //
                // In this case we must decode `PtgNameX` into a *function identifier* (no workbook
                // prefix or quoting), otherwise the rendered formula won't be parseable.
                let is_udf_call = namex_is_udf_call(input);

                match format_namex_ref(ixti, iname, is_udf_call, ctx) {
                    Ok(txt) => {
                        if is_value_class && !is_udf_call {
                            // Like value-class range/name tokens, a value-class NameX can represent
                            // legacy implicit intersection. Emit an explicit `@` to preserve scalar
                            // semantics.
                            stack.push(ExprFragment {
                                text: format!("@{txt}"),
                                precedence: 70,
                                contains_union: false,
                                is_missing: false,
                            });
                        } else {
                            stack.push(ExprFragment::new(txt));
                        }
                    }
                    Err(err) => {
                        push_warning(&mut warnings, err, &mut warnings_suppressed);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                    }
                }
            }
            // PtgRef (2D)
            0x24 | 0x44 | 0x64 => {
                if input.len() < 4 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let row = u16::from_le_bytes([input[0], input[1]]);
                let col = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];
                stack.push(ExprFragment::new(format_cell_ref(row, col)));
            }
            // PtgArea (2D)
            0x25 | 0x45 | 0x65 => {
                if input.len() < 8 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let row1 = u16::from_le_bytes([input[0], input[1]]);
                let row2 = u16::from_le_bytes([input[2], input[3]]);
                let col1 = u16::from_le_bytes([input[4], input[5]]);
                let col2 = u16::from_le_bytes([input[6], input[7]]);
                input = &input[8..];
                let is_single_cell =
                    row1 == row2 && (col1 & COL_INDEX_MASK) == (col2 & COL_INDEX_MASK);
                let is_value_class = (ptg & 0x60) == 0x40;

                // Prefer formatting explicit single-cell areas as cell refs. This matches Excel's
                // canonical printing and avoids emitting degenerate `A1:A1` ranges when the
                // relative flags differ between endpoints.
                let area = if is_single_cell {
                    if (col1 & RELATIVE_MASK) != (col2 & RELATIVE_MASK) {
                        push_warning(
                            &mut warnings,
                            format!(
                                "BIFF8 single-cell area has mismatched relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                    format_cell_ref(row1, col1)
                } else {
                    format_area_ref_ptg_area(
                        row1,
                        col1,
                        row2,
                        col2,
                        &mut warnings,
                        &mut warnings_suppressed,
                    )
                };

                if is_value_class && !is_single_cell {
                    // Legacy implicit intersection: Excel encodes this by using a value-class range
                    // token; modern formula text uses an explicit `@` operator.
                    stack.push(ExprFragment {
                        text: format!("@{area}"),
                        precedence: 70,
                        contains_union: false,
                        is_missing: false,
                    });
                } else {
                    stack.push(ExprFragment::new(area));
                }
            }
            // PtgMem* tokens: no-op for printing, but consume payload to keep offsets aligned.
            //
            // Payload: [cce: u16][rgce: cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if input.len() < 2 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let cce = u16::from_le_bytes([input[0], input[1]]) as usize;
                input = &input[2..];
                if input.len() < cce {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }

                // `PtgMem*` tokens contain nested subexpressions that are not directly rendered in
                // the output formula text, but their nested streams can still contain `PtgArray`
                // tokens that consume data blocks from the trailing `rgcb` stream.
                //
                // If we skip the nested stream without consuming its referenced `rgcb` blocks, any
                // later visible `PtgArray` tokens will decode against the wrong `rgcb` offset.
                if let Some(rgcb) = rgcb {
                    let sub = &input[..cce];
                    if let Err(err) = consume_rgcb_arrays_in_subexpression(
                        sub,
                        rgcb,
                        &mut rgcb_pos,
                        &mut warnings,
                        &mut warnings_suppressed,
                    ) {
                        push_warning(
                            &mut warnings,
                            format!(
                                "failed to scan PtgMem* subexpression for nested PtgArray constants: {err}"
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                }

                input = &input[cce..];
            }
            // PtgRefErr / PtgRefErrN: [rw: u16][col: u16]
            //
            // Some documentation refers to the relative-reference error tokens as `PtgRefErrN` /
            // `PtgAreaErrN`; in BIFF8 the ptg id is the same regardless of context.
            0x2A | 0x4A | 0x6A => {
                if input.len() < 4 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[4..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr / PtgAreaErrN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                if input.len() < 8 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[8..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN: [rw: u16][col: u16]
            0x2C | 0x4C | 0x6C => {
                if input.len() < 4 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let row_raw = u16::from_le_bytes([input[0], input[1]]);
                let col_field = u16::from_le_bytes([input[2], input[3]]);
                input = &input[4..];
                if base_is_default
                    && !warned_default_base_for_relative
                    && (col_field & RELATIVE_MASK) != 0
                {
                    push_warning(
                        &mut warnings,
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)",
                        &mut warnings_suppressed,
                    );
                    warned_default_base_for_relative = true;
                }
                stack.push(ExprFragment::new(decode_ref_n(
                    row_raw,
                    col_field,
                    base,
                    &mut warnings,
                    &mut warnings_suppressed,
                    "PtgRefN",
                )));
            }
            // PtgAreaN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2D | 0x4D | 0x6D => {
                if input.len() < 8 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let is_value_class = (ptg & 0x60) == 0x40;
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
                    push_warning(
                        &mut warnings,
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)",
                        &mut warnings_suppressed,
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
                    &mut warnings_suppressed,
                    "PtgAreaN",
                );
                if area == "#REF!" {
                    stack.push(ExprFragment::new(area));
                    continue;
                }

                let is_single_cell = !area.contains(':');
                if is_value_class && !is_single_cell {
                    stack.push(ExprFragment {
                        text: format!("@{area}"),
                        precedence: 70,
                        contains_union: false,
                        is_missing: false,
                    });
                } else {
                    stack.push(ExprFragment::new(area));
                }
            }
            // PtgRef3d
            0x3A | 0x5A | 0x7A => {
                if input.len() < 6 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row = u16::from_le_bytes([input[2], input[3]]);
                let col = u16::from_le_bytes([input[4], input[5]]);
                input = &input[6..];

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        push_warning(&mut warnings, err, &mut warnings_suppressed);
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
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let is_value_class = (ptg & 0x60) == 0x40;
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row1 = u16::from_le_bytes([input[2], input[3]]);
                let row2 = u16::from_le_bytes([input[4], input[5]]);
                let col1 = u16::from_le_bytes([input[6], input[7]]);
                let col2 = u16::from_le_bytes([input[8], input[9]]);
                input = &input[10..];

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        push_warning(&mut warnings, err, &mut warnings_suppressed);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
                    }
                };

                let is_single_cell =
                    row1 == row2 && (col1 & COL_INDEX_MASK) == (col2 & COL_INDEX_MASK);
                let area = if is_single_cell {
                    if (col1 & RELATIVE_MASK) != (col2 & RELATIVE_MASK) {
                        push_warning(
                            &mut warnings,
                            format!(
                                "BIFF8 3D single-cell area has mismatched relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                    format_cell_ref(row1, col1)
                } else {
                    format_area_ref_ptg_area(
                        row1,
                        col1,
                        row2,
                        col2,
                        &mut warnings,
                        &mut warnings_suppressed,
                    )
                };

                if is_value_class && !is_single_cell {
                    stack.push(ExprFragment {
                        text: format!("@{sheet_prefix}{area}"),
                        precedence: 70,
                        contains_union: false,
                        is_missing: false,
                    });
                } else {
                    stack.push(ExprFragment::new(format!("{sheet_prefix}{area}")));
                }
            }
            // PtgRefErr3d: consume payload and emit `#REF!`.
            //
            // Payload matches `PtgRef3d`: [ixti: u16][row: u16][col: u16]
            0x3C | 0x5C | 0x7C => {
                if input.len() < 6 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[6..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr3d: consume payload and emit `#REF!`.
            //
            // Payload matches `PtgArea3d`: [ixti: u16][row1: u16][row2: u16][col1: u16][col2: u16]
            0x3D | 0x5D | 0x7D => {
                if input.len() < 10 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                input = &input[10..];
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN3d (relative/absolute 3D reference): [ixti: u16][rw: u16][col: u16]
            0x3E | 0x5E | 0x7E => {
                if input.len() < 6 {
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let ixti = u16::from_le_bytes([input[0], input[1]]);
                let row_raw = u16::from_le_bytes([input[2], input[3]]);
                let col_field = u16::from_le_bytes([input[4], input[5]]);
                input = &input[6..];

                if base_is_default
                    && !warned_default_base_for_relative
                    && (col_field & RELATIVE_MASK) != 0
                {
                    push_warning(
                        &mut warnings,
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)",
                        &mut warnings_suppressed,
                    );
                    warned_default_base_for_relative = true;
                }

                let cell = decode_ref_n(
                    row_raw,
                    col_field,
                    base,
                    &mut warnings,
                    &mut warnings_suppressed,
                    "PtgRefN3d",
                );
                if cell == "#REF!" {
                    stack.push(ExprFragment::new(cell));
                    continue;
                }

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        push_warning(&mut warnings, err, &mut warnings_suppressed);
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
                    push_warning(
                        &mut warnings,
                        "unexpected end of rgce stream",
                        &mut warnings_suppressed,
                    );
                    return unsupported(ptg, warnings, &mut warnings_suppressed);
                }
                let is_value_class = (ptg & 0x60) == 0x40;
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
                    push_warning(
                        &mut warnings,
                        "relative reference tokens are interpreted relative to A1 (no base cell provided)",
                        &mut warnings_suppressed,
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
                    &mut warnings_suppressed,
                    "PtgAreaN3d",
                );
                if area == "#REF!" {
                    stack.push(ExprFragment::new(area));
                    continue;
                }

                let sheet_prefix = match format_sheet_ref(ixti, ctx) {
                    Ok(v) => v,
                    Err(err) => {
                        push_warning(&mut warnings, err, &mut warnings_suppressed);
                        stack.push(ExprFragment::new("#REF!".to_string()));
                        continue;
                    }
                };
                let is_single_cell = !area.contains(':');
                if is_value_class && !is_single_cell {
                    stack.push(ExprFragment {
                        text: format!("@{sheet_prefix}{area}"),
                        precedence: 70,
                        contains_union: false,
                        is_missing: false,
                    });
                } else {
                    stack.push(ExprFragment::new(format!("{sheet_prefix}{area}")));
                }
            }
            other => {
                push_warning(
                    &mut warnings,
                    format!("unsupported rgce token 0x{other:02X}"),
                    &mut warnings_suppressed,
                );
                return unsupported(other, warnings, &mut warnings_suppressed);
            }
        }
    }

    let text = match stack.len() {
        0 => String::new(),
        1 => stack.pop().expect("len checked").text,
        _ => {
            push_warning(
                &mut warnings,
                format!("rgce decode ended with {} expressions on stack", stack.len()),
                &mut warnings_suppressed,
            );
            stack.pop().expect("non-empty").text
        }
    };

    DecodeRgceResult { text, warnings }
}

fn decode_array_constant(
    rgcb: &[u8],
    pos: &mut usize,
    warnings: &mut Vec<String>,
    suppressed: &mut bool,
) -> Option<String> {
    // BIFF8 array constant payload stream stored as trailing `rgcb` bytes. This matches the
    // structure used by BIFF12-era formats at a high level:
    //   [cols_minus1: u16][rows_minus1: u16][values...]
    // where each value begins with a type byte:
    //   0x00 = empty
    //   0x01 = number (f64)
    //   0x02 = string ([cch: u16][utf16 chars...])
    //   0x04 = bool ([b: u8])
    //   0x10 = error ([code: u8])
    const MAX_ARRAY_CELLS: usize = 4096;

    let mut i = *pos;
    if rgcb.len().saturating_sub(i) < 4 {
        return None;
    }

    let cols_minus1 = u16::from_le_bytes([rgcb[i], rgcb[i + 1]]) as usize;
    let rows_minus1 = u16::from_le_bytes([rgcb[i + 2], rgcb[i + 3]]) as usize;
    i += 4;

    let cols = cols_minus1.saturating_add(1);
    let rows = rows_minus1.saturating_add(1);
    if cols == 0 || rows == 0 {
        return None;
    }
    if cols.saturating_mul(rows) > MAX_ARRAY_CELLS {
        push_warning(
            warnings,
            format!(
                "array constant is too large to decode (rows={rows}, cols={cols}); rendering #UNKNOWN!"
            ),
            suppressed,
        );
        return None;
    }

    let mut row_texts = Vec::with_capacity(rows);
    for _ in 0..rows {
        let mut col_texts = Vec::with_capacity(cols);
        for _ in 0..cols {
            if i >= rgcb.len() {
                return None;
            }
            let ty = rgcb[i];
            i += 1;
            match ty {
                0x00 => col_texts.push(String::new()),
                0x01 => {
                    if rgcb.len().saturating_sub(i) < 8 {
                        return None;
                    }
                    let mut bytes = [0u8; 8];
                    bytes.copy_from_slice(&rgcb[i..i + 8]);
                    i += 8;
                    col_texts.push(f64::from_le_bytes(bytes).to_string());
                }
                0x02 => {
                    if rgcb.len().saturating_sub(i) < 2 {
                        return None;
                    }
                    let cch = u16::from_le_bytes([rgcb[i], rgcb[i + 1]]) as usize;
                    i += 2;
                    let byte_len = cch.checked_mul(2)?;
                    if rgcb.len().saturating_sub(i) < byte_len {
                        return None;
                    }
                    let raw = &rgcb[i..i + byte_len];
                    i += byte_len;

                    let mut units = Vec::with_capacity(cch);
                    for chunk in raw.chunks_exact(2) {
                        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                    }
                    let s = String::from_utf16_lossy(&units);
                    let escaped = s.replace('"', "\"\"");
                    col_texts.push(format!("\"{escaped}\""));
                }
                0x04 => {
                    if rgcb.len().saturating_sub(i) < 1 {
                        return None;
                    }
                    let b = rgcb[i];
                    i += 1;
                    col_texts.push(if b == 0 { "FALSE" } else { "TRUE" }.to_string());
                }
                0x10 => {
                    if rgcb.len().saturating_sub(i) < 1 {
                        return None;
                    }
                    let code = rgcb[i];
                    i += 1;
                    let lit = match code {
                        0x00 => "#NULL!",
                        0x07 => "#DIV/0!",
                        0x0F => "#VALUE!",
                        0x17 => "#REF!",
                        0x1D => "#NAME?",
                        0x24 => "#NUM!",
                        0x2A => "#N/A",
                        0x2B => "#GETTING_DATA",
                        0x2C => "#SPILL!",
                        0x2D => "#CALC!",
                        0x2E => "#FIELD!",
                        0x2F => "#CONNECT!",
                        0x30 => "#BLOCKED!",
                        0x31 => "#UNKNOWN!",
                        _ => {
                            push_warning(
                                warnings,
                                format!("unknown error code 0x{code:02X} in array constant"),
                                suppressed,
                            );
                            "#UNKNOWN!"
                        }
                    };
                    col_texts.push(lit.to_string());
                }
                other => {
                    push_warning(
                        warnings,
                        format!("unsupported array constant element type 0x{other:02X}"),
                        suppressed,
                    );
                    return None;
                }
            }
        }
        row_texts.push(col_texts.join(","));
    }

    *pos = i;
    Some(format!("{{{}}}", row_texts.join(";")))
}

/// Scan a nested BIFF8 token subexpression (e.g. the payload of `PtgMemFunc`) and advance the
/// `rgcb` cursor for any `PtgArray` tokens encountered.
///
/// `PtgMem*` tokens are non-printing, but their nested streams can still contain `PtgArray`, which
/// consumes an array-constant block from the trailing `rgcb` stream. If we skip the nested stream
/// without consuming its referenced `rgcb` bytes, later visible `PtgArray` tokens will decode
/// against the wrong `rgcb` block.
fn consume_rgcb_arrays_in_subexpression(
    rgce: &[u8],
    rgcb: &[u8],
    rgcb_pos: &mut usize,
    warnings: &mut Vec<String>,
    suppressed: &mut bool,
) -> Result<(), String> {
    fn inner(
        input: &[u8],
        rgcb: &[u8],
        rgcb_pos: &mut usize,
        warnings: &mut Vec<String>,
        suppressed: &mut bool,
    ) -> Result<(), String> {
        let mut i = 0usize;
        while i < input.len() {
            let ptg = *input
                .get(i)
                .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
            i = i
                .checked_add(1)
                .ok_or_else(|| "rgce offset overflow".to_string())?;

            match ptg {
                // PtgExp / PtgTbl: [rw: u16][col: u16]
                0x01 | 0x02 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // Fixed-width / no-payload operators and punctuation.
                0x03..=0x16 | 0x2F => {}
                // PtgStr: ShortXLUnicodeString (variable).
                0x17 => {
                    let remaining = input
                        .get(i..)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let (_s, consumed) = strings::parse_biff8_short_string(remaining, 1252)
                        .map_err(|e| format!("failed to parse PtgStr: {e}"))?;
                    i = i
                        .checked_add(consumed)
                        .ok_or_else(|| "rgce offset overflow".to_string())?;
                    if i > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                }
                // PtgExtend* tokens: [etpg: u8][payload...]
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = *input
                        .get(i)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    i += 1;
                    match etpg {
                        0x19 => {
                            // PtgList: fixed 12-byte payload.
                            if i + 12 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 12;
                        }
                        _ => {
                            // Opaque 5-byte payload (see decoder heuristics).
                            //
                            // The ptg itself is followed by 5 bytes; since we consumed the first
                            // one as the "etpg" discriminator above, skip the remaining 4 bytes.
                            if i + 4 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 4;
                        }
                    }
                }
                // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
                0x19 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let grbit = input[i];
                    let w_attr = u16::from_le_bytes([input[i + 1], input[i + 2]]);
                    i += 3;
                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if grbit & T_ATTR_CHOOSE != 0 {
                        let needed = (w_attr as usize)
                            .checked_mul(2)
                            .ok_or_else(|| "PtgAttr jump table length overflow".to_string())?;
                        if i + needed > input.len() {
                            return Err("unexpected end of rgce stream".to_string());
                        }
                        i += needed;
                    }
                }
                // PtgErr / PtgBool: 1 byte.
                0x1C | 0x1D => {
                    if i + 1 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 1;
                }
                // PtgInt: 2 bytes.
                0x1E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgNum: 8 bytes.
                0x1F => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgArray: [unused: 7 bytes] + serialized array constant stored in rgcb.
                0x20 | 0x40 | 0x60 => {
                    if i + 7 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 7;
                    if decode_array_constant(rgcb, rgcb_pos, warnings, suppressed).is_none() {
                        return Err("failed to decode PtgArray constant from rgcb".to_string());
                    }
                }
                // PtgFunc: 2 bytes.
                0x21 | 0x41 | 0x61 => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgFuncVar: 3 bytes.
                0x22 | 0x42 | 0x62 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 3;
                }
                // PtgName: 6 bytes.
                0x23 | 0x43 | 0x63 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef: 4 bytes.
                0x24 | 0x44 | 0x64 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgArea: 8 bytes.
                0x25 | 0x45 | 0x65 => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgMem* tokens: [cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let cce = u16::from_le_bytes([input[i], input[i + 1]]) as usize;
                    i += 2;
                    let sub = input
                        .get(i..i + cce)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    inner(sub, rgcb, rgcb_pos, warnings, suppressed)?;
                    i += cce;
                }
                // PtgRefErr: 4 bytes.
                0x2A | 0x4A | 0x6A => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaErr: 8 bytes.
                0x2B | 0x4B | 0x6B => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgRefN: 4 bytes.
                0x2C | 0x4C | 0x6C => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaN: 8 bytes.
                0x2D | 0x4D | 0x6D => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgNameX: 6 bytes.
                0x39 | 0x59 | 0x79 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef3d: 6 bytes.
                0x3A | 0x5A | 0x7A => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgArea3d: 10 bytes.
                0x3B | 0x5B | 0x7B => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                // PtgRefErr3d: 6 bytes.
                0x3C | 0x5C | 0x7C => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaErr3d: 10 bytes.
                0x3D | 0x5D | 0x7D => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                // PtgRefN3d: 6 bytes.
                0x3E | 0x5E | 0x7E => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaN3d: 10 bytes.
                0x3F | 0x5F | 0x7F => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                other => {
                    // Unknown/unsupported token. We can't safely skip it without knowing its
                    // payload length. Be conservative and stop scanning so we don't desync.
                    push_warning(
                        warnings,
                        format!(
                            "unsupported rgce token 0x{other:02X} while scanning PtgMem* subexpression"
                        ),
                        suppressed,
                    );
                    return Ok(());
                }
            }
        }
        Ok(())
    }

    inner(rgce, rgcb, rgcb_pos, warnings, suppressed)
}

fn unsupported(ptg: u8, warnings: Vec<String>, suppressed: &mut bool) -> DecodeRgceResult {
    let mut warnings = warnings;
    let msg = format!("unsupported rgce token 0x{ptg:02X}");
    if !warnings.iter().any(|w| w == &msg) {
        push_warning(&mut warnings, msg, suppressed);
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
    suppressed: &mut bool,
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
        push_warning(
            warnings,
            format!(
                "{ptg_name} produced out-of-bounds reference: base=({},{}), {row_desc}, {col_desc} -> #REF!",
                base.row, base.col
            ),
            suppressed,
        );
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
    suppressed: &mut bool,
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
        push_warning(
            warnings,
            format!(
                "{ptg_name} produced out-of-bounds area: base=({},{}), {row1_desc}, {row2_desc}, {col1_desc}, {col2_desc} -> #REF!",
                base.row, base.col
            ),
            suppressed,
        );
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

// -------------------------------------------------------------------------------------------------
// BIFF8 shared formula materialization
// -------------------------------------------------------------------------------------------------

/// Analysis of a BIFF8 shared-formula `rgce` stream.
#[derive(Debug, Clone, Copy, Default)]
pub(crate) struct Biff8SharedFormulaRgceAnalysis {
    /// Whether the rgce contains any relative-offset ptgs (`PtgRefN`/`PtgAreaN`, including 3D variants).
    pub(crate) has_refn_or_arean: bool,
    /// Whether the rgce contains any absolute reference ptgs (`PtgRef`/`PtgArea`/`PtgRef3d`/`PtgArea3d`)
    /// that have row/col-relative flags set.
    pub(crate) has_abs_refs_with_relative_flags: bool,
}

/// Best-effort scan of an `rgce` stream to detect token classes relevant for shared-formula
/// materialization.
///
/// This is used to decide when to materialize a shared formula definition (`SHRFMLA.rgce`) before
/// decoding it for a follower cell.
pub(crate) fn analyze_biff8_shared_formula_rgce(
    rgce: &[u8],
) -> Result<Biff8SharedFormulaRgceAnalysis, String> {
    fn inner(input: &[u8], out: &mut Biff8SharedFormulaRgceAnalysis) -> Result<(), String> {
        let mut i = 0usize;
        while i < input.len() {
            let ptg = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
            i = i.saturating_add(1);

            match ptg {
                // PtgExp / PtgTbl: [rw: u16][col: u16]
                0x01 | 0x02 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // Fixed-width / no-payload operators and punctuation.
                0x03..=0x16 | 0x2F => {}
                // PtgStr: ShortXLUnicodeString (variable).
                0x17 => {
                    let remaining = input.get(i..).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let (_s, consumed) = strings::parse_biff8_short_string(remaining, 1252)
                        .map_err(|e| format!("failed to parse PtgStr: {e}"))?;
                    i = i
                        .checked_add(consumed)
                        .ok_or_else(|| "rgce offset overflow".to_string())?;
                    if i > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                }
                // PtgExtend* tokens (ptg=0x18 variants): [etpg: u8][payload...]
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    i += 1;
                    match etpg {
                        0x19 => {
                            // PtgList: fixed 12-byte payload.
                            if i + 12 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 12;
                        }
                        _ => {
                            // Opaque 5-byte payload (see decoder heuristics).
                            //
                            // The ptg itself is followed by 5 bytes; since we consumed the first
                            // one as the "etpg" discriminator above, skip the remaining 4 bytes.
                            if i + 4 > input.len() {
                                return Err("unexpected end of rgce stream".to_string());
                            }
                            i += 4;
                        }
                    }
                }
                // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
                0x19 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let grbit = input[i];
                    let w_attr = u16::from_le_bytes([input[i + 1], input[i + 2]]);
                    i += 3;
                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if grbit & T_ATTR_CHOOSE != 0 {
                        let needed = (w_attr as usize)
                            .checked_mul(2)
                            .ok_or_else(|| "PtgAttr jump table length overflow".to_string())?;
                        if i + needed > input.len() {
                            return Err("unexpected end of rgce stream".to_string());
                        }
                        i += needed;
                    }
                }
                // PtgErr / PtgBool: 1 byte.
                0x1C | 0x1D => {
                    if i + 1 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 1;
                }
                // PtgInt: 2 bytes.
                0x1E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgNum: 8 bytes.
                0x1F => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgArray: 7 bytes.
                0x20 | 0x40 | 0x60 => {
                    if i + 7 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 7;
                }
                // PtgFunc: 2 bytes.
                0x21 | 0x41 | 0x61 => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 2;
                }
                // PtgFuncVar: 3 bytes.
                0x22 | 0x42 | 0x62 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 3;
                }
                // PtgName: 6 bytes.
                0x23 | 0x43 | 0x63 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef: [rw: u16][col: u16]
                0x24 | 0x44 | 0x64 => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let col_field = u16::from_le_bytes([input[i + 2], input[i + 3]]);
                    if (col_field & RELATIVE_MASK) != 0 {
                        out.has_abs_refs_with_relative_flags = true;
                    }
                    i += 4;
                }
                // PtgArea: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
                0x25 | 0x45 | 0x65 => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let col1 = u16::from_le_bytes([input[i + 4], input[i + 5]]);
                    let col2 = u16::from_le_bytes([input[i + 6], input[i + 7]]);
                    if (col1 & RELATIVE_MASK) != 0 || (col2 & RELATIVE_MASK) != 0 {
                        out.has_abs_refs_with_relative_flags = true;
                    }
                    i += 8;
                }
                // PtgMem* tokens: [cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let cce = u16::from_le_bytes([input[i], input[i + 1]]) as usize;
                    i += 2;
                    let sub = input
                        .get(i..i + cce)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    inner(sub, out)?;
                    i += cce;
                }
                // PtgRefErr: 4 bytes.
                0x2A | 0x4A | 0x6A => {
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaErr: 8 bytes.
                0x2B | 0x4B | 0x6B => {
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgRefN: 4 bytes.
                0x2C | 0x4C | 0x6C => {
                    out.has_refn_or_arean = true;
                    if i + 4 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 4;
                }
                // PtgAreaN: 8 bytes.
                0x2D | 0x4D | 0x6D => {
                    out.has_refn_or_arean = true;
                    if i + 8 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 8;
                }
                // PtgNameX: 6 bytes.
                0x39 | 0x59 | 0x79 => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgRef3d: [ixti: u16][rw: u16][col: u16]
                0x3A | 0x5A | 0x7A => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let col_field = u16::from_le_bytes([input[i + 4], input[i + 5]]);
                    if (col_field & RELATIVE_MASK) != 0 {
                        out.has_abs_refs_with_relative_flags = true;
                    }
                    i += 6;
                }
                // PtgArea3d: [ixti: u16][rw1: u16][rw2: u16][col1: u16][col2: u16]
                0x3B | 0x5B | 0x7B => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let col1 = u16::from_le_bytes([input[i + 6], input[i + 7]]);
                    let col2 = u16::from_le_bytes([input[i + 8], input[i + 9]]);
                    if (col1 & RELATIVE_MASK) != 0 || (col2 & RELATIVE_MASK) != 0 {
                        out.has_abs_refs_with_relative_flags = true;
                    }
                    i += 10;
                }
                // PtgRefErr3d: 6 bytes.
                0x3C | 0x5C | 0x7C => {
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaErr3d: 10 bytes.
                0x3D | 0x5D | 0x7D => {
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                // PtgRefN3d: 6 bytes.
                0x3E | 0x5E | 0x7E => {
                    out.has_refn_or_arean = true;
                    if i + 6 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 6;
                }
                // PtgAreaN3d: 10 bytes.
                0x3F | 0x5F | 0x7F => {
                    out.has_refn_or_arean = true;
                    if i + 10 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    i += 10;
                }
                other => {
                    // Unknown/unsupported token. We can't safely skip it without knowing its payload
                    // length, but analysis is only a heuristic. Be conservative and assume the
                    // shared formula might need materialization, then stop scanning.
                    out.has_abs_refs_with_relative_flags = true;
                    let _ = other;
                    return Ok(());
                }
            }
        }
        Ok(())
    }

    let mut out = Biff8SharedFormulaRgceAnalysis::default();
    inner(rgce, &mut out)?;
    Ok(out)
}

/// Materialize a BIFF8 shared formula `rgce` stream for a follower cell.
///
/// Some BIFF8 writers store shared formula definitions using absolute reference ptgs (`PtgRef`,
/// `PtgArea`, `PtgRef3d`, `PtgArea3d`) with the row/col-relative bits set. When Excel fills the
/// shared formula across its range, those stored coordinates are shifted by `(delta_row, delta_col)`
/// for each follower cell. This helper implements that shifting, producing a cell-specific `rgce`.
///
/// When a shifted coordinate is out of bounds, the corresponding token is converted to a `PtgRefErr*`
/// / `PtgAreaErr*` error token (preserving the ptg class), matching Excel's behavior.
pub(crate) fn materialize_biff8_shared_formula_rgce(
    base_rgce: &[u8],
    base_cell: CellCoord,
    target_cell: CellCoord,
) -> Result<Vec<u8>, String> {
    const MAX_ROW: i64 = u16::MAX as i64;
    const MAX_COL: i64 = BIFF8_MAX_COL0;

    let delta_row = target_cell.row as i64 - base_cell.row as i64;
    let delta_col = target_cell.col as i64 - base_cell.col as i64;

    fn inner(input: &[u8], delta_row: i64, delta_col: i64) -> Result<Vec<u8>, String> {
        let mut out = Vec::with_capacity(input.len());
        let mut i = 0usize;
        while i < input.len() {
            let ptg = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
            i += 1;

            match ptg {
                // PtgExp / PtgTbl: [rw: u16][col: u16]
                0x01 | 0x02 => {
                    let payload = input
                        .get(i..i + 4)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 4;
                }
                // Fixed-width / no-payload operators and punctuation.
                0x03..=0x16 | 0x2F => out.push(ptg),
                // PtgStr (ShortXLUnicodeString).
                0x17 => {
                    let remaining = input
                        .get(i..)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let (_s, consumed) = strings::parse_biff8_short_string(remaining, 1252)
                        .map_err(|e| format!("failed to parse PtgStr: {e}"))?;
                    let payload = input
                        .get(i..i + consumed)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += consumed;
                }
                // PtgExtend / PtgExtendV / PtgExtendA (+0x60): [etpg: u8][payload...]
                0x18 | 0x38 | 0x58 | 0x78 => {
                    let etpg = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    i += 1;
                    out.push(ptg);
                    out.push(etpg);
                    match etpg {
                        0x19 => {
                            let payload = input
                                .get(i..i + 12)
                                .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                            out.extend_from_slice(payload);
                            i += 12;
                        }
                        _ => {
                            let payload = input
                                // Like the rgce decoder, treat unknown Ptg18 variants as an opaque
                                // 5-byte payload following the ptg (including the subtype byte).
                                // Since we already consumed `etpg`, only 4 bytes remain.
                                .get(i..i + 4)
                                .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                            out.extend_from_slice(payload);
                            i += 4;
                        }
                    }
                }
                // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table.
                0x19 => {
                    if i + 3 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let grbit = input[i];
                    let w_attr = u16::from_le_bytes([input[i + 1], input[i + 2]]);
                    out.push(ptg);
                    out.extend_from_slice(&input[i..i + 3]);
                    i += 3;

                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if grbit & T_ATTR_CHOOSE != 0 {
                        let needed = (w_attr as usize)
                            .checked_mul(2)
                            .ok_or_else(|| "PtgAttr jump table length overflow".to_string())?;
                        let payload = input
                            .get(i..i + needed)
                            .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                        out.extend_from_slice(payload);
                        i += needed;
                    }
                }
                // PtgErr / PtgBool: 1 byte.
                0x1C | 0x1D => {
                    let b = *input.get(i).ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.push(b);
                    i += 1;
                }
                // PtgInt: 2 bytes.
                0x1E => {
                    let payload = input
                        .get(i..i + 2)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 2;
                }
                // PtgNum: 8 bytes.
                0x1F => {
                    let payload = input
                        .get(i..i + 8)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 8;
                }
                // PtgArray: [unused: 7 bytes] (array data in rgcb).
                0x20 | 0x40 | 0x60 => {
                    let payload = input
                        .get(i..i + 7)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 7;
                }
                // PtgFunc: [iftab: u16]
                0x21 | 0x41 | 0x61 => {
                    let payload = input
                        .get(i..i + 2)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 2;
                }
                // PtgFuncVar: [argc: u8][iftab: u16]
                0x22 | 0x42 | 0x62 => {
                    let payload = input
                        .get(i..i + 3)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 3;
                }
                // PtgName: [nameId: u32][reserved: u16]
                0x23 | 0x43 | 0x63 => {
                    let payload = input
                        .get(i..i + 6)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 6;
                }
                // PtgNameX: [ixti: u16][iname: u16][reserved: u16]
                0x39 | 0x59 | 0x79 => {
                    let payload = input
                        .get(i..i + 6)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 6;
                }
                // PtgRef: [rw: u16][col: u16]
                0x24 | 0x44 | 0x64 => {
                    let payload = input
                        .get(i..i + 4)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let row_raw = u16::from_le_bytes([payload[0], payload[1]]) as i64;
                    let col_field = u16::from_le_bytes([payload[2], payload[3]]);
                    let col_raw = (col_field & COL_INDEX_MASK) as i64;
                    let row_rel = (col_field & ROW_RELATIVE_BIT) != 0;
                    let col_rel = (col_field & COL_RELATIVE_BIT) != 0;

                    let new_row = if row_rel { row_raw + delta_row } else { row_raw };
                    let new_col = if col_rel { col_raw + delta_col } else { col_raw };

                    if new_row < 0 || new_row > MAX_ROW || new_col < 0 || new_col > MAX_COL {
                        out.push(ptg.saturating_add(0x06)); // PtgRef* -> PtgRefErr*
                        out.extend_from_slice(payload);
                        i += 4;
                        continue;
                    }

                    let new_col_field =
                        ((new_col as u16) & COL_INDEX_MASK) | (col_field & RELATIVE_MASK);
                    out.push(ptg);
                    out.extend_from_slice(&(new_row as u16).to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                    i += 4;
                }
                // PtgArea: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
                0x25 | 0x45 | 0x65 => {
                    let payload = input
                        .get(i..i + 8)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let row1_raw = u16::from_le_bytes([payload[0], payload[1]]) as i64;
                    let row2_raw = u16::from_le_bytes([payload[2], payload[3]]) as i64;
                    let col1_field = u16::from_le_bytes([payload[4], payload[5]]);
                    let col2_field = u16::from_le_bytes([payload[6], payload[7]]);

                    let col1_raw = (col1_field & COL_INDEX_MASK) as i64;
                    let col2_raw = (col2_field & COL_INDEX_MASK) as i64;

                    let row1_rel = (col1_field & ROW_RELATIVE_BIT) != 0;
                    let col1_rel = (col1_field & COL_RELATIVE_BIT) != 0;
                    let row2_rel = (col2_field & ROW_RELATIVE_BIT) != 0;
                    let col2_rel = (col2_field & COL_RELATIVE_BIT) != 0;

                    let new_row1 = if row1_rel { row1_raw + delta_row } else { row1_raw };
                    let new_col1 = if col1_rel { col1_raw + delta_col } else { col1_raw };
                    let new_row2 = if row2_rel { row2_raw + delta_row } else { row2_raw };
                    let new_col2 = if col2_rel { col2_raw + delta_col } else { col2_raw };

                    if new_row1 < 0
                        || new_row1 > MAX_ROW
                        || new_row2 < 0
                        || new_row2 > MAX_ROW
                        || new_col1 < 0
                        || new_col1 > MAX_COL
                        || new_col2 < 0
                        || new_col2 > MAX_COL
                    {
                        out.push(ptg.saturating_add(0x06)); // PtgArea* -> PtgAreaErr*
                        out.extend_from_slice(payload);
                        i += 8;
                        continue;
                    }

                    let new_col1_field =
                        ((new_col1 as u16) & COL_INDEX_MASK) | (col1_field & RELATIVE_MASK);
                    let new_col2_field =
                        ((new_col2 as u16) & COL_INDEX_MASK) | (col2_field & RELATIVE_MASK);

                    out.push(ptg);
                    out.extend_from_slice(&(new_row1 as u16).to_le_bytes());
                    out.extend_from_slice(&(new_row2 as u16).to_le_bytes());
                    out.extend_from_slice(&new_col1_field.to_le_bytes());
                    out.extend_from_slice(&new_col2_field.to_le_bytes());
                    i += 8;
                }
                // PtgMem* tokens: [cce: u16][rgce: cce bytes]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    if i + 2 > input.len() {
                        return Err("unexpected end of rgce stream".to_string());
                    }
                    let cce = u16::from_le_bytes([input[i], input[i + 1]]) as usize;
                    i += 2;
                    let sub = input
                        .get(i..i + cce)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let materialized = inner(sub, delta_row, delta_col)?;

                    out.push(ptg);
                    out.extend_from_slice(&(cce as u16).to_le_bytes());
                    out.extend_from_slice(&materialized);
                    i += cce;
                }
                // PtgRefErr: [rw: u16][col: u16]
                0x2A | 0x4A | 0x6A => {
                    let payload = input
                        .get(i..i + 4)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 4;
                }
                // PtgAreaErr: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
                0x2B | 0x4B | 0x6B => {
                    let payload = input
                        .get(i..i + 8)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 8;
                }
                // PtgRefN: relative offsets; copy through.
                0x2C | 0x4C | 0x6C => {
                    let payload = input
                        .get(i..i + 4)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 4;
                }
                // PtgAreaN: relative offsets; copy through.
                0x2D | 0x4D | 0x6D => {
                    let payload = input
                        .get(i..i + 8)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 8;
                }
                // PtgRef3d: [ixti: u16][rw: u16][col: u16]
                0x3A | 0x5A | 0x7A => {
                    let payload = input
                        .get(i..i + 6)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let ixti = u16::from_le_bytes([payload[0], payload[1]]);
                    let row_raw = u16::from_le_bytes([payload[2], payload[3]]) as i64;
                    let col_field = u16::from_le_bytes([payload[4], payload[5]]);
                    let col_raw = (col_field & COL_INDEX_MASK) as i64;
                    let row_rel = (col_field & ROW_RELATIVE_BIT) != 0;
                    let col_rel = (col_field & COL_RELATIVE_BIT) != 0;

                    let new_row = if row_rel { row_raw + delta_row } else { row_raw };
                    let new_col = if col_rel { col_raw + delta_col } else { col_raw };

                    if new_row < 0 || new_row > MAX_ROW || new_col < 0 || new_col > MAX_COL {
                        out.push(ptg.saturating_add(0x02)); // PtgRef3d* -> PtgRefErr3d*
                        out.extend_from_slice(payload);
                        i += 6;
                        continue;
                    }

                    let new_col_field =
                        ((new_col as u16) & COL_INDEX_MASK) | (col_field & RELATIVE_MASK);
                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&(new_row as u16).to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                    i += 6;
                }
                // PtgArea3d: [ixti: u16][rw1: u16][rw2: u16][col1: u16][col2: u16]
                0x3B | 0x5B | 0x7B => {
                    let payload = input
                        .get(i..i + 10)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    let ixti = u16::from_le_bytes([payload[0], payload[1]]);
                    let row1_raw = u16::from_le_bytes([payload[2], payload[3]]) as i64;
                    let row2_raw = u16::from_le_bytes([payload[4], payload[5]]) as i64;
                    let col1_field = u16::from_le_bytes([payload[6], payload[7]]);
                    let col2_field = u16::from_le_bytes([payload[8], payload[9]]);

                    let col1_raw = (col1_field & COL_INDEX_MASK) as i64;
                    let col2_raw = (col2_field & COL_INDEX_MASK) as i64;

                    let row1_rel = (col1_field & ROW_RELATIVE_BIT) != 0;
                    let col1_rel = (col1_field & COL_RELATIVE_BIT) != 0;
                    let row2_rel = (col2_field & ROW_RELATIVE_BIT) != 0;
                    let col2_rel = (col2_field & COL_RELATIVE_BIT) != 0;

                    let new_row1 = if row1_rel { row1_raw + delta_row } else { row1_raw };
                    let new_col1 = if col1_rel { col1_raw + delta_col } else { col1_raw };
                    let new_row2 = if row2_rel { row2_raw + delta_row } else { row2_raw };
                    let new_col2 = if col2_rel { col2_raw + delta_col } else { col2_raw };

                    if new_row1 < 0
                        || new_row1 > MAX_ROW
                        || new_row2 < 0
                        || new_row2 > MAX_ROW
                        || new_col1 < 0
                        || new_col1 > MAX_COL
                        || new_col2 < 0
                        || new_col2 > MAX_COL
                    {
                        out.push(ptg.saturating_add(0x02)); // PtgArea3d* -> PtgAreaErr3d*
                        out.extend_from_slice(payload);
                        i += 10;
                        continue;
                    }

                    let new_col1_field =
                        ((new_col1 as u16) & COL_INDEX_MASK) | (col1_field & RELATIVE_MASK);
                    let new_col2_field =
                        ((new_col2 as u16) & COL_INDEX_MASK) | (col2_field & RELATIVE_MASK);

                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&(new_row1 as u16).to_le_bytes());
                    out.extend_from_slice(&(new_row2 as u16).to_le_bytes());
                    out.extend_from_slice(&new_col1_field.to_le_bytes());
                    out.extend_from_slice(&new_col2_field.to_le_bytes());
                    i += 10;
                }
                // PtgRefErr3d: [ixti: u16][rw: u16][col: u16]
                0x3C | 0x5C | 0x7C => {
                    let payload = input
                        .get(i..i + 6)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 6;
                }
                // PtgAreaErr3d: [ixti: u16][rw1: u16][rw2: u16][col1: u16][col2: u16]
                0x3D | 0x5D | 0x7D => {
                    let payload = input
                        .get(i..i + 10)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 10;
                }
                // PtgRefN3d: relative offsets; copy through.
                0x3E | 0x5E | 0x7E => {
                    let payload = input
                        .get(i..i + 6)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 6;
                }
                // PtgAreaN3d: relative offsets; copy through.
                0x3F | 0x5F | 0x7F => {
                    let payload = input
                        .get(i..i + 10)
                        .ok_or_else(|| "unexpected end of rgce stream".to_string())?;
                    out.push(ptg);
                    out.extend_from_slice(payload);
                    i += 10;
                }
                other => {
                    return Err(format!(
                        "unsupported rgce token 0x{other:02X} while materializing shared formula rgce"
                    ));
                }
            }
        }

        Ok(out)
    }

    let out = inner(base_rgce, delta_row, delta_col)?;
    Ok(out)
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
    suppressed: &mut bool,
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
        push_warning(
            warnings,
            format!(
                "BIFF8 area matches both whole-row and whole-column patterns (rwFirst={row1}, rwLast={row2}, colFirst=0x{col1:04X}, colLast=0x{col2:04X}); rendering as explicit A1-style area"
            ),
            suppressed,
        );
        return format_area_ref(row1, col1, row2, col2);
    }

    if is_whole_row {
        let row_rel = if row1_rel != row2_rel {
            push_warning(
                warnings,
                format!(
                    "BIFF8 whole-row area has mismatched row-relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
                ),
                suppressed,
            );
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
            push_warning(
                warnings,
                format!(
                    "BIFF8 whole-column area has mismatched col-relative flags (colFirst=0x{col1:04X}, colLast=0x{col2:04X}); using first"
                ),
                suppressed,
            );
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
    if extern_name.is_empty() {
        return Err(format!(
            "PtgNameX references empty EXTERNNAME iname={iname} for SUPBOOK index {supbook_index} (ixti={ixti})"
        ));
    }

    if is_function {
        return Ok(extern_name.clone());
    }

    match sb.kind {
        // Internal workbook EXTERNNAME (rare, but some producers use PtgNameX even for internal
        // sheet-scoped names). If a usable internal sheet prefix is available via EXTERNSHEET,
        // include it so the rendered formula matches Excel’s `Sheet1!Name` style.
        SupBookKind::Internal => {
            if sheet_ref_available {
                if let Ok(prefix) = format_sheet_ref(ixti, ctx) {
                    if extern_name.eq_ignore_ascii_case("TRUE")
                        || extern_name.eq_ignore_ascii_case("FALSE")
                        || extern_name.starts_with('#')
                        || extern_name.starts_with('\'')
                    {
                        return Err(format!(
                            "PtgNameX external name `{extern_name}` cannot be rendered parseably after a sheet prefix (ixti={ixti}, iname={iname})"
                        ));
                    }
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
                    if extern_name.eq_ignore_ascii_case("TRUE")
                        || extern_name.eq_ignore_ascii_case("FALSE")
                        || extern_name.starts_with('#')
                        || extern_name.starts_with('\'')
                    {
                        return Err(format!(
                            "PtgNameX external name `{extern_name}` cannot be rendered parseably after a sheet prefix (ixti={ixti}, iname={iname})"
                        ));
                    }
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
    // Some producers may include an absolute/relative path in the SUPBOOK virtPath. For formula
    // rendering we want a best-effort workbook *basename*.
    //
    // Some producers may also already include brackets; normalize to a single set.
    //
    // Workbook names may contain literal `[` characters (non-nesting), and literal `]` characters
    // must be escaped in Excel formula text as `]]`. Ensure the formatted workbook prefix is both
    // parseable and round-trips through our formula engine.
    let without_nuls = workbook.replace('\0', "");
    let trimmed_full = without_nuls.trim();
    let has_wrapper_brackets = trimmed_full.starts_with('[') && trimmed_full.ends_with(']');

    let basename = trimmed_full
        .rsplit(['\\', '/'])
        .next()
        .unwrap_or(trimmed_full);
    let mut inner = basename.trim();

    // Be permissive about bracket placement: we sometimes see values like:
    // - "[Book.xlsx]"
    // - "Book.xlsx"
    // - "[C:\\path\\Book.xlsx]" (bracketed full path; the `[` is lost when taking the basename)
    //
    // Only strip a leading/trailing bracket when the *full* input appears to be wrapped, so we
    // don't drop legitimate `[` / `]` characters that are part of the workbook name.
    if has_wrapper_brackets {
        if let Some(stripped) = inner.strip_prefix('[') {
            inner = stripped;
        }
        if let Some(stripped) = inner.strip_suffix(']') {
            inner = stripped;
        }
    }

    // Escape literal `]` characters inside the workbook name by doubling them (`]]`), matching
    // Excel's external workbook prefix syntax.
    let escaped = if inner.contains(']') {
        inner.replace(']', "]]")
    } else {
        inner.to_string()
    };
    format!("[{escaped}]")
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StructuredRefItem {
    All,
    Data,
    Headers,
    Totals,
    ThisRow,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum StructuredColumns {
    All,
    Single(String),
    Range { start: String, end: String },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct PtgListDecoded {
    table_id: u32,
    flags: u32,
    col_first: u32,
    col_last: u32,
}

fn decode_ptg_list_payload_best_effort(payload: &[u8; 12]) -> PtgListDecoded {
    // There are multiple observed encodings for the 12-byte PtgList payload (table refs /
    // structured references). Try a handful of plausible layouts and prefer the one that looks
    // most consistent without workbook table metadata.
    //
    // Layout A (u32 + 4*u16):
    //   [table_id: u32][flags: u16][col_first: u16][col_last: u16][reserved: u16]
    //
    // Layout B (u32 + 2*u32):
    //   [table_id: u32][col_first_raw: u32][col_last_raw: u32]
    //   where `col_first_raw` packs `[col_first: u16][flags: u16]` (and `col_last_raw` packs
    //   `[col_last: u16][reserved: u16]`).
    //
    // Layout C (3*u32):
    //   [table_id: u32][flags: u32][col_spec: u32]
    //   where `col_spec` packs `[col_first: u16][col_last: u16]`.
    let table_id = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);

    let flags_a = u16::from_le_bytes([payload[4], payload[5]]) as u32;
    let col_first_a = u16::from_le_bytes([payload[6], payload[7]]) as u32;
    let col_last_a = u16::from_le_bytes([payload[8], payload[9]]) as u32;

    let col_first_raw = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
    let col_last_raw = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
    let col_first_b = (col_first_raw & 0xFFFF) as u32;
    let flags_b = (col_first_raw >> 16) & 0xFFFF;
    let col_last_b = (col_last_raw & 0xFFFF) as u32;

    let flags_c = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
    let col_spec_c = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
    let col_first_c = (col_spec_c & 0xFFFF) as u32;
    let col_last_c = ((col_spec_c >> 16) & 0xFFFF) as u32;

    let candidates = [
        PtgListDecoded {
            table_id,
            flags: flags_a,
            col_first: col_first_a,
            col_last: col_last_a,
        },
        PtgListDecoded {
            table_id,
            flags: flags_b,
            col_first: col_first_b,
            col_last: col_last_b,
        },
        PtgListDecoded {
            table_id,
            flags: flags_c,
            col_first: col_first_c,
            col_last: col_last_c,
        },
        // Layout D: treat the middle/end u32s as raw column ids with no separate flags.
        PtgListDecoded {
            table_id,
            flags: 0,
            col_first: col_first_raw,
            col_last: col_last_raw,
        },
    ];

    *candidates
        .iter()
        .max_by_key(|cand| score_ptg_list_candidate(cand))
        .expect("non-empty")
}

fn score_ptg_list_candidate(cand: &PtgListDecoded) -> i32 {
    let mut score = 0i32;

    // Table ids are typically non-zero and small.
    if cand.table_id != 0 {
        score += 1;
    }

    let col_first = cand.col_first;
    let col_last = cand.col_last;

    // Column id `0` is treated as a sentinel for "all columns"; seeing it on only one side is
    // usually a sign we've chosen the wrong payload layout.
    if (col_first == 0) ^ (col_last == 0) {
        score -= 50;
    }

    // Most table column ids are small-ish.
    if col_first <= 16_384 {
        score += 1;
    }
    if col_last <= 16_384 {
        score += 1;
    }

    // Prefer flags that fit in the lower 16 bits (where the documented item bits live).
    if cand.flags <= 0xFFFF {
        score += 1;
    }

    score
}

fn structured_ref_is_single_cell(item: Option<StructuredRefItem>, columns: &StructuredColumns) -> bool {
    match (item, columns) {
        (Some(StructuredRefItem::ThisRow), StructuredColumns::Single(_)) => true,
        (Some(StructuredRefItem::Headers), StructuredColumns::Single(_)) => true,
        (Some(StructuredRefItem::Totals), StructuredColumns::Single(_)) => true,
        _ => false,
    }
}

fn format_structured_ref(
    table_name: Option<&str>,
    item: Option<StructuredRefItem>,
    columns: &StructuredColumns,
) -> String {
    // This-row shorthand: `[@Col]`, `[@]`, and `[@[Col1]:[Col2]]`.
    if matches!(item, Some(StructuredRefItem::ThisRow)) {
        match columns {
            StructuredColumns::Single(col) => {
                return format!("[@{}]", escape_structured_ref_bracket_content(col));
            }
            StructuredColumns::All => return "[@]".to_string(),
            StructuredColumns::Range { start, end } => {
                let start = escape_structured_ref_bracket_content(start);
                let end = escape_structured_ref_bracket_content(end);
                return format!("[@[{start}]:[{end}]]");
            }
        }
    }

    let table = table_name.unwrap_or("");

    // Item-only selections: `Table1[#All]`, `Table1[#Headers]`, etc.
    if columns == &StructuredColumns::All {
        if let Some(item) = item {
            return format!("{table}[{}]", structured_ref_item_literal(item));
        }
        // Default row selector with no column selection: treat as `[#Data]`.
        return format!("{table}[#Data]");
    }

    // Single-column selection with default/data item: `Table1[Col]`.
    if matches!(item, None | Some(StructuredRefItem::Data)) {
        match columns {
            StructuredColumns::Single(col) => {
                return format!("{table}[{}]", escape_structured_ref_bracket_content(col));
            }
            StructuredColumns::Range { start, end } => {
                let start = escape_structured_ref_bracket_content(start);
                let end = escape_structured_ref_bracket_content(end);
                return format!("{table}[[{start}]:[{end}]]");
            }
            StructuredColumns::All => {}
        }
    }

    // General nested form: `Table1[[#Headers],[Col]]` or `Table1[[#Headers],[Col1]:[Col2]]`.
    let item = item.expect("handled None above");
    match columns {
        StructuredColumns::Single(col) => {
            let col = escape_structured_ref_bracket_content(col);
            format!("{table}[[{}],[{col}]]", structured_ref_item_literal(item))
        }
        StructuredColumns::Range { start, end } => {
            let start = escape_structured_ref_bracket_content(start);
            let end = escape_structured_ref_bracket_content(end);
            format!(
                "{table}[[{}],[{start}]:[{end}]]",
                structured_ref_item_literal(item)
            )
        }
        StructuredColumns::All => unreachable!("handled above"),
    }
}

fn structured_ref_item_literal(item: StructuredRefItem) -> &'static str {
    match item {
        StructuredRefItem::All => "#All",
        StructuredRefItem::Data => "#Data",
        StructuredRefItem::Headers => "#Headers",
        StructuredRefItem::Totals => "#Totals",
        StructuredRefItem::ThisRow => "#This Row",
    }
}

fn escape_structured_ref_bracket_content(s: &str) -> String {
    if !s.contains(']') {
        return s.to_string();
    }
    s.replace(']', "]]")
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
        let ast = parse_formula(expr, ParseOptions::default()).unwrap_or_else(|err| {
            panic!(
                "expected decoded expression to be parseable, expr={expr:?}, err={err:?}"
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
        let mut suppressed = false;
        crate::parse_print_area_refers_to(sheet_name, expr, &mut warnings, &mut suppressed)
            .expect("parse print area defined name");
    }

    fn assert_print_titles_parseable(sheet_name: &str, expr: &str) {
        let mut warnings = Vec::<crate::ImportWarning>::new();
        let mut suppressed = false;
        crate::parse_print_titles_refers_to(sheet_name, expr, &mut warnings, &mut suppressed)
            .expect("parse print titles defined name");
    }

    #[test]
    fn print_area_parsing_accepts_explicit_implicit_intersection() {
        assert_print_area_parseable("Sheet1", "@Sheet1!$A$1:$B$2");
    }

    #[test]
    fn print_titles_parsing_accepts_explicit_implicit_intersection() {
        assert_print_titles_parseable("Sheet1", "@Sheet1!$1:$1");
    }

    #[test]
    fn formats_cell_ref_no_dollars() {
        assert_eq!(format_cell_ref_no_dollars(0, 0), "A1");
        assert_eq!(format_cell_ref_no_dollars(1, 1), "B2");
        assert_eq!(format_cell_ref_no_dollars(0, 26), "AA1");
    }

    #[test]
    fn formats_external_workbook_names_with_literal_brackets() {
        // Workbook names may contain literal `[` characters without escaping.
        assert_eq!(
            format_external_workbook_name("A1[Name.xls"),
            "[A1[Name.xls]"
        );
        assert_parseable("=[A1[Name.xls]Sheet1!A1");

        // Leading `[` characters are preserved (they are not treated as nested prefixes).
        assert_eq!(
            format_external_workbook_name("[LeadingBracket.xls"),
            "[[LeadingBracket.xls]"
        );
        assert_parseable("=[[LeadingBracket.xls]Sheet1!A1");

        // Literal `]` characters inside the workbook name are escaped in Excel formulas as `]]`.
        assert_eq!(
            format_external_workbook_name("Book[Name].xls"),
            "[Book[Name]].xls]"
        );
        assert_parseable("=[Book[Name]].xls]Sheet1!A1");

        // Preserve trailing `]` characters (which must become `]]` inside the prefix).
        assert_eq!(format_external_workbook_name("Book.xls]"), "[Book.xls]]]");

        // Some producers wrap full paths in brackets. Drop the path and normalize the wrapper.
        assert_eq!(
            format_external_workbook_name("[C:\\path\\Book[Name].xls]"),
            "[Book[Name]].xls]"
        );
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
    fn decodes_ptgexp_and_ptgtbl_as_unknown_error_literals() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgExp is a shared-formula token that we cannot resolve in NAME decoding.
        // Ensure we consume its 4-byte payload and keep decoding the remaining tokens.
        //
        // rgce: PtgExp(payload=0) ; PtgInt 1 ; PtgAdd
        let rgce = [
            0x01, // PtgExp
            0x00, 0x00, 0x00, 0x00, // payload
            0x1E, 0x01, 0x00, // 1
            0x03, // +
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!+1");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0x01")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);

        // PtgTbl is another token used by shared formulas; treat it similarly.
        let rgce = [
            0x02, // PtgTbl
            0x00, 0x00, 0x00, 0x00, // payload
            0x1E, 0x01, 0x00, // 1
            0x03, // +
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#UNKNOWN!+1");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("0x02")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_array_as_unknown_error_literal_and_continues() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // 1 + {array_constant} where the array constant is stored out-of-band in rgcb (not
        // available to the NAME rgce decoder). We should still keep the result parseable.
        let rgce = [
            0x1E, 0x01, 0x00, // PtgInt 1
            0x20, // PtgArray
            0, 0, 0, 0, 0, 0, 0, // 7-byte opaque header
            0x03, // PtgAdd
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "1+#UNKNOWN!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgArray")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_array_using_rgcb_to_array_literal() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // 1 + {1,2;3,4}
        let rgce = [
            0x1E, 0x01, 0x00, // PtgInt 1
            0x20, // PtgArray
            0, 0, 0, 0, 0, 0, 0, // 7-byte opaque header
            0x03, // PtgAdd
        ];

        // BIFF8 array constant: [cols_minus1: u16][rows_minus1: u16] + values row-major.
        let mut rgcb = Vec::<u8>::new();
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
        rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 rows
        for n in [1.0f64, 2.0, 3.0, 4.0] {
            rgcb.push(0x01); // number
            rgcb.extend_from_slice(&n.to_le_bytes());
        }

        let decoded = decode_biff8_rgce_with_base_and_rgcb(&rgce, &rgcb, &ctx, None);
        assert_eq!(decoded.text, "1+{1,2;3,4}");
        assert!(
            decoded.warnings.is_empty(),
            "expected no warnings, got {:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_array_in_ptgmemfunc_subexpression_advances_rgcb_cursor() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgMemFunc (non-printing) containing a nested PtgArray, followed by a visible PtgArray.
        //
        // The nested array consumes the *first* array constant block in rgcb; the visible array
        // should therefore decode from the *second* block.
        let mut rgce = Vec::<u8>::new();
        rgce.push(0x29); // PtgMemFunc
        rgce.extend_from_slice(&8u16.to_le_bytes()); // cce = PtgArray (1) + 7-byte header
        rgce.push(0x20); // nested PtgArray
        rgce.extend_from_slice(&[0u8; 7]);
        rgce.push(0x20); // visible PtgArray
        rgce.extend_from_slice(&[0u8; 7]);

        // Two 1x1 array constants: {1} then {2}.
        let mut rgcb = Vec::<u8>::new();
        for n in [1.0f64, 2.0] {
            rgcb.extend_from_slice(&0u16.to_le_bytes()); // 1 col
            rgcb.extend_from_slice(&0u16.to_le_bytes()); // 1 row
            rgcb.push(0x01); // number
            rgcb.extend_from_slice(&n.to_le_bytes());
        }

        let decoded = decode_biff8_rgce_with_base_and_rgcb(&rgce, &rgcb, &ctx, None);
        assert_eq!(decoded.text, "{2}");
        assert!(
            decoded.warnings.is_empty(),
            "expected no warnings, got {:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_array_string_element_escapes_quotes() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let rgce = [
            0x20, // PtgArray
            0, 0, 0, 0, 0, 0, 0, // 7-byte opaque header
        ];

        // {"a""b"}
        let mut rgcb = Vec::<u8>::new();
        rgcb.extend_from_slice(&0u16.to_le_bytes()); // 1 col
        rgcb.extend_from_slice(&0u16.to_le_bytes()); // 1 row
        rgcb.push(0x02); // string
        rgcb.extend_from_slice(&3u16.to_le_bytes()); // cch
        for ch in ['a', '"', 'b'] {
            rgcb.extend_from_slice(&(ch as u16).to_le_bytes());
        }

        let decoded = decode_biff8_rgce_with_base_and_rgcb(&rgce, &rgcb, &ctx, None);
        assert_eq!(decoded.text, "{\"a\"\"b\"}");
        assert!(
            decoded.warnings.is_empty(),
            "expected no warnings, got {:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg18_opaque_payload_token_as_no_op() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Some `.xls` files include an unknown ptg=0x18 token with a 5-byte payload. It should be
        // safe to skip for printing as long as we consume the payload.
        let rgce = [
            0x18, 0x11, 0x22, 0x33, 0x44, 0x55, // ptg=0x18 + 5-byte payload
            0x1E, 0x01, 0x00, // PtgInt 1
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "1");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("Ptg18")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn analyzes_shared_formula_rgce_with_ptg18_opaque_payload() {
        // Ensure our shared-formula rgce analysis stays aligned when an unknown Ptg18 token
        // appears with a 5-byte payload.
        let rgce = vec![
            0x18, 0x11, 0x22, 0x33, 0x44, 0x55, // ptg=0x18 + 5-byte opaque payload
            0x24, 0x00, 0x00, 0x00, 0xC0, // PtgRef A1 (row+col relative flags)
        ];
        let analysis =
            analyze_biff8_shared_formula_rgce(&rgce).expect("analyze shared formula rgce");
        assert!(
            analysis.has_abs_refs_with_relative_flags,
            "expected relative flags to be detected"
        );
        assert!(
            !analysis.has_refn_or_arean,
            "expected no RefN/AreaN tokens to be detected"
        );
    }

    #[test]
    fn materializes_shared_formula_rgce_with_ptg18_opaque_payload() {
        // Ensure shared-formula materialization can copy through an unknown Ptg18 token with a
        // 5-byte payload without corrupting token alignment.
        //
        // Base formula (B1): [opaque][A1]+1
        // Follower materialization (B2): [opaque][A2]+1
        let base_cell = CellCoord::new(0, 1); // B1
        let target_cell = CellCoord::new(1, 1); // B2
        let rgce = vec![
            0x18, 0x11, 0x22, 0x33, 0x44, 0x55, // ptg=0x18 + 5-byte opaque payload
            0x24, 0x00, 0x00, 0x00, 0xC0, // PtgRef A1 (row+col relative flags)
            0x1E, 0x01, 0x00, // PtgInt 1
            0x03, // PtgAdd
        ];

        let materialized = materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell)
            .expect("materialize shared formula rgce");

        let expected = vec![
            0x18, 0x11, 0x22, 0x33, 0x44, 0x55, // opaque token preserved
            0x24, 0x01, 0x00, 0x00, 0xC0, // A2
            0x1E, 0x01, 0x00, // 1
            0x03, // +
        ];
        assert_eq!(materialized, expected);
    }

    #[test]
    fn materializes_shared_formula_ref_col_oob_to_referr_variants() {
        // When a `PtgRef*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 2D error ptg (`PtgRefErr*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(0, 1); // delta_col=+1
        let col_field = encode_col_field(0x3FFF, true, false); // col=0x3FFF with col-relative flag
        for &ptg_ref in &[0x24_u8, 0x44, 0x64] {
            let mut rgce = Vec::new();
            rgce.push(ptg_ref);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row=0
            rgce.extend_from_slice(&col_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_ref + 0x06, "ptg={ptg_ref:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_ref_row_oob_to_referr_variants() {
        // When a `PtgRef*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 2D error ptg (`PtgRefErr*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(1, 0); // delta_row=+1
        let col_field = encode_col_field(0, false, true); // row-relative flag set
        for &ptg_ref in &[0x24_u8, 0x44, 0x64] {
            let mut rgce = Vec::new();
            rgce.push(ptg_ref);
            rgce.extend_from_slice(&u16::MAX.to_le_bytes()); // row=65535
            rgce.extend_from_slice(&col_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_ref + 0x06, "ptg={ptg_ref:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_area_col_oob_to_areaerr_variants() {
        // When a `PtgArea*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 2D error ptg (`PtgAreaErr*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(0, 1); // delta_col=+1
        let col1_field = encode_col_field(0x3FFE, true, false);
        let col2_field = encode_col_field(0x3FFF, true, false); // endpoint shifts OOB
        for &ptg_area in &[0x25_u8, 0x45, 0x65] {
            let mut rgce = Vec::new();
            rgce.push(ptg_area);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row1
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row2
            rgce.extend_from_slice(&col1_field.to_le_bytes());
            rgce.extend_from_slice(&col2_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_area + 0x06, "ptg={ptg_area:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_area_row_oob_to_areaerr_variants() {
        // When a `PtgArea*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 2D error ptg (`PtgAreaErr*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(1, 0); // delta_row=+1
        // Keep the first row fixed in-bounds; shift only the second endpoint out-of-bounds.
        let col1_field = encode_col_field(0, false, false);
        let col2_field = encode_col_field(0, false, true); // row-relative flag set on row2
        for &ptg_area in &[0x25_u8, 0x45, 0x65] {
            let mut rgce = Vec::new();
            rgce.push(ptg_area);
            rgce.extend_from_slice(&(u16::MAX - 1).to_le_bytes()); // row1=65534
            rgce.extend_from_slice(&u16::MAX.to_le_bytes()); // row2=65535 (shifts OOB)
            rgce.extend_from_slice(&col1_field.to_le_bytes());
            rgce.extend_from_slice(&col2_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_area + 0x06, "ptg={ptg_area:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_ref3d_col_oob_to_referr3d_variants() {
        // When a `PtgRef3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 3D error ptg (`PtgRefErr3d*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(0, 1); // delta_col=+1
        let col_field = encode_col_field(0x3FFF, true, false); // col=0x3FFF with col-relative flag
        for &ptg_ref3d in &[0x3A_u8, 0x5A, 0x7A] {
            let mut rgce = Vec::new();
            rgce.push(ptg_ref3d);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row=0
            rgce.extend_from_slice(&col_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_ref3d + 0x02, "ptg={ptg_ref3d:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_ref3d_row_oob_to_referr3d_variants() {
        // When a `PtgRef3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 3D error ptg (`PtgRefErr3d*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(1, 0); // delta_row=+1
        let col_field = encode_col_field(0, false, true); // row-relative flag set
        for &ptg_ref3d in &[0x3A_u8, 0x5A, 0x7A] {
            let mut rgce = Vec::new();
            rgce.push(ptg_ref3d);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            rgce.extend_from_slice(&u16::MAX.to_le_bytes()); // row=65535
            rgce.extend_from_slice(&col_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_ref3d + 0x02, "ptg={ptg_ref3d:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_area3d_col_oob_to_areaerr3d_variants() {
        // When a `PtgArea3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 3D error ptg (`PtgAreaErr3d*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(0, 1); // delta_col=+1
        let col1_field = encode_col_field(0x3FFE, true, false);
        let col2_field = encode_col_field(0x3FFF, true, false); // endpoint shifts OOB
        for &ptg_area3d in &[0x3B_u8, 0x5B, 0x7B] {
            let mut rgce = Vec::new();
            rgce.push(ptg_area3d);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row1
            rgce.extend_from_slice(&0u16.to_le_bytes()); // row2
            rgce.extend_from_slice(&col1_field.to_le_bytes());
            rgce.extend_from_slice(&col2_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_area3d + 0x02, "ptg={ptg_area3d:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn materializes_shared_formula_area3d_row_oob_to_areaerr3d_variants() {
        // When a `PtgArea3d*` token shifts out of BIFF8 bounds during shared-formula materialization,
        // the materializer must emit the 3D error ptg (`PtgAreaErr3d*`), preserving token width.
        let base_cell = CellCoord::new(0, 0);
        let target_cell = CellCoord::new(1, 0); // delta_row=+1
        // Keep the first row fixed in-bounds; shift only the second endpoint out-of-bounds.
        let col1_field = encode_col_field(0, false, false);
        let col2_field = encode_col_field(0, false, true); // row-relative flag set on row2
        for &ptg_area3d in &[0x3B_u8, 0x5B, 0x7B] {
            let mut rgce = Vec::new();
            rgce.push(ptg_area3d);
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
            rgce.extend_from_slice(&(u16::MAX - 1).to_le_bytes()); // row1=65534
            rgce.extend_from_slice(&u16::MAX.to_le_bytes()); // row2=65535 (shifts OOB)
            rgce.extend_from_slice(&col1_field.to_le_bytes());
            rgce.extend_from_slice(&col2_field.to_le_bytes());
            let out =
                materialize_biff8_shared_formula_rgce(&rgce, base_cell, target_cell).unwrap();
            assert_eq!(out[0], ptg_area3d + 0x02, "ptg={ptg_area3d:02X}");
            assert_eq!(&out[1..], &rgce[1..], "payload should be preserved");
        }
    }

    #[test]
    fn decodes_ptg_list_structured_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgExtend(etpg=0x19) table ref with placeholder table/column ids.
        //
        // Payload layout (12 bytes):
        //   [table_id: u32][flags: u16][col_first: u16][col_last: u16][reserved: u16]
        let table_id = 1u32;
        let flags = 0u16; // default/data
        let col_first = 2u16;
        let col_last = 2u16;
        let reserved = 0u16;
        let rgce = [
            vec![0x18, 0x19], // PtgExtend + etpg=PtgList
            table_id.to_le_bytes().to_vec(),
            flags.to_le_bytes().to_vec(),
            col_first.to_le_bytes().to_vec(),
            col_last.to_le_bytes().to_vec(),
            reserved.to_le_bytes().to_vec(),
        ]
        .concat();

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Table1[Column2]");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_list_structured_ref_value_class_adds_at() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Value-class PtgExtend (0x38) for a multi-cell structured ref should render with `@`.
        let table_id = 1u32;
        let flags = 0u16; // default/data
        let col_first = 2u16;
        let col_last = 2u16;
        let reserved = 0u16;
        let rgce = [
            vec![0x38, 0x19], // PtgExtendV + etpg=PtgList
            table_id.to_le_bytes().to_vec(),
            flags.to_le_bytes().to_vec(),
            col_first.to_le_bytes().to_vec(),
            col_last.to_le_bytes().to_vec(),
            reserved.to_le_bytes().to_vec(),
        ]
        .concat();

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "@Table1[Column2]");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
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
    fn decodes_value_class_ptgname_with_explicit_implicit_intersection() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = vec![DefinedNameMeta {
            name: "MyName".to_string(),
            scope_sheet: None,
        }];
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgNameV (value class): name_id=1, reserved bytes=0.
        let rgce = [0x43, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "@MyName");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
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
    fn renders_unknown_ptgfunc_ids_as_parseable_function_calls_best_effort_unary() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Best-effort: unknown PtgFunc with one stack argument should be rendered as a unary call.
        // _UNKNOWN_FUNC_0xFFFF(1):
        //   PtgInt 1
        //   PtgFunc iftab=0xFFFF (unknown)
        let rgce = [0x1E, 0x01, 0x00, 0x21, 0xFF, 0xFF];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "_UNKNOWN_FUNC_0xFFFF(1)");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgFunc")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn renders_unknown_ptgfunc_ids_with_empty_stack_as_parseable_function_calls() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Unknown PtgFunc with no arguments on stack => `_UNKNOWN_FUNC_0xFFFF()`.
        let rgce = [0x21, 0xFF, 0xFF];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "_UNKNOWN_FUNC_0xFFFF()");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("PtgFunc")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_value_class_ptgarea_with_explicit_implicit_intersection() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgAreaV A1:B2 (all components relative).
        let row1 = 0u16;
        let row2 = 1u16;
        let col1 = encode_col_field(0, true, true);
        let col2 = encode_col_field(1, true, true);
        let rgce = [
            0x45, // PtgAreaV
            row1.to_le_bytes()[0],
            row1.to_le_bytes()[1],
            row2.to_le_bytes()[0],
            row2.to_le_bytes()[1],
            col1.to_le_bytes()[0],
            col1.to_le_bytes()[1],
            col2.to_le_bytes()[0],
            col2.to_le_bytes()[1],
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "@A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_value_class_ptgarea3d_with_explicit_implicit_intersection() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgArea3dV ixti=0 A1:B2.
        let ixti = 0u16;
        let row1 = 0u16;
        let row2 = 1u16;
        let col1 = encode_col_field(0, true, true);
        let col2 = encode_col_field(1, true, true);
        let rgce = [
            0x5B, // PtgArea3dV
            ixti.to_le_bytes()[0],
            ixti.to_le_bytes()[1],
            row1.to_le_bytes()[0],
            row1.to_le_bytes()[1],
            row2.to_le_bytes()[0],
            row2.to_le_bytes()[1],
            col1.to_le_bytes()[0],
            col1.to_le_bytes()[1],
            col2.to_le_bytes()[0],
            col2.to_le_bytes()[1],
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "@Sheet1!A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
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
    fn decodes_value_class_ptg_area_n_with_explicit_implicit_intersection() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgAreaNV (value class): (0,0)-(1,1) with all components relative => @A1:B2.
        let row_first = 0u16;
        let row_last = 1u16;
        let col_first = encode_col_field(0, true, true);
        let col_last = encode_col_field(1, true, true);
        let rgce = [
            0x4D, // PtgAreaNV
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
        assert_eq!(decoded.text, "@A1:B2");
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
    fn decodes_value_class_ptg_area_n3d_with_explicit_implicit_intersection() {
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // PtgAreaN3dV ixti=0 with all components relative => @Sheet1!A1:B2 (base cell A1).
        let base = CellCoord::new(0, 0);
        let ixti = 0u16;
        let row_first = 0u16;
        let row_last = 1u16;
        let col_first = encode_col_field(0, true, true);
        let col_last = encode_col_field(1, true, true);
        let rgce = [
            0x5F, // PtgAreaN3dV
            ixti.to_le_bytes()[0],
            ixti.to_le_bytes()[1],
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
        assert_eq!(decoded.text, "@Sheet1!A1:B2");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
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
    fn decodes_ptg_namex_value_class_external_workbook_workbook_scoped_name_with_at() {
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

        // PtgNameXV (value class).
        let rgce = [0x59, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "@'[Book2.xlsx]MyName'");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
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
    fn ptg_namex_internal_sheet_scoped_reserved_boolean_name_falls_back_to_ref() {
        // Like PtgName, PtgNameX can refer to external (or internal) names. If the external name is
        // `TRUE`/`FALSE`, it cannot be rendered as `Sheet1!TRUE` because the parser will lex `TRUE`
        // as a boolean literal after a sheet prefix. Ensure we fall back to a parseable placeholder.
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
            extern_names: vec!["TRUE".to_string()],
        }];

        let ctx = RgceDecodeContext {
            codepage: 1252,
            sheet_names: &sheet_names,
            externsheet: &externsheet,
            supbooks: &supbooks,
            defined_names: &defined_names,
        };

        // PtgNameX (ixti=0, iname=1).
        let rgce = [0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#REF!");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("cannot be rendered parseably")),
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
    fn decodes_ptg_area3d_external_workbook_sheet_range_with_quoting() {
        let sheet_names: Vec<String> = Vec::new();
        // External workbook sheet span (itabFirst != itabLast).
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 1,
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
                sheet_names: vec!["Sheet 1".to_string(), "Sheet3".to_string()],
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

        // '[Book2.xlsx]Sheet 1:Sheet3'!A1:B2
        let mut rgce = Vec::new();
        rgce.push(0x3B); // PtgArea3d
        rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti=0
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rowFirst=0
        rgce.extend_from_slice(&1u16.to_le_bytes()); // rowLast=1
        rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst=A relative
        rgce.extend_from_slice(&0xC001u16.to_le_bytes()); // colLast=B relative

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'[Book2.xlsx]Sheet 1:Sheet3'!A1:B2");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref3d_external_workbook_sheet_range_ref() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 2,
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
                sheet_names: vec![
                    "SheetA".to_string(),
                    "SheetB".to_string(),
                    "SheetC".to_string(),
                ],
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

        // '[Book2.xlsx]SheetA:SheetC'!A1
        let col_field = encode_col_field(0, true, true);
        let rgce = [
            0x3A, // PtgRef3d
            0x00, 0x00, // ixti=0
            0x00, 0x00, // row=0
            col_field.to_le_bytes()[0],
            col_field.to_le_bytes()[1],
        ];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'[Book2.xlsx]SheetA:SheetC'!A1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref3d_external_workbook_sheet_ref_strips_paths_from_virtpath() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        // Some producers store `SUPBOOK.virtPath` as a path (or already bracketed workbook name).
        // Even if workbook_name metadata is missing, we should render the workbook basename.
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
                // Bracketed full path (e.g. `[C:\work\Book2.xlsx]`) is not canonical Excel output,
                // but it appears in some writer implementations.
                virt_path: r"[C:\work\Book2.xlsx]".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: None,
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
    fn decodes_ptg_ref3d_with_internal_supbook_marker() {
        // Some writers reference the internal workbook via an explicit SUPBOOK entry (virtPath=0x0001)
        // instead of using iSupBook=0 in EXTERNSHEET. Ensure we treat that as an internal sheet ref.
        let sheet_names: Vec<String> = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        let supbooks = vec![
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0002}".to_string(),
                kind: SupBookKind::Other,
                workbook_name: None,
                sheet_names: Vec::new(),
                extern_names: Vec::new(),
            },
            SupBookInfo {
                ctab: 0,
                virt_path: "\u{0001}".to_string(),
                kind: SupBookKind::Internal,
                workbook_name: None,
                sheet_names: Vec::new(),
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

        // Sheet1!$A$1.
        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "Sheet1!$A$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_ref3d_external_workbook_sheet_ref_strips_bracketed_full_path_from_virtpath() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        // Some producers store `SUPBOOK.virtPath` as a fully bracketed path like:
        // `[C:\work\Book2.xlsx]`. When extracting the basename, the leading `[` is lost. Ensure we
        // still normalize this to `[Book2.xlsx]` for formula rendering.
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
                virt_path: r"[C:\work\Book2.xlsx]".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: None,
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
    fn decodes_ptg_namex_value_class_udf_function_name_does_not_add_at() {
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

        // MyFunc(1) encoded as: [PtgInt 1][PtgNameXV][PtgFuncVar argc=2 iftab=0x00FF]
        let rgce = vec![
            0x1E, 0x01, 0x00, // PtgInt 1
            0x59, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, // PtgNameXV (ixti=0, iname=1)
            0x22, 0x02, 0xFF, 0x00, // PtgFuncVar(argc=2, iftab=0x00FF)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "MyFunc(1)");
        assert!(decoded.warnings.is_empty(), "warnings={:?}", decoded.warnings);
        assert_parseable(&decoded.text);
    }

    #[test]
    fn decodes_ptg_namex_udf_function_name_skipping_ptgattr_before_ptgfuncvar_00ff() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = vec![ExternSheetEntry {
            supbook: 1,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();

        // Deliberately use an ExternalWorkbook SUPBOOK here: if `PtgNameX` is not recognized as a
        // function name, the decoder would include a workbook+sheet prefix (`'[Book]Sheet'!Name`)
        // and the rendered UDF call would no longer match Excel's canonical `Name(args...)` form.
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

        // MyFunc(1) encoded as: [PtgInt 1][PtgNameX][PtgAttr][PtgFuncVar argc=2 iftab=0x00FF]
        let rgce = vec![
            0x1E, 0x01, 0x00, // PtgInt 1
            0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, // PtgNameX (ixti=0, iname=1)
            0x19, 0x00, 0x00, 0x00, // PtgAttr (no-op; should be skipped by NameX UDF detection)
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
    fn decodes_ptg_namex_external_workbook_workbook_scoped_name_strips_paths_from_virtpath() {
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
                virt_path: r"C:\work\Book2.xlsx".to_string(),
                kind: SupBookKind::ExternalWorkbook,
                workbook_name: None,
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

        // Workbook-scoped external name (sheet indices are negative, so we fall back to `[Book]Name`).
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
    fn wraps_union_operator_when_used_as_function_argument_among_others() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // SUM((A1,B1),C1):
        // - The first argument is a union expression `A1,B1`, which must be parenthesized so the
        //   comma is not interpreted as an argument separator.
        //
        // rgce:
        //   PtgRef A1
        //   PtgRef B1
        //   PtgUnion
        //   PtgRef C1
        //   PtgFuncVar argc=2 iftab=4 (SUM)
        let a1_col = encode_col_field(0, true, true);
        let b1_col = encode_col_field(1, true, true);
        let c1_col = encode_col_field(2, true, true);
        let rgce = vec![
            0x24, 0x00, 0x00, a1_col.to_le_bytes()[0], a1_col.to_le_bytes()[1], // A1
            0x24, 0x00, 0x00, b1_col.to_le_bytes()[0], b1_col.to_le_bytes()[1], // B1
            0x10, // union operator
            0x24, 0x00, 0x00, c1_col.to_le_bytes()[0], c1_col.to_le_bytes()[1], // C1
            0x22, 0x02, 0x04, 0x00, // SUM(argc=2)
        ];

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "SUM((A1,B1),C1)");
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
    fn ptgname_sheet_scoped_reserved_boolean_name_falls_back_to_name_error() {
        // A sheet-scoped name of "TRUE"/"FALSE" cannot be rendered as `Sheet1!TRUE` because the
        // parser lexes `TRUE`/`FALSE` as booleans (not identifiers) after a sheet prefix. Ensure we
        // fall back to a parseable Excel error literal.
        let sheet_names = vec!["Sheet1".to_string()];
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names = vec![DefinedNameMeta {
            name: "TRUE".to_string(),
            scope_sheet: Some(0),
        }];
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Sheet-scoped name (id=1).
        let rgce = [0x23, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00];
        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "#NAME?");
        assert!(
            decoded.warnings.iter().any(|w| w.contains("cannot be rendered parseably")),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
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

    #[test]
    fn quotes_sheet_names_that_look_like_booleans() {
        // Sheet name "TRUE" must be quoted or it will lex as a boolean literal instead of a sheet
        // identifier.
        let rgce = [0x3Au8, 0, 0, 0, 0, 0, 0];

        let sheet_names = vec!["TRUE".to_string()];
        let externsheet = vec![ExternSheetEntry {
            supbook: 0,
            itab_first: 0,
            itab_last: 0,
        }];
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert_eq!(decoded.text, "'TRUE'!$A$1");
        assert!(
            decoded.warnings.is_empty(),
            "warnings={:?}",
            decoded.warnings
        );
        assert_parseable(&decoded.text);
    }

    #[test]
    fn caps_rgce_decode_warnings() {
        let sheet_names: Vec<String> = Vec::new();
        let externsheet: Vec<ExternSheetEntry> = Vec::new();
        let defined_names: Vec<DefinedNameMeta> = Vec::new();
        let ctx = empty_ctx(&sheet_names, &externsheet, &defined_names);

        // Repeated `PtgExp` tokens are valid BIFF8 tokens but unexpected in NAME rgce streams. When
        // present in malformed/corrupt inputs they can occur in long runs, generating one warning
        // per token.
        let reps = MAX_RGCE_WARNINGS + 25;
        let mut rgce = Vec::with_capacity(reps * 5);
        for _ in 0..reps {
            rgce.extend_from_slice(&[0x01, 0, 0, 0, 0]);
        }

        let decoded = decode_biff8_rgce(&rgce, &ctx);
        assert!(
            decoded.warnings.len() <= MAX_RGCE_WARNINGS + 1,
            "warning cap exceeded: len={} warnings={:?}",
            decoded.warnings.len(),
            decoded.warnings
        );
        assert!(
            decoded
                .warnings
                .iter()
                .any(|w| w == RGCE_WARNINGS_SUPPRESSED_MESSAGE),
            "expected suppression message, warnings={:?}",
            decoded.warnings
        );
    }
}
