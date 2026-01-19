//! BIFF12 `rgce` (formula token stream) codec.
//!
//! The decoder is best-effort and is primarily used for diagnostics and for surfacing formula
//! text when reading XLSB files.
//!
//! The encoder is intentionally small-scope: it supports enough of Excel's formula language to
//! round-trip common patterns while we build out full compatibility.

use crate::errors::{xlsb_error_code_from_literal, xlsb_error_literal};
use crate::workbook_context::{NameScope, WorkbookContext};
use formula_biff::ptg_list::{decode_ptg_list_payload_candidates, PtgListDecoded};
use formula_biff::structured_refs::{
    estimated_structured_ref_len, push_structured_ref, structured_columns_placeholder_from_ids,
    structured_ref_is_single_cell, structured_ref_item_from_flags, KNOWN_FLAGS_MASK, StructuredColumns,
    StructuredRefItem,
};
use formula_model::external_refs::{format_external_key, format_external_span_key};
#[cfg(feature = "write")]
use formula_model::external_refs::format_external_workbook_key;
use formula_model::{
    push_escaped_excel_double_quote_char, push_escaped_excel_single_quotes,
    push_excel_single_quoted_identifier, sheet_name_needs_quotes_a1,
};
use formula_model::sheet_name_eq_case_insensitive;
use thiserror::Error;

/// Structured `rgce` decode failure with ptg id + offset.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeError {
    /// Encountered a ptg we do not (yet) handle.
    UnknownPtg { offset: usize, ptg: u8 },
    /// Not enough bytes remained in the rgce stream to decode the current token.
    UnexpectedEof {
        offset: usize,
        ptg: u8,
        needed: usize,
        remaining: usize,
    },
    /// The ptg required more stack items than were available.
    StackUnderflow { offset: usize, ptg: u8 },
    /// A ptg referenced a constant we don't know how to display.
    InvalidConstant { offset: usize, ptg: u8, value: u8 },
    /// Decoding exceeded the maximum output size derived from the input length.
    OutputTooLarge {
        offset: usize,
        ptg: u8,
        max_len: usize,
    },
    /// After decoding, the expression stack didn't contain exactly one item.
    StackNotSingular {
        offset: usize,
        ptg: u8,
        stack_len: usize,
    },
}

impl DecodeError {
    pub fn offset(&self) -> usize {
        match *self {
            DecodeError::UnknownPtg { offset, .. } => offset,
            DecodeError::UnexpectedEof { offset, .. } => offset,
            DecodeError::StackUnderflow { offset, .. } => offset,
            DecodeError::InvalidConstant { offset, .. } => offset,
            DecodeError::OutputTooLarge { offset, .. } => offset,
            DecodeError::StackNotSingular { offset, .. } => offset,
        }
    }

    pub fn ptg(&self) -> Option<u8> {
        match *self {
            DecodeError::UnknownPtg { ptg, .. } => Some(ptg),
            DecodeError::UnexpectedEof { ptg, .. } => Some(ptg),
            DecodeError::StackUnderflow { ptg, .. } => Some(ptg),
            DecodeError::InvalidConstant { ptg, .. } => Some(ptg),
            DecodeError::OutputTooLarge { ptg, .. } => Some(ptg),
            DecodeError::StackNotSingular { ptg, .. } => Some(ptg),
        }
    }
}

impl std::fmt::Display for DecodeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match *self {
            DecodeError::UnknownPtg { offset, ptg } => {
                write!(f, "unknown ptg=0x{ptg:02X} at rgce offset {offset}")
            }
            DecodeError::UnexpectedEof {
                offset,
                ptg,
                needed,
                remaining,
            } => write!(
                f,
                "unexpected eof decoding ptg=0x{ptg:02X} at rgce offset {offset} (needed {needed} bytes, remaining {remaining})"
            ),
            DecodeError::StackUnderflow { offset, ptg } => {
                write!(f, "stack underflow decoding ptg=0x{ptg:02X} at rgce offset {offset}")
            }
            DecodeError::InvalidConstant { offset, ptg, value } => write!(
                f,
                "invalid constant 0x{value:02X} decoding ptg=0x{ptg:02X} at rgce offset {offset}"
            ),
            DecodeError::OutputTooLarge { offset, ptg, max_len } => write!(
                f,
                "formula decode exceeded max_len={max_len} decoding ptg=0x{ptg:02X} at rgce offset {offset}"
            ),
            DecodeError::StackNotSingular {
                offset,
                ptg,
                stack_len,
            } => write!(
                f,
                "formula decoded with stack_len={stack_len} at rgce offset {offset} (ptg=0x{ptg:02X}, expected 1)"
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

/// A non-fatal issue encountered while decoding an rgce token stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeFailureKind {
    UnknownPtg,
    UnexpectedEof,
    StackUnderflow,
    InvalidConstant,
    OutputTooLarge,
    StackNotSingular,
}

/// A non-fatal issue encountered while decoding an rgce token stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DecodeWarning {
    /// Encountered an unknown/extended error code in a `PtgErr` token.
    ///
    /// The decoder will emit `#UNKNOWN!` in the output formula text and continue.
    UnknownErrorCode { code: u8, offset: usize },
    /// Encountered an unknown/extended error code inside an array constant (`PtgArray` / `rgcb`).
    ///
    /// The decoder will emit `#UNKNOWN!` in the output formula text and continue.
    UnknownArrayErrorCode { code: u8, offset: usize },
    /// Encountered unknown flag bits while decoding an Excel structured reference (table ref).
    ///
    /// The decoder will ignore unknown bits and emit best-effort formula text.
    UnknownStructuredRefFlags { flags: u32, offset: usize },
    /// The decoder encountered a hard failure and could not produce formula text.
    ///
    /// When this warning is present, the returned formula text may be missing (`None`) or may
    /// contain best-effort placeholders (e.g. `_UNKNOWN_FUNC_0XFFFF(...)`).
    ///
    /// This is surfaced as a warning (rather than bubbling up an error) because the
    /// [`decode_formula_rgce*`] APIs are intended to be best-effort for diagnostics.
    DecodeFailed {
        kind: DecodeFailureKind,
        offset: usize,
        ptg: u8,
    },
}

/// Result of decoding an rgce token stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodedFormula {
    /// Best-effort decoded Excel formula text (without the leading `=`).
    pub text: Option<String>,
    /// Any non-fatal decode warnings.
    pub warnings: Vec<DecodeWarning>,
}

/// Best-effort decode of an XLSB `rgce` token stream to Excel formula text.
pub fn decode_formula_rgce(rgce: &[u8]) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, &[], None, None)
}

/// Best-effort decode of an XLSB `rgce` token stream to Excel formula text, using trailing `rgcb`
/// data blocks referenced by certain ptgs (e.g. `PtgArray`).
pub fn decode_formula_rgce_with_rgcb(rgce: &[u8], rgcb: &[u8]) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, rgcb, None, None)
}

/// Best-effort decode of an XLSB `rgce` token stream, using workbook context.
pub fn decode_formula_rgce_with_context(rgce: &[u8], ctx: &WorkbookContext) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, &[], Some(ctx), None)
}

/// Best-effort decode of an XLSB `rgce` token stream, using workbook context and trailing `rgcb`
/// data blocks referenced by certain ptgs (e.g. `PtgArray`).
pub fn decode_formula_rgce_with_context_and_rgcb(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &WorkbookContext,
) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, rgcb, Some(ctx), None)
}

/// Best-effort decode of an `rgce` token stream using the (0-indexed) origin cell for
/// relative-reference tokens like `PtgRefN` / `PtgAreaN`.
pub fn decode_formula_rgce_with_base(rgce: &[u8], base: CellCoord) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, &[], None, Some(base))
}

/// Best-effort decode of an `rgce` token stream using trailing `rgcb` data blocks and the
/// (0-indexed) origin cell for relative-reference tokens like `PtgRefN` / `PtgAreaN`.
pub fn decode_formula_rgce_with_rgcb_and_base(
    rgce: &[u8],
    rgcb: &[u8],
    base: CellCoord,
) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, rgcb, None, Some(base))
}

/// Best-effort decode of an `rgce` token stream using both workbook context (for 3D refs / names)
/// and a base cell (for relative-reference tokens like `PtgRefN` / `PtgAreaN`).
pub fn decode_formula_rgce_with_context_and_base(
    rgce: &[u8],
    ctx: &WorkbookContext,
    base: CellCoord,
) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, &[], Some(ctx), Some(base))
}

/// Best-effort decode of an `rgce` token stream using workbook context, trailing `rgcb` data
/// blocks, and a base cell (for relative-reference tokens like `PtgRefN` / `PtgAreaN`).
pub fn decode_formula_rgce_with_context_and_rgcb_and_base(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &WorkbookContext,
    base: CellCoord,
) -> DecodedFormula {
    decode_formula_rgce_impl(rgce, rgcb, Some(ctx), Some(base))
}

fn decode_formula_rgce_impl(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: Option<&WorkbookContext>,
    base: Option<CellCoord>,
) -> DecodedFormula {
    let mut warnings = Vec::new();
    let text = match decode_rgce_impl(rgce, rgcb, ctx, base, Some(&mut warnings)) {
        Ok(text) => Some(text),
        Err(e) => {
            let (kind, offset, ptg) = match e {
                DecodeError::UnknownPtg { offset, ptg } => {
                    (DecodeFailureKind::UnknownPtg, offset, ptg)
                }
                DecodeError::UnexpectedEof { offset, ptg, .. } => {
                    (DecodeFailureKind::UnexpectedEof, offset, ptg)
                }
                DecodeError::StackUnderflow { offset, ptg } => {
                    (DecodeFailureKind::StackUnderflow, offset, ptg)
                }
                DecodeError::InvalidConstant { offset, ptg, .. } => {
                    (DecodeFailureKind::InvalidConstant, offset, ptg)
                }
                DecodeError::OutputTooLarge { offset, ptg, .. } => {
                    (DecodeFailureKind::OutputTooLarge, offset, ptg)
                }
                DecodeError::StackNotSingular { offset, ptg, .. } => {
                    (DecodeFailureKind::StackNotSingular, offset, ptg)
                }
            };
            warnings.push(DecodeWarning::DecodeFailed { kind, offset, ptg });
            None
        }
    };
    DecodedFormula { text, warnings }
}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`).
pub fn decode_rgce(rgce: &[u8]) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, &[], None, None, None)
}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`),
/// using trailing `rgcb` data blocks referenced by certain ptgs (e.g. `PtgArray`).
pub fn decode_rgce_with_rgcb(rgce: &[u8], rgcb: &[u8]) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, rgcb, None, None, None)
}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`),
/// using workbook context to resolve sheet indices (`ixti`) and defined names.
pub fn decode_rgce_with_context(rgce: &[u8], ctx: &WorkbookContext) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, &[], Some(ctx), None, None)
}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`),
/// using workbook context and trailing `rgcb` data blocks referenced by certain ptgs (e.g.
/// `PtgArray`).
pub fn decode_rgce_with_context_and_rgcb(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &WorkbookContext,
) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, rgcb, Some(ctx), None, None)
}

/// Decode an `rgce` token stream using the (0-indexed) origin cell for relative-reference tokens
/// like `PtgRefN` / `PtgAreaN`.
pub fn decode_rgce_with_base(rgce: &[u8], base: CellCoord) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, &[], None, Some(base), None)
}

/// Decode an `rgce` token stream using trailing `rgcb` data blocks and the (0-indexed) origin cell
/// for relative-reference tokens like `PtgRefN` / `PtgAreaN`.
pub fn decode_rgce_with_rgcb_and_base(
    rgce: &[u8],
    rgcb: &[u8],
    base: CellCoord,
) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, rgcb, None, Some(base), None)
}

/// Decode an `rgce` token stream using both workbook context (for 3D refs / names) and a base cell
/// (for relative-reference tokens like `PtgRefN` / `PtgAreaN`).
pub fn decode_rgce_with_context_and_base(
    rgce: &[u8],
    ctx: &WorkbookContext,
    base: CellCoord,
) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, &[], Some(ctx), Some(base), None)
}

/// Decode an `rgce` token stream using workbook context, trailing `rgcb` data blocks, and a base
/// cell (for relative-reference tokens like `PtgRefN` / `PtgAreaN`).
pub fn decode_rgce_with_context_and_rgcb_and_base(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: &WorkbookContext,
    base: CellCoord,
) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, rgcb, Some(ctx), Some(base), None)
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
    let contains_union = args.iter().any(|arg| arg.contains_union);

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
        // The union operator uses `,`, which is also the function argument separator. To make
        // decoded formulas round-trip through `formula-engine`, wrap any arg containing union in
        // parentheses (Excel's canonical form, e.g. `SUM((A1,B1))`).
        if arg.contains_union {
            // If the argument is already explicitly parenthesized (via `PtgParen`), avoid adding
            // an extra set of parentheses that `formula-engine` would discard during
            // serialization.
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

fn fmt_sheet_name(out: &mut String, sheet: &str) {
    formula_model::push_sheet_name_a1(out, sheet);
}

fn format_sheet_prefix(first: &str, last: &str) -> String {
    let mut out = String::new();
    if first == last {
        fmt_sheet_name(&mut out, first);
    } else {
        // Excel's canonical text format for a 3D sheet range is `Sheet1:Sheet3!A1`, but
        // `formula-engine` only recognizes a single identifier token before `!`. Since `:` is
        // invalid in Excel sheet names, we can round-trip through text by always emitting the
        // combined prefix as one quoted identifier: `'Sheet1:Sheet3'!A1`.
        out.push('\'');
        push_escaped_excel_single_quotes(&mut out, first);
        out.push(':');
        push_escaped_excel_single_quotes(&mut out, last);
        out.push('\'');
    }
    out.push('!');
    out
}

fn format_external_sheet_prefix(book: &str, first: &str, last: &str) -> String {
    let same_sheet = sheet_name_eq_case_insensitive(first, last);
    if !same_sheet {
        // Keep external workbook 3D spans parseable by `formula-engine` by emitting a single
        // quoted identifier before `!` (e.g. `'[Book.xlsx]Sheet1:Sheet3'!A1`).
        //
        // Excel sheet names cannot contain ':' so this representation is unambiguous.
        let combined = format_external_span_key(book, first, last);
        let mut out = String::new();
        push_excel_single_quoted_identifier(&mut out, &combined);
        out.push('!');
        return out;
    }

    if sheet_name_needs_quotes_a1(first) {
        let combined = format_external_key(book, first);
        let mut out = String::new();
        push_excel_single_quoted_identifier(&mut out, &combined);
        out.push('!');
        out
    } else {
        let mut out = format_external_key(book, first);
        out.push('!');
        out
    }
}

fn remaining_len(buf: &[u8], i: usize) -> usize {
    buf.len().checked_sub(i).unwrap_or(0)
}

fn decode_rgce_impl(
    rgce: &[u8],
    rgcb: &[u8],
    ctx: Option<&WorkbookContext>,
    base: Option<CellCoord>,
    mut warnings: Option<&mut Vec<DecodeWarning>>,
) -> Result<String, DecodeError> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    // Prevent pathological expansion (e.g. from future token support).
    const MAX_OUTPUT_FACTOR: usize = 10;
    // Some ptgs (notably `PtgArray`) reference additional data stored in the trailing `rgcb`
    // buffer. Include it when deriving an upper bound for decoded output so we don't reject
    // legitimate array constants whose `rgce` stream is tiny but `rgcb` is not.
    let max_len = rgce
        .len()
        .checked_add(rgcb.len())
        .and_then(|n| n.checked_mul(MAX_OUTPUT_FACTOR))
        .unwrap_or(usize::MAX);

    let mut i = 0usize;
    let mut last_ptg_offset = 0usize;
    let mut last_ptg = rgce.get(0).copied().unwrap_or(0);
    let mut rgcb_pos = 0usize;

    let mut stack: Vec<ExprFragment> = Vec::new();

    fn parenthesize(mut text: String) -> String {
        text.reserve(2);
        text.insert(0, '(');
        text.push(')');
        text
    }

    fn maybe_parenthesize(expr: ExprFragment, required_prec: u8) -> String {
        if expr.precedence < required_prec && !expr.is_missing {
            parenthesize(expr.text)
        } else {
            expr.text
        }
    }

    while i < rgce.len() {
        let ptg_offset = i;
        let Some(&ptg) = rgce.get(i) else {
            debug_assert!(false, "rgce cursor out of bounds (i={i}, len={})", rgce.len());
            return Err(DecodeError::UnexpectedEof {
                offset: ptg_offset,
                ptg: last_ptg,
                needed: 1,
                remaining: 0,
            });
        };
        i += 1;

        last_ptg_offset = ptg_offset;
        last_ptg = ptg;

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };
                let Some(prec) = binary_precedence(ptg) else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                if stack.len() < 2 {
                    return Err(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    });
                }
                let right = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                let left = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;

                let contains_union = left.contains_union || right.contains_union || ptg == 0x10;
                let left_s = maybe_parenthesize(left, prec);
                let right_s = maybe_parenthesize(right, prec);

                let mut text = String::new();
                let _ = text.try_reserve(left_s.len() + op.len() + right_s.len());
                text.push_str(&left_s);
                text.push_str(op);
                text.push_str(&right_s);

                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union,
                    is_missing: false,
                });
            }
            // Unary +/-.
            0x12 | 0x13 => {
                let op = if ptg == 0x12 { "+" } else { "-" };
                let expr = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                let prec = 70;
                let contains_union = expr.contains_union;
                let inner = maybe_parenthesize(expr, prec);
                let mut text = String::new();
                let _ = text.try_reserve(op.len() + inner.len());
                text.push_str(op);
                text.push_str(&inner);
                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union,
                    is_missing: false,
                });
            }
            // Percent postfix.
            0x14 => {
                let expr = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                let prec = 60;
                let contains_union = expr.contains_union;
                let mut text = maybe_parenthesize(expr, prec);
                text.push('%');
                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union,
                    is_missing: false,
                });
            }
            // Spill range postfix (`#`).
            PTG_SPILL => {
                let expr = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                let prec = 60;
                let contains_union = expr.contains_union;
                let mut text = maybe_parenthesize(expr, prec);
                text.push('#');
                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union,
                    is_missing: false,
                });
            }
            // Explicit parentheses.
            0x15 => {
                let expr = stack.pop().ok_or(DecodeError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                stack.push(ExprFragment {
                    text: parenthesize(expr.text),
                    precedence: 100,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Missing arg.
            0x16 => stack.push(ExprFragment::missing()),
            // PtgStr: [cch: u16][utf16 chars...]
            0x17 => {
                let remaining = remaining_len(rgce, i);
                if remaining < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining,
                    });
                }
                let len_end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: remaining_len(rgce, i),
                })?;
                let len_bytes = rgce.get(i..len_end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: remaining_len(rgce, i),
                })?;
                let cch = u16::from_le_bytes([len_bytes[0], len_bytes[1]]) as usize;
                i = len_end;

                // `cch` is a u16 widened to usize, so `* 2` cannot overflow.
                let needed = cch * 2;
                let remaining = remaining_len(rgce, i);
                if remaining < needed {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed,
                        remaining,
                    });
                }
                let end = i.checked_add(needed).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed,
                    remaining: remaining_len(rgce, i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed,
                    remaining: remaining_len(rgce, i),
                })?;
                i = end;

                // Excel escapes embedded quotes by doubling them inside the literal.
                let iter = raw
                    .chunks_exact(2)
                    .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                let mut lit = String::new();
                let _ = lit.try_reserve(cch + 2);
                lit.push('"');
                for decoded in std::char::decode_utf16(iter) {
                    match decoded {
                        Ok(ch) => push_escaped_excel_double_quote_char(&mut lit, ch),
                        Err(_) => push_escaped_excel_double_quote_char(&mut lit, '\u{FFFD}'),
                    }
                }
                lit.push('"');
                stack.push(ExprFragment::new(lit));
            }
            0x18 | 0x38 | 0x58 => {
                // PtgExtend / PtgExtendV / PtgExtendA.
                //
                // MS-XLSB encodes newer operand tokens (including structured references / table
                // refs) using `PtgExtend` followed by an `etpg` subtype byte.
                //
                // We currently decode the variants required for Excel structured references
                // (tables), which appear as `etpg=0x19` (`PtgList` in documentation).
                let remaining = remaining_len(rgce, i);
                if remaining < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining,
                    });
                }
                let Some(&etpg) = rgce.get(i) else {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: remaining_len(rgce, i),
                    });
                };
                i += 1;

                match etpg {
                    0x19 => {
                        // PtgList (structured reference / table ref).
                        //
                        // Excel uses a fixed 12-byte payload. The exact layout is documented in
                        // MS-XLSB, but in practice there are multiple observed encodings in the
                        // wild. We decode in a best-effort way by trying a handful of plausible
                        // interpretations and preferring the one that matches available workbook
                        // context (table/column name lookups).
                        // Some XLSB producers appear to insert extra prefix bytes before the
                        // canonical 12-byte payload. Use the same best-effort alignment heuristic
                        // as shared-formula materialization so we can skip the token correctly and
                        // keep the rgce stream aligned.
                        // When no workbook context is provided, still apply the column/table-id
                        // plausibility heuristic by scoring candidates against an empty context.
                        // This improves decoding robustness for "weird" payload alignments even
                        // when table metadata is unavailable.
                        let default_ctx = WorkbookContext::default();
                        let ctx_for_scoring = Some(ctx.unwrap_or(&default_ctx));

                        let remaining = remaining_len(rgce, i);
                        let tail = rgce.get(i..).unwrap_or(&[]);
                        let Some(payload_len) = ptg_list_payload_len_best_effort(tail, ctx_for_scoring)
                        else {
                            return Err(DecodeError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: 12,
                                remaining,
                            });
                        };
                        let offset = payload_len.checked_sub(12).unwrap_or(0);
                        if remaining < payload_len {
                            return Err(DecodeError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: payload_len,
                                remaining,
                            });
                        }

                        let core_start = i.checked_add(offset).ok_or(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: payload_len,
                            remaining,
                        })?;
                        let core_end = core_start.checked_add(12).ok_or(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: payload_len,
                            remaining,
                        })?;
                        let mut payload = [0u8; 12];
                        let core = rgce.get(core_start..core_end).ok_or(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: payload_len,
                            remaining,
                        })?;
                        payload.copy_from_slice(core);
                        i = i.checked_add(payload_len).ok_or(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: payload_len,
                            remaining,
                        })?;

                        let decoded =
                            decode_ptg_list_payload_best_effort(&payload, ctx_for_scoring);

                        // Interpret row/item flags. We intentionally accept unknown bits and
                        // continue decoding.
                        let flags16 = (decoded.flags & 0xFFFF) as u16;
                        let unknown = flags16 & !KNOWN_FLAGS_MASK;
                        if unknown != 0 {
                            if let Some(warnings) = warnings.as_deref_mut() {
                                warnings.push(DecodeWarning::UnknownStructuredRefFlags {
                                    flags: decoded.flags,
                                    offset: ptg_offset,
                                });
                            }
                        }

                        let item = structured_ref_item_from_flags(flags16);

                        let table_name = ctx
                            .and_then(|ctx| ctx.table_name(decoded.table_id))
                            .map(|s| s.to_string())
                            .unwrap_or_else(|| format!("Table{}", decoded.table_id));

                        let col_first = decoded.col_first;
                        let col_last = decoded.col_last;

                        let columns = if let Some(ctx) = ctx {
                            if col_first == 0 && col_last == 0 {
                                StructuredColumns::All
                            } else if col_first == col_last {
                                let name = ctx
                                    .table_column_name(decoded.table_id, col_first)
                                    .map(|s| s.to_string())
                                    .unwrap_or_else(|| format!("Column{col_first}"));
                                StructuredColumns::Single(name)
                            } else {
                                let start = ctx
                                    .table_column_name(decoded.table_id, col_first)
                                    .map(|s| s.to_string())
                                    .unwrap_or_else(|| format!("Column{col_first}"));
                                let end = ctx
                                    .table_column_name(decoded.table_id, col_last)
                                    .map(|s| s.to_string())
                                    .unwrap_or_else(|| format!("Column{col_last}"));
                                StructuredColumns::Range { start, end }
                            }
                        } else {
                            structured_columns_placeholder_from_ids(col_first, col_last)
                        };

                        let display_table_name = match item {
                            Some(StructuredRefItem::ThisRow) => None,
                            _ => Some(table_name.as_str()),
                        };

                        let mut prec = 100;
                        let is_value_class = ptg == 0x38;
                        let needs_at = is_value_class && !structured_ref_is_single_cell(item, &columns);
                        let mut out = String::new();
                        if let Some(cap) = estimated_structured_ref_len(display_table_name, item, &columns)
                            .checked_add(needs_at as usize)
                        {
                            let _ = out.try_reserve(cap);
                        }
                        if needs_at {
                            // Like value-class range/name tokens, Excel uses value-class list
                            // tokens to represent legacy implicit intersection.
                            prec = 70;
                            out.push('@');
                        }
                        push_structured_ref(display_table_name, item, &columns, &mut out);

                        stack.push(ExprFragment {
                            text: out,
                            precedence: prec,
                            contains_union: false,
                            is_missing: false,
                        });
                    }
                    _ => {
                        return Err(DecodeError::UnknownPtg {
                            offset: ptg_offset,
                            ptg,
                        })
                    }
                }
            }
            0x19 => {
                // PtgAttr: [grbit: u8][wAttr: u16]
                //
                // Excel uses `PtgAttr` for multiple attributes. Most are evaluation hints or
                // formatting metadata that do not affect the reconstructed formula text, but some
                // do. In particular, `tAttrSum` is used for an optimization where `SUM(A1:A10)` is
                // encoded as `PtgArea` + `PtgAttr(tAttrSum)` (no explicit `PtgFuncVar(SUM)` token).
                let remaining = remaining_len(rgce, i);
                if remaining < 3 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining,
                    });
                }
                let end = i.checked_add(3).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: remaining_len(rgce, i),
                })?;
                let hdr = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: remaining_len(rgce, i),
                })?;
                let grbit = hdr[0];
                let w_attr = u16::from_le_bytes([hdr[1], hdr[2]]);
                i = end;

                const T_ATTR_VOLATILE: u8 = 0x01;
                const T_ATTR_IF: u8 = 0x02;
                const T_ATTR_CHOOSE: u8 = 0x04;
                const T_ATTR_SKIP: u8 = 0x08;
                const T_ATTR_SUM: u8 = 0x10;
                const T_ATTR_SPACE: u8 = 0x40;
                const T_ATTR_SEMI: u8 = 0x80;

                if grbit & T_ATTR_SUM != 0 {
                    let a = stack.pop().ok_or(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?;
                    stack.push(format_function_call("SUM", vec![a]));
                } else if grbit & T_ATTR_CHOOSE != 0 {
                    // `tAttrChoose` is followed by a jump table of `u16` offsets used for
                    // short-circuit evaluation.
                    //
                    // We don't need it for printing, but we must consume it so subsequent tokens
                    // stay aligned.
                    // `w_attr` is a u16 widened to usize, so `* 2` cannot overflow.
                    let needed = (w_attr as usize) * 2;
                    let remaining = remaining_len(rgce, i);
                    if remaining < needed {
                        return Err(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed,
                            remaining,
                        });
                    }
                    i = i.checked_add(needed).ok_or(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed,
                        remaining: remaining_len(rgce, i),
                    })?;
                } else {
                    // Ignore other attributes for printing, but keep the constants referenced so
                    // this doesn't accidentally get treated as dead code.
                    let _ = grbit
                        & (T_ATTR_VOLATILE | T_ATTR_IF | T_ATTR_SKIP | T_ATTR_SPACE | T_ATTR_SEMI);
                }
            }
            // PtgErr: [err: u8]
            0x1C => {
                let remaining = remaining_len(rgce, i);
                if remaining < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining,
                    });
                }
                let code_offset = i;
                let Some(&err) = rgce.get(i) else {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: remaining_len(rgce, i),
                    });
                };
                i += 1;

                let text = match xlsb_error_literal(err) {
                    Some(lit) => lit,
                    None => {
                        if let Some(warnings) = warnings.as_deref_mut() {
                            warnings.push(DecodeWarning::UnknownErrorCode {
                                code: err,
                                offset: code_offset,
                            });
                        }
                        // `formula-engine` can parse `#UNKNOWN!`, which lets us preserve the full
                        // formula string even when Excel introduces new internal error ids.
                        "#UNKNOWN!"
                    }
                };
                stack.push(ExprFragment::new(text.to_string()));
            }
            // PtgBool: [b: u8]
            0x1D => {
                let remaining = remaining_len(rgce, i);
                if remaining < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining,
                    });
                }
                let Some(&b) = rgce.get(i) else {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: remaining_len(rgce, i),
                    });
                };
                i += 1;
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // PtgInt: [n: u16]
            0x1E => {
                let remaining = remaining_len(rgce, i);
                if remaining < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining,
                    });
                }
                let end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: remaining_len(rgce, i),
                })?;
                let n_bytes = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: remaining_len(rgce, i),
                })?;
                let n = u16::from_le_bytes([n_bytes[0], n_bytes[1]]);
                i = end;
                stack.push(ExprFragment::new(n.to_string()));
            }
            // PtgNum: [f64]
            0x1F => {
                let remaining = remaining_len(rgce, i);
                if remaining < 8 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining,
                    });
                }
                let mut bytes = [0u8; 8];
                let end = i.checked_add(8).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 8,
                    remaining: remaining_len(rgce, i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 8,
                    remaining: remaining_len(rgce, i),
                })?;
                bytes.copy_from_slice(raw);
                i = end;
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgArray: [unused: 7 bytes] + serialized array constant stored in rgcb.
            0x20 | 0x40 | 0x60 => {
                let remaining = remaining_len(rgce, i);
                if remaining < 7 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 7,
                        remaining,
                    });
                }
                i += 7;

                let arr = decode_array_constant(rgcb, &mut rgcb_pos, warnings.as_deref_mut())
                    .ok_or(DecodeError::InvalidConstant {
                        offset: ptg_offset,
                        ptg,
                        value: 0xFF,
                    })?;
                stack.push(ExprFragment::new(arr));
            }
            // PtgRef: [row: u32][col: u16 (with relative flags in high bits)]
            0x24 | 0x44 | 0x64 => {
                let remaining = remaining_len(rgce, i);
                if remaining < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining,
                    });
                }

                let end = i.checked_add(6).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: remaining_len(rgce, i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: remaining_len(rgce, i),
                })?;
                let row0 = u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]);
                let row = u64::from(row0) + 1;
                let flags = raw[5];
                let col = u16::from_le_bytes([raw[4], flags & 0x3F]);
                i = end;

                stack.push(ExprFragment::new(format_cell_ref(row, col as u32, flags)));
            }
            // PtgArea: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x25 | 0x45 | 0x65 => {
                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let end = i.checked_add(12).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 12,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 12,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let row_first0 = u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]);
                let row_last0 = u32::from_le_bytes([raw[4], raw[5], raw[6], raw[7]]);
                let col_first = u16::from_le_bytes([raw[8], raw[9]]);
                let col_last = u16::from_le_bytes([raw[10], raw[11]]);
                i = end;

                let is_single_cell = row_first0 == row_last0
                    && (col_first & COL_INDEX_MASK) == (col_last & COL_INDEX_MASK);
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                let mut prec = 100;
                if is_value_class && !is_single_cell {
                    // Legacy implicit intersection: Excel encodes this by using a value-class
                    // range token; modern formula text uses an explicit `@` operator.
                    text.push('@');
                    prec = 70;
                }
                const MAX_ROW: u32 = 1_048_575;
                let col_first_idx = col_first & COL_INDEX_MASK;
                let col_last_idx = col_last & COL_INDEX_MASK;
                if row_first0 == 0 && row_last0 == MAX_ROW {
                    // Column range: `A:C` / `A:A`.
                    push_col_ref_from_field(&mut text, col_first);
                    text.push(':');
                    push_col_ref_from_field(&mut text, col_last);
                } else if col_first_idx == 0 && col_last_idx == COL_INDEX_MASK {
                    // Row range: `1:3` / `1:1`.
                    push_row_ref_from_field(&mut text, row_first0, col_first);
                    text.push(':');
                    push_row_ref_from_field(&mut text, row_last0, col_last);
                } else if is_single_cell {
                    push_cell_ref_from_field(&mut text, row_first0, col_first);
                } else {
                    push_cell_ref_from_field(&mut text, row_first0, col_first);
                    text.push(':');
                    push_cell_ref_from_field(&mut text, row_last0, col_last);
                }
                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgMem* tokens: no-op for printing, but consume payload to keep offsets aligned.
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                // MS-XLSB/BIFF: u16 cce (size of a subexpression, used by the evaluator).
                let cce_end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let cce_bytes = rgce.get(i..cce_end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let cce = u16::from_le_bytes([cce_bytes[0], cce_bytes[1]]) as usize;
                i = cce_end;
                if rgce.len().saturating_sub(i) < cce {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                // `PtgMem*` embeds a nested token stream (`cce` bytes) that is not printed, but
                // it can still contain `PtgArray` tokens that consume trailing `rgcb` bytes.
                // Advance the `rgcb` cursor through any array-constant blocks referenced by the
                // nested stream so later `PtgArray` tokens stay aligned.
                if !rgcb.is_empty() {
                    let subexpr_end = i.checked_add(cce).ok_or(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    })?;
                    let subexpr = rgce.get(i..subexpr_end).ok_or(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    })?;
                    consume_rgcb_arrays_in_subexpression(
                        subexpr,
                        rgcb,
                        &mut rgcb_pos,
                        i,
                        ctx,
                    )?;
                }
                i = i.checked_add(cce).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: cce,
                    remaining: rgce.len().saturating_sub(i),
                })?;
            }
            // PtgRefErr: [row: u32][col: u16]
            0x2A | 0x4A | 0x6A => {
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 6;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 12;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgRefN: [row_off: i32][col_off: i16]
            0x2C | 0x4C | 0x6C => {
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let Some(base) = base else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(6).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let row_off = i32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]) as i64;
                let col_off = i16::from_le_bytes([raw[4], raw[5]]) as i64;
                i = end;

                const MAX_ROW: i64 = 1_048_575;
                const MAX_COL: i64 = COL_INDEX_MASK as i64;
                let abs_row = base.row as i64 + row_off;
                let abs_col = base.col as i64 + col_off;
                if abs_row < 0 || abs_row > MAX_ROW || abs_col < 0 || abs_col > MAX_COL {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let col_field = encode_col_field(abs_col as u32, false, false);
                    stack.push(ExprFragment::new(format_cell_ref_from_field(
                        abs_row as u32,
                        col_field,
                    )));
                }
            }
            // PtgAreaN: [rowFirst_off: i32][rowLast_off: i32][colFirst_off: i16][colLast_off: i16]
            0x2D | 0x4D | 0x6D => {
                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let Some(base) = base else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(12).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 12,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 12,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let row_first_off = i32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]) as i64;
                let row_last_off = i32::from_le_bytes([raw[4], raw[5], raw[6], raw[7]]) as i64;
                let col_first_off = i16::from_le_bytes([raw[8], raw[9]]) as i64;
                let col_last_off = i16::from_le_bytes([raw[10], raw[11]]) as i64;
                i = end;

                const MAX_ROW: i64 = 1_048_575;
                const MAX_COL: i64 = COL_INDEX_MASK as i64;
                let abs_row_first = base.row as i64 + row_first_off;
                let abs_row_last = base.row as i64 + row_last_off;
                let abs_col_first = base.col as i64 + col_first_off;
                let abs_col_last = base.col as i64 + col_last_off;

                if abs_row_first < 0
                    || abs_row_first > MAX_ROW
                    || abs_row_last < 0
                    || abs_row_last > MAX_ROW
                    || abs_col_first < 0
                    || abs_col_first > MAX_COL
                    || abs_col_last < 0
                    || abs_col_last > MAX_COL
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let is_single_cell =
                        abs_row_first == abs_row_last && abs_col_first == abs_col_last;
                    let is_value_class = (ptg & 0x60) == 0x40;

                    let mut text = String::new();
                    let mut prec = 100;
                    if is_value_class && !is_single_cell {
                        text.push('@');
                        prec = 70;
                    }
                    if is_single_cell {
                        push_cell_ref_from_field(
                            &mut text,
                            abs_row_first as u32,
                            encode_col_field(abs_col_first as u32, false, false),
                        );
                    } else {
                        push_cell_ref_from_field(
                            &mut text,
                            abs_row_first as u32,
                            encode_col_field(abs_col_first as u32, false, false),
                        );
                        text.push(':');
                        push_cell_ref_from_field(
                            &mut text,
                            abs_row_last as u32,
                            encode_col_field(abs_col_last as u32, false, false),
                        );
                    }
                    stack.push(ExprFragment {
                        text,
                        precedence: prec,
                        contains_union: false,
                        is_missing: false,
                    });
                }
            }
            // PtgRef3d: [ixti: u16][row: u32][col: u16]
            0x3A | 0x5A | 0x7A => {
                if rgce.len().saturating_sub(i) < 8 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let Some(ctx) = ctx else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(8).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 8,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 8,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let ixti = u16::from_le_bytes([raw[0], raw[1]]);
                let row0 = u32::from_le_bytes([raw[2], raw[3], raw[4], raw[5]]);
                let col_field = u16::from_le_bytes([raw[6], raw[7]]);
                i = end;

                let (workbook, first, last) =
                    ctx.extern_sheet_target(ixti)
                        .ok_or(DecodeError::UnknownPtg {
                            offset: ptg_offset,
                            ptg,
                        })?;
                let prefix = match workbook {
                    None => format_sheet_prefix(first, last),
                    Some(book) => format_external_sheet_prefix(book, first, last),
                };
                let mut text = prefix;
                push_cell_ref_from_field(&mut text, row0, col_field);
                stack.push(ExprFragment::new(text));
            }
            // PtgArea3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3B | 0x5B | 0x7B => {
                if rgce.len().saturating_sub(i) < 14 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 14,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let Some(ctx) = ctx else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(14).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 14,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 14,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let ixti = u16::from_le_bytes([raw[0], raw[1]]);
                let row_first0 = u32::from_le_bytes([raw[2], raw[3], raw[4], raw[5]]);
                let row_last0 = u32::from_le_bytes([raw[6], raw[7], raw[8], raw[9]]);
                let col_first = u16::from_le_bytes([raw[10], raw[11]]);
                let col_last = u16::from_le_bytes([raw[12], raw[13]]);
                i = end;

                let (workbook, first, last) =
                    ctx.extern_sheet_target(ixti)
                        .ok_or(DecodeError::UnknownPtg {
                            offset: ptg_offset,
                            ptg,
                        })?;
                let prefix = match workbook {
                    None => format_sheet_prefix(first, last),
                    Some(book) => format_external_sheet_prefix(book, first, last),
                };

                let is_single_cell = row_first0 == row_last0
                    && (col_first & COL_INDEX_MASK) == (col_last & COL_INDEX_MASK);
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                let mut prec = 100;
                if is_value_class && !is_single_cell {
                    text.push('@');
                    prec = 70;
                }
                text.push_str(&prefix);
                const MAX_ROW: u32 = 1_048_575;
                let col_first_idx = col_first & COL_INDEX_MASK;
                let col_last_idx = col_last & COL_INDEX_MASK;
                if row_first0 == 0 && row_last0 == MAX_ROW {
                    push_col_ref_from_field(&mut text, col_first);
                    text.push(':');
                    push_col_ref_from_field(&mut text, col_last);
                } else if col_first_idx == 0 && col_last_idx == COL_INDEX_MASK {
                    push_row_ref_from_field(&mut text, row_first0, col_first);
                    text.push(':');
                    push_row_ref_from_field(&mut text, row_last0, col_last);
                } else if is_single_cell {
                    push_cell_ref_from_field(&mut text, row_first0, col_first);
                } else {
                    push_cell_ref_from_field(&mut text, row_first0, col_first);
                    text.push(':');
                    push_cell_ref_from_field(&mut text, row_last0, col_last);
                }
                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgRefErr3d: [ixti: u16][row: u32][col: u16]
            0x3C | 0x5C | 0x7C => {
                if rgce.len().saturating_sub(i) < 8 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 8;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3D | 0x5D | 0x7D => {
                if rgce.len().saturating_sub(i) < 14 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 14,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 14;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgName: [nameId: u32][reserved: u16]
            0x23 | 0x43 | 0x63 => {
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let Some(ctx) = ctx else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(6).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 6,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let name_id = u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]);
                i = end; // nameId + reserved

                let def = ctx
                    .name_definition(name_id)
                    .ok_or(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    })?;

                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                let mut prec = 100;
                if is_value_class {
                    // Like value-class range tokens, a value-class name can require legacy
                    // implicit intersection (e.g. when the name refers to a multi-cell range).
                    // Emit an explicit `@` so the formula text preserves scalar semantics.
                    text.push('@');
                    prec = 70;
                }

                match &def.scope {
                    NameScope::Workbook => text.push_str(&def.name),
                    NameScope::Sheet(sheet) => {
                        fmt_sheet_name(&mut text, sheet);
                        text.push('!');
                        text.push_str(&def.name);
                    }
                }

                stack.push(ExprFragment {
                    text,
                    precedence: prec,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgNameX: [ixti: u16][nameIndex: u16]
            0x39 | 0x59 | 0x79 => {
                if rgce.len().saturating_sub(i) < 4 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 4,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let Some(ctx) = ctx else {
                    return Err(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                let end = i.checked_add(4).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 4,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 4,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let ixti = u16::from_le_bytes([raw[0], raw[1]]);
                let name_index = u16::from_le_bytes([raw[2], raw[3]]);
                i = end;

                let txt = ctx
                    .format_namex(ixti, name_index)
                    .ok_or(DecodeError::UnknownPtg {
                        offset: ptg_offset,
                        ptg,
                    })?;
                stack.push(ExprFragment::new(txt));
            }
            // PtgFunc: [iftab: u16] (argument count is implicit and fixed for the function).
            0x21 | 0x41 | 0x61 => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let iftab = u16::from_le_bytes([raw[0], raw[1]]);
                i = end;

                // `PtgFunc` does not store argc; it is implicit and requires fixed-arity function
                // metadata to decode.
                //
                // For strict decode APIs (`decode_rgce*`), missing/variable arity metadata is a
                // hard error.
                //
                // For best-effort decode APIs (`decode_formula_rgce*`), fall back to a parseable
                // call expression and keep decoding.
                let argc: usize;
                let name_owned;
                let name: &str = match formula_biff::function_spec_from_id(iftab) {
                    Some(spec) if spec.min_args == spec.max_args => {
                        argc = spec.min_args as usize;
                        spec.name
                    }
                    Some(spec) => {
                        if warnings.is_none() {
                            return Err(DecodeError::UnknownPtg {
                                offset: ptg_offset,
                                ptg,
                            });
                        }

                        if let Some(w) = warnings.as_deref_mut() {
                            w.push(DecodeWarning::DecodeFailed {
                                kind: DecodeFailureKind::UnknownPtg,
                                offset: ptg_offset,
                                ptg,
                            });
                        }

                        // Best-effort: if we have metadata but it's variable-arity, assume the
                        // minimum argument count (common when older Excel versions encoded
                        // fixed-arity calls that later gained optional args).
                        let min = spec.min_args as usize;
                        argc = if stack.len() < min {
                            stack.len().min(1)
                        } else {
                            min
                        };
                        spec.name
                    }
                    None => {
                        if warnings.is_none() {
                            return Err(DecodeError::UnknownPtg {
                                offset: ptg_offset,
                                ptg,
                            });
                        }

                        if let Some(w) = warnings.as_deref_mut() {
                            w.push(DecodeWarning::DecodeFailed {
                                kind: DecodeFailureKind::UnknownPtg,
                                offset: ptg_offset,
                                ptg,
                            });
                        }

                        // Best-effort: no argc metadata; assume unary if possible.
                        argc = stack.len().min(1);
                        match function_name(iftab) {
                            Some(name) => name,
                            None => {
                                name_owned = format!("_UNKNOWN_FUNC_0X{iftab:04X}");
                                &name_owned
                            }
                        }
                    }
                };

                if stack.len() < argc {
                    return Err(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    });
                }

                let mut args = Vec::new();
                let _ = args.try_reserve_exact(argc);
                for _ in 0..argc {
                    args.push(stack.pop().ok_or(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?);
                }
                args.reverse();
                stack.push(format_function_call(name, args));
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let end = i.checked_add(3).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let raw = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let argc = raw[0] as usize;
                let iftab = u16::from_le_bytes([raw[1], raw[2]]);
                i = end;

                if stack.len() < argc {
                    return Err(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    });
                }

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name (typically from `PtgNameX`), followed by args.
                if iftab == 0x00FF {
                    if argc == 0 {
                        return Err(DecodeError::StackUnderflow {
                            offset: ptg_offset,
                            ptg,
                        });
                    }

                    let func_name = stack.pop().ok_or(DecodeError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?.text;
                    let mut args = Vec::new();
                    let _ = args.try_reserve_exact(argc.saturating_sub(1));
                    for _ in 0..argc.saturating_sub(1) {
                        args.push(stack.pop().ok_or(DecodeError::StackUnderflow {
                            offset: ptg_offset,
                            ptg,
                        })?);
                    }
                    args.reverse();
                    stack.push(format_function_call(&func_name, args));
                } else {
                    let name_owned;
                    let name = match function_name(iftab) {
                        Some(name) => name,
                        None => {
                            if warnings.is_none() {
                                return Err(DecodeError::UnknownPtg {
                                    offset: ptg_offset,
                                    ptg,
                                });
                            }

                            if let Some(w) = warnings.as_deref_mut() {
                                w.push(DecodeWarning::DecodeFailed {
                                    kind: DecodeFailureKind::UnknownPtg,
                                    offset: ptg_offset,
                                    ptg,
                                });
                            }

                            name_owned = format!("_UNKNOWN_FUNC_0X{iftab:04X}");
                            &name_owned
                        }
                    };

                    let mut args = Vec::new();
                    let _ = args.try_reserve_exact(argc);
                    for _ in 0..argc {
                        args.push(stack.pop().ok_or(DecodeError::StackUnderflow {
                            offset: ptg_offset,
                            ptg,
                        })?);
                    }
                    args.reverse();
                    stack.push(format_function_call(name, args));
                }
            }
            _ => {
                return Err(DecodeError::UnknownPtg {
                    offset: ptg_offset,
                    ptg,
                })
            }
        }

        if stack.last().is_some_and(|s| s.text.len() > max_len) {
            return Err(DecodeError::OutputTooLarge {
                offset: ptg_offset,
                ptg,
                max_len,
            });
        }
    }

    if stack.len() == 1 {
        Ok(stack
            .pop()
            .map(|v| v.text)
            .unwrap_or_else(|| {
                debug_assert!(false, "stack length checked");
                String::new()
            }))
    } else if warnings.is_some() {
        // Best-effort mode: keep the decode parseable for diagnostics by returning the
        // top-of-stack fragment, even if the token stream left extra items on the stack.
        // Surface the mismatch via a structured warning so callers can detect partial output.
        if let Some(w) = warnings.as_deref_mut() {
            w.push(DecodeWarning::DecodeFailed {
                kind: DecodeFailureKind::StackNotSingular,
                offset: last_ptg_offset,
                ptg: last_ptg,
            });
        }
        Ok(stack.pop().map(|v| v.text).unwrap_or_default())
    } else {
        Err(DecodeError::StackNotSingular {
            offset: last_ptg_offset,
            ptg: last_ptg,
            stack_len: stack.len(),
        })
    }
}

fn decode_array_constant(
    rgcb: &[u8],
    pos: &mut usize,
    mut warnings: Option<&mut Vec<DecodeWarning>>,
) -> Option<String> {
    // MS-XLSB 2.5.198.8 PtgArray references an Array constant serialized in `rgcb`.
    // The exact structure differs from the BIFF8-era format (larger row/col counts),
    // but at a high level it is:
    //
    //   [cols_minus1: u16][rows_minus1: u16][values...]
    //
    // Values are stored row-major and each starts with a type byte:
    //   0x00 = empty
    //   0x01 = number (f64)
    //   0x02 = string ([cch: u16][utf16 chars...])
    //   0x04 = bool ([b: u8])
    //   0x10 = error ([code: u8])
    //
    // We decode a minimal subset that is sufficient for common array constants.

    let mut i = *pos;
    if remaining_len(rgcb, i) < 4 {
        return None;
    }

    let hdr_end = i.checked_add(4)?;
    let hdr = rgcb.get(i..hdr_end)?;
    let cols_minus1 = u16::from_le_bytes([hdr[0], hdr[1]]) as usize;
    let rows_minus1 = u16::from_le_bytes([hdr[2], hdr[3]]) as usize;
    i = hdr_end;

    // Values are u16 widened to usize, so `+ 1` cannot overflow.
    let cols = cols_minus1 + 1;
    let rows = rows_minus1 + 1;

    use core::fmt::Write as _;

    let mut out = String::new();
    out.push('{');

    for row in 0..rows {
        if row > 0 {
            out.push(';');
        }
        for col in 0..cols {
            if col > 0 {
                out.push(',');
            }
            if i >= rgcb.len() {
                return None;
            }
            let ty = *rgcb.get(i)?;
            i += 1;
            match ty {
                0x00 => {}
                0x01 => {
                    if remaining_len(rgcb, i) < 8 {
                        return None;
                    }
                    let mut bytes = [0u8; 8];
                    let end = i.checked_add(8)?;
                    let raw = rgcb.get(i..end)?;
                    bytes.copy_from_slice(raw);
                    i = end;
                    write!(&mut out, "{}", f64::from_le_bytes(bytes)).ok()?;
                }
                0x02 => {
                    if remaining_len(rgcb, i) < 2 {
                        return None;
                    }
                    let len_end = i.checked_add(2)?;
                    let len_bytes = rgcb.get(i..len_end)?;
                    let cch = u16::from_le_bytes([len_bytes[0], len_bytes[1]]) as usize;
                    i = len_end;
                    let byte_len = cch.checked_mul(2)?;
                    if remaining_len(rgcb, i) < byte_len {
                        return None;
                    }
                    let end = i.checked_add(byte_len)?;
                    let raw = rgcb.get(i..end)?;
                    i = end;
                    out.push('"');
                    let iter = raw
                        .chunks_exact(2)
                        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
                    for decoded in std::char::decode_utf16(iter) {
                        match decoded {
                            Ok(ch) => push_escaped_excel_double_quote_char(&mut out, ch),
                            Err(_) => push_escaped_excel_double_quote_char(&mut out, '\u{FFFD}'),
                        }
                    }
                    out.push('"');
                }
                0x04 => {
                    if remaining_len(rgcb, i) < 1 {
                        return None;
                    }
                    let b = *rgcb.get(i)?;
                    i += 1;
                    out.push_str(if b == 0 { "FALSE" } else { "TRUE" });
                }
                0x10 => {
                    if remaining_len(rgcb, i) < 1 {
                        return None;
                    }
                    let code_offset = i;
                    let code = *rgcb.get(i)?;
                    i += 1;
                    match xlsb_error_literal(code) {
                        Some(lit) => out.push_str(lit),
                        None => {
                            if let Some(warnings) = warnings.as_deref_mut() {
                                warnings.push(DecodeWarning::UnknownArrayErrorCode {
                                    code,
                                    offset: code_offset,
                                });
                            }
                            out.push_str("#UNKNOWN!");
                        }
                    }
                }
                _ => return None,
            }
        }
    }

    *pos = i;
    out.push('}');
    Some(out)
}

/// Scan a nested BIFF12 token subexpression (e.g. the payload of `PtgMemFunc`) and advance the
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
    rgce_base_offset: usize,
    ctx: Option<&WorkbookContext>,
) -> Result<(), DecodeError> {
    fn has_remaining(buf: &[u8], i: usize, needed: usize) -> bool {
        remaining_len(buf, i) >= needed
    }

    let mut i = 0usize;
    while i < rgce.len() {
        let ptg_offset = rgce_base_offset.checked_add(i).unwrap_or(usize::MAX);
        let Some(&ptg) = rgce.get(i) else {
            debug_assert!(false, "rgce cursor out of bounds (i={i}, len={})", rgce.len());
            return Err(DecodeError::UnexpectedEof {
                offset: ptg_offset,
                ptg: 0,
                needed: 1,
                remaining: 0,
            });
        };
        i += 1;

        match ptg {
            // PtgExp / PtgTbl: [row: u16][col: u16]
            //
            // These tokens are used for shared formulas / data tables and can appear inside
            // non-printing `PtgMem*` subexpressions. We don't decode them to text here, but we
            // must skip their payload to keep scanning aligned so we can still find any nested
            // `PtgArray` tokens.
            0x01 | 0x02 => {
                if !has_remaining(rgce, i, 4) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 4,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 4;
            }

            // PtgArray (any class): [unused: 7 bytes] + array constant in rgcb.
            0x20 | 0x40 | 0x60 => {
                if !has_remaining(rgce, i, 7) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 7,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 7;
                let _ = decode_array_constant(rgcb, rgcb_pos, None).ok_or(
                    DecodeError::InvalidConstant {
                        offset: ptg_offset,
                        ptg,
                        value: 0xFF,
                    },
                )?;
            }

            // Binary operators and simple operators with no payload.
            0x03..=0x16 | 0x2F => {}

            // PtgStr: [cch: u16][utf16 chars...]
            0x17 => {
                if !has_remaining(rgce, i, 2) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let len_bytes = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let cch = u16::from_le_bytes([len_bytes[0], len_bytes[1]]) as usize;
                i = end;
                // `cch` is a u16 widened to usize, so `* 2` cannot overflow.
                let byte_len = cch * 2;
                if !has_remaining(rgce, i, byte_len) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: byte_len,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += byte_len;
            }

            // PtgExtend* (structured refs): [etpg: u8][payload...]
            0x18 | 0x38 | 0x58 => {
                if !has_remaining(rgce, i, 1) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let Some(&etpg) = rgce.get(i) else {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                };
                i += 1;
                match etpg {
                    // etpg=0x19 is the structured reference payload (PtgList).
                    //
                    // MS-XLSB documents a fixed 12-byte payload, but some producers emit extra
                    // prefix bytes (e.g. 2/4 bytes). Use the same best-effort alignment heuristic
                    // as the main decoder so we can keep scanning aligned and still find later
                    // nested `PtgArray` tokens.
                    0x19 => {
                        let default_ctx = WorkbookContext::default();
                        let ctx_for_scoring = Some(ctx.unwrap_or(&default_ctx));

                        let remaining = rgce.len().saturating_sub(i);
                        let tail = rgce.get(i..).unwrap_or(&[]);
                        let Some(payload_len) = ptg_list_payload_len_best_effort(tail, ctx_for_scoring)
                        else {
                            return Err(DecodeError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: 12,
                                remaining,
                            });
                        };
                        if remaining < payload_len {
                            return Err(DecodeError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: payload_len,
                                remaining,
                            });
                        }
                        i += payload_len;
                    }
                    // Unknown extend subtype: stop scanning to avoid desync/false positives.
                    _ => break,
                }
            }

            // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
            0x19 => {
                if !has_remaining(rgce, i, 3) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let end = i.checked_add(3).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let hdr = rgce.get(i..end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 3,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let grbit = hdr[0];
                let w_attr = u16::from_le_bytes([hdr[1], hdr[2]]) as usize;
                i = end;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    // `w_attr` is a u16 widened to usize, so `* 2` cannot overflow.
                    let needed = w_attr * 2;
                    if !has_remaining(rgce, i, needed) {
                        return Err(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed,
                            remaining: rgce.len().saturating_sub(i),
                        });
                    }
                    i += needed;
                }
            }

            // PtgErr: [code: u8]
            0x1C => {
                if !has_remaining(rgce, i, 1) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 1;
            }
            // PtgBool: [b: u8]
            0x1D => {
                if !has_remaining(rgce, i, 1) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 1;
            }
            // PtgInt: [n: u16]
            0x1E => {
                if !has_remaining(rgce, i, 2) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 2;
            }
            // PtgNum: [f64]
            0x1F => {
                if !has_remaining(rgce, i, 8) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 8;
            }

            // PtgFunc: [iftab: u16]
            0x21 | 0x41 | 0x61 => {
                if !has_remaining(rgce, i, 2) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 2;
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if !has_remaining(rgce, i, 3) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 3;
            }

            // PtgName: [nameIndex: u32][unused: u16]
            0x23 | 0x43 | 0x63 => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 6;
            }

            // PtgRef: [row: u32][col: u16]
            0x24 | 0x44 | 0x64 => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 6;
            }
            // PtgArea: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x25 | 0x45 | 0x65 => {
                if !has_remaining(rgce, i, 12) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 12;
            }

            // PtgMem* tokens: [cce: u16][subexpression...]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if !has_remaining(rgce, i, 2) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cce_end = i.checked_add(2).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let cce_bytes = rgce.get(i..cce_end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 2,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let cce = u16::from_le_bytes([cce_bytes[0], cce_bytes[1]]) as usize;
                i = cce_end;
                if !has_remaining(rgce, i, cce) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let subexpr_end = i.checked_add(cce).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: cce,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                let subexpr = rgce.get(i..subexpr_end).ok_or(DecodeError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: cce,
                    remaining: rgce.len().saturating_sub(i),
                })?;
                consume_rgcb_arrays_in_subexpression(
                    subexpr,
                    rgcb,
                    rgcb_pos,
                    rgce_base_offset.checked_add(i).unwrap_or(usize::MAX),
                    ctx,
                )?;
                i = subexpr_end;
            }

            // PtgRefErr: [row: u32][col: u16]
            0x2A | 0x4A | 0x6A => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 6;
            }
            // PtgAreaErr: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                if !has_remaining(rgce, i, 12) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 12;
            }

            // PtgRefN: [row_off: i32][col_off: i16]
            0x2C | 0x4C | 0x6C => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 6;
            }
            // PtgAreaN: [rowFirst_off: i32][rowLast_off: i32][colFirst_off: i16][colLast_off: i16]
            0x2D | 0x4D | 0x6D => {
                if !has_remaining(rgce, i, 12) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 12;
            }

            // PtgNameX: [ixti: u16][nameIndex: u16]
            0x39 | 0x59 | 0x79 => {
                if !has_remaining(rgce, i, 4) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 4,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 4;
            }

            // PtgRef3d: [ixti: u16][row: u32][col: u16]
            0x3A | 0x5A | 0x7A => {
                if !has_remaining(rgce, i, 8) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 8;
            }
            // PtgArea3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3B | 0x5B | 0x7B => {
                if !has_remaining(rgce, i, 14) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 14,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 14;
            }
            // PtgRefErr3d: [ixti: u16][row: u32][col: u16]
            0x3C | 0x5C | 0x7C => {
                if !has_remaining(rgce, i, 8) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 8;
            }
            // PtgAreaErr3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3D | 0x5D | 0x7D => {
                if !has_remaining(rgce, i, 14) {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 14,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 14;
            }

            // Unknown ptg: stop scanning to avoid desync/false positives.
            _ => break,
        }
    }

    Ok(())
}

fn format_cell_ref(row1: u64, col: u32, flags: u8) -> String {
    let mut out = String::new();
    let abs_col = flags & 0x80 != 0x80;
    let abs_row = flags & 0x40 != 0x40;
    formula_model::push_a1_cell_ref_row1(row1, col, abs_col, abs_row, &mut out);
    out
}

fn push_cell_ref_from_field(out: &mut String, row0: u32, col_field: u16) {
    let row1 = u64::from(row0) + 1;
    let col = (col_field & 0x3FFF) as u32;
    let col_relative = (col_field & 0x8000) == 0x8000;
    let row_relative = (col_field & 0x4000) == 0x4000;
    formula_model::push_a1_cell_ref_row1(row1, col, !col_relative, !row_relative, out);
}

fn format_cell_ref_from_field(row0: u32, col_field: u16) -> String {
    let mut out = String::new();
    push_cell_ref_from_field(&mut out, row0, col_field);
    out
}

fn push_col_ref_from_field(out: &mut String, col_field: u16) {
    let col = (col_field & COL_INDEX_MASK) as u32;
    let col_relative = (col_field & COL_RELATIVE_MASK) == COL_RELATIVE_MASK;
    formula_model::push_a1_col_ref(col, !col_relative, out);
}

fn push_row_ref_from_field(out: &mut String, row0: u32, col_field: u16) {
    let row1 = u64::from(row0) + 1;
    let row_relative = (col_field & ROW_RELATIVE_MASK) == ROW_RELATIVE_MASK;
    formula_model::push_a1_row_ref_row1(row1, !row_relative, out);
}

fn decode_ptg_list_payload_best_effort(
    payload: &[u8; 12],
    ctx: Option<&WorkbookContext>,
) -> PtgListDecoded {
    // There are multiple "in the wild" encodings for the 12-byte PtgList payload (table refs /
    // structured references). We try a handful of plausible layouts and prefer the one that
    // resolves cleanly against the provided workbook context.
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

    let mut candidates = decode_ptg_list_payload_candidates(payload);

    if let Some(ctx) = ctx {
        candidates.sort_by_key(|cand| std::cmp::Reverse(score_ptg_list_candidate(cand, ctx)));
    }

    candidates[0]
}

/// Best-effort determine how many bytes a `PtgExtend` structured reference payload (`etpg=0x19`,
/// `PtgList`) consumes.
///
/// MS-XLSB documents a fixed 12-byte payload for structured references. In practice, some XLSB
/// producers appear to insert extra prefix bytes before the 12-byte core payload (e.g. alignment
/// padding or undocumented fields).
///
/// The formula decoder is able to interpret multiple observed core payload layouts (A/B/C). Shared
/// formula materialization must also be able to *skip* the payload correctly to keep the rgce stream
/// aligned, even when the core payload is not at the canonical start position.
///
/// This helper chooses the most plausible core payload alignment using the same context-based
/// scoring heuristics as the decoder, and returns the total number of bytes to consume (prefix +
/// core). The caller should copy the raw bytes verbatim; structured references do not embed
/// relative row/col offsets that require shifting during shared-formula materialization.
pub(crate) fn ptg_list_payload_len_best_effort(
    data: &[u8],
    ctx: Option<&WorkbookContext>,
) -> Option<usize> {
    const CORE_LEN: usize = 12;
    if data.len() < CORE_LEN {
        return None;
    }

    // Common "prefix padding" sizes seen in other BIFF12 record layouts are 2 and 4 bytes. Prefer
    // the canonical alignment (0) unless the context-based scoring strongly suggests otherwise.
    const OFFSETS: [usize; 3] = [0, 2, 4];

    let mut best_score = i32::MIN;
    let mut best_len = CORE_LEN;

    for &offset in &OFFSETS {
        let Some(end) = offset.checked_add(CORE_LEN) else {
            continue;
        };
        let Some(window) = data.get(offset..end) else {
            continue;
        };

        let mut payload = [0u8; CORE_LEN];
        payload.copy_from_slice(window);
        let decoded = decode_ptg_list_payload_best_effort(&payload, ctx);

        // Reuse the decoder's scoring heuristic when workbook context is available. Without
        // context, fall back to choosing the canonical alignment.
        let mut score = if let Some(ctx) = ctx {
            score_ptg_list_candidate(&decoded, ctx)
        } else {
            0
        };

        // Strongly prefer canonical alignment.
        score -= (offset as i32) * 10;

        // Treat non-zero prefix bytes as a strong signal that the candidate alignment is wrong;
        // most observed padding prefixes are zero-filled.
        if offset > 0 && data[..offset].iter().any(|&b| b != 0) {
            score -= 1_000;
        }

        if score > best_score {
            best_score = score;
            best_len = offset + CORE_LEN;
        }
    }

    Some(best_len)
}

fn score_ptg_list_candidate(cand: &PtgListDecoded, ctx: &WorkbookContext) -> i32 {
    let mut score = 0i32;

    if ctx.table_name(cand.table_id).is_some() {
        score += 100;
    }

    let col_first = cand.col_first;
    let col_last = cand.col_last;

    // Column id `0` is treated as a sentinel for "all columns"; seeing it on only one side is
    // usually a sign we've chosen the wrong payload layout.
    if (col_first == 0) ^ (col_last == 0) {
        score -= 50;
    }

    if col_first == 0 && col_last == 0 {
        score += 1;
        return score;
    }

    if col_first != 0 && ctx.table_column_name(cand.table_id, col_first).is_some() {
        score += 10;
    }
    if col_last != 0 && ctx.table_column_name(cand.table_id, col_last).is_some() {
        score += 10;
    }

    // Table column ids are typically small.
    if col_first <= 16_384 {
        score += 1;
    }
    if col_last <= 16_384 {
        score += 1;
    }

    score
}

fn function_name(iftab: u16) -> Option<&'static str> {
    formula_biff::function_id_to_name(iftab)
}

/// 0-indexed cell coordinate used as the base for relative reference encoding.
///
/// Note: BIFF12 (XLSB) stores row/col indices directly in `PtgRef`/`PtgArea` tokens and uses
/// flags to indicate relative vs absolute behavior when formulas are moved/copied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellCoord {
    pub row: u32,
    pub col: u32,
}

impl CellCoord {
    pub const fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }
}

/// Encoded formula token stream.
///
/// XLSB stores the primary token stream as `rgce`. Some tokens (arrays) require a trailing
/// `rgcb` buffer. We keep both so callers can emit valid records.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncodedRgce {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum EncodeError {
    #[error("invalid formula: {0}")]
    Parse(String),
    #[error("unsupported expression in BIFF12 encoder: {0}")]
    Unsupported(&'static str),
    #[error("unknown sheet reference: {0}")]
    UnknownSheet(String),
    #[error("unknown name: {name}")]
    UnknownName { name: String },
    #[error("unknown function: {name}")]
    UnknownFunction { name: String },
}

/// Encode an A1-style formula to an XLSB `rgce` token stream using the full `formula-engine` AST
/// and workbook context for 3D references, defined names, and add-in/UDF calls.
///
/// The input may include a leading `=`; it will be ignored.
#[cfg(feature = "write")]
pub fn encode_rgce_with_context_ast(
    formula: &str,
    ctx: &WorkbookContext,
    base: CellCoord,
) -> Result<EncodedRgce, EncodeError> {
    encode_ast::encode_rgce_with_context_ast(formula, ctx, base)
}

/// Encode an A1-style formula to an XLSB `rgce` token stream using workbook context and a sheet
/// name for resolving table-less structured references like `[@Col]`.
///
/// The input may include a leading `=`; it will be ignored.
#[cfg(feature = "write")]
pub fn encode_rgce_with_context_ast_in_sheet(
    formula: &str,
    ctx: &WorkbookContext,
    sheet: &str,
    base: CellCoord,
) -> Result<EncodedRgce, EncodeError> {
    encode_ast::encode_rgce_with_context_ast_in_sheet(formula, ctx, sheet, base)
}

#[cfg(feature = "write")]
mod encode_ast {
    use super::{
        format_external_key, format_external_workbook_key,
        push_utf16le_u16_len_with_rollback,
        ptg_with_class, ArrayConst, ArrayElem, CellCoord, EncodeError, EncodedRgce, PtgClass,
        WorkbookContext, COL_INDEX_MASK, COL_RELATIVE_MASK, PTG_AREA, PTG_AREA3D, PTG_FUNCVAR,
        PTG_NAME, PTG_NAMEX, PTG_REF, PTG_REF3D, PTG_SPILL, PTG_UMINUS, PTG_UPLUS,
        ROW_RELATIVE_MASK,
    };
    use crate::errors::xlsb_error_code_from_literal;

    use formula_engine as fe;
    use fe::structured_refs::{parse_structured_ref, StructuredColumns, StructuredRefItem};
    use formula_model::{column_label_to_index, sheet_name_eq_case_insensitive};

    const PTG_ERR: u8 = 0x1C;
    const PTG_BOOL: u8 = 0x1D;
    const PTG_INT: u8 = 0x1E;
    const PTG_NUM: u8 = 0x1F;
    const PTG_STR: u8 = 0x17;
    const PTG_FUNC: u8 = 0x21;
    const PTG_MISS_ARG: u8 = 0x16;
    const PTG_PERCENT: u8 = 0x14;

    // Binary operators.
    const PTG_ADD: u8 = 0x03;
    const PTG_SUB: u8 = 0x04;
    const PTG_MUL: u8 = 0x05;
    const PTG_DIV: u8 = 0x06;
    const PTG_POW: u8 = 0x07;
    const PTG_CONCAT: u8 = 0x08;
    const PTG_LT: u8 = 0x09;
    const PTG_LE: u8 = 0x0A;
    const PTG_EQ: u8 = 0x0B;
    const PTG_GT: u8 = 0x0C;
    const PTG_GE: u8 = 0x0D;
    const PTG_NE: u8 = 0x0E;
    const PTG_INTERSECT: u8 = 0x0F;
    const PTG_UNION: u8 = 0x10;
    const PTG_RANGE: u8 = 0x11;

    #[derive(Debug, Clone, PartialEq, Eq)]
    enum SheetSpec {
        Current,
        Single(String),
        Range(String, String),
    }

    #[derive(Debug, Clone, PartialEq)]
    struct CellRefInfo {
        sheet: SheetSpec,
        row: u32,
        col: u32,
        abs_row: bool,
        abs_col: bool,
    }

    pub(super) fn encode_rgce_with_context_ast(
        formula: &str,
        ctx: &WorkbookContext,
        base: CellCoord,
    ) -> Result<EncodedRgce, EncodeError> {
        encode_rgce_with_context_ast_impl(formula, ctx, None, base)
    }

    pub(super) fn encode_rgce_with_context_ast_in_sheet(
        formula: &str,
        ctx: &WorkbookContext,
        sheet: &str,
        base: CellCoord,
    ) -> Result<EncodedRgce, EncodeError> {
        encode_rgce_with_context_ast_impl(formula, ctx, Some(sheet), base)
    }

    fn encode_rgce_with_context_ast_impl(
        formula: &str,
        ctx: &WorkbookContext,
        sheet: Option<&str>,
        base: CellCoord,
    ) -> Result<EncodedRgce, EncodeError> {
        let ast = fe::parse_formula(formula, fe::ParseOptions::default()).map_err(|e| {
            EncodeError::Parse(format!(
                "{} (span {}..{})",
                e.message, e.span.start, e.span.end
            ))
        })?;

        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        emit_expr(&ast.expr, ctx, sheet, base, &mut rgce, &mut rgcb)?;
        Ok(EncodedRgce { rgce, rgcb })
    }

    fn emit_expr(
        expr: &fe::Expr,
        ctx: &WorkbookContext,
        sheet: Option<&str>,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        rgcb: &mut Vec<u8>,
    ) -> Result<(), EncodeError> {
        match expr {
            fe::Expr::Number(raw) => {
                let n: f64 = raw
                    .parse()
                    .map_err(|_| EncodeError::Parse(format!("invalid number literal: {raw}")))?;
                emit_number(n, rgce);
            }
            fe::Expr::String(s) => {
                let start_len = rgce.len();
                rgce.push(PTG_STR);
                push_utf16le_u16_len_with_rollback(rgce, start_len, s, "string literal too long")?;
            }
            fe::Expr::Boolean(b) => {
                rgce.push(PTG_BOOL);
                rgce.push(u8::from(*b));
            }
            fe::Expr::Error(raw) => {
                let code = xlsb_error_code_from_literal(raw).ok_or_else(|| {
                    EncodeError::Parse(format!("unsupported error literal: {raw}"))
                })?;
                rgce.push(PTG_ERR);
                rgce.push(code);
            }
            fe::Expr::Missing => {
                rgce.push(PTG_MISS_ARG);
            }
            fe::Expr::CellRef(r) => {
                emit_cell_ref(r, ctx, base, rgce, PtgClass::Ref)?;
            }
            fe::Expr::NameRef(name) => {
                emit_defined_name(name, ctx, rgce, PtgClass::Ref)?;
            }
            fe::Expr::Array(arr) => {
                let array_const = array_literal_to_const(arr)?;
                // Reuse the existing `PtgArray` + rgcb serialization.
                super::emit_array(&array_const, rgce, rgcb)?;
            }
            fe::Expr::FunctionCall(call) => {
                for arg in &call.args {
                    if matches!(arg, fe::Expr::Missing) {
                        rgce.push(PTG_MISS_ARG);
                    } else {
                        emit_expr(arg, ctx, sheet, base, rgce, rgcb)?;
                    }
                }
                match emit_function_call(&call.name, call.args.len(), ctx, rgce) {
                    Ok(()) => {}
                    Err(EncodeError::UnknownFunction { .. }) => {
                        // `formula-engine` parses some callable-reference patterns like `A1(1)`
                        // as a `FunctionCall` with the name `A1`. In Excel this syntax is used to
                        // call a cell containing a `LAMBDA`.
                        //
                        // If we can't resolve the identifier as a built-in function, extern name,
                        // or defined name, treat it as a cell reference and encode using the
                        // UDF sentinel (iftab=255) with a `PtgRef` callee.
                        if let Some(callee) = cell_ref_info_from_function_name(&call.name)? {
                            emit_cell_ref_info(&callee, ctx, rgce, PtgClass::Ref)?;
                            emit_call_udf_sentinel(call.args.len(), rgce)?;
                        } else {
                            return Err(EncodeError::UnknownFunction {
                                name: call.name.name_upper.clone(),
                            });
                        }
                    }
                    Err(e) => return Err(e),
                }
            }
            fe::Expr::Call(call) => {
                // Encode "callable expression" invocations such as:
                // - `Sheet1!MyLambda(1,2)` (sheet-scoped defined name call)
                // - `A1(3)` (cell containing a LAMBDA)
                // - `LAMBDA(x,x+1)(3)` (inline lambda invocation)
                //
                // BIFF uses the UDF sentinel (iftab=255) with the "function name" stored as the
                // top-of-stack item (callee expression) immediately before `PtgFuncVar`.
                for arg in &call.args {
                    if matches!(arg, fe::Expr::Missing) {
                        rgce.push(PTG_MISS_ARG);
                    } else {
                        emit_expr(arg, ctx, sheet, base, rgce, rgcb)?;
                    }
                }

                emit_expr(&call.callee, ctx, sheet, base, rgce, rgcb)?;
                emit_call_udf_sentinel(call.args.len(), rgce)?;
            }
            fe::Expr::Unary(u) => match u.op {
                fe::UnaryOp::ImplicitIntersection => {
                    emit_reference_expr(&u.expr, ctx, sheet, base, rgce, rgcb, PtgClass::Value)?;
                }
                fe::UnaryOp::Plus => {
                    emit_expr(&u.expr, ctx, sheet, base, rgce, rgcb)?;
                    rgce.push(PTG_UPLUS);
                }
                fe::UnaryOp::Minus => {
                    emit_expr(&u.expr, ctx, sheet, base, rgce, rgcb)?;
                    rgce.push(PTG_UMINUS);
                }
            },
            fe::Expr::Postfix(p) => {
                emit_expr(&p.expr, ctx, sheet, base, rgce, rgcb)?;
                match p.op {
                    fe::PostfixOp::Percent => rgce.push(PTG_PERCENT),
                    fe::PostfixOp::SpillRange => rgce.push(PTG_SPILL),
                }
            }
            fe::Expr::Binary(b) => match b.op {
                fe::BinaryOp::Range => emit_range_binary(b, ctx, sheet, base, rgce, rgcb)?,
                _ => {
                    emit_expr(&b.left, ctx, sheet, base, rgce, rgcb)?;
                    emit_expr(&b.right, ctx, sheet, base, rgce, rgcb)?;
                    rgce.push(match b.op {
                        fe::BinaryOp::Add => PTG_ADD,
                        fe::BinaryOp::Sub => PTG_SUB,
                        fe::BinaryOp::Mul => PTG_MUL,
                        fe::BinaryOp::Div => PTG_DIV,
                        fe::BinaryOp::Pow => PTG_POW,
                        fe::BinaryOp::Concat => PTG_CONCAT,
                        fe::BinaryOp::Lt => PTG_LT,
                        fe::BinaryOp::Le => PTG_LE,
                        fe::BinaryOp::Eq => PTG_EQ,
                        fe::BinaryOp::Gt => PTG_GT,
                        fe::BinaryOp::Ge => PTG_GE,
                        fe::BinaryOp::Ne => PTG_NE,
                        fe::BinaryOp::Intersect => PTG_INTERSECT,
                        fe::BinaryOp::Union => PTG_UNION,
                        fe::BinaryOp::Range => PTG_RANGE,
                    });
                }
            },
            fe::Expr::ColRef(r) => {
                emit_col_ref(r, ctx, base, rgce, PtgClass::Ref)?;
            }
            fe::Expr::RowRef(r) => {
                emit_row_ref(r, ctx, base, rgce, PtgClass::Ref)?;
            }
            fe::Expr::StructuredRef(r) => {
                emit_structured_ref(r, ctx, sheet, base, rgce, PtgClass::Ref)?;
            }
            fe::Expr::FieldAccess(_) => {
                return Err(EncodeError::Unsupported("field access"));
            }
        }

        Ok(())
    }

    fn emit_number(n: f64, rgce: &mut Vec<u8>) {
        if n.fract() == 0.0 && (0.0..=65535.0).contains(&n) {
            rgce.push(PTG_INT);
            rgce.extend_from_slice(&(n as u16).to_le_bytes());
        } else {
            rgce.push(PTG_NUM);
            rgce.extend_from_slice(&n.to_le_bytes());
        }
    }

    // --- Structured references / table refs (PtgExtend etpg=0x19) -----------------------------

    const PTG_EXTEND: u8 = 0x18;
    const ETPG_LIST: u8 = 0x19;

    const FLAG_ALL: u16 = 0x0001;
    const FLAG_HEADERS: u16 = 0x0002;
    const FLAG_DATA: u16 = 0x0004;
    const FLAG_TOTALS: u16 = 0x0008;
    const FLAG_THIS_ROW: u16 = 0x0010;

    fn emit_structured_ref(
        r: &fe::StructuredRef,
        ctx: &WorkbookContext,
        sheet: Option<&str>,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        if r.workbook.is_some() {
            // BIFF12 structured references (`PtgList`) encode only a table id + column selectors.
            // Cross-workbook structured references would require a different encoding that ties the
            // table to an external workbook/sheet. Until the workbook context can resolve those,
            // keep rejecting them.
            return Err(EncodeError::Unsupported(
                "workbook-qualified structured references",
            ));
        }

        let sheet_qualifier = match r.sheet.as_ref() {
            None => None,
            Some(sheet_ref) => match sheet_ref.as_single_sheet() {
                Some(name) => Some(name),
                None => {
                    return Err(EncodeError::Unsupported(
                        "3D sheet-range structured references",
                    ))
                }
            },
        };

        let table_id = match r.table.as_deref() {
            Some(table_name) => {
                let table_id = ctx
                    .table_id_by_name(table_name)
                    .ok_or_else(|| EncodeError::Parse(format!("unknown table: {table_name}")))?;

                if let Some(sheet_name) = sheet_qualifier {
                    if let Some(false) = ctx.table_is_on_sheet(table_id, sheet_name) {
                        return Err(EncodeError::Parse(format!(
                            "structured reference table '{table_name}' is not on sheet '{sheet_name}'",
                        )));
                    }
                }

                table_id
            }
            None => {
                if let Some(sheet_name) = sheet_qualifier {
                    // `Sheet1![@Col]`-style references: use the explicit sheet qualifier along with
                    // the origin cell to infer the containing table.
                    if let Some(table_id) = ctx.table_id_for_cell(sheet_name, base.row, base.col) {
                        table_id
                    } else if let Some(table_id) = ctx.single_table_id() {
                        // If we know the sheet containing the (only) table and it does not match
                        // the sheet qualifier, treat the reference as invalid. BIFF12 `PtgList`
                        // tokens do not encode sheet qualifiers, so this avoids silently dropping
                        // a mismatched qualifier and changing semantics.
                        if let Some(false) = ctx.table_is_on_sheet(table_id, sheet_name) {
                            let table_name = ctx.table_name(table_id).unwrap_or("Table");
                            return Err(EncodeError::Parse(format!(
                                "structured reference table '{table_name}' is not on sheet '{sheet_name}'",
                            )));
                        }

                        // If the workbook context has a registered table range on this sheet, then
                        // failing to find the table by cell implies the base cell is outside the
                        // table. `[@Col]`-style structured references require the origin cell to
                        // be inside the table range, so error rather than guessing.
                        if let Some(true) = ctx.table_is_on_sheet(table_id, sheet_name) {
                            return Err(EncodeError::Parse(format!(
                                "cannot infer table for structured reference without an explicit table name at '{sheet_name}'!R{}C{} (cell must be inside exactly one table)",
                                base.row.checked_add(1).unwrap_or(u32::MAX),
                                base.col.checked_add(1).unwrap_or(u32::MAX)
                            )));
                        }

                        table_id
                    } else {
                        return Err(EncodeError::Parse(format!(
                            "cannot infer table for structured reference without an explicit table name at '{sheet_name}'!R{}C{} (cell must be inside exactly one table)",
                            base.row.checked_add(1).unwrap_or(u32::MAX),
                            base.col.checked_add(1).unwrap_or(u32::MAX)
                        )));
                    }
                } else {
                    // Always prefer the "single table in workbook" heuristic when possible. This
                    // is the only inference path when we lack sheet+range information, and is
                    // consistent with Excel's `[@Col]` / `[@]` shorthand in workbooks where the
                    // target table is unambiguous.
                    if let Some(table_id) = ctx.single_table_id() {
                        table_id
                    } else if let Some(sheet) = sheet {
                        ctx.table_id_for_cell(sheet, base.row, base.col).ok_or_else(|| {
                            EncodeError::Parse(format!(
                                "cannot infer table for structured reference without an explicit table name at '{sheet}'!R{}C{} (cell must be inside exactly one table)",
                                base.row.checked_add(1).unwrap_or(u32::MAX),
                                base.col.checked_add(1).unwrap_or(u32::MAX)
                            ))
                        })?
                    } else {
                        return Err(EncodeError::Parse(
                            "structured references without an explicit table name are ambiguous; specify TableName[...]"
                                .to_string(),
                        ));
                    }
                }
            }
        };

        let (flags, col_first, col_last) = encode_structured_ref_spec(r, table_id, ctx)?;

        rgce.push(ptg_with_class(PTG_EXTEND, class));
        rgce.push(ETPG_LIST);
        rgce.extend_from_slice(&table_id.to_le_bytes());
        rgce.extend_from_slice(&flags.to_le_bytes());
        rgce.extend_from_slice(&col_first.to_le_bytes());
        rgce.extend_from_slice(&col_last.to_le_bytes());
        rgce.extend_from_slice(&0u16.to_le_bytes());
        Ok(())
    }

    fn encode_structured_ref_spec(
        r: &fe::StructuredRef,
        table_id: u32,
        ctx: &WorkbookContext,
    ) -> Result<(u16, u16, u16), EncodeError> {
        // `formula-engine` lexes/parses structured references and stores the bracket contents as
        // `r.spec`. Re-parse via the authoritative structured-ref parser to avoid drift in
        // edge cases (nested brackets vs escaped `]]`).
        let mut text = String::new();
        if let Some(table) = &r.table {
            text.push_str(table);
        }
        text.push('[');
        text.push_str(&r.spec);
        text.push(']');

        let (sref, end) = parse_structured_ref(&text, 0)
            .ok_or_else(|| EncodeError::Parse(format!("invalid structured reference: {text}")))?;
        if end != text.len() {
            return Err(EncodeError::Parse(format!(
                "invalid structured reference: {text}"
            )));
        }
        if sref.table_name.as_deref() != r.table.as_deref() {
            return Err(EncodeError::Parse(format!(
                "invalid structured reference: {text}"
            )));
        }

        let mut flags = 0u16;
        for item in &sref.items {
            flags |= match item {
                StructuredRefItem::All => FLAG_ALL,
                StructuredRefItem::Headers => FLAG_HEADERS,
                StructuredRefItem::Data => FLAG_DATA,
                StructuredRefItem::Totals => FLAG_TOTALS,
                StructuredRefItem::ThisRow => FLAG_THIS_ROW,
            };
        }
        if (flags & FLAG_THIS_ROW) != 0 && (flags & !FLAG_THIS_ROW) != 0 {
            return Err(EncodeError::Unsupported(
                "structured references combining #This Row with other items",
            ));
        }

        let (col_first, col_last) = match &sref.columns {
            StructuredColumns::All => (0u16, 0u16),
            StructuredColumns::Single(col) => {
                let id = structured_ref_column_id(ctx, table_id, col)?;
                (id, id)
            }
            StructuredColumns::Range { start, end } => {
                let first = structured_ref_column_id(ctx, table_id, start)?;
                let last = structured_ref_column_id(ctx, table_id, end)?;
                (first, last)
            }
            StructuredColumns::Multi(_) => {
                return Err(EncodeError::Unsupported(
                    "structured references selecting multiple non-contiguous columns",
                ));
            }
        };

        Ok((flags, col_first, col_last))
    }

    fn structured_ref_column_id(
        ctx: &WorkbookContext,
        table_id: u32,
        name: &str,
    ) -> Result<u16, EncodeError> {
        let name = name.trim();
        let col_id = ctx
            .table_column_id_by_name(table_id, name)
            .ok_or_else(|| EncodeError::Parse(format!("unknown table column: {name}")))?;
        u16::try_from(col_id).map_err(|_| {
            EncodeError::Parse(format!(
                "table column id {col_id} is out of range for BIFF12"
            ))
        })
    }

    fn emit_reference_expr(
        expr: &fe::Expr,
        ctx: &WorkbookContext,
        sheet: Option<&str>,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        _rgcb: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        match expr {
            fe::Expr::CellRef(r) => emit_cell_ref(r, ctx, base, rgce, class),
            fe::Expr::ColRef(r) => emit_col_ref(r, ctx, base, rgce, class),
            fe::Expr::RowRef(r) => emit_row_ref(r, ctx, base, rgce, class),
            fe::Expr::NameRef(name) => emit_defined_name(name, ctx, rgce, class),
            fe::Expr::StructuredRef(r) => emit_structured_ref(r, ctx, sheet, base, rgce, class),
            fe::Expr::Binary(b) if b.op == fe::BinaryOp::Range => {
                if let Some(area) = area_ref_from_range_operands(&b.left, &b.right, ctx, base)? {
                    return emit_area_ref_info(&area.0, &area.1, &area.2, ctx, rgce, class);
                }

                Err(EncodeError::Parse(
                    "implicit intersection (@) is only supported on simple references".to_string(),
                ))
            }
            _ => Err(EncodeError::Parse(
                "implicit intersection (@) is only supported on references".to_string(),
            )),
        }
    }

    fn emit_range_binary(
        b: &fe::BinaryExpr,
        ctx: &WorkbookContext,
        sheet: Option<&str>,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        rgcb: &mut Vec<u8>,
    ) -> Result<(), EncodeError> {
        // Prefer encoding simple references as `PtgArea`/`PtgArea3d` for Excel-compatible rgce.
        if let Some(area) = area_ref_from_range_operands(&b.left, &b.right, ctx, base)? {
            return emit_area_ref_info(&area.0, &area.1, &area.2, ctx, rgce, PtgClass::Ref);
        }

        // Fallback: encode as the `:` operator.
        emit_expr(&b.left, ctx, sheet, base, rgce, rgcb)?;
        emit_expr(&b.right, ctx, sheet, base, rgce, rgcb)?;
        rgce.push(PTG_RANGE);
        Ok(())
    }

    fn area_ref_from_range_operands(
        left: &fe::Expr,
        right: &fe::Expr,
        _ctx: &WorkbookContext,
        base: CellCoord,
    ) -> Result<Option<(CellRefInfo, CellRefInfo, SheetSpec)>, EncodeError> {
        match (left, right) {
            (fe::Expr::CellRef(a_ref), fe::Expr::CellRef(b_ref)) => {
                let a = cell_ref_info_from_cell_ref(a_ref, base)?;
                let b = cell_ref_info_from_cell_ref(b_ref, base)?;
                let merged = merge_sheets(&a.sheet, &b.sheet).ok_or_else(|| {
                    EncodeError::Parse("range operands refer to different sheets".to_string())
                })?;
                Ok(Some((a, b, merged)))
            }
            (fe::Expr::ColRef(a_ref), fe::Expr::ColRef(b_ref)) => {
                // Column ranges like `A:C` / `A:A`.
                const MAX_ROW: u32 = 1_048_575;
                let (col_a, abs_col_a) = coord_to_a1_index(&a_ref.col, base.col)?;
                let (col_b, abs_col_b) = coord_to_a1_index(&b_ref.col, base.col)?;

                let sheet_a = sheet_spec_from_ref_prefix(&a_ref.workbook, &a_ref.sheet);
                let sheet_b = sheet_spec_from_ref_prefix(&b_ref.workbook, &b_ref.sheet);
                let merged = merge_sheets(&sheet_a, &sheet_b).ok_or_else(|| {
                    EncodeError::Parse("range operands refer to different sheets".to_string())
                })?;

                let a = CellRefInfo {
                    sheet: sheet_a,
                    row: 0,
                    col: col_a,
                    abs_row: true,
                    abs_col: abs_col_a,
                };
                let b = CellRefInfo {
                    sheet: sheet_b,
                    row: MAX_ROW,
                    col: col_b,
                    abs_row: true,
                    abs_col: abs_col_b,
                };
                Ok(Some((a, b, merged)))
            }
            (fe::Expr::RowRef(a_ref), fe::Expr::RowRef(b_ref)) => {
                // Row ranges like `1:3` / `1:1`.
                const MAX_COL: u32 = COL_INDEX_MASK as u32;
                let (row_a, abs_row_a) = coord_to_a1_index(&a_ref.row, base.row)?;
                let (row_b, abs_row_b) = coord_to_a1_index(&b_ref.row, base.row)?;

                let sheet_a = sheet_spec_from_ref_prefix(&a_ref.workbook, &a_ref.sheet);
                let sheet_b = sheet_spec_from_ref_prefix(&b_ref.workbook, &b_ref.sheet);
                let merged = merge_sheets(&sheet_a, &sheet_b).ok_or_else(|| {
                    EncodeError::Parse("range operands refer to different sheets".to_string())
                })?;

                let a = CellRefInfo {
                    sheet: sheet_a,
                    row: row_a,
                    col: 0,
                    abs_row: abs_row_a,
                    abs_col: true,
                };
                let b = CellRefInfo {
                    sheet: sheet_b,
                    row: row_b,
                    col: MAX_COL,
                    abs_row: abs_row_b,
                    abs_col: true,
                };
                Ok(Some((a, b, merged)))
            }
            _ => Ok(None),
        }
    }

    fn cell_ref_info_from_cell_ref(
        r: &fe::CellRef,
        base: CellCoord,
    ) -> Result<CellRefInfo, EncodeError> {
        let (row, abs_row) = coord_to_a1_index(&r.row, base.row)?;
        let (col, abs_col) = coord_to_a1_index(&r.col, base.col)?;
        let sheet = sheet_spec_from_ref_prefix(&r.workbook, &r.sheet);
        Ok(CellRefInfo {
            sheet,
            row,
            col,
            abs_row,
            abs_col,
        })
    }

    fn sheet_spec_from_ref_prefix(
        workbook: &Option<String>,
        sheet: &Option<fe::SheetRef>,
    ) -> SheetSpec {
        match (workbook.as_ref(), sheet.as_ref()) {
            (None, None) => SheetSpec::Current,
            (None, Some(fe::SheetRef::Sheet(sheet))) => SheetSpec::Single(sheet.clone()),
            (None, Some(fe::SheetRef::SheetRange { start, end })) => {
                SheetSpec::Range(start.clone(), end.clone())
            }
            (Some(book), None) => SheetSpec::Single(format_external_workbook_key(book)),
            (Some(book), Some(fe::SheetRef::Sheet(sheet))) => {
                SheetSpec::Single(format_external_key(book, sheet))
            }
            (Some(book), Some(fe::SheetRef::SheetRange { start, end })) => {
                SheetSpec::Range(format_external_key(book, start), format_external_key(book, end))
            }
        }
    }

    fn merge_sheets(a: &SheetSpec, b: &SheetSpec) -> Option<SheetSpec> {
        match (a, b) {
            (SheetSpec::Current, SheetSpec::Current) => Some(SheetSpec::Current),
            (SheetSpec::Current, other) | (other, SheetSpec::Current) => Some(other.clone()),
            (SheetSpec::Single(a), SheetSpec::Single(b))
                if sheet_name_eq_case_insensitive(a, b) =>
            {
                Some(SheetSpec::Single(a.clone()))
            }
            (SheetSpec::Range(af, al), SheetSpec::Range(bf, bl))
                if sheet_name_eq_case_insensitive(af, bf)
                    && sheet_name_eq_case_insensitive(al, bl) =>
            {
                Some(SheetSpec::Range(af.clone(), al.clone()))
            }
            _ => None,
        }
    }

    fn emit_cell_ref(
        r: &fe::CellRef,
        ctx: &WorkbookContext,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        let info = cell_ref_info_from_cell_ref(r, base)?;
        emit_cell_ref_info(&info, ctx, rgce, class)
    }

    fn emit_col_ref(
        r: &fe::ColRef,
        ctx: &WorkbookContext,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        // Column references like `A:A` are represented in BIFF as areas spanning the entire row
        // range for the given column.
        const MAX_ROW: u32 = 1_048_575;
        let (col, abs_col) = coord_to_a1_index(&r.col, base.col)?;
        let sheet = sheet_spec_from_ref_prefix(&r.workbook, &r.sheet);

        let a = CellRefInfo {
            sheet: sheet.clone(),
            row: 0,
            col,
            abs_row: true,
            abs_col,
        };
        let b = CellRefInfo {
            sheet: sheet.clone(),
            row: MAX_ROW,
            col,
            abs_row: true,
            abs_col,
        };
        emit_area_ref_info(&a, &b, &sheet, ctx, rgce, class)
    }

    fn emit_row_ref(
        r: &fe::RowRef,
        ctx: &WorkbookContext,
        base: CellCoord,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        // Row references like `1:1` are represented in BIFF as areas spanning the entire column
        // range for the given row.
        const MAX_COL: u32 = COL_INDEX_MASK as u32;
        let (row, abs_row) = coord_to_a1_index(&r.row, base.row)?;
        let sheet = sheet_spec_from_ref_prefix(&r.workbook, &r.sheet);

        let a = CellRefInfo {
            sheet: sheet.clone(),
            row,
            col: 0,
            abs_row,
            abs_col: true,
        };
        let b = CellRefInfo {
            sheet: sheet.clone(),
            row,
            col: MAX_COL,
            abs_row,
            abs_col: true,
        };
        emit_area_ref_info(&a, &b, &sheet, ctx, rgce, class)
    }

    fn emit_cell_ref_info(
        r: &CellRefInfo,
        ctx: &WorkbookContext,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        match &r.sheet {
            SheetSpec::Current => {
                rgce.push(ptg_with_class(PTG_REF, class));
                rgce.extend_from_slice(&r.row.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(r.col, r.abs_row, r.abs_col).to_le_bytes(),
                );
            }
            SheetSpec::Single(sheet) => {
                let ixti = ctx
                    .extern_sheet_range_index(sheet, sheet)
                    .ok_or_else(|| EncodeError::UnknownSheet(sheet.clone()))?;
                rgce.push(ptg_with_class(PTG_REF3D, class));
                rgce.extend_from_slice(&ixti.to_le_bytes());
                rgce.extend_from_slice(&r.row.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(r.col, r.abs_row, r.abs_col).to_le_bytes(),
                );
            }
            SheetSpec::Range(first, last) => {
                let ixti = ctx
                    .extern_sheet_range_index(first, last)
                    .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;
                rgce.push(ptg_with_class(PTG_REF3D, class));
                rgce.extend_from_slice(&ixti.to_le_bytes());
                rgce.extend_from_slice(&r.row.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(r.col, r.abs_row, r.abs_col).to_le_bytes(),
                );
            }
        }
        Ok(())
    }

    fn emit_area_ref_info(
        a: &CellRefInfo,
        b: &CellRefInfo,
        sheet: &SheetSpec,
        ctx: &WorkbookContext,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        let row_first = a.row.min(b.row);
        let row_last = a.row.max(b.row);
        let col_first = a.col.min(b.col);
        let col_last = a.col.max(b.col);

        // Preserve absolute flags when the input endpoints are not in canonical top-left/bottom-right
        // order, and when one dimension is degenerate but still uses mixed abs markers.
        let rows_equal = a.row == b.row;
        let cols_equal = a.col == b.col;
        let (row_first_from_a, col_first_from_a, row_last_from_a, col_last_from_a) =
            if rows_equal && cols_equal {
                (true, true, false, false)
            } else if rows_equal {
                let col_first_from_a = a.col < b.col;
                let col_last_from_a = a.col > b.col;
                (
                    col_first_from_a,
                    col_first_from_a,
                    col_last_from_a,
                    col_last_from_a,
                )
            } else if cols_equal {
                let row_first_from_a = a.row < b.row;
                let row_last_from_a = a.row > b.row;
                (
                    row_first_from_a,
                    row_first_from_a,
                    row_last_from_a,
                    row_last_from_a,
                )
            } else {
                (a.row < b.row, a.col < b.col, a.row > b.row, a.col > b.col)
            };

        let abs_row_first = if row_first_from_a {
            a.abs_row
        } else {
            b.abs_row
        };
        let abs_col_first = if col_first_from_a {
            a.abs_col
        } else {
            b.abs_col
        };
        let abs_row_last = if row_last_from_a {
            a.abs_row
        } else {
            b.abs_row
        };
        let abs_col_last = if col_last_from_a {
            a.abs_col
        } else {
            b.abs_col
        };

        match sheet {
            SheetSpec::Current => {
                rgce.push(ptg_with_class(PTG_AREA, class));
                rgce.extend_from_slice(&row_first.to_le_bytes());
                rgce.extend_from_slice(&row_last.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(col_first, abs_row_first, abs_col_first).to_le_bytes(),
                );
                rgce.extend_from_slice(
                    &encode_col_field(col_last, abs_row_last, abs_col_last).to_le_bytes(),
                );
            }
            SheetSpec::Single(name) => {
                let ixti = ctx
                    .extern_sheet_range_index(name, name)
                    .ok_or_else(|| EncodeError::UnknownSheet(name.clone()))?;
                rgce.push(ptg_with_class(PTG_AREA3D, class));
                rgce.extend_from_slice(&ixti.to_le_bytes());
                rgce.extend_from_slice(&row_first.to_le_bytes());
                rgce.extend_from_slice(&row_last.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(col_first, abs_row_first, abs_col_first).to_le_bytes(),
                );
                rgce.extend_from_slice(
                    &encode_col_field(col_last, abs_row_last, abs_col_last).to_le_bytes(),
                );
            }
            SheetSpec::Range(first, last) => {
                let ixti = ctx
                    .extern_sheet_range_index(first, last)
                    .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;
                rgce.push(ptg_with_class(PTG_AREA3D, class));
                rgce.extend_from_slice(&ixti.to_le_bytes());
                rgce.extend_from_slice(&row_first.to_le_bytes());
                rgce.extend_from_slice(&row_last.to_le_bytes());
                rgce.extend_from_slice(
                    &encode_col_field(col_first, abs_row_first, abs_col_first).to_le_bytes(),
                );
                rgce.extend_from_slice(
                    &encode_col_field(col_last, abs_row_last, abs_col_last).to_le_bytes(),
                );
            }
        }

        Ok(())
    }

    fn emit_defined_name(
        name: &fe::NameRef,
        ctx: &WorkbookContext,
        rgce: &mut Vec<u8>,
        class: PtgClass,
    ) -> Result<(), EncodeError> {
        // 3D sheet ranges for defined names (`Sheet1:Sheet3!MyName`) require `PtgNameX` with an
        // `ixti` pointing at the sheet span and a `nameIndex` referencing an ExternName entry.
        //
        // This is distinct from workbook-scoped and single-sheet-scoped defined names, which are
        // encoded using `PtgName` (defined name index).
        if let Some(fe::SheetRef::SheetRange { start, end }) = name.sheet.as_ref() {
            let (first_key, last_key) = match name.workbook.as_deref() {
                Some(book) => (format_external_key(book, start), format_external_key(book, end)),
                None => (start.clone(), end.clone()),
            };

            let ixti = ctx
                .extern_sheet_range_index(&first_key, &last_key)
                .ok_or_else(|| EncodeError::UnknownSheet(format!("{start}:{end}")))?;

            let name_index = ctx
                .namex_defined_name_index_for_ixti(ixti, &name.name)
                .ok_or_else(|| EncodeError::UnknownName {
                    name: name.name.clone(),
                })?;

            rgce.push(ptg_with_class(PTG_NAMEX, class));
            rgce.extend_from_slice(&ixti.to_le_bytes());
            rgce.extend_from_slice(&name_index.to_le_bytes());
            return Ok(());
        }

        let sheet = match name.sheet.as_ref() {
            None => None,
            Some(fe::SheetRef::Sheet(sheet)) => Some(sheet.as_str()),
            // SheetRange handled above.
            Some(fe::SheetRef::SheetRange { .. }) => None,
        };

        // Prefer workbook-defined names over NameX extern names when the formula text has no
        // workbook prefix.
        if name.workbook.is_none() {
            if let Some(idx) = ctx.name_index(&name.name, sheet) {
                rgce.push(ptg_with_class(PTG_NAME, class));
                rgce.extend_from_slice(&idx.to_le_bytes());
                rgce.extend_from_slice(&0u16.to_le_bytes());
                return Ok(());
            }
        }

        // Fallback: encode as a `PtgNameX` external name if present in the workbook's SupBook /
        // ExternName tables (e.g. add-in names).
        if let Some((ixti, name_index)) = ctx.namex_ref(name.workbook.as_deref(), sheet, &name.name)
        {
            rgce.push(ptg_with_class(PTG_NAMEX, class));
            rgce.extend_from_slice(&ixti.to_le_bytes());
            rgce.extend_from_slice(&name_index.to_le_bytes());
            return Ok(());
        }

        Err(EncodeError::UnknownName {
            name: name.name.clone(),
        })
    }

    fn emit_function_call(
        name: &fe::FunctionName,
        argc: usize,
        ctx: &WorkbookContext,
        rgce: &mut Vec<u8>,
    ) -> Result<(), EncodeError> {
        let upper = name.name_upper.as_str();

        if let Some(iftab) =
            formula_biff::function_name_to_id_uppercase(upper).filter(|id| *id != 0x00FF)
        {
            let spec = formula_biff::function_spec_from_id(iftab)
                .ok_or_else(|| EncodeError::Parse(format!("unknown function id: {iftab}")))?;

            if argc < spec.min_args as usize || argc > spec.max_args as usize {
                return Err(EncodeError::Parse(format!(
                    "invalid argument count for {upper}: got {argc}, expected {}..={}",
                    spec.min_args, spec.max_args
                )));
            }

            if spec.min_args == spec.max_args {
                // Fixed-arity -> PtgFunc.
                rgce.push(PTG_FUNC);
                rgce.extend_from_slice(&iftab.to_le_bytes());
                return Ok(());
            }

            // Variable arity -> PtgFuncVar.
            let argc_u8: u8 = argc
                .try_into()
                .map_err(|_| EncodeError::Parse("too many function arguments".to_string()))?;
            rgce.push(PTG_FUNCVAR);
            rgce.push(argc_u8);
            rgce.extend_from_slice(&iftab.to_le_bytes());
            return Ok(());
        }

        // Add-in / UDF call pattern: args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
        if let Some((ixti, name_index)) = ctx.namex_function_ref(upper) {
            let argc_total = argc
                .checked_add(1)
                .ok_or_else(|| EncodeError::Parse("too many function arguments".to_string()))?;
            let argc_total_u8: u8 = argc_total
                .try_into()
                .map_err(|_| EncodeError::Parse("too many function arguments".to_string()))?;

            rgce.push(PTG_NAMEX);
            rgce.extend_from_slice(&ixti.to_le_bytes());
            rgce.extend_from_slice(&name_index.to_le_bytes());

            rgce.push(PTG_FUNCVAR);
            rgce.push(argc_total_u8);
            rgce.extend_from_slice(&0x00FFu16.to_le_bytes());
            return Ok(());
        }

        // Workbook-defined names can also be invoked like functions (e.g. a LAMBDA stored in a
        // name). BIFF encodes this the same way as UDF calls, but uses `PtgName` (defined name
        // index) instead of `PtgNameX` (extern name).
        if let Some(name_index) = ctx.name_index(upper, None) {
            let argc_total = argc
                .checked_add(1)
                .ok_or_else(|| EncodeError::Parse("too many function arguments".to_string()))?;
            let argc_total_u8: u8 = argc_total
                .try_into()
                .map_err(|_| EncodeError::Parse("too many function arguments".to_string()))?;

            rgce.push(PTG_NAME);
            rgce.extend_from_slice(&name_index.to_le_bytes());
            rgce.extend_from_slice(&0u16.to_le_bytes());

            rgce.push(PTG_FUNCVAR);
            rgce.push(argc_total_u8);
            rgce.extend_from_slice(&0x00FFu16.to_le_bytes());
            return Ok(());
        }

        Err(EncodeError::UnknownFunction {
            name: upper.to_string(),
        })
    }

    fn emit_call_udf_sentinel(arg_count: usize, rgce: &mut Vec<u8>) -> Result<(), EncodeError> {
        let argc_total = arg_count
            .checked_add(1)
            .ok_or_else(|| EncodeError::Parse("too many function arguments".to_string()))?;
        let argc_total_u8: u8 = argc_total
            .try_into()
            .map_err(|_| EncodeError::Parse("too many function arguments".to_string()))?;

        rgce.push(PTG_FUNCVAR);
        rgce.push(argc_total_u8);
        rgce.extend_from_slice(&0x00FFu16.to_le_bytes());
        Ok(())
    }

    fn cell_ref_info_from_function_name(
        name: &fe::FunctionName,
    ) -> Result<Option<CellRefInfo>, EncodeError> {
        let s = name.name_upper.as_str();
        let bytes = s.as_bytes();
        let mut i = 0usize;

        let abs_col = if bytes.get(i) == Some(&b'$') {
            i += 1;
            true
        } else {
            false
        };

        let col_start = i;
        while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
            i += 1;
        }
        if i == col_start {
            return Ok(None);
        }
        let col_letters = &s[col_start..i];

        let abs_row = if bytes.get(i) == Some(&b'$') {
            i += 1;
            true
        } else {
            false
        };

        let row_start = i;
        while i < bytes.len() && bytes[i].is_ascii_digit() {
            i += 1;
        }
        if i == row_start || i != bytes.len() {
            // Not a plain A1 reference.
            return Ok(None);
        }

        let col: u32 = match column_label_to_index(col_letters) {
            Ok(v) => v,
            Err(_) => return Ok(None),
        };
        let row1: u32 = match s[row_start..i].parse() {
            Ok(v) if v != 0 => v,
            _ => return Ok(None),
        };
        let row = row1 - 1;

        // Keep bounds aligned with BIFF12's allowed grid size.
        const MAX_ROW: u32 = 1_048_575;
        if row > MAX_ROW || col > COL_INDEX_MASK as u32 {
            return Ok(None);
        }

        Ok(Some(CellRefInfo {
            sheet: SheetSpec::Current,
            row,
            col,
            abs_row,
            abs_col,
        }))
    }

    fn coord_to_a1_index(coord: &fe::Coord, base_axis: u32) -> Result<(u32, bool), EncodeError> {
        match coord {
            fe::Coord::A1 { index, abs } => Ok((*index, *abs)),
            fe::Coord::Offset(delta) => {
                let idx = base_axis as i32 + *delta;
                if idx < 0 {
                    return Err(EncodeError::Parse(
                        "relative reference moved before A1".to_string(),
                    ));
                }
                Ok((idx as u32, false))
            }
        }
    }

    fn encode_col_field(col: u32, abs_row: bool, abs_col: bool) -> u16 {
        let mut v = (col as u16) & COL_INDEX_MASK;
        if !abs_row {
            v |= ROW_RELATIVE_MASK;
        }
        if !abs_col {
            v |= COL_RELATIVE_MASK;
        }
        v
    }

    fn array_literal_to_const(arr: &fe::ArrayLiteral) -> Result<ArrayConst, EncodeError> {
        if arr.rows.is_empty() {
            return Err(EncodeError::Parse(
                "array literal cannot be empty".to_string(),
            ));
        }
        let cols = arr.rows[0].len();
        if cols == 0 {
            return Err(EncodeError::Parse(
                "array literal cannot be empty".to_string(),
            ));
        }
        if arr.rows.iter().any(|r| r.len() != cols) {
            return Err(EncodeError::Parse(
                "array literal rows must have the same number of columns".to_string(),
            ));
        }

        let mut rows = Vec::new();
        let _ = rows.try_reserve_exact(arr.rows.len());
        for row in &arr.rows {
            let mut out_row = Vec::new();
            let _ = out_row.try_reserve_exact(row.len());
            for el in row {
                out_row.push(array_elem_from_expr(el)?);
            }
            rows.push(out_row);
        }

        Ok(ArrayConst { rows })
    }

    fn array_elem_from_expr(expr: &fe::Expr) -> Result<ArrayElem, EncodeError> {
        match expr {
            fe::Expr::Missing => Ok(ArrayElem::Empty),
            fe::Expr::Number(raw) => {
                let n: f64 = raw
                    .parse()
                    .map_err(|_| EncodeError::Parse(format!("invalid number literal: {raw}")))?;
                Ok(ArrayElem::Number(n))
            }
            fe::Expr::String(s) => Ok(ArrayElem::Str(s.clone())),
            fe::Expr::Boolean(b) => Ok(ArrayElem::Bool(*b)),
            fe::Expr::Error(raw) => {
                let code = xlsb_error_code_from_literal(raw).ok_or_else(|| {
                    EncodeError::Parse(format!("unsupported error literal: {raw}"))
                })?;
                Ok(ArrayElem::Error(code))
            }
            fe::Expr::Unary(u) if u.op == fe::UnaryOp::Plus => match u.expr.as_ref() {
                fe::Expr::Number(raw) => {
                    let n: f64 = raw.parse().map_err(|_| {
                        EncodeError::Parse(format!("invalid number literal: {raw}"))
                    })?;
                    Ok(ArrayElem::Number(n))
                }
                _ => Err(EncodeError::Parse(
                    "unsupported unary '+' in array literal".to_string(),
                )),
            },
            fe::Expr::Unary(u) if u.op == fe::UnaryOp::Minus => match u.expr.as_ref() {
                fe::Expr::Number(raw) => {
                    let n: f64 = raw.parse().map_err(|_| {
                        EncodeError::Parse(format!("invalid number literal: {raw}"))
                    })?;
                    Ok(ArrayElem::Number(-n))
                }
                _ => Err(EncodeError::Parse(
                    "unsupported unary '-' in array literal".to_string(),
                )),
            },
            _ => Err(EncodeError::Parse(
                "unsupported expression in array literal".to_string(),
            )),
        }
    }

    // Keep module-local helpers below.
}

/// Encode an A1-style formula to an XLSB `rgce` token stream using workbook context for 3D
/// references and defined names.
///
/// The input may include a leading `=`; it will be ignored.
pub fn encode_rgce_with_context(
    formula: &str,
    ctx: &WorkbookContext,
    _base: CellCoord,
) -> Result<EncodedRgce, EncodeError> {
    let body = formula.strip_prefix('=').unwrap_or(formula);
    let mut parser = FormulaParser::new(body);
    let expr = parser.parse().map_err(EncodeError::Parse)?;

    let mut rgce = Vec::new();
    let mut rgcb = Vec::new();
    emit_expr(&expr, ctx, &mut rgce, &mut rgcb)?;
    Ok(EncodedRgce { rgce, rgcb })
}

const PTG_ADD: u8 = 0x03;
const PTG_SUB: u8 = 0x04;
const PTG_MUL: u8 = 0x05;
const PTG_DIV: u8 = 0x06;
const PTG_POW: u8 = 0x07;
const PTG_CONCAT: u8 = 0x08;
const PTG_LT: u8 = 0x09;
const PTG_LE: u8 = 0x0A;
const PTG_EQ: u8 = 0x0B;
const PTG_GT: u8 = 0x0C;
const PTG_GE: u8 = 0x0D;
const PTG_NE: u8 = 0x0E;
const PTG_INTERSECT: u8 = 0x0F;
const PTG_UNION: u8 = 0x10;
const PTG_RANGE: u8 = 0x11;
const PTG_UPLUS: u8 = 0x12;
const PTG_UMINUS: u8 = 0x13;
const PTG_PERCENT: u8 = 0x14;
const PTG_MISSARG: u8 = 0x16;
const PTG_STR: u8 = 0x17;
const PTG_ERR: u8 = 0x1C;
const PTG_BOOL: u8 = 0x1D;
const PTG_INT: u8 = 0x1E;
const PTG_NUM: u8 = 0x1F;
const PTG_ARRAY: u8 = 0x20;

const PTG_FUNCVAR: u8 = 0x22;
const PTG_NAME: u8 = 0x23;
const PTG_REF: u8 = 0x24;
const PTG_AREA: u8 = 0x25;
const PTG_REF3D: u8 = 0x3A;
const PTG_AREA3D: u8 = 0x3B;
const PTG_NAMEX: u8 = 0x39;
const PTG_SPILL: u8 = 0x2F;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PtgClass {
    Ref,
    Value,
    #[allow(dead_code)]
    Array,
}

fn ptg_with_class(base: u8, class: PtgClass) -> u8 {
    match class {
        PtgClass::Ref => base,
        PtgClass::Value => base.wrapping_add(0x20),
        PtgClass::Array => base.wrapping_add(0x40),
    }
}

const COL_RELATIVE_MASK: u16 = 0x8000;
const ROW_RELATIVE_MASK: u16 = 0x4000;
const COL_INDEX_MASK: u16 = 0x3FFF;

#[derive(Clone, Debug, PartialEq)]
enum Expr {
    Missing,
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    Ref(Ref),
    Name(NameRef),
    Array(ArrayConst),
    Func {
        name: String,
        args: Vec<Expr>,
    },
    SpillRange(Box<Expr>),
    Percent(Box<Expr>),
    Unary {
        op: UnaryOp,
        expr: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
}

#[derive(Clone, Debug, PartialEq)]
struct ArrayConst {
    rows: Vec<Vec<ArrayElem>>,
}

#[derive(Clone, Debug, PartialEq)]
enum ArrayElem {
    Empty,
    Number(f64),
    Bool(bool),
    Str(String),
    Error(u8),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum UnaryOp {
    Plus,
    Minus,
    ImplicitIntersection,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Lt,
    Le,
    Eq,
    Gt,
    Ge,
    Ne,
    Intersect,
    Union,
    Range,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum SheetSpec {
    Single(String),
    Range(String, String),
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct CellRef {
    row: u32, // 0-indexed
    col: u32, // 0-indexed
    abs_row: bool,
    abs_col: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum RefKind {
    Cell(CellRef),
    Area(CellRef, CellRef),
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Ref {
    sheet: Option<SheetSpec>,
    kind: RefKind,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct NameRef {
    sheet: Option<String>,
    name: String,
}

fn emit_expr(
    expr: &Expr,
    ctx: &WorkbookContext,
    rgce: &mut Vec<u8>,
    rgcb: &mut Vec<u8>,
) -> Result<(), EncodeError> {
    match expr {
        Expr::Missing => rgce.push(PTG_MISSARG),
        Expr::Number(n) => emit_number(*n, rgce),
        Expr::String(s) => {
            let start_len = rgce.len();
            rgce.push(PTG_STR);
            push_utf16le_u16_len_with_rollback(rgce, start_len, s, "string literal too long")?;
        }
        Expr::Bool(b) => {
            rgce.push(PTG_BOOL);
            rgce.push(u8::from(*b));
        }
        Expr::Error(code) => {
            rgce.push(PTG_ERR);
            rgce.push(*code);
        }
        Expr::Ref(r) => emit_ref(r, ctx, rgce, PtgClass::Ref)?,
        Expr::Name(n) => emit_name(n, ctx, rgce, PtgClass::Ref)?,
        Expr::Array(a) => emit_array(a, rgce, rgcb)?,
        Expr::Func { name, args } => {
            for arg in args {
                emit_expr(arg, ctx, rgce, rgcb)?;
            }
            emit_func(name, args.len(), ctx, rgce)?;
        }
        Expr::SpillRange(inner) => {
            emit_expr(inner, ctx, rgce, rgcb)?;
            rgce.push(PTG_SPILL);
        }
        Expr::Percent(inner) => {
            emit_expr(inner, ctx, rgce, rgcb)?;
            rgce.push(PTG_PERCENT);
        }
        Expr::Unary { op, expr } => {
            match op {
                UnaryOp::ImplicitIntersection => match expr.as_ref() {
                    // Encode `@` by emitting value-class reference tokens. This matches Excel's
                    // legacy implicit-intersection encoding, and round-trips through
                    // `decode_rgce*` as an explicit `@`.
                    Expr::Ref(r) => emit_ref(r, ctx, rgce, PtgClass::Value)?,
                    Expr::Name(n) => emit_name(n, ctx, rgce, PtgClass::Value)?,
                    _ => {
                        return Err(EncodeError::Parse(
                            "implicit intersection (@) is only supported on references".to_string(),
                        ))
                    }
                },
                UnaryOp::Plus => {
                    emit_expr(expr, ctx, rgce, rgcb)?;
                    rgce.push(PTG_UPLUS);
                }
                UnaryOp::Minus => {
                    emit_expr(expr, ctx, rgce, rgcb)?;
                    rgce.push(PTG_UMINUS);
                }
            }
        }
        Expr::Binary { op, left, right } => {
            emit_expr(left, ctx, rgce, rgcb)?;
            emit_expr(right, ctx, rgce, rgcb)?;
            rgce.push(match op {
                BinaryOp::Add => PTG_ADD,
                BinaryOp::Sub => PTG_SUB,
                BinaryOp::Mul => PTG_MUL,
                BinaryOp::Div => PTG_DIV,
                BinaryOp::Pow => PTG_POW,
                BinaryOp::Concat => PTG_CONCAT,
                BinaryOp::Lt => PTG_LT,
                BinaryOp::Le => PTG_LE,
                BinaryOp::Eq => PTG_EQ,
                BinaryOp::Gt => PTG_GT,
                BinaryOp::Ge => PTG_GE,
                BinaryOp::Ne => PTG_NE,
                BinaryOp::Intersect => PTG_INTERSECT,
                BinaryOp::Union => PTG_UNION,
                BinaryOp::Range => PTG_RANGE,
            });
        }
    }
    Ok(())
}

fn emit_number(n: f64, out: &mut Vec<u8>) {
    if n.fract() == 0.0 && (0.0..=65535.0).contains(&n) {
        out.push(PTG_INT);
        out.extend_from_slice(&(n as u16).to_le_bytes());
    } else {
        out.push(PTG_NUM);
        out.extend_from_slice(&n.to_le_bytes());
    }
}

fn emit_array(
    array: &ArrayConst,
    rgce: &mut Vec<u8>,
    rgcb: &mut Vec<u8>,
) -> Result<(), EncodeError> {
    rgce.push(ptg_with_class(PTG_ARRAY, PtgClass::Array));
    rgce.extend_from_slice(&[0u8; 7]); // reserved
    encode_array_constant(array, rgcb)
}

fn push_utf16le_u16_len_with_rollback(
    out: &mut Vec<u8>,
    start_len: usize,
    s: &str,
    err_msg: &'static str,
) -> Result<(), EncodeError> {
    let len_pos = out.len();
    out.extend_from_slice(&0u16.to_le_bytes()); // backpatched

    let mut cch: u16 = 0;
    for unit in s.encode_utf16() {
        cch = cch.checked_add(1).ok_or_else(|| {
            out.truncate(start_len);
            EncodeError::Parse(err_msg.to_string())
        })?;
        out.extend_from_slice(&unit.to_le_bytes());
    }

    let len_end = len_pos.checked_add(2).ok_or_else(|| {
        out.truncate(start_len);
        EncodeError::Parse(err_msg.to_string())
    })?;
    let Some(dst) = out.get_mut(len_pos..len_end) else {
        out.truncate(start_len);
        return Err(EncodeError::Parse(err_msg.to_string()));
    };
    dst.copy_from_slice(&cch.to_le_bytes());
    Ok(())
}

fn encode_array_constant(array: &ArrayConst, rgcb: &mut Vec<u8>) -> Result<(), EncodeError> {
    let rows = array.rows.len();
    let cols = array.rows.first().map(|r| r.len()).unwrap_or(0);
    if rows == 0 || cols == 0 {
        return Err(EncodeError::Parse(
            "array constant cannot be empty".to_string(),
        ));
    }
    if array.rows.iter().any(|r| r.len() != cols) {
        return Err(EncodeError::Parse(
            "array constant rows must have the same number of columns".to_string(),
        ));
    }

    let cols_minus1 = u16::try_from(cols - 1)
        .map_err(|_| EncodeError::Parse("array constant is too wide".to_string()))?;
    let rows_minus1 = u16::try_from(rows - 1)
        .map_err(|_| EncodeError::Parse("array constant is too tall".to_string()))?;

    rgcb.extend_from_slice(&cols_minus1.to_le_bytes());
    rgcb.extend_from_slice(&rows_minus1.to_le_bytes());

    for row in &array.rows {
        for elem in row {
            match elem {
                ArrayElem::Empty => rgcb.push(0x00),
                ArrayElem::Number(n) => {
                    rgcb.push(0x01);
                    rgcb.extend_from_slice(&n.to_le_bytes());
                }
                ArrayElem::Str(s) => {
                    let start_len = rgcb.len();
                    rgcb.push(0x02);
                    push_utf16le_u16_len_with_rollback(
                        rgcb,
                        start_len,
                        s,
                        "array string literal is too long",
                    )?;
                }
                ArrayElem::Bool(b) => {
                    rgcb.push(0x04);
                    rgcb.push(if *b { 1 } else { 0 });
                }
                ArrayElem::Error(code) => {
                    rgcb.push(0x10);
                    rgcb.push(*code);
                }
            }
        }
    }
    Ok(())
}

fn emit_func(
    name: &str,
    argc: usize,
    ctx: &WorkbookContext,
    out: &mut Vec<u8>,
) -> Result<(), EncodeError> {
    let name = name.trim();
    let mut buf = [0u8; 64];
    let upper_owned: String;
    let upper: &str = if name.len() <= buf.len() {
        for (dst, src) in buf[..name.len()].iter_mut().zip(name.as_bytes()) {
            *dst = src.to_ascii_uppercase();
        }
        match std::str::from_utf8(&buf[..name.len()]) {
            Ok(s) => s,
            Err(_) => {
                debug_assert!(false, "ASCII uppercasing preserves UTF-8");
                name
            }
        }
    } else {
        upper_owned = name.to_ascii_uppercase();
        &upper_owned
    };

    // Built-in functions.
    //
    // Note: Excel encodes "future" (forward-compatible) functions as user-defined calls (iftab=255)
    // paired with a name token. We do not currently support that encoding path, so we explicitly
    // avoid treating `iftab=0x00FF` as a built-in here.
    if let Some(iftab) =
        formula_biff::function_name_to_id_uppercase(upper).filter(|id| *id != 0x00FF)
    {
        if argc > u8::MAX as usize {
            return Err(EncodeError::Parse(
                "too many function arguments".to_string(),
            ));
        }
        out.push(PTG_FUNCVAR);
        out.push(argc as u8);
        out.extend_from_slice(&iftab.to_le_bytes());
        return Ok(());
    }

    // Add-in / UDF call pattern: args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
    if let Some((ixti, name_index)) = ctx.namex_function_ref(upper) {
        let argc_total = argc
            .checked_add(1)
            .ok_or_else(|| EncodeError::Parse("too many function arguments".to_string()))?;
        if argc_total > u8::MAX as usize {
            return Err(EncodeError::Parse(
                "too many function arguments".to_string(),
            ));
        }

        out.push(PTG_NAMEX);
        out.extend_from_slice(&ixti.to_le_bytes());
        out.extend_from_slice(&name_index.to_le_bytes());

        out.push(PTG_FUNCVAR);
        out.push(argc_total as u8);
        out.extend_from_slice(&0x00FFu16.to_le_bytes());
        return Ok(());
    }

    Err(EncodeError::UnknownFunction {
        name: upper.to_string(),
    })
}

fn emit_name(
    name: &NameRef,
    ctx: &WorkbookContext,
    out: &mut Vec<u8>,
    class: PtgClass,
) -> Result<(), EncodeError> {
    // Sheet-span defined-name references are encoded via `PtgNameX` with an `ixti` that points at
    // the sheet range (mirrors 3D ref encoding). The decoded/canonical text form uses a single
    // quoted identifier prefix (e.g. `'Sheet1:Sheet3'!MyName`), which the small-scope formula
    // parser represents as a "sheet" string containing `:`.
    //
    // Excel sheet names cannot contain `:`, so this representation is unambiguous.
    if let Some(sheet) = name.sheet.as_deref() {
        if let Some((first, last)) = sheet.split_once(':') {
            let ixti = extern_sheet_range_index_with_fallback(ctx, first, last)
                .ok_or_else(|| EncodeError::UnknownSheet(sheet.to_string()))?;
            let name_index = ctx
                .namex_defined_name_index_for_ixti(ixti, &name.name)
                .ok_or_else(|| EncodeError::UnknownName {
                    name: name.name.clone(),
                })?;

            out.push(ptg_with_class(PTG_NAMEX, class));
            out.extend_from_slice(&ixti.to_le_bytes());
            out.extend_from_slice(&name_index.to_le_bytes());
            return Ok(());
        }
    }

    let idx = ctx
        .name_index(&name.name, name.sheet.as_deref())
        .ok_or_else(|| EncodeError::UnknownName {
            name: name.name.clone(),
        })?;

    out.push(ptg_with_class(PTG_NAME, class));
    out.extend_from_slice(&idx.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    Ok(())
}

fn emit_ref(
    r: &Ref,
    ctx: &WorkbookContext,
    out: &mut Vec<u8>,
    class: PtgClass,
) -> Result<(), EncodeError> {
    match (&r.sheet, &r.kind) {
        (None, RefKind::Cell(cell)) => {
            out.push(ptg_with_class(PTG_REF, class));
            emit_cell_ref_fields(cell, out);
        }
        (None, RefKind::Area(a, b)) => {
            out.push(ptg_with_class(PTG_AREA, class));
            emit_area_fields(a, b, out);
        }
        (Some(sheet), RefKind::Cell(cell)) => {
            let (first, last) = match sheet {
                SheetSpec::Single(s) => match s.split_once(':') {
                    // `formula-engine` only parses a single identifier before `!`, so 3D sheet
                    // ranges round-trip through text as a single "sheet" string containing `:`
                    // (e.g. `'Sheet1:Sheet3'!A1`). Since `:` is invalid in Excel sheet names, this
                    // split is unambiguous.
                    Some((first, last)) => (first, last),
                    None => (s.as_str(), s.as_str()),
                },
                SheetSpec::Range(a, b) => (a.as_str(), b.as_str()),
            };
            let ixti = extern_sheet_range_index_with_fallback(ctx, first, last)
                .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;

            out.push(ptg_with_class(PTG_REF3D, class));
            out.extend_from_slice(&ixti.to_le_bytes());
            emit_cell_ref_fields(cell, out);
        }
        (Some(sheet), RefKind::Area(a, b)) => {
            let (first, last) = match sheet {
                SheetSpec::Single(s) => match s.split_once(':') {
                    // See the `RefKind::Cell` branch above.
                    Some((first, last)) => (first, last),
                    None => (s.as_str(), s.as_str()),
                },
                SheetSpec::Range(a, b) => (a.as_str(), b.as_str()),
            };
            let ixti = extern_sheet_range_index_with_fallback(ctx, first, last)
                .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;

            out.push(ptg_with_class(PTG_AREA3D, class));
            out.extend_from_slice(&ixti.to_le_bytes());
            emit_area_fields(a, b, out);
        }
    }
    Ok(())
}

fn extern_sheet_range_index_with_fallback(
    ctx: &WorkbookContext,
    first: &str,
    last: &str,
) -> Option<u16> {
    // Fast path.
    if let Some(ixti) = ctx.extern_sheet_range_index(first, last) {
        return Some(ixti);
    }

    // For external workbook refs, our decoded text format uses the Excel form
    // `'[Book]Sheet1:Sheet3'!A1` (workbook prefix appears once). The AST encoder
    // (`sheet_spec_from_ref_prefix`) expands this into `"[Book]Sheet1"` and `"[Book]Sheet3"`.
    // Try both representations so callers can register ExternSheet entries either way.
    let first_prefix = formula_model::external_refs::split_external_workbook_prefix(first);
    let last_prefix = formula_model::external_refs::split_external_workbook_prefix(last);

    if let (Some((prefix, _)), None) = (first_prefix, last_prefix) {
        let mut last_with_prefix = prefix.to_string();
        last_with_prefix.push_str(last);
        if let Some(ixti) = ctx.extern_sheet_range_index(first, &last_with_prefix) {
            return Some(ixti);
        }
    }

    if let (None, Some((prefix, last_sheet))) = (first_prefix, last_prefix) {
        let mut first_with_prefix = prefix.to_string();
        first_with_prefix.push_str(first);
        if let Some(ixti) = ctx.extern_sheet_range_index(&first_with_prefix, last) {
            return Some(ixti);
        }
        if let Some(ixti) = ctx.extern_sheet_range_index(&first_with_prefix, last_sheet) {
            return Some(ixti);
        }
    }

    if let (Some((prefix1, _)), Some((prefix2, last_sheet))) = (first_prefix, last_prefix) {
        if sheet_name_eq_case_insensitive(prefix1, prefix2) {
            if let Some(ixti) = ctx.extern_sheet_range_index(first, last_sheet) {
                return Some(ixti);
            }
        }
    }

    None
}

fn emit_cell_ref_fields(cell: &CellRef, out: &mut Vec<u8>) {
    out.extend_from_slice(&cell.row.to_le_bytes());
    out.extend_from_slice(&encode_col_field(cell.col, cell.abs_row, cell.abs_col).to_le_bytes());
}

fn emit_area_fields(a: &CellRef, b: &CellRef, out: &mut Vec<u8>) {
    let row_first = a.row.min(b.row);
    let row_last = a.row.max(b.row);
    let col_first = a.col.min(b.col);
    let col_last = a.col.max(b.col);

    // `rgce` stores row/col absolute flags on the corner payloads (in the `colFirst`/`colLast`
    // fields). When a range is written in non-canonical order (e.g. `B$1:$A2`) the "first" corner
    // in the token stream is still the top-left corner (min row/col) and the "last" corner is the
    // bottom-right (max row/col).
    //
    // Preserve absolute flags by selecting them from the appropriate input reference for each
    // corner coordinate:
    // - For cross-corner ranges (e.g. `B$1:$A2`), the top-left corner combines the smaller row from
    //   one endpoint with the smaller column from the other.
    // - For degenerate ranges where one dimension is equal (e.g. `A1:B$1`, `A1:$A$2`), Excel can
    //   still preserve mixed absolute markers, so use the other dimension as a stable tie-breaker.
    let rows_equal = a.row == b.row;
    let cols_equal = a.col == b.col;
    let (row_first_from_a, col_first_from_a, row_last_from_a, col_last_from_a) =
        if rows_equal && cols_equal {
            // Single-cell area (A1:A1): preserve both endpoints' flags if they differ by assigning
            // the first corner to `a` and the last corner to `b`.
            (true, true, false, false)
        } else if rows_equal {
            // Horizontal range: use column ordering to decide which endpoint supplies the row flags.
            let col_first_from_a = a.col < b.col;
            let col_last_from_a = a.col > b.col;
            (
                col_first_from_a,
                col_first_from_a,
                col_last_from_a,
                col_last_from_a,
            )
        } else if cols_equal {
            // Vertical range: use row ordering to decide which endpoint supplies the column flags.
            let row_first_from_a = a.row < b.row;
            let row_last_from_a = a.row > b.row;
            (
                row_first_from_a,
                row_first_from_a,
                row_last_from_a,
                row_last_from_a,
            )
        } else {
            // General rectangle: rows/cols both differ, so each dimension can be resolved
            // independently.
            (a.row < b.row, a.col < b.col, a.row > b.row, a.col > b.col)
        };

    let abs_row_first = if row_first_from_a {
        a.abs_row
    } else {
        b.abs_row
    };
    let abs_col_first = if col_first_from_a {
        a.abs_col
    } else {
        b.abs_col
    };
    let abs_row_last = if row_last_from_a {
        a.abs_row
    } else {
        b.abs_row
    };
    let abs_col_last = if col_last_from_a {
        a.abs_col
    } else {
        b.abs_col
    };

    out.extend_from_slice(&row_first.to_le_bytes());
    out.extend_from_slice(&row_last.to_le_bytes());
    out.extend_from_slice(&encode_col_field(col_first, abs_row_first, abs_col_first).to_le_bytes());
    out.extend_from_slice(&encode_col_field(col_last, abs_row_last, abs_col_last).to_le_bytes());
}

fn encode_col_field(col: u32, abs_row: bool, abs_col: bool) -> u16 {
    let mut v = (col as u16) & COL_INDEX_MASK;
    if !abs_row {
        v |= ROW_RELATIVE_MASK;
    }
    if !abs_col {
        v |= COL_RELATIVE_MASK;
    }
    v
}

struct FormulaParser<'a> {
    input: &'a str,
    pos: usize,
}

impl<'a> FormulaParser<'a> {
    fn new(input: &'a str) -> Self {
        Self { input, pos: 0 }
    }

    fn parse(&mut self) -> Result<Expr, String> {
        self.skip_ws();
        let expr = self.parse_expr(false)?;
        self.skip_ws();
        if self.pos < self.input.len() {
            return Err(format!("unexpected trailing input at byte {}", self.pos));
        }
        Ok(expr)
    }

    fn parse_expr(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        self.parse_comparison(stop_at_comma)
    }

    fn parse_comparison(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_concat(stop_at_comma)?;
        loop {
            self.skip_ws();
            let rest = &self.input[self.pos..];
            let op = if rest.starts_with("<=") {
                self.pos += 2;
                Some(BinaryOp::Le)
            } else if rest.starts_with(">=") {
                self.pos += 2;
                Some(BinaryOp::Ge)
            } else if rest.starts_with("<>") {
                self.pos += 2;
                Some(BinaryOp::Ne)
            } else if rest.starts_with('<') {
                self.pos += 1;
                Some(BinaryOp::Lt)
            } else if rest.starts_with('>') {
                self.pos += 1;
                Some(BinaryOp::Gt)
            } else if rest.starts_with('=') {
                self.pos += 1;
                Some(BinaryOp::Eq)
            } else {
                None
            };

            let Some(op) = op else { break };
            let rhs = self.parse_concat(stop_at_comma)?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_concat(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_add_sub(stop_at_comma)?;
        loop {
            self.skip_ws();
            if self.peek_char() != Some('&') {
                break;
            }
            self.next_char();
            let rhs = self.parse_add_sub(stop_at_comma)?;
            expr = Expr::Binary {
                op: BinaryOp::Concat,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_add_sub(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_mul_div(stop_at_comma)?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('+') => BinaryOp::Add,
                Some('-') => BinaryOp::Sub,
                _ => break,
            };
            self.next_char();
            let rhs = self.parse_mul_div(stop_at_comma)?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_mul_div(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_unary(stop_at_comma)?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('*') => BinaryOp::Mul,
                Some('/') => BinaryOp::Div,
                _ => break,
            };
            self.next_char();
            let rhs = self.parse_power(stop_at_comma)?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_power(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        // In Excel, exponentiation binds tighter than unary +/- (e.g. `-2^2` == `-(2^2)`).
        let expr = self.parse_ref_union(stop_at_comma)?;
        self.skip_ws();
        if self.peek_char() != Some('^') {
            return Ok(expr);
        }
        // Excel exponentiation is right-associative.
        self.next_char();
        let rhs = self.parse_unary(stop_at_comma)?;
        Ok(Expr::Binary {
            op: BinaryOp::Pow,
            left: Box::new(expr),
            right: Box::new(rhs),
        })
    }

    fn parse_unary(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        self.skip_ws();
        match self.peek_char() {
            Some('+') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_unary(stop_at_comma)?),
                })
            }
            Some('-') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::Minus,
                    expr: Box::new(self.parse_unary(stop_at_comma)?),
                })
            }
            Some('@') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(self.parse_unary(stop_at_comma)?),
                })
            }
            _ => self.parse_power(stop_at_comma),
        }
    }

    fn parse_ref_union(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_ref_intersect(stop_at_comma)?;
        if stop_at_comma {
            return Ok(expr);
        }
        loop {
            self.skip_ws();
            if self.peek_char() != Some(',') {
                break;
            }
            self.next_char();
            let rhs = self.parse_ref_intersect(stop_at_comma)?;
            expr = Expr::Binary {
                op: BinaryOp::Union,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_ref_intersect(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_ref_range(stop_at_comma)?;
        loop {
            let had_ws = self.skip_ws();
            if !had_ws {
                break;
            }
            if !self.is_intersection_rhs_start() {
                break;
            }
            let rhs = self.parse_ref_range(stop_at_comma)?;
            expr = Expr::Binary {
                op: BinaryOp::Intersect,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_ref_range(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_postfix(stop_at_comma)?;
        loop {
            let after_ws = self.peek_non_ws_pos();
            if after_ws >= self.input.len() {
                break;
            }
            if self.input[after_ws..].starts_with(':') {
                // Commit any whitespace.
                self.pos = after_ws;
                self.next_char();
                // Allow whitespace after ':'.
                self.skip_ws();
                let rhs = self.parse_postfix(stop_at_comma)?;
                expr = Expr::Binary {
                    op: BinaryOp::Range,
                    left: Box::new(expr),
                    right: Box::new(rhs),
                };
                continue;
            }
            break;
        }
        Ok(expr)
    }

    fn parse_postfix(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        let mut expr = self.parse_primary(stop_at_comma)?;
        loop {
            match self.peek_char() {
                Some('#') => {
                    self.next_char();
                    expr = Expr::SpillRange(Box::new(expr));
                }
                Some('%') => {
                    self.next_char();
                    expr = Expr::Percent(Box::new(expr));
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_primary(&mut self, stop_at_comma: bool) -> Result<Expr, String> {
        self.skip_ws();
        let _ = stop_at_comma;
        Ok(match self.peek_char() {
            Some('{') => self.parse_array_literal()?,
            Some('(') => {
                self.next_char();
                // Parentheses allow union operators, even inside function argument lists.
                let expr = self.parse_expr(false)?;
                self.skip_ws();
                if self.next_char() != Some(')') {
                    return Err("expected ')'".to_string());
                }
                expr
            }
            Some('"') => Expr::String(self.parse_string_literal()?),
            Some('#') => Expr::Error(self.parse_error_literal()?),
            Some(ch) if ch.is_ascii_digit() => {
                // Row-range references like `1:3` start with digits, which would otherwise be
                // interpreted as a numeric literal. Try parsing a row range first so we can emit
                // Excel-compatible `PtgArea` tokens for `1:3` / `1:1`.
                if let Some(r) = self.try_parse_ref(None)? {
                    Expr::Ref(r)
                } else {
                    self.parse_number()?
                }
            }
            Some('.') => self.parse_number()?,
            Some('$') => self.parse_ident_or_ref()?,
            Some('[') => self.parse_ident_or_ref()?,
            Some('\'') => self.parse_ident_or_ref()?,
            Some(ch) if is_ident_start(ch) => self.parse_ident_or_ref()?,
            _ => return Err("unexpected token".to_string()),
        })
    }

    fn parse_array_literal(&mut self) -> Result<Expr, String> {
        if self.next_char() != Some('{') {
            return Err("expected '{'".to_string());
        }

        let mut rows: Vec<Vec<ArrayElem>> = Vec::new();
        loop {
            self.skip_ws();

            let mut row: Vec<ArrayElem> = Vec::new();
            row.push(self.parse_array_elem_or_empty()?);
            self.skip_ws();

            while self.peek_char() == Some(',') {
                self.next_char();
                row.push(self.parse_array_elem_or_empty()?);
                self.skip_ws();
            }

            rows.push(row);

            match self.peek_char() {
                Some(';') => {
                    self.next_char();
                    continue;
                }
                Some('}') => {
                    self.next_char();
                    break;
                }
                _ => return Err("expected ';' or '}' in array literal".to_string()),
            }
        }

        if rows.is_empty() {
            return Err("array literal cannot be empty".to_string());
        }

        let cols = rows[0].len();
        if cols == 0 {
            return Err("array literal cannot be empty".to_string());
        }
        if rows.iter().any(|r| r.len() != cols) {
            return Err("array literal rows must have the same number of columns".to_string());
        }

        Ok(Expr::Array(ArrayConst { rows }))
    }

    fn parse_array_elem_or_empty(&mut self) -> Result<ArrayElem, String> {
        self.skip_ws();
        match self.peek_char() {
            Some(',') | Some(';') | Some('}') => Ok(ArrayElem::Empty),
            _ => self.parse_array_elem(),
        }
    }

    fn parse_array_elem(&mut self) -> Result<ArrayElem, String> {
        self.skip_ws();
        match self.peek_char() {
            Some('"') => Ok(ArrayElem::Str(self.parse_string_literal()?)),
            Some('#') => Ok(ArrayElem::Error(self.parse_error_literal()?)),
            Some('+') => {
                self.next_char();
                self.skip_ws();
                match self.parse_number()? {
                    Expr::Number(n) => Ok(ArrayElem::Number(n)),
                    _ => Err("expected number".to_string()),
                }
            }
            Some('-') => {
                self.next_char();
                self.skip_ws();
                match self.parse_number()? {
                    Expr::Number(n) => Ok(ArrayElem::Number(-n)),
                    _ => Err("expected number".to_string()),
                }
            }
            Some(ch) if ch.is_ascii_digit() || ch == '.' => match self.parse_number()? {
                Expr::Number(n) => Ok(ArrayElem::Number(n)),
                _ => Err("expected number".to_string()),
            },
            Some(ch) if is_ident_start(ch) => {
                let ident = self
                    .parse_identifier()?
                    .ok_or_else(|| "expected identifier".to_string())?;
                if ident.eq_ignore_ascii_case("TRUE") {
                    Ok(ArrayElem::Bool(true))
                } else if ident.eq_ignore_ascii_case("FALSE") {
                    Ok(ArrayElem::Bool(false))
                } else {
                    Err(format!("unexpected identifier in array literal: {ident}"))
                }
            }
            _ => Err("unexpected token in array literal".to_string()),
        }
    }

    fn parse_string_literal(&mut self) -> Result<String, String> {
        if self.next_char() != Some('"') {
            return Err("expected string literal".to_string());
        }

        let mut out = String::new();
        loop {
            match self.next_char() {
                Some('"') => {
                    if self.peek_char() == Some('"') {
                        self.next_char();
                        out.push('"');
                        continue;
                    }
                    break;
                }
                Some(ch) => out.push(ch),
                None => return Err("unterminated string literal".to_string()),
            }
        }
        Ok(out)
    }

    fn parse_error_literal(&mut self) -> Result<u8, String> {
        let start = self.pos;
        if self.next_char() != Some('#') {
            return Err("expected error literal".to_string());
        }
        while let Some(ch) = self.peek_char() {
            if matches!(ch, '_' | '/' | '!' | '?') || ch.is_ascii_alphanumeric() {
                self.next_char();
            } else {
                break;
            }
        }
        let raw = &self.input[start..self.pos];
        xlsb_error_code_from_literal(raw).ok_or_else(|| format!("unknown error literal: {raw}"))
    }

    fn parse_ident_or_ref(&mut self) -> Result<Expr, String> {
        if let Some(expr) = self.try_parse_sheet_qualified()? {
            return Ok(expr);
        }

        if let Some(r) = self.try_parse_ref(None)? {
            return Ok(Expr::Ref(r));
        }

        let ident = self
            .parse_identifier()?
            .ok_or_else(|| "expected identifier".to_string())?;

        // Function calls allow optional whitespace between the identifier and `(` (e.g. `SUM (A1)`).
        // But whitespace is also meaningful for the intersection operator, so only consume it if we
        // actually see an opening paren.
        let lparen_pos = self.peek_non_ws_pos();
        if lparen_pos < self.input.len() && self.input[lparen_pos..].starts_with('(') {
            self.pos = lparen_pos;
            self.next_char();
            let mut args = Vec::new();
            self.skip_ws();
            if self.peek_char() != Some(')') {
                loop {
                    // Excel allows empty arguments, which are encoded as `PtgMissArg` in rgce.
                    // For example: `DISC(...,)` leaves the optional `basis` argument blank.
                    if matches!(self.peek_char(), Some(',') | Some(')')) {
                        args.push(Expr::Missing);
                    } else {
                        // Commas delimit arguments at the top level. Union expressions using `,`
                        // must be parenthesized (e.g. `(A1,B1)`).
                        args.push(self.parse_expr(true)?);
                    }
                    self.skip_ws();
                    match self.peek_char() {
                        Some(',') => {
                            self.next_char();
                            self.skip_ws();
                        }
                        Some(')') => break,
                        _ => return Err("expected ',' or ')'".to_string()),
                    }
                }
            }
            self.next_char();
            Ok(Expr::Func { name: ident, args })
        } else {
            if ident.eq_ignore_ascii_case("TRUE") {
                Ok(Expr::Bool(true))
            } else if ident.eq_ignore_ascii_case("FALSE") {
                Ok(Expr::Bool(false))
            } else {
                Ok(Expr::Name(NameRef {
                    sheet: None,
                    name: ident,
                }))
            }
        }
    }

    fn try_parse_sheet_qualified(&mut self) -> Result<Option<Expr>, String> {
        let start = self.pos;
        let Some(first_sheet) = self.parse_sheet_name()? else {
            return Ok(None);
        };
        self.skip_ws();

        let sheet_spec = match self.peek_char() {
            Some('!') => {
                self.next_char();
                Some(SheetSpec::Single(first_sheet))
            }
            Some(':') => {
                self.next_char();
                let Some(second_sheet) = self.parse_sheet_name()? else {
                    // If we can't parse a second sheet name, treat this as a normal `:` range
                    // operator (e.g. `A1:$A$2`) rather than a sheet span.
                    self.pos = start;
                    return Ok(None);
                };
                self.skip_ws();
                if self.peek_char() != Some('!') {
                    self.pos = start;
                    return Ok(None);
                }
                self.next_char();
                Some(SheetSpec::Range(first_sheet, second_sheet))
            }
            _ => {
                self.pos = start;
                None
            }
        };

        let Some(sheet_spec) = sheet_spec else {
            return Ok(None);
        };

        if let Some(r) = self.try_parse_ref(Some(sheet_spec.clone()))? {
            return Ok(Some(Expr::Ref(r)));
        }

        let name = self
            .parse_identifier()?
            .ok_or_else(|| "expected name after sheet qualifier".to_string())?;

        let sheet_name = match sheet_spec {
            SheetSpec::Single(s) => Some(s),
            SheetSpec::Range(first, last) => Some(format!("{first}:{last}")),
        };

        Ok(Some(Expr::Name(NameRef {
            sheet: sheet_name,
            name,
        })))
    }

    fn try_parse_ref(&mut self, sheet: Option<SheetSpec>) -> Result<Option<Ref>, String> {
        let start = self.pos;
        if let Some(a) = self.parse_cell_ref()? {
            let after_a = self.pos;

            // Area references like `A1:B2` allow optional whitespace around the `:` operator. But
            // whitespace is also significant for the intersection operator, so only consume it if
            // the next non-whitespace character is actually `:`.
            let colon_pos = self.peek_non_ws_pos();
            if colon_pos < self.input.len() && self.input[colon_pos..].starts_with(':') {
                // Commit whitespace and consume ':'.
                self.pos = colon_pos;
                self.next_char();
                // Allow whitespace after ':'.
                self.skip_ws();

                if let Some(b) = self.parse_cell_ref()? {
                    return Ok(Some(Ref {
                        sheet,
                        kind: RefKind::Area(a, b),
                    }));
                }

                // Not a simple area ref (`A1:B2`). Leave the `:` operator to be handled as a
                // general range expression (`PtgRange`) by higher-precedence parsing.
                self.pos = after_a;
            }

            return Ok(Some(Ref {
                sheet,
                kind: RefKind::Cell(a),
            }));
        }
        self.pos = start;

        // Column ranges like `A:C` / `A:A`.
        if let Some((col_a, abs_col_a)) = self.parse_col_ref()? {
            let colon_pos = self.peek_non_ws_pos();
            if colon_pos < self.input.len() && self.input[colon_pos..].starts_with(':') {
                self.pos = colon_pos;
                self.next_char();
                self.skip_ws();
                if let Some((col_b, abs_col_b)) = self.parse_col_ref()? {
                    const MAX_ROW: u32 = 1_048_575;
                    return Ok(Some(Ref {
                        sheet,
                        kind: RefKind::Area(
                            CellRef {
                                row: 0,
                                col: col_a,
                                abs_row: true,
                                abs_col: abs_col_a,
                            },
                            CellRef {
                                row: MAX_ROW,
                                col: col_b,
                                abs_row: true,
                                abs_col: abs_col_b,
                            },
                        ),
                    }));
                }
            }
            // Only treat this as a column reference if both sides are valid column refs. Otherwise
            // roll back so the caller can interpret it as a defined name, number, etc.
            self.pos = start;
        }

        // Row ranges like `1:3` / `1:1`.
        if let Some((row_a, abs_row_a)) = self.parse_row_ref()? {
            let colon_pos = self.peek_non_ws_pos();
            if colon_pos < self.input.len() && self.input[colon_pos..].starts_with(':') {
                self.pos = colon_pos;
                self.next_char();
                self.skip_ws();
                if let Some((row_b, abs_row_b)) = self.parse_row_ref()? {
                    const MAX_COL: u32 = COL_INDEX_MASK as u32;
                    return Ok(Some(Ref {
                        sheet,
                        kind: RefKind::Area(
                            CellRef {
                                row: row_a,
                                col: 0,
                                abs_row: abs_row_a,
                                abs_col: true,
                            },
                            CellRef {
                                row: row_b,
                                col: MAX_COL,
                                abs_row: abs_row_b,
                                abs_col: true,
                            },
                        ),
                    }));
                }
            }
            self.pos = start;
        }

        Ok(None)
    }

    fn parse_number(&mut self) -> Result<Expr, String> {
        let start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() || ch == '.' {
                self.next_char();
            } else {
                break;
            }
        }
        // Optional scientific notation suffix: `E[+|-]?digits`.
        let exp_start = self.pos;
        if matches!(self.peek_char(), Some('e' | 'E')) {
            self.next_char();
            if matches!(self.peek_char(), Some('+' | '-')) {
                self.next_char();
            }
            let digits_start = self.pos;
            while let Some(ch) = self.peek_char() {
                if ch.is_ascii_digit() {
                    self.next_char();
                } else {
                    break;
                }
            }
            // If we didn't consume any exponent digits, roll back and treat the `E` as a separate
            // token (the formula will error later, but this avoids silently accepting invalid
            // numeric literals like `1E`).
            if self.pos == digits_start {
                self.pos = exp_start;
            }
        }
        let s = &self.input[start..self.pos];
        let n: f64 = s
            .parse()
            .map_err(|_| format!("invalid number literal: {s}"))?;
        Ok(Expr::Number(n))
    }

    fn parse_col_ref(&mut self) -> Result<Option<(u32, bool)>, String> {
        let start = self.pos;
        let abs_col = self.consume_if('$');
        let col_start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphabetic() {
                self.next_char();
            } else {
                break;
            }
        }
        let col_label = &self.input[col_start..self.pos];
        if col_label.is_empty() {
            self.pos = start;
            return Ok(None);
        }
        if col_label.len() > 3 {
            self.pos = start;
            return Ok(None);
        }

        let col = formula_model::column_label_to_index_lenient(col_label)
            .map_err(|_| "invalid column label".to_string())?;
        if col > COL_INDEX_MASK as u32 {
            self.pos = start;
            return Ok(None);
        }

        Ok(Some((col, abs_col)))
    }

    fn parse_row_ref(&mut self) -> Result<Option<(u32, bool)>, String> {
        let start = self.pos;
        let abs_row = self.consume_if('$');
        let row_start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() {
                self.next_char();
            } else {
                break;
            }
        }
        let row_str = &self.input[row_start..self.pos];
        if row_str.is_empty() {
            self.pos = start;
            return Ok(None);
        }

        let row1: u32 = row_str.parse().map_err(|_| "invalid row".to_string())?;
        if row1 == 0 {
            self.pos = start;
            return Ok(None);
        }
        let row = row1 - 1;
        if row > 1_048_575 {
            self.pos = start;
            return Ok(None);
        }

        Ok(Some((row, abs_row)))
    }

    fn parse_cell_ref(&mut self) -> Result<Option<CellRef>, String> {
        let start = self.pos;
        let abs_col = self.consume_if('$');
        let col_start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphabetic() {
                self.next_char();
            } else {
                break;
            }
        }
        let col_label = &self.input[col_start..self.pos];
        if col_label.is_empty() {
            self.pos = start;
            return Ok(None);
        }
        if col_label.len() > 3 {
            self.pos = start;
            return Ok(None);
        }
        let abs_row = self.consume_if('$');
        let row_start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() {
                self.next_char();
            } else {
                break;
            }
        }
        let row_str = &self.input[row_start..self.pos];
        if row_str.is_empty() {
            self.pos = start;
            return Ok(None);
        }

        let col = formula_model::column_label_to_index_lenient(col_label)
            .map_err(|_| "invalid column label".to_string())?;
        let row1: u32 = row_str.parse().map_err(|_| "invalid row".to_string())?;
        if row1 == 0 {
            self.pos = start;
            return Ok(None);
        }
        let row = row1 - 1;

        if col > 16383 || row > 1048575 {
            self.pos = start;
            return Ok(None);
        }

        Ok(Some(CellRef {
            row,
            col,
            abs_row,
            abs_col,
        }))
    }

    fn parse_sheet_name(&mut self) -> Result<Option<String>, String> {
        self.skip_ws();
        match self.peek_char() {
            Some('\'') => self.parse_quoted_sheet_name().map(Some),
            Some('[') => {
                let start = self.pos;
                let parsed = self.parse_external_sheet_name()?;
                if parsed.is_none() {
                    self.pos = start;
                }
                Ok(parsed)
            }
            Some(ch) if is_ident_start(ch) => Ok(self.parse_identifier()?),
            _ => Ok(None),
        }
    }

    fn parse_external_sheet_name(&mut self) -> Result<Option<String>, String> {
        if self.peek_char() != Some('[') {
            return Ok(None);
        }

        let start = self.pos;
        let end = match formula_model::external_refs::find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(
            self.input,
            start,
        ) {
            Some(end) => end,
            None => {
                // If we saw at least one unescaped `]` but none were followed by a plausible sheet
                // name token, this is not an external workbook sheet prefix (it may be a structured
                // ref like `[@Col]`).
                if formula_model::external_refs::find_external_workbook_prefix_end(self.input, start)
                    .is_some()
                {
                    return Ok(None);
                }
                return Err("unterminated external workbook prefix".to_string());
            }
        };
        let Some(book_start) = start.checked_add(1) else {
            return Err("unterminated external workbook prefix".to_string());
        };
        let Some(book_end) = end.checked_sub(1) else {
            return Err("unterminated external workbook prefix".to_string());
        };
        if book_start >= book_end {
            return Ok(None);
        }

        let book = self
            .input
            .get(book_start..book_end)
            .ok_or_else(|| "unterminated external workbook prefix".to_string())?;
        self.pos = end;
        self.skip_ws();

        let sheet = match self.peek_char() {
            Some('\'') => Some(self.parse_quoted_sheet_name()?),
            Some(ch) if is_ident_start(ch) => self.parse_identifier()?,
            _ => None,
        };
        let Some(sheet) = sheet else {
            self.pos = start;
            return Ok(None);
        };

        Ok(Some(format_external_key(book, &sheet)))
    }

    fn parse_quoted_sheet_name(&mut self) -> Result<String, String> {
        if self.next_char() != Some('\'') {
            return Err("expected \"'\"".to_string());
        }
        let mut out = String::new();
        loop {
            match self.next_char() {
                Some('\'') => {
                    if self.peek_char() == Some('\'') {
                        self.next_char();
                        out.push('\'');
                        continue;
                    }
                    break;
                }
                Some(ch) => out.push(ch),
                None => return Err("unterminated quoted sheet name".to_string()),
            }
        }
        Ok(out)
    }

    fn parse_identifier(&mut self) -> Result<Option<String>, String> {
        self.skip_ws();
        let start = self.pos;
        let Some(ch) = self.peek_char() else {
            return Ok(None);
        };
        if !is_ident_start(ch) {
            return Ok(None);
        }
        self.next_char();
        while let Some(ch) = self.peek_char() {
            if is_ident_continue(ch) {
                self.next_char();
            } else {
                break;
            }
        }
        Ok(Some(self.input[start..self.pos].to_string()))
    }

    fn peek_non_ws_pos(&self) -> usize {
        let mut i = self.pos;
        while i < self.input.len() {
            let Some(rest) = self.input.get(i..) else {
                break;
            };
            let Some(ch) = rest.chars().next() else {
                break;
            };
            if ch.is_whitespace() {
                i += ch.len_utf8();
            } else {
                break;
            }
        }
        i
    }

    fn looks_like_row_range_start(&self) -> bool {
        let mut i = self.pos;
        while i < self.input.len() {
            let Some(rest) = self.input.get(i..) else {
                break;
            };
            let Some(ch) = rest.chars().next() else {
                break;
            };
            if ch.is_ascii_digit() {
                i += ch.len_utf8();
            } else {
                break;
            }
        }
        if i == self.pos {
            return false;
        }
        while i < self.input.len() {
            let Some(rest) = self.input.get(i..) else {
                break;
            };
            let Some(ch) = rest.chars().next() else {
                break;
            };
            if ch.is_whitespace() {
                i += ch.len_utf8();
            } else {
                break;
            }
        }
        i < self.input.len() && self.input.get(i..).is_some_and(|rest| rest.starts_with(':'))
    }

    fn is_intersection_rhs_start(&self) -> bool {
        match self.peek_char() {
            Some('$' | '[' | '\'' | '(') => true,
            Some(ch) if ch.is_ascii_digit() => self.looks_like_row_range_start(),
            Some(ch) if is_ident_start(ch) => true,
            _ => false,
        }
    }

    fn skip_ws(&mut self) -> bool {
        let start = self.pos;
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.next_char();
            } else {
                break;
            }
        }
        self.pos != start
    }

    fn consume_if(&mut self, expected: char) -> bool {
        if self.peek_char() == Some(expected) {
            self.next_char();
            true
        } else {
            false
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.pos..].chars().next()
    }

    fn next_char(&mut self) -> Option<char> {
        let ch = self.peek_char()?;
        self.pos += ch.len_utf8();
        Some(ch)
    }
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_'
}

fn is_ident_continue(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '.'
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_sheet_name_accepts_canonical_external_workbook_with_bracketed_path_components() {
        let mut parser = FormulaParser::new(r"[C:\[foo]\Book.xlsx]Sheet1");
        let parsed = parser.parse_sheet_name().expect("parse");
        assert_eq!(parsed, Some(r"[C:\[foo]\Book.xlsx]Sheet1".to_string()));
    }

    #[test]
    fn parse_sheet_name_accepts_external_workbook_names_containing_lbracket() {
        let mut parser = FormulaParser::new("[A1[Name.xlsx]Sheet1");
        let parsed = parser.parse_sheet_name().expect("parse");
        assert_eq!(parsed, Some("[A1[Name.xlsx]Sheet1".to_string()));
    }

    #[test]
    fn parse_sheet_name_rejects_structured_refs_starting_with_bracket() {
        let mut parser = FormulaParser::new("[@Col2]");
        let parsed = parser.parse_sheet_name().expect("parse");
        assert_eq!(parsed, None);
    }

    #[test]
    fn parse_sheet_name_rejects_bracket_only_tokens() {
        let mut parser = FormulaParser::new("[Book.xlsx]");
        let parsed = parser.parse_sheet_name().expect("parse");
        assert_eq!(parsed, None);
    }
}

// Column-label parsing is centralized in `formula-model` to ensure consistent A1 grammar handling
// and overflow safety.
