//! BIFF12 `rgce` (formula token stream) codec.
//!
//! The decoder is best-effort and is primarily used for diagnostics and for surfacing formula
//! text when reading XLSB files.
//!
//! The encoder is intentionally small-scope: it supports enough of Excel's formula language to
//! round-trip common patterns while we build out full compatibility.

use crate::errors::xlsb_error_literal;
use crate::format::push_column_label;
use crate::formula_text::escape_excel_string_literal;
use crate::workbook_context::{NameScope, WorkbookContext};
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
    OutputTooLarge { offset: usize, ptg: u8, max_len: usize },
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
    let text = decode_rgce_impl(rgce, rgcb, ctx, base, Some(&mut warnings)).ok();
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
    // Based on `formula-engine`'s `fmt_sheet_name`: be conservative, because quoting is always
    // accepted by Excel while the unquoted form is only valid for identifier-like names.
    if sheet_name_needs_quotes(sheet) {
        out.push('\'');
        for ch in sheet.chars() {
            if ch == '\'' {
                out.push('\'');
                out.push('\'');
            } else {
                out.push(ch);
            }
        }
        out.push('\'');
    } else {
        out.push_str(sheet);
    }
}

fn format_sheet_prefix(first: &str, last: &str) -> String {
    let mut out = String::new();
    if first == last {
        fmt_sheet_name(&mut out, first);
    } else if sheet_name_needs_quotes(first) || sheet_name_needs_quotes(last) {
        // Excel quotes the combined `Sheet1:Sheet3` prefix as a single string.
        out.push('\'');
        for ch in format!("{first}:{last}").chars() {
            if ch == '\'' {
                out.push('\'');
                out.push('\'');
            } else {
                out.push(ch);
            }
        }
        out.push('\'');
    } else {
        out.push_str(first);
        out.push(':');
        out.push_str(last);
    }
    out.push('!');
    out
}

fn sheet_name_needs_quotes(sheet: &str) -> bool {
    if sheet.is_empty() {
        return true;
    }
    if sheet
        .chars()
        .any(|c| c.is_whitespace() || matches!(c, '!' | '\''))
    {
        return true;
    }
    sheet_part_needs_quotes(sheet)
}

fn sheet_part_needs_quotes(sheet: &str) -> bool {
    debug_assert!(!sheet.is_empty());

    if sheet.eq_ignore_ascii_case("TRUE") || sheet.eq_ignore_ascii_case("FALSE") {
        return true;
    }

    // Quote sheet names that look like cell refs (e.g. `A1`, `$B$2`).
    if starts_like_a1_cell_ref(sheet) {
        return true;
    }

    !is_valid_sheet_ident(sheet)
}

fn is_valid_sheet_ident(ident: &str) -> bool {
    let mut chars = ident.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z') {
        return false;
    }
    chars.all(is_sheet_ident_cont_char)
}

fn is_sheet_ident_cont_char(c: char) -> bool {
    matches!(
        c,
        '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'
    )
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

fn col_from_a1(letters: &str) -> Option<u32> {
    let mut col: u32 = 0;
    for (i, ch) in letters.chars().enumerate() {
        let v = (ch.to_ascii_uppercase() as u8).wrapping_sub(b'A') as u32;
        if v >= 26 {
            return None;
        }
        col = col * 26 + v + 1;
        if i >= 3 {
            return None;
        }
    }
    Some(col - 1)
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
        .saturating_add(rgcb.len())
        .saturating_mul(MAX_OUTPUT_FACTOR);

    let mut i = 0usize;
    let mut last_ptg_offset = 0usize;
    let mut last_ptg = rgce[0];
    let mut rgcb_pos = 0usize;

    let mut stack: Vec<ExprFragment> = Vec::new();

    while i < rgce.len() {
        let ptg_offset = i;
        let ptg = rgce[i];
        i += 1;

        last_ptg_offset = ptg_offset;
        last_ptg = ptg;

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };
                let prec = binary_precedence(ptg).expect("precedence for binary ops");

                if stack.len() < 2 {
                    return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                }
                let right = stack.pop().expect("len checked");
                let left = stack.pop().expect("len checked");

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
            // Unary +/-.
            0x12 | 0x13 => {
                let op = if ptg == 0x12 { "+" } else { "-" };
                let expr = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                let prec = 70;
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
                let expr = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                let prec = 60;
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
            PTG_SPILL => {
                let expr = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                let prec = 60;
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
                let expr = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                stack.push(ExprFragment {
                    text: format!("({})", expr.text),
                    precedence: 100,
                    contains_union: expr.contains_union,
                    is_missing: false,
                });
            }
            // Missing arg.
            0x16 => stack.push(ExprFragment::missing()),
            // PtgStr: [cch: u16][utf16 chars...]
            0x17 => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cch = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;

                let needed = cch.saturating_mul(2);
                if rgce.len().saturating_sub(i) < needed {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let raw = &rgce[i..i + needed];
                i += needed;

                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }

                // Excel escapes embedded quotes by doubling them inside the literal.
                let s = String::from_utf16_lossy(&units);
                let escaped = escape_excel_string_literal(&s);
                stack.push(ExprFragment::new(format!("\"{escaped}\"")));
            }
            0x19 => {
                // PtgAttr: [grbit: u8][wAttr: u16]
                //
                // Excel uses `PtgAttr` for multiple attributes. Most are evaluation hints or
                // formatting metadata that do not affect the reconstructed formula text, but some
                // do. In particular, `tAttrSum` is used for an optimization where `SUM(A1:A10)` is
                // encoded as `PtgArea` + `PtgAttr(tAttrSum)` (no explicit `PtgFuncVar(SUM)` token).
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let grbit = rgce[i];
                let w_attr = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]);
                i += 3;

                const T_ATTR_VOLATILE: u8 = 0x01;
                const T_ATTR_IF: u8 = 0x02;
                const T_ATTR_CHOOSE: u8 = 0x04;
                const T_ATTR_SKIP: u8 = 0x08;
                const T_ATTR_SUM: u8 = 0x10;
                const T_ATTR_SPACE: u8 = 0x40;
                const T_ATTR_SEMI: u8 = 0x80;

                if grbit & T_ATTR_SUM != 0 {
                    let a = stack
                        .pop()
                        .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                    stack.push(format_function_call("SUM", vec![a]));
                } else if grbit & T_ATTR_CHOOSE != 0 {
                    // `tAttrChoose` is followed by a jump table of `u16` offsets used for
                    // short-circuit evaluation.
                    //
                    // We don't need it for printing, but we must consume it so subsequent tokens
                    // stay aligned.
                    let needed = (w_attr as usize).saturating_mul(2);
                    if rgce.len().saturating_sub(i) < needed {
                        return Err(DecodeError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed,
                            remaining: rgce.len().saturating_sub(i),
                        });
                    }
                    i += needed;
                } else {
                    // Ignore other attributes for printing, but keep the constants referenced so
                    // this doesn't accidentally get treated as dead code.
                    let _ = grbit
                        & (T_ATTR_VOLATILE | T_ATTR_IF | T_ATTR_SKIP | T_ATTR_SPACE | T_ATTR_SEMI);
                }
            }
            // PtgErr: [err: u8]
            0x1C => {
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let code_offset = i;
                let err = rgce[i];
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
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let b = rgce[i];
                i += 1;
                stack.push(ExprFragment::new(
                    if b == 0 { "FALSE" } else { "TRUE" }.to_string(),
                ));
            }
            // PtgInt: [n: u16]
            0x1E => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let n = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                i += 2;
                stack.push(ExprFragment::new(n.to_string()));
            }
            // PtgNum: [f64]
            0x1F => {
                if rgce.len().saturating_sub(i) < 8 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&rgce[i..i + 8]);
                i += 8;
                stack.push(ExprFragment::new(f64::from_le_bytes(bytes).to_string()));
            }
            // PtgArray: [unused: 7 bytes] + serialized array constant stored in rgcb.
            0x20 | 0x40 | 0x60 => {
                if rgce.len().saturating_sub(i) < 7 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 7,
                        remaining: rgce.len().saturating_sub(i),
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
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let row0 = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                let row = (row0 as u64).saturating_add(1);
                let flags = rgce[i + 5];
                let col = u16::from_le_bytes([rgce[i + 4], flags & 0x3F]);
                i += 6;

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

                let row_first0 =
                    u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                let row_last0 = u32::from_le_bytes([
                    rgce[i + 4],
                    rgce[i + 5],
                    rgce[i + 6],
                    rgce[i + 7],
                ]);
                let col_first = u16::from_le_bytes([rgce[i + 8], rgce[i + 9]]);
                let col_last = u16::from_le_bytes([rgce[i + 10], rgce[i + 11]]);
                i += 12;

                let a = format_cell_ref_from_field(row_first0, col_first);
                let b = format_cell_ref_from_field(row_last0, col_last);
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
                if is_single_cell {
                    text.push_str(&a);
                } else {
                    text.push_str(&a);
                    text.push(':');
                    text.push_str(&b);
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
                let cce = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;
                if rgce.len().saturating_sub(i) < cce {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += cce;
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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let row_off =
                    i32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) as i64;
                let col_off = i16::from_le_bytes([rgce[i + 4], rgce[i + 5]]) as i64;
                i += 6;

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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let row_first_off =
                    i32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) as i64;
                let row_last_off = i32::from_le_bytes([
                    rgce[i + 4],
                    rgce[i + 5],
                    rgce[i + 6],
                    rgce[i + 7],
                ]) as i64;
                let col_first_off = i16::from_le_bytes([rgce[i + 8], rgce[i + 9]]) as i64;
                let col_last_off = i16::from_le_bytes([rgce[i + 10], rgce[i + 11]]) as i64;
                i += 12;

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
                    let col_first = encode_col_field(abs_col_first as u32, false, false);
                    let col_last = encode_col_field(abs_col_last as u32, false, false);
                    let a = format_cell_ref_from_field(abs_row_first as u32, col_first);
                    let b = format_cell_ref_from_field(abs_row_last as u32, col_last);

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
                        text.push_str(&a);
                    } else {
                        text.push_str(&a);
                        text.push(':');
                        text.push_str(&b);
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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let row0 = u32::from_le_bytes([rgce[i + 2], rgce[i + 3], rgce[i + 4], rgce[i + 5]]);
                let col_field = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                i += 8;

                let (first, last) = ctx
                    .extern_sheet_names(ixti)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;

                let prefix = format_sheet_prefix(first, last);
                let cell = format_cell_ref_from_field(row0, col_field);
                stack.push(ExprFragment::new(format!("{prefix}{cell}")));
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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let row_first0 =
                    u32::from_le_bytes([rgce[i + 2], rgce[i + 3], rgce[i + 4], rgce[i + 5]]);
                let row_last0 =
                    u32::from_le_bytes([rgce[i + 6], rgce[i + 7], rgce[i + 8], rgce[i + 9]]);
                let col_first = u16::from_le_bytes([rgce[i + 10], rgce[i + 11]]);
                let col_last = u16::from_le_bytes([rgce[i + 12], rgce[i + 13]]);
                i += 14;

                let (first, last) = ctx
                    .extern_sheet_names(ixti)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;
                let prefix = format_sheet_prefix(first, last);

                let a = format_cell_ref_from_field(row_first0, col_first);
                let b = format_cell_ref_from_field(row_last0, col_last);
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
                if is_single_cell {
                    text.push_str(&a);
                } else {
                    text.push_str(&a);
                    text.push(':');
                    text.push_str(&b);
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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let name_id =
                    u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                i += 4;
                i += 2; // reserved

                let def = ctx
                    .name_definition(name_id)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;

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
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                };

                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let name_index = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                i += 4;

                let txt = ctx
                    .format_namex(ixti, name_index)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;
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

                let iftab = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                i += 2;

                let spec = formula_biff::function_spec_from_id(iftab)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;
                if spec.min_args != spec.max_args {
                    return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg });
                }

                let argc = spec.min_args as usize;
                if stack.len() < argc {
                    return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
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
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let argc = rgce[i] as usize;
                let iftab = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]);
                i += 3;

                if stack.len() < argc {
                    return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                }

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name (typically from `PtgNameX`), followed by args.
                if iftab == 0x00FF {
                    if argc == 0 {
                        return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                    }

                    let func_name = stack.pop().expect("len checked").text;
                    let mut args = Vec::with_capacity(argc - 1);
                    for _ in 0..argc.saturating_sub(1) {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format_function_call(&func_name, args));
                } else {
                    let name =
                        function_name(iftab).ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;

                    let mut args = Vec::with_capacity(argc);
                    for _ in 0..argc {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format_function_call(name, args));
                }
            }
            _ => return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg }),
        }

        if stack
            .last()
            .is_some_and(|s| s.text.len() > max_len)
        {
            return Err(DecodeError::OutputTooLarge {
                offset: ptg_offset,
                ptg,
                max_len,
            });
        }
    }

    if stack.len() == 1 {
        Ok(stack.pop().expect("len checked").text)
    } else {
        Err(DecodeError::StackNotSingular {
            offset: last_ptg_offset,
            ptg: last_ptg,
            stack_len: stack.len(),
        })
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

fn error_literal(code: u8) -> Option<&'static str> {
    match code {
        0x00 => Some("#NULL!"),
        0x07 => Some("#DIV/0!"),
        0x0F => Some("#VALUE!"),
        0x17 => Some("#REF!"),
        0x1D => Some("#NAME?"),
        0x24 => Some("#NUM!"),
        0x2A => Some("#N/A"),
        0x2B => Some("#GETTING_DATA"),
        _ => None,
    }
}

fn error_code_from_literal(literal: &str) -> Option<u8> {
    match literal.trim().to_ascii_uppercase().as_str() {
        "#NULL!" => Some(0x00),
        "#DIV/0!" => Some(0x07),
        "#VALUE!" => Some(0x0F),
        "#REF!" => Some(0x17),
        "#NAME?" => Some(0x1D),
        "#NUM!" => Some(0x24),
        "#N/A" => Some(0x2A),
        "#GETTING_DATA" => Some(0x2B),
        _ => None,
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
                    col_texts.push(format!("\"{}\"", escape_excel_string(&s)));
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
                    let code_offset = i;
                    let code = rgcb[i];
                    i += 1;
                    match error_literal(code) {
                        Some(lit) => col_texts.push(lit.to_string()),
                        None => {
                            if let Some(warnings) = warnings.as_deref_mut() {
                                warnings.push(DecodeWarning::UnknownArrayErrorCode {
                                    code,
                                    offset: code_offset,
                                });
                            }
                            col_texts.push("#UNKNOWN!".to_string());
                        }
                    }
                }
                _ => return None,
            }
        }
        row_texts.push(col_texts.join(","));
    }

    *pos = i;
    Some(format!("{{{}}}", row_texts.join(";")))
}

fn format_cell_ref(row1: u64, col: u32, flags: u8) -> String {
    let mut out = String::new();
    if flags & 0x80 != 0x80 {
        out.push('$');
    }
    push_column_label(col, &mut out);
    if flags & 0x40 != 0x40 {
        out.push('$');
    }
    out.push_str(&row1.to_string());
    out
}

fn format_cell_ref_from_field(row0: u32, col_field: u16) -> String {
    let row1 = (row0 as u64).saturating_add(1);
    let col = (col_field & 0x3FFF) as u32;
    let col_relative = (col_field & 0x8000) == 0x8000;
    let row_relative = (col_field & 0x4000) == 0x4000;

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
    #[error("unknown sheet reference: {0}")]
    UnknownSheet(String),
    #[error("unknown name: {name}")]
    UnknownName { name: String },
    #[error("unknown function: {name}")]
    UnknownFunction { name: String },
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
    Ok(EncodedRgce {
        rgce,
        rgcb,
    })
}

const PTG_ADD: u8 = 0x03;
const PTG_SUB: u8 = 0x04;
const PTG_MUL: u8 = 0x05;
const PTG_DIV: u8 = 0x06;
const PTG_UPLUS: u8 = 0x12;
const PTG_UMINUS: u8 = 0x13;
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
    Number(f64),
    Ref(Ref),
    Name(NameRef),
    Array(ArrayConst),
    Func { name: String, args: Vec<Expr> },
    SpillRange(Box<Expr>),
    Unary { op: UnaryOp, expr: Box<Expr> },
    Binary { op: BinaryOp, left: Box<Expr>, right: Box<Expr> },
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
        Expr::Number(n) => emit_number(*n, rgce),
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

fn emit_array(array: &ArrayConst, rgce: &mut Vec<u8>, rgcb: &mut Vec<u8>) -> Result<(), EncodeError> {
    rgce.push(ptg_with_class(PTG_ARRAY, PtgClass::Array));
    rgce.extend_from_slice(&[0u8; 7]); // reserved
    encode_array_constant(array, rgcb)
}

fn encode_array_constant(array: &ArrayConst, rgcb: &mut Vec<u8>) -> Result<(), EncodeError> {
    let rows = array.rows.len();
    let cols = array.rows.first().map(|r| r.len()).unwrap_or(0);
    if rows == 0 || cols == 0 {
        return Err(EncodeError::Parse("array constant cannot be empty".to_string()));
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
                    rgcb.push(0x02);
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let len: u16 = units
                        .len()
                        .try_into()
                        .map_err(|_| EncodeError::Parse("array string literal is too long".to_string()))?;
                    rgcb.extend_from_slice(&len.to_le_bytes());
                    for u in units {
                        rgcb.extend_from_slice(&u.to_le_bytes());
                    }
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
    let upper = name.trim().to_ascii_uppercase();

    // Built-in functions.
    //
    // Note: Excel encodes "future" (forward-compatible) functions as user-defined calls (iftab=255)
    // paired with a name token. We do not currently support that encoding path, so we explicitly
    // avoid treating `iftab=0x00FF` as a built-in here.
    if let Some(iftab) = formula_biff::function_name_to_id(&upper).filter(|id| *id != 0x00FF) {
        if argc > u8::MAX as usize {
            return Err(EncodeError::Parse("too many function arguments".to_string()));
        }
        out.push(PTG_FUNCVAR);
        out.push(argc as u8);
        out.extend_from_slice(&iftab.to_le_bytes());
        return Ok(());
    }

    // Add-in / UDF call pattern: args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
    if let Some((ixti, name_index)) = ctx.namex_function_ref(&upper) {
        let argc_total = argc
            .checked_add(1)
            .ok_or_else(|| EncodeError::Parse("too many function arguments".to_string()))?;
        if argc_total > u8::MAX as usize {
            return Err(EncodeError::Parse("too many function arguments".to_string()));
        }

        out.push(PTG_NAMEX);
        out.extend_from_slice(&ixti.to_le_bytes());
        out.extend_from_slice(&name_index.to_le_bytes());

        out.push(PTG_FUNCVAR);
        out.push(argc_total as u8);
        out.extend_from_slice(&0x00FFu16.to_le_bytes());
        return Ok(());
    }

    Err(EncodeError::UnknownFunction { name: upper })
}

fn emit_name(
    name: &NameRef,
    ctx: &WorkbookContext,
    out: &mut Vec<u8>,
    class: PtgClass,
) -> Result<(), EncodeError> {
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

fn emit_ref(r: &Ref, ctx: &WorkbookContext, out: &mut Vec<u8>, class: PtgClass) -> Result<(), EncodeError> {
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
                SheetSpec::Single(s) => (s.as_str(), s.as_str()),
                SheetSpec::Range(a, b) => (a.as_str(), b.as_str()),
            };
            let ixti = ctx
                .extern_sheet_range_index(first, last)
                .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;

            out.push(ptg_with_class(PTG_REF3D, class));
            out.extend_from_slice(&ixti.to_le_bytes());
            emit_cell_ref_fields(cell, out);
        }
        (Some(sheet), RefKind::Area(a, b)) => {
            let (first, last) = match sheet {
                SheetSpec::Single(s) => (s.as_str(), s.as_str()),
                SheetSpec::Range(a, b) => (a.as_str(), b.as_str()),
            };
            let ixti = ctx
                .extern_sheet_range_index(first, last)
                .ok_or_else(|| EncodeError::UnknownSheet(format!("{first}:{last}")))?;

            out.push(ptg_with_class(PTG_AREA3D, class));
            out.extend_from_slice(&ixti.to_le_bytes());
            emit_area_fields(a, b, out);
        }
    }
    Ok(())
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

    out.extend_from_slice(&row_first.to_le_bytes());
    out.extend_from_slice(&row_last.to_le_bytes());
    out.extend_from_slice(&encode_col_field(col_first, a.abs_row, a.abs_col).to_le_bytes());
    out.extend_from_slice(&encode_col_field(col_last, b.abs_row, b.abs_col).to_le_bytes());
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
        let expr = self.parse_add_sub()?;
        self.skip_ws();
        if self.pos < self.input.len() {
            return Err(format!("unexpected trailing input at byte {}", self.pos));
        }
        Ok(expr)
    }

    fn parse_add_sub(&mut self) -> Result<Expr, String> {
        let mut expr = self.parse_mul_div()?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('+') => BinaryOp::Add,
                Some('-') => BinaryOp::Sub,
                _ => break,
            };
            self.next_char();
            let rhs = self.parse_mul_div()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_mul_div(&mut self) -> Result<Expr, String> {
        let mut expr = self.parse_unary()?;
        loop {
            self.skip_ws();
            let op = match self.peek_char() {
                Some('*') => BinaryOp::Mul,
                Some('/') => BinaryOp::Div,
                _ => break,
            };
            self.next_char();
            let rhs = self.parse_unary()?;
            expr = Expr::Binary {
                op,
                left: Box::new(expr),
                right: Box::new(rhs),
            };
        }
        Ok(expr)
    }

    fn parse_unary(&mut self) -> Result<Expr, String> {
        self.skip_ws();
        match self.peek_char() {
            Some('+') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            Some('-') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::Minus,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            Some('@') => {
                self.next_char();
                Ok(Expr::Unary {
                    op: UnaryOp::ImplicitIntersection,
                    expr: Box::new(self.parse_unary()?),
                })
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, String> {
        self.skip_ws();
        let mut expr = match self.peek_char() {
            Some('{') => self.parse_array_literal()?,
            Some('(') => {
                self.next_char();
                let expr = self.parse_add_sub()?;
                self.skip_ws();
                if self.next_char() != Some(')') {
                    return Err("expected ')'".to_string());
                }
                expr
            }
            Some(ch) if ch.is_ascii_digit() || ch == '.' => self.parse_number()?,
            Some('\'') => self.parse_ident_or_ref()?,
            Some(ch) if is_ident_start(ch) => self.parse_ident_or_ref()?,
            _ => return Err("unexpected token".to_string()),
        };

        self.skip_ws();
        while self.peek_char() == Some('#') {
            self.next_char();
            expr = Expr::SpillRange(Box::new(expr));
            self.skip_ws();
        }

        Ok(expr)
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
                match ident.to_ascii_uppercase().as_str() {
                    "TRUE" => Ok(ArrayElem::Bool(true)),
                    "FALSE" => Ok(ArrayElem::Bool(false)),
                    _ => Err(format!("unexpected identifier in array literal: {ident}")),
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
        while let Some(ch) = self.peek_char() {
            if matches!(ch, ',' | ';' | '}' ) || ch.is_whitespace() {
                break;
            }
            self.next_char();
        }
        let raw = &self.input[start..self.pos];
        error_code_from_literal(raw).ok_or_else(|| format!("unknown error literal: {raw}"))
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
        self.skip_ws();
        if self.peek_char() == Some('(') {
            self.next_char();
            let mut args = Vec::new();
            self.skip_ws();
            if self.peek_char() != Some(')') {
                loop {
                    args.push(self.parse_add_sub()?);
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
            Ok(Expr::Name(NameRef { sheet: None, name: ident }))
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
                let second_sheet = self
                    .parse_sheet_name()?
                    .ok_or_else(|| "expected sheet name after ':'".to_string())?;
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
            SheetSpec::Range(_, _) => None,
        };

        Ok(Some(Expr::Name(NameRef { sheet: sheet_name, name })))
    }

    fn try_parse_ref(&mut self, sheet: Option<SheetSpec>) -> Result<Option<Ref>, String> {
        let start = self.pos;
        let Some(a) = self.parse_cell_ref()? else {
            self.pos = start;
            return Ok(None);
        };
        self.skip_ws();
        if self.peek_char() == Some(':') {
            self.next_char();
            let b = self
                .parse_cell_ref()?
                .ok_or_else(|| "expected cell reference after ':'".to_string())?;
            return Ok(Some(Ref { sheet, kind: RefKind::Area(a, b) }));
        }
        Ok(Some(Ref { sheet, kind: RefKind::Cell(a) }))
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
        let s = &self.input[start..self.pos];
        let n: f64 = s.parse().map_err(|_| format!("invalid number literal: {s}"))?;
        Ok(Expr::Number(n))
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

        let col = col_label_to_index(col_label).ok_or_else(|| "invalid column label".to_string())?;
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

        Ok(Some(CellRef { row, col, abs_row, abs_col }))
    }

    fn parse_sheet_name(&mut self) -> Result<Option<String>, String> {
        self.skip_ws();
        match self.peek_char() {
            Some('\'') => self.parse_quoted_sheet_name().map(Some),
            Some(ch) if is_ident_start(ch) => Ok(self.parse_identifier()?),
            _ => Ok(None),
        }
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

    fn skip_ws(&mut self) {
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.next_char();
            } else {
                break;
            }
        }
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

fn col_label_to_index(label: &str) -> Option<u32> {
    let mut col: u32 = 0;
    for ch in label.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let upper = ch.to_ascii_uppercase() as u32;
        col = col * 26 + (upper - 'A' as u32 + 1);
    }
    if col == 0 {
        return None;
    }
    Some(col - 1)
}
