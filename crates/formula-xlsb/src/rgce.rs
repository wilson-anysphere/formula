//! BIFF12 `rgce` (formula token stream) codec.
//!
//! The decoder is best-effort and is primarily used for diagnostics and for surfacing formula
//! text when reading XLSB files.
//!
//! The encoder is intentionally small-scope: it supports enough of Excel's formula language to
//! round-trip common patterns while we build out full compatibility.

use crate::format::push_column_label;
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

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`).
pub fn decode_rgce(rgce: &[u8]) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, None)
}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`),
/// using workbook context to resolve sheet indices (`ixti`) and defined names.
pub fn decode_rgce_with_context(rgce: &[u8], ctx: &WorkbookContext) -> Result<String, DecodeError> {
    decode_rgce_impl(rgce, Some(ctx))
}

fn decode_rgce_impl(rgce: &[u8], ctx: Option<&WorkbookContext>) -> Result<String, DecodeError> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    // Prevent pathological expansion (e.g. from future token support).
    const MAX_OUTPUT_FACTOR: usize = 10;
    let max_len = rgce.len().saturating_mul(MAX_OUTPUT_FACTOR);

    let mut i = 0usize;
    let mut last_ptg_offset = 0usize;
    let mut last_ptg = rgce[0];

    let mut stack: Vec<String> = Vec::new();

    while i < rgce.len() {
        let ptg_offset = i;
        let ptg = rgce[i];
        i += 1;

        last_ptg_offset = ptg_offset;
        last_ptg = ptg;

        match ptg {
            0x03..=0x11 => {
                if stack.len() < 2 {
                    return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                }
                let b = stack.pop().expect("len checked");
                let a = stack.pop().expect("len checked");
                let op = match ptg {
                    0x03 => "+",
                    0x04 => "-",
                    0x05 => "*",
                    0x06 => "/",
                    0x07 => "^",
                    0x08 => "&",
                    0x09 => "<",
                    0x0A => "<=",
                    0x0B => "=",
                    0x0C => ">",
                    0x0D => ">=",
                    0x0E => "<>",
                    0x0F => " ",
                    0x10 => ",",
                    0x11 => ":",
                    _ => unreachable!("ptg matched by range"),
                };
                stack.push(format!("{a}{op}{b}"));
            }
            0x12 => {
                let a = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                stack.push(format!("+{a}"));
            }
            0x13 => {
                let a = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                stack.push(format!("-{a}"));
            }
            0x14 => {
                let a = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                stack.push(format!("{a}%"));
            }
            0x15 => {
                let a = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                stack.push(format!("({a})"));
            }
            0x16 => {
                // PtgMissArg
                stack.push(String::new());
            }
            0x17 => {
                // PtgStr: [cch: u16][utf16 chars...]
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

                stack.push(format!("\"{}\"", String::from_utf16_lossy(&units)));
            }
            0x19 => {
                // PtgAttr: [grbit: u8][unused: u16]
                //
                // Commonly used for spacing/optimization and does not affect expression semantics
                // for our best-effort display purposes.
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 3;
            }
            0x1C => {
                // PtgErr: [err: u8]
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let err = rgce[i];
                i += 1;

                let text = match err {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => {
                        return Err(DecodeError::InvalidConstant {
                            offset: ptg_offset,
                            ptg,
                            value: err,
                        });
                    }
                };
                stack.push(text.to_string());
            }
            0x1D => {
                // PtgBool: [b: u8]
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
                stack.push(if b == 0 { "FALSE" } else { "TRUE" }.to_string());
            }
            0x1E => {
                // PtgInt: [n: u16]
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
                stack.push(n.to_string());
            }
            0x1F => {
                // PtgNum: [f64]
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
                stack.push(f64::from_le_bytes(bytes).to_string());
            }
            0x24 | 0x44 | 0x64 => {
                // PtgRef: [row: u32][col: u16 (with relative flags in high bits)]
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

                stack.push(format_cell_ref(row, col as u32, flags));
            }
            0x25 | 0x45 | 0x65 => {
                // PtgArea: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let row_first0 = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
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

                let mut out = String::new();
                if is_value_class && !is_single_cell {
                    // Legacy implicit intersection: Excel encodes this by using a value-class
                    // range token; modern formula text uses an explicit `@` operator.
                    out.push('@');
                }
                if is_single_cell {
                    out.push_str(&a);
                } else {
                    out.push_str(&a);
                    out.push(':');
                    out.push_str(&b);
                }
                stack.push(out);
            }
            0x2C | 0x4C | 0x6C => {
                // PtgRefN: [row: u32][col: u16]
                //
                // In BIFF, *N tokens are typically used for relative references in contexts like
                // defined names and shared formulas. In BIFF12, Excel still stores row/col fields
                // plus relative flags; for now we decode them the same way as `PtgRef`.
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let row0 = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                let col_field = u16::from_le_bytes([rgce[i + 4], rgce[i + 5]]);
                i += 6;
                stack.push(format_cell_ref_from_field(row0, col_field));
            }
            0x2D | 0x4D | 0x6D => {
                // PtgAreaN: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                //
                // Same caveat as `PtgRefN` above; decode as absolute row/col with relative flags.
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

                let mut out = String::new();
                if is_value_class && !is_single_cell {
                    out.push('@');
                }
                if is_single_cell {
                    out.push_str(&a);
                } else {
                    out.push_str(&a);
                    out.push(':');
                    out.push_str(&b);
                }
                stack.push(out);
            }
            0x3A | 0x5A | 0x7A => {
                // PtgRef3d: [ixti: u16][row: u32][col: u16]
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
                let prefix = if first == last {
                    format!("{first}!")
                } else {
                    format!("{first}:{last}!")
                };
                let cell = format_cell_ref_from_field(row0, col_field);
                stack.push(format!("{prefix}{cell}"));
            }
            0x3B | 0x5B | 0x7B => {
                // PtgArea3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
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
                let row_first0 = u32::from_le_bytes([rgce[i + 2], rgce[i + 3], rgce[i + 4], rgce[i + 5]]);
                let row_last0 = u32::from_le_bytes([rgce[i + 6], rgce[i + 7], rgce[i + 8], rgce[i + 9]]);
                let col_first = u16::from_le_bytes([rgce[i + 10], rgce[i + 11]]);
                let col_last = u16::from_le_bytes([rgce[i + 12], rgce[i + 13]]);
                i += 14;

                let (first, last) = ctx
                    .extern_sheet_names(ixti)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;
                let prefix = if first == last {
                    format!("{first}!")
                } else {
                    format!("{first}:{last}!")
                };

                let a = format_cell_ref_from_field(row_first0, col_first);
                let b = format_cell_ref_from_field(row_last0, col_last);
                let is_single_cell = row_first0 == row_last0
                    && (col_first & COL_INDEX_MASK) == (col_last & COL_INDEX_MASK);
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut out = String::new();
                if is_value_class && !is_single_cell {
                    out.push('@');
                }
                out.push_str(&prefix);
                if is_single_cell {
                    out.push_str(&a);
                } else {
                    out.push_str(&a);
                    out.push(':');
                    out.push_str(&b);
                }
                stack.push(out);
            }
            0x23 | 0x43 | 0x63 => {
                // PtgName: [nameId: u32][reserved: u16]
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

                let name_id = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                i += 4;
                // reserved
                i += 2;

                let def = ctx
                    .name_definition(name_id)
                    .ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;

                let txt = match &def.scope {
                    NameScope::Workbook => def.name.clone(),
                    NameScope::Sheet(sheet) => format!("{sheet}!{}", def.name),
                };
                stack.push(txt);
            }
            0x39 | 0x59 | 0x79 => {
                // PtgNameX: [ixti: u16][nameIndex: u16]
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
                stack.push(txt);
            }
            0x21 | 0x41 | 0x61 => {
                // PtgFunc: [iftab: u16] (argument count is implicit and fixed for the function).
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

                stack.push(format!("{}({})", spec.name, args.join(",")));
            }
            0x22 | 0x42 | 0x62 => {
                // PtgFuncVar: [argc: u8][iftab: u16]
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

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name (typically from `PtgNameX`), followed by args.
                if iftab == 0x00FF {
                    if argc == 0 {
                        return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                    }
                    if stack.len() < argc {
                        return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                    }

                    let func_name = stack.pop().expect("len checked");
                    let mut args = Vec::with_capacity(argc - 1);
                    for _ in 0..argc.saturating_sub(1) {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();
                    stack.push(format!("{func_name}({})", args.join(",")));
                } else {
                    let name =
                        function_name(iftab).ok_or(DecodeError::UnknownPtg { offset: ptg_offset, ptg })?;

                    if stack.len() < argc {
                        return Err(DecodeError::StackUnderflow { offset: ptg_offset, ptg });
                    }

                    let mut args = Vec::with_capacity(argc);
                    for _ in 0..argc {
                        args.push(stack.pop().expect("len checked"));
                    }
                    args.reverse();

                    stack.push(format!("{name}({})", args.join(",")));
                }
            }
            _ => return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg }),
        }

        if stack
            .last()
            .is_some_and(|s| s.len() > max_len)
        {
            return Err(DecodeError::OutputTooLarge {
                offset: ptg_offset,
                ptg,
                max_len,
            });
        }
    }

    if stack.len() == 1 {
        Ok(stack.pop().expect("len checked"))
    } else {
        Err(DecodeError::StackNotSingular {
            offset: last_ptg_offset,
            ptg: last_ptg,
            stack_len: stack.len(),
        })
    }
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

    let mut out = Vec::new();
    emit_expr(&expr, ctx, &mut out)?;
    Ok(EncodedRgce {
        rgce: out,
        rgcb: Vec::new(),
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

const PTG_FUNCVAR: u8 = 0x22;
const PTG_NAME: u8 = 0x23;
const PTG_REF: u8 = 0x24;
const PTG_AREA: u8 = 0x25;
const PTG_REF3D: u8 = 0x3A;
const PTG_AREA3D: u8 = 0x3B;
const PTG_NAMEX: u8 = 0x39;

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
    Func { name: String, args: Vec<Expr> },
    Unary { op: UnaryOp, expr: Box<Expr> },
    Binary { op: BinaryOp, left: Box<Expr>, right: Box<Expr> },
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

fn emit_expr(expr: &Expr, ctx: &WorkbookContext, out: &mut Vec<u8>) -> Result<(), EncodeError> {
    match expr {
        Expr::Number(n) => emit_number(*n, out),
        Expr::Ref(r) => emit_ref(r, ctx, out, PtgClass::Ref)?,
        Expr::Name(n) => emit_name(n, ctx, out, PtgClass::Ref)?,
        Expr::Func { name, args } => {
            for arg in args {
                emit_expr(arg, ctx, out)?;
            }
            emit_func(name, args.len(), ctx, out)?;
        }
        Expr::Unary { op, expr } => {
            match op {
                UnaryOp::ImplicitIntersection => match expr.as_ref() {
                    // Encode `@` by emitting value-class reference tokens. This matches Excel's
                    // legacy implicit-intersection encoding, and round-trips through
                    // `decode_rgce*` as an explicit `@`.
                    Expr::Ref(r) => emit_ref(r, ctx, out, PtgClass::Value)?,
                    Expr::Name(n) => emit_name(n, ctx, out, PtgClass::Value)?,
                    _ => {
                        return Err(EncodeError::Parse(
                            "implicit intersection (@) is only supported on references".to_string(),
                        ))
                    }
                },
                UnaryOp::Plus => {
                    emit_expr(expr, ctx, out)?;
                    out.push(PTG_UPLUS);
                }
                UnaryOp::Minus => {
                    emit_expr(expr, ctx, out)?;
                    out.push(PTG_UMINUS);
                }
            }
        }
        Expr::Binary { op, left, right } => {
            emit_expr(left, ctx, out)?;
            emit_expr(right, ctx, out)?;
            out.push(match op {
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

fn emit_func(
    name: &str,
    argc: usize,
    ctx: &WorkbookContext,
    out: &mut Vec<u8>,
) -> Result<(), EncodeError> {
    let upper = name.trim().to_ascii_uppercase();

    // Built-in functions (currently very small subset).
    if let Some(iftab) = match upper.as_str() {
        "SUM" => Some(0x0004u16),
        _ => None,
    } {
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
        match self.peek_char() {
            Some('(') => {
                self.next_char();
                let expr = self.parse_add_sub()?;
                self.skip_ws();
                if self.next_char() != Some(')') {
                    return Err("expected ')'".to_string());
                }
                Ok(expr)
            }
            Some(ch) if ch.is_ascii_digit() || ch == '.' => self.parse_number(),
            Some('\'') => self.parse_ident_or_ref(),
            Some(ch) if is_ident_start(ch) => self.parse_ident_or_ref(),
            _ => Err("unexpected token".to_string()),
        }
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
