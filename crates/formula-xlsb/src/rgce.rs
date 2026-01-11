//! Minimal `rgce` (BIFF12 formula token stream) decoder.
//!
//! This is intentionally incomplete: it exists primarily to aid diagnostics and to
//! accelerate support work by making failures actionable (ptg id + byte offset).

use crate::format::push_column_label;

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
    /// After decoding, the expression stack didn't contain exactly one item.
    StackNotSingular { offset: usize, stack_len: usize },
}

impl DecodeError {
    pub fn offset(&self) -> usize {
        match *self {
            DecodeError::UnknownPtg { offset, .. } => offset,
            DecodeError::UnexpectedEof { offset, .. } => offset,
            DecodeError::StackUnderflow { offset, .. } => offset,
            DecodeError::InvalidConstant { offset, .. } => offset,
            DecodeError::StackNotSingular { offset, .. } => offset,
        }
    }

    pub fn ptg(&self) -> Option<u8> {
        match *self {
            DecodeError::UnknownPtg { ptg, .. } => Some(ptg),
            DecodeError::UnexpectedEof { ptg, .. } => Some(ptg),
            DecodeError::StackUnderflow { ptg, .. } => Some(ptg),
            DecodeError::InvalidConstant { ptg, .. } => Some(ptg),
            DecodeError::StackNotSingular { .. } => None,
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
            DecodeError::StackNotSingular { offset, stack_len } => write!(
                f,
                "formula decoded with stack_len={stack_len} at rgce offset {offset} (expected 1)"
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Decode an `rgce` token stream into best-effort Excel formula text (without leading `=`).
pub fn decode_rgce(rgce: &[u8]) -> Result<String, DecodeError> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    let mut i = 0usize;
    let mut stack: Vec<usize> = Vec::new();
    let mut formula = String::with_capacity(rgce.len());

    while i < rgce.len() {
        let ptg_offset = i;
        let ptg = rgce[i];
        i += 1;

        match ptg {
            0x03..=0x11 => {
                let e2_start = stack
                    .pop()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                let e2 = formula.split_off(e2_start);
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
                    _ => unreachable!(),
                };
                formula.push_str(op);
                formula.push_str(&e2);
            }
            0x12 => {
                let &e = stack
                    .last()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                formula.insert(e, '+');
            }
            0x13 => {
                let &e = stack
                    .last()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                formula.insert(e, '-');
            }
            0x14 => {
                formula.push('%');
            }
            0x15 => {
                let &e = stack
                    .last()
                    .ok_or(DecodeError::StackUnderflow { offset: ptg_offset, ptg })?;
                formula.insert(e, '(');
                formula.push(')');
            }
            0x16 => {
                // PtgMissArg
                stack.push(formula.len());
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

                stack.push(formula.len());
                formula.push('"');
                formula.push_str(&String::from_utf16_lossy(&units));
                formula.push('"');
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

                stack.push(formula.len());
                formula.push_str(text);
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

                stack.push(formula.len());
                formula.push_str(if b == 0 { "FALSE" } else { "TRUE" });
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

                stack.push(formula.len());
                formula.push_str(&n.to_string());
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

                stack.push(formula.len());
                formula.push_str(&f64::from_le_bytes(bytes).to_string());
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

                let row = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) + 1;
                let flags = rgce[i + 5];
                let col = u16::from_le_bytes([rgce[i + 4], flags & 0x3F]);
                i += 6;

                stack.push(formula.len());
                if flags & 0x80 != 0x80 {
                    formula.push('$');
                }
                push_column_label(col as u32, &mut formula);
                if flags & 0x40 != 0x40 {
                    formula.push('$');
                }
                formula.push_str(&row.to_string());
            }
            _ => return Err(DecodeError::UnknownPtg { offset: ptg_offset, ptg }),
        }
    }

    if stack.len() == 1 {
        Ok(formula)
    } else {
        Err(DecodeError::StackNotSingular {
            offset: rgce.len(),
            stack_len: stack.len(),
        })
    }
}

