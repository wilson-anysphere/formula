use crate::function_ids::{function_id_to_name, function_spec_from_id};
use crate::structured_refs::{
    format_structured_ref, structured_ref_is_single_cell, StructuredColumns, StructuredRefItem,
};

/// Structured `rgce` decode failure with ptg id + offset.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum DecodeRgceError {
    /// Encountered a ptg we do not (yet) handle.
    #[error("unsupported ptg=0x{ptg:02X} at rgce offset {offset}")]
    UnsupportedToken { offset: usize, ptg: u8 },
    /// Not enough bytes remained in the rgce stream to decode the current token.
    #[error("unexpected eof decoding ptg=0x{ptg:02X} at rgce offset {offset} (needed {needed} bytes, remaining {remaining})")]
    UnexpectedEof {
        offset: usize,
        ptg: u8,
        needed: usize,
        remaining: usize,
    },
    /// The ptg required more stack items than were available.
    #[error("stack underflow decoding ptg=0x{ptg:02X} at rgce offset {offset}")]
    StackUnderflow { offset: usize, ptg: u8 },
    /// A ptg referenced a function id we don't know how to display.
    #[error(
        "unknown function id=0x{func_id:04X} decoding ptg=0x{ptg:02X} at rgce offset {offset}"
    )]
    UnknownFunctionId {
        offset: usize,
        ptg: u8,
        func_id: u16,
    },
    /// Failed to decode a UTF-16 string literal payload.
    #[error("invalid utf16 string payload decoding ptg=0x{ptg:02X} at rgce offset {offset}")]
    InvalidUtf16 { offset: usize, ptg: u8 },
    /// Decoding exceeded the maximum output size derived from the input length.
    #[error("formula decode exceeded max_len={max_len} decoding ptg=0x{ptg:02X} at rgce offset {offset}")]
    OutputTooLarge {
        offset: usize,
        ptg: u8,
        max_len: usize,
    },
    /// After decoding, the expression stack didn't contain exactly one item.
    #[error("formula decoded with stack_len={stack_len} at rgce offset {offset} (ptg=0x{ptg:02X}, expected 1)")]
    StackNotSingular {
        offset: usize,
        ptg: u8,
        stack_len: usize,
    },
}

impl DecodeRgceError {
    pub fn offset(&self) -> usize {
        match *self {
            DecodeRgceError::UnsupportedToken { offset, .. } => offset,
            DecodeRgceError::UnexpectedEof { offset, .. } => offset,
            DecodeRgceError::StackUnderflow { offset, .. } => offset,
            DecodeRgceError::UnknownFunctionId { offset, .. } => offset,
            DecodeRgceError::InvalidUtf16 { offset, .. } => offset,
            DecodeRgceError::OutputTooLarge { offset, .. } => offset,
            DecodeRgceError::StackNotSingular { offset, .. } => offset,
        }
    }

    pub fn ptg(&self) -> Option<u8> {
        match *self {
            DecodeRgceError::UnsupportedToken { ptg, .. } => Some(ptg),
            DecodeRgceError::UnexpectedEof { ptg, .. } => Some(ptg),
            DecodeRgceError::StackUnderflow { ptg, .. } => Some(ptg),
            DecodeRgceError::UnknownFunctionId { ptg, .. } => Some(ptg),
            DecodeRgceError::InvalidUtf16 { ptg, .. } => Some(ptg),
            DecodeRgceError::OutputTooLarge { ptg, .. } => Some(ptg),
            DecodeRgceError::StackNotSingular { ptg, .. } => Some(ptg),
        }
    }
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

/// Decode a BIFF12 array constant from the `rgcb` payload stream.
///
/// BIFF12 stores array constant values in a separate trailing `rgcb` data stream, referenced by
/// `PtgArray` tokens in the main `rgce` token stream.
///
/// Layout (MS-XLSB 2.5.198.8 PtgArray):
/// `[cols_minus1:u16][rows_minus1:u16][values...]`
///
/// Values are stored row-major and each starts with a type byte:
/// - `0x00` = empty
/// - `0x01` = number (`f64`)
/// - `0x02` = string (`[cch:u16][utf16 chars...]`)
/// - `0x04` = bool (`[b:u8]`)
/// - `0x10` = error (`[code:u8]`)
fn decode_array_constant(
    rgcb: &[u8],
    pos: &mut usize,
    ptg_offset: usize,
    ptg: u8,
) -> Result<String, DecodeRgceError> {
    let mut i = *pos;
    if rgcb.len().saturating_sub(i) < 4 {
        return Err(DecodeRgceError::UnexpectedEof {
            offset: ptg_offset,
            ptg,
            needed: 4,
            remaining: rgcb.len().saturating_sub(i),
        });
    }

    let cols_minus1 = u16::from_le_bytes([rgcb[i], rgcb[i + 1]]) as usize;
    let rows_minus1 = u16::from_le_bytes([rgcb[i + 2], rgcb[i + 3]]) as usize;
    i += 4;

    let cols = cols_minus1.saturating_add(1);
    let rows = rows_minus1.saturating_add(1);

    let mut row_texts = Vec::with_capacity(rows);
    for _ in 0..rows {
        let mut col_texts = Vec::with_capacity(cols);
        for _ in 0..cols {
            if rgcb.len().saturating_sub(i) < 1 {
                return Err(DecodeRgceError::UnexpectedEof {
                    offset: ptg_offset,
                    ptg,
                    needed: 1,
                    remaining: rgcb.len().saturating_sub(i),
                });
            }
            let ty = rgcb[i];
            i += 1;
            match ty {
                0x00 => col_texts.push(String::new()),
                0x01 => {
                    if rgcb.len().saturating_sub(i) < 8 {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: 8,
                            remaining: rgcb.len().saturating_sub(i),
                        });
                    }
                    let mut bytes = [0u8; 8];
                    bytes.copy_from_slice(&rgcb[i..i + 8]);
                    i += 8;
                    col_texts.push(f64::from_le_bytes(bytes).to_string());
                }
                0x02 => {
                    if rgcb.len().saturating_sub(i) < 2 {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: 2,
                            remaining: rgcb.len().saturating_sub(i),
                        });
                    }
                    let cch = u16::from_le_bytes([rgcb[i], rgcb[i + 1]]) as usize;
                    i += 2;
                    let byte_len = cch.saturating_mul(2);
                    if rgcb.len().saturating_sub(i) < byte_len {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: byte_len,
                            remaining: rgcb.len().saturating_sub(i),
                        });
                    }
                    let raw = &rgcb[i..i + byte_len];
                    i += byte_len;

                    let mut units = Vec::with_capacity(cch);
                    for chunk in raw.chunks_exact(2) {
                        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                    }
                    let s =
                        String::from_utf16(&units).map_err(|_| DecodeRgceError::InvalidUtf16 {
                            offset: ptg_offset,
                            ptg,
                        })?;
                    let escaped = s.replace('"', "\"\"");
                    col_texts.push(format!("\"{escaped}\""));
                }
                0x04 => {
                    if rgcb.len().saturating_sub(i) < 1 {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: 1,
                            remaining: rgcb.len().saturating_sub(i),
                        });
                    }
                    let b = rgcb[i];
                    i += 1;
                    col_texts.push(if b == 0 { "FALSE" } else { "TRUE" }.to_string());
                }
                0x10 => {
                    if rgcb.len().saturating_sub(i) < 1 {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed: 1,
                            remaining: rgcb.len().saturating_sub(i),
                        });
                    }
                    let code = rgcb[i];
                    i += 1;
                    let text = match code {
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
                        _ => "#UNKNOWN!",
                    };
                    col_texts.push(text.to_string());
                }
                _ => {
                    // Unknown array constant element type.
                    return Err(DecodeRgceError::UnsupportedToken {
                        offset: ptg_offset,
                        ptg,
                    });
                }
            }
        }
        row_texts.push(col_texts.join(","));
    }

    *pos = i;
    Ok(format!("{{{}}}", row_texts.join(";")))
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
) -> Result<(), DecodeRgceError> {
    fn has_remaining(buf: &[u8], i: usize, needed: usize) -> bool {
        buf.len().saturating_sub(i) >= needed
    }

    let mut i = 0usize;
    while i < rgce.len() {
        let ptg_offset = rgce_base_offset.saturating_add(i);
        let ptg = rgce[i];
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 7,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 7;
                let _ = decode_array_constant(rgcb, rgcb_pos, ptg_offset, ptg)?;
            }

            // Binary operators and simple operators with no payload.
            0x03..=0x16 | 0x2F => {}

            // PtgStr: [cch: u16][utf16 chars...]
            0x17 => {
                if !has_remaining(rgce, i, 2) {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cch = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;
                let byte_len = cch.saturating_mul(2);
                if !has_remaining(rgce, i, byte_len) {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let etpg = rgce[i];
                i += 1;
                match etpg {
                    // etpg=0x19 is the structured reference payload (PtgList): fixed 12 bytes.
                    0x19 => {
                        if !has_remaining(rgce, i, 12) {
                            return Err(DecodeRgceError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: 12,
                                remaining: rgce.len().saturating_sub(i),
                            });
                        }
                        i += 12;
                    }
                    // Unknown extend subtype: stop scanning to avoid desync/false positives.
                    _ => break,
                }
            }

            // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
            0x19 => {
                if !has_remaining(rgce, i, 3) {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let grbit = rgce[i];
                let w_attr = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]) as usize;
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = w_attr.saturating_mul(2);
                    if !has_remaining(rgce, i, needed) {
                        return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 3;
            }

            // PtgName: [nameId: u32][reserved: u16]
            0x23 | 0x43 | 0x63 => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cce = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;
                if !has_remaining(rgce, i, cce) {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: cce,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                consume_rgcb_arrays_in_subexpression(
                    &rgce[i..i + cce],
                    rgcb,
                    rgcb_pos,
                    rgce_base_offset.saturating_add(i),
                )?;
                i += cce;
            }

            // PtgRefErr: [row: u32][col: u16]
            0x2A | 0x4A | 0x6A => {
                if !has_remaining(rgce, i, 6) {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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

/// Best-effort decode of a BIFF12 `rgce` token stream into formula text.
///
/// The returned string does **not** include a leading `=`.
pub fn decode_rgce(rgce: &[u8]) -> Result<String, DecodeRgceError> {
    decode_rgce_impl(rgce, None, None)
}

/// Best-effort decode of a BIFF12 `rgce` token stream into formula text, using a trailing `rgcb`
/// payload stream to decode array constants (`PtgArray`).
///
/// The returned string does **not** include a leading `=`.
pub fn decode_rgce_with_rgcb(rgce: &[u8], rgcb: &[u8]) -> Result<String, DecodeRgceError> {
    decode_rgce_impl(rgce, Some(rgcb), None)
}

/// Best-effort decode of a BIFF12 `rgce` token stream into formula text, using a base cell for
/// relative-reference tokens.
///
/// Excel encodes certain formulas (notably in shared formulas) using relative-reference tokens
/// like `PtgRefN` / `PtgAreaN` that store offsets from the formula's origin cell. This helper
/// converts those relative offsets into A1-style references using the provided base coordinates.
///
/// `base_row0` and `base_col0` are **0-indexed** cell coordinates (`A1` is `(0, 0)`).
///
/// The returned string does **not** include a leading `=`.
pub fn decode_rgce_with_base(
    rgce: &[u8],
    base_row0: u32,
    base_col0: u32,
) -> Result<String, DecodeRgceError> {
    decode_rgce_impl(rgce, None, Some((base_row0, base_col0)))
}

fn decode_rgce_impl(
    rgce: &[u8],
    rgcb: Option<&[u8]>,
    base: Option<(u32, u32)>,
) -> Result<String, DecodeRgceError> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    // Prevent pathological output expansion (e.g. from malformed tokens, or from future token
    // support that expands to workbook-context-backed names).
    const MAX_OUTPUT_FACTOR: usize = 10;
    // Clamp the decoded output to a conservative upper bound to avoid allocating huge strings.
    const MAX_OUTPUT_LEN: usize = 1_000_000;
    // Some ptgs (notably `PtgArray`) reference additional data stored in the trailing `rgcb`
    // buffer. Include it when deriving an upper bound so we don't reject legitimate array
    // constants whose `rgce` stream is tiny but `rgcb` is not.
    let max_len = rgce
        .len()
        .saturating_add(rgcb.map_or(0, |b| b.len()))
        .saturating_mul(MAX_OUTPUT_FACTOR)
        .min(MAX_OUTPUT_LEN);

    let mut i = 0usize;
    let mut rgcb_pos = 0usize;
    let mut stack: Vec<ExprFragment> = Vec::new();
    let mut last_ptg_offset = 0usize;
    let mut last_ptg = rgce[0];

    while i < rgce.len() {
        let ptg_offset = i;
        let ptg = rgce[i];
        last_ptg_offset = ptg_offset;
        last_ptg = ptg;
        i += 1;

        match ptg {
            // Binary operators.
            0x03..=0x11 => {
                let Some(op) = op_str(ptg) else {
                    return Err(DecodeRgceError::UnsupportedToken {
                        offset: ptg_offset,
                        ptg,
                    });
                };
                let prec = binary_precedence(ptg).ok_or(DecodeRgceError::UnsupportedToken {
                    offset: ptg_offset,
                    ptg,
                })?;

                let right = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
                let left = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;

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
                let expr = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
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
                let expr = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
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
                let expr = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
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
                let expr = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                    offset: ptg_offset,
                    ptg,
                })?;
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
            // String literal.
            0x17 => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cch = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;
                let byte_len = cch.saturating_mul(2);
                if rgce.len().saturating_sub(i) < byte_len {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: byte_len,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let raw = &rgce[i..i + byte_len];
                i += byte_len;

                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                // BIFF strings are UTF-16LE, but real-world files can contain malformed UTF-16.
                // Stay best-effort (matching `formula-xlsb`) by decoding lossily instead of
                // aborting the entire formula decode.
                let s = String::from_utf16_lossy(&units);
                let escaped = s.replace('"', "\"\"");
                stack.push(ExprFragment::new(format!("\"{escaped}\"")));
            }
            // PtgExtend / PtgExtendV / PtgExtendA.
            //
            // MS-XLSB encodes newer operand tokens (including structured references / table refs)
            // as a `PtgExtend*` token followed by an `etpg` subtype byte.
            //
            // We support the structured reference subtype (`etpg=0x19`, "PtgList") so formulas
            // extracted from real XLSB files can be decoded.
            0x18 | 0x38 | 0x58 => {
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 1,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let etpg = rgce[i];
                i += 1;

                match etpg {
                    // MS-XLSB 2.5.198.51 PtgList (structured reference / table ref).
                    //
                    // Canonical 12-byte payload:
                    //   [table_id: u32]
                    //   [flags: u16]
                    //   [col_first: u16]
                    //   [col_last: u16]
                    //   [reserved: u16]
                    0x19 => {
                        if rgce.len().saturating_sub(i) < 12 {
                            return Err(DecodeRgceError::UnexpectedEof {
                                offset: ptg_offset,
                                ptg,
                                needed: 12,
                                remaining: rgce.len().saturating_sub(i),
                            });
                        }

                        // Excel uses a fixed 12-byte payload. The canonical layout is documented
                        // in MS-XLSB, but in practice there are multiple observed encodings in the
                        // wild (different field packing/ordering). Decode in a best-effort way by
                        // trying a handful of plausible interpretations and choosing the most
                        // likely one based on simple heuristics.
                        let mut payload = [0u8; 12];
                        payload.copy_from_slice(&rgce[i..i + 12]);
                        i += 12;

                        let decoded = decode_ptg_list_payload_best_effort(&payload);

                        let table_id = decoded.table_id;
                        let flags16 = (decoded.flags & 0xFFFF) as u16;
                        let col_first = decoded.col_first;
                        let col_last = decoded.col_last;

                        // Best-effort: map table/column IDs to placeholder names (we don't have
                        // workbook context in this crate).
                        let table_name = format!("Table{table_id}");
                        let columns = structured_columns_from_ids(col_first, col_last);

                        let item = structured_ref_item_from_flags(flags16);
                        let display_table_name = match item {
                            Some(StructuredRefItem::ThisRow) => None,
                            _ => Some(table_name.as_str()),
                        };

                        let mut text = format_structured_ref(display_table_name, item, &columns);

                        let mut precedence = 100;
                        let is_value_class = ptg == 0x38;
                        if is_value_class && !structured_ref_is_single_cell(item, &columns) {
                            // Value-class list tokens represent legacy implicit intersection,
                            // mirroring PtgAreaV behavior.
                            text = format!("@{text}");
                            precedence = 70;
                        }

                        stack.push(ExprFragment {
                            text,
                            precedence,
                            contains_union: false,
                            is_missing: false,
                        });
                    }
                    _ => {
                        return Err(DecodeRgceError::UnsupportedToken {
                            offset: ptg_offset,
                            ptg,
                        })
                    }
                }
            }
            // PtgAttr: [grbit: u8][wAttr: u16] + optional payloads.
            //
            // Most attributes are evaluation hints or formatting metadata. We treat them as
            // non-printing tokens, but must consume their payload so later ptgs stay aligned.
            //
            // Excel also uses `tAttrSum` as an optimization where `SUM(A1:A10)` is encoded as:
            //   PtgArea(A1:A10) + PtgAttr(tAttrSum)
            // with no explicit `PtgFuncVar(SUM)` token.
            0x19 => {
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let grbit = rgce[i];
                let w_attr = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]);
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                const T_ATTR_SUM: u8 = 0x10;

                if grbit & T_ATTR_SUM != 0 {
                    let arg = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?;
                    let mut text = String::new();
                    text.push_str("SUM(");
                    if !arg.is_missing {
                        if arg.contains_union {
                            text.push('(');
                            text.push_str(&arg.text);
                            text.push(')');
                        } else {
                            text.push_str(&arg.text);
                        }
                    }
                    text.push(')');
                    stack.push(ExprFragment::new(text));
                }

                if grbit & T_ATTR_CHOOSE != 0 {
                    // `tAttrChoose` is followed by a jump table of `u16` offsets (wAttr entries).
                    let needed = (w_attr as usize).saturating_mul(2);
                    if rgce.len().saturating_sub(i) < needed {
                        return Err(DecodeRgceError::UnexpectedEof {
                            offset: ptg_offset,
                            ptg,
                            needed,
                            remaining: rgce.len().saturating_sub(i),
                        });
                    }
                    i += needed;
                }
            }
            // Error literal.
            0x1C => {
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    0x2C => "#SPILL!",
                    0x2D => "#CALC!",
                    0x2E => "#FIELD!",
                    0x2F => "#CONNECT!",
                    0x30 => "#BLOCKED!",
                    0x31 => "#UNKNOWN!",
                    // Best-effort forward-compatibility with newer Excel error codes.
                    _ => "#UNKNOWN!",
                };
                stack.push(ExprFragment::new(text.to_string()));
            }
            // Bool literal.
            0x1D => {
                if rgce.len().saturating_sub(i) < 1 {
                    return Err(DecodeRgceError::UnexpectedEof {
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
            // Int literal.
            0x1E => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeRgceError::UnexpectedEof {
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
            // Num literal.
            0x1F => {
                if rgce.len().saturating_sub(i) < 8 {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                let Some(rgcb) = rgcb else {
                    // Keep `decode_rgce` behavior unchanged: without rgcb, PtgArray is unsupported.
                    return Err(DecodeRgceError::UnsupportedToken {
                        offset: ptg_offset,
                        ptg,
                    });
                };
                if rgce.len().saturating_sub(i) < 7 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 7,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                i += 7;

                let arr = decode_array_constant(rgcb, &mut rgcb_pos, ptg_offset, ptg)?;
                stack.push(ExprFragment::new(arr));
            }
            // PtgFunc
            0x21 | 0x41 | 0x61 => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let func_id = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                i += 2;

                let Some(spec) = function_spec_from_id(func_id) else {
                    return Err(DecodeRgceError::UnknownFunctionId {
                        offset: ptg_offset,
                        ptg,
                        func_id,
                    });
                };
                if spec.min_args != spec.max_args {
                    return Err(DecodeRgceError::UnknownFunctionId {
                        offset: ptg_offset,
                        ptg,
                        func_id,
                    });
                }

                let argc = spec.min_args as usize;
                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?);
                }
                args.reverse();

                let mut text = String::new();
                text.push_str(spec.name);
                text.push('(');
                for (i, arg) in args.into_iter().enumerate() {
                    if i > 0 {
                        text.push(',');
                    }
                    if arg.is_missing {
                        continue;
                    }
                    if arg.contains_union {
                        text.push('(');
                        text.push_str(&arg.text);
                        text.push(')');
                    } else {
                        text.push_str(&arg.text);
                    }
                }
                text.push(')');

                stack.push(ExprFragment::new(text));
            }
            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                if rgce.len().saturating_sub(i) < 3 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 3,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let argc = rgce[i] as usize;
                let func_id = u16::from_le_bytes([rgce[i + 1], rgce[i + 2]]);
                i += 3;

                // Excel uses a sentinel function id for user-defined functions: the top-of-stack
                // item is the function name expression (typically from `PtgNameX`), followed by
                // args.
                if func_id == 0x00FF {
                    if argc == 0 {
                        return Err(DecodeRgceError::StackUnderflow {
                            offset: ptg_offset,
                            ptg,
                        });
                    }

                    let func_name = stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?;
                    // Use the decoded name token text as the function name. When we don't have
                    // workbook context for `PtgNameX`, we emit a stable placeholder identifier
                    // (`ExternName_IXTI<ixti>_N<idx>`) that remains parseable by Excel formula
                    // parsers (avoid `:` / `{}`).
                    let func_name_text = func_name.text;
                    let mut args = Vec::with_capacity(argc.saturating_sub(1));
                    for _ in 0..argc.saturating_sub(1) {
                        args.push(stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                            offset: ptg_offset,
                            ptg,
                        })?);
                    }
                    args.reverse();

                    let mut text = String::new();
                    text.push_str(&func_name_text);
                    text.push('(');
                    for (i, arg) in args.into_iter().enumerate() {
                        if i > 0 {
                            text.push(',');
                        }
                        if arg.is_missing {
                            continue;
                        }
                        if arg.contains_union {
                            text.push('(');
                            text.push_str(&arg.text);
                            text.push(')');
                        } else {
                            text.push_str(&arg.text);
                        }
                    }
                    text.push(')');

                    stack.push(ExprFragment::new(text));
                    continue;
                }

                let name =
                    function_id_to_name(func_id).ok_or(DecodeRgceError::UnknownFunctionId {
                        offset: ptg_offset,
                        ptg,
                        func_id,
                    })?;

                let mut args = Vec::with_capacity(argc);
                for _ in 0..argc {
                    args.push(stack.pop().ok_or(DecodeRgceError::StackUnderflow {
                        offset: ptg_offset,
                        ptg,
                    })?);
                }
                args.reverse();

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
                    if arg.contains_union {
                        text.push('(');
                        text.push_str(&arg.text);
                        text.push(')');
                    } else {
                        text.push_str(&arg.text);
                    }
                }
                text.push(')');

                stack.push(ExprFragment::new(text));
            }
            // PtgName: [nameId: u32][reserved: u16]
            0x23 | 0x43 | 0x63 => {
                let remaining = rgce.len().saturating_sub(i);
                if remaining < 6 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining,
                    });
                }

                let name_id = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]);
                // Skip `[nameId: u32][reserved: u16]`.
                i = i.saturating_add(6);

                // Best-effort: we don't have workbook name context in this crate, so emit a stable
                // placeholder that is parseable as an Excel identifier.
                let is_value_class = (ptg & 0x60) == 0x40;
                let mut text = String::new();
                let mut precedence = 100;
                if is_value_class {
                    // Value-class names can require legacy implicit intersection (e.g. when the
                    // underlying defined name refers to a multi-cell range). Without workbook
                    // context we can't know, so conservatively emit `@` to preserve scalar
                    // semantics.
                    text.push('@');
                    precedence = 70;
                }
                text.push_str(&format!("Name_{name_id}"));

                stack.push(ExprFragment {
                    text,
                    precedence,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgNameX: [ixti: u16][nameIndex: u16]
            0x39 | 0x59 | 0x79 => {
                let remaining = rgce.len().saturating_sub(i);
                if remaining < 4 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 4,
                        remaining,
                    });
                }

                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let name_index = u16::from_le_bytes([rgce[i + 2], rgce[i + 3]]);
                i = i.saturating_add(4);

                // Best-effort: emit a stable placeholder identifier for the extern name.
                //
                // Excel add-in / UDF calls typically reference extern names via `PtgNameX`
                // followed by `PtgFuncVar(0x00FF)`. Keep the format stable for tests and
                // downstream diagnostics, and ensure it stays parseable as an Excel identifier.
                let is_value_class = (ptg & 0x60) == 0x40;
                let mut text = String::new();
                let mut precedence = 100;
                if is_value_class {
                    text.push('@');
                    precedence = 70;
                }
                text.push_str(&format!("ExternName_IXTI{ixti}_N{name_index}"));

                stack.push(ExprFragment {
                    text,
                    precedence,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgRef
            0x24 | 0x44 | 0x64 => {
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let row = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) + 1;
                let col = u16::from_le_bytes([rgce[i + 4], rgce[i + 5] & 0x3F]) as u32;
                let flags = rgce[i + 5];
                i += 6;

                let mut text = String::new();
                if flags & 0x80 == 0 {
                    text.push('$');
                }
                push_column(col, &mut text);
                if flags & 0x40 == 0 {
                    text.push('$');
                }
                text.push_str(&row.to_string());
                stack.push(ExprFragment::new(text));
            }
            // PtgArea
            0x25 | 0x45 | 0x65 => {
                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let row1 = u32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) + 1;
                let row2 =
                    u32::from_le_bytes([rgce[i + 4], rgce[i + 5], rgce[i + 6], rgce[i + 7]]) + 1;
                let col1 = u16::from_le_bytes([rgce[i + 8], rgce[i + 9] & 0x3F]) as u32;
                let col2 = u16::from_le_bytes([rgce[i + 10], rgce[i + 11] & 0x3F]) as u32;
                let flags1 = rgce[i + 9];
                let flags2 = rgce[i + 11];
                i += 12;

                let mut start = String::new();
                if flags1 & 0x80 == 0 {
                    start.push('$');
                }
                push_column(col1, &mut start);
                if flags1 & 0x40 == 0 {
                    start.push('$');
                }
                start.push_str(&row1.to_string());

                let mut end = String::new();
                if flags2 & 0x80 == 0 {
                    end.push('$');
                }
                push_column(col2, &mut end);
                if flags2 & 0x40 == 0 {
                    end.push('$');
                }
                end.push_str(&row2.to_string());

                let is_single_cell = row1 == row2 && col1 == col2;
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                if is_value_class && !is_single_cell {
                    // Preserve legacy implicit intersection semantics.
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
            // PtgMem* tokens: non-printing, but consume payload bytes to keep parsing aligned.
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if rgce.len().saturating_sub(i) < 2 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 2,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let cce = u16::from_le_bytes([rgce[i], rgce[i + 1]]) as usize;
                i += 2;
                if rgce.len().saturating_sub(i) < cce {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                if let Some(rgcb) = rgcb {
                    consume_rgcb_arrays_in_subexpression(
                        &rgce[i..i + cce],
                        rgcb,
                        &mut rgcb_pos,
                        i,
                    )?;
                }
                i += cce;
            }
            // PtgRefErr: [row: u32][col: u16]
            0x2A | 0x4A | 0x6A => {
                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeRgceError::UnexpectedEof {
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
                    return Err(DecodeRgceError::UnexpectedEof {
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
                let Some((base_row0, base_col0)) = base else {
                    return Err(DecodeRgceError::UnsupportedToken {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                if rgce.len().saturating_sub(i) < 6 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 6,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let row_off =
                    i32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) as i64;
                let col_off = i16::from_le_bytes([rgce[i + 4], rgce[i + 5]]) as i64;
                i += 6;

                const MAX_ROW0: i64 = 1_048_575;
                const MAX_COL0: i64 = 0x3FFF;
                let abs_row0 = base_row0 as i64 + row_off;
                let abs_col0 = base_col0 as i64 + col_off;
                if abs_row0 < 0 || abs_row0 > MAX_ROW0 || abs_col0 < 0 || abs_col0 > MAX_COL0 {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    stack.push(ExprFragment::new(format_cell_ref_a1(
                        abs_row0 as u32,
                        abs_col0 as u32,
                    )));
                }
            }
            // PtgAreaN: [rowFirst_off: i32][rowLast_off: i32][colFirst_off: i16][colLast_off: i16]
            0x2D | 0x4D | 0x6D => {
                let Some((base_row0, base_col0)) = base else {
                    return Err(DecodeRgceError::UnsupportedToken {
                        offset: ptg_offset,
                        ptg,
                    });
                };

                if rgce.len().saturating_sub(i) < 12 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 12,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }

                let row1_off =
                    i32::from_le_bytes([rgce[i], rgce[i + 1], rgce[i + 2], rgce[i + 3]]) as i64;
                let row2_off =
                    i32::from_le_bytes([rgce[i + 4], rgce[i + 5], rgce[i + 6], rgce[i + 7]]) as i64;
                let col1_off = i16::from_le_bytes([rgce[i + 8], rgce[i + 9]]) as i64;
                let col2_off = i16::from_le_bytes([rgce[i + 10], rgce[i + 11]]) as i64;
                i += 12;

                const MAX_ROW0: i64 = 1_048_575;
                const MAX_COL0: i64 = 0x3FFF;
                let abs_row1 = base_row0 as i64 + row1_off;
                let abs_row2 = base_row0 as i64 + row2_off;
                let abs_col1 = base_col0 as i64 + col1_off;
                let abs_col2 = base_col0 as i64 + col2_off;

                if abs_row1 < 0
                    || abs_row1 > MAX_ROW0
                    || abs_row2 < 0
                    || abs_row2 > MAX_ROW0
                    || abs_col1 < 0
                    || abs_col1 > MAX_COL0
                    || abs_col2 < 0
                    || abs_col2 > MAX_COL0
                {
                    stack.push(ExprFragment::new("#REF!".to_string()));
                } else {
                    let start = format_cell_ref_a1(abs_row1 as u32, abs_col1 as u32);
                    let end = format_cell_ref_a1(abs_row2 as u32, abs_col2 as u32);

                    let is_single_cell = abs_row1 == abs_row2 && abs_col1 == abs_col2;
                    let is_value_class = (ptg & 0x60) == 0x40;

                    let mut text = String::new();
                    if is_value_class && !is_single_cell {
                        // Preserve legacy implicit intersection semantics.
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
            }
            // PtgRef3d: [ixti: u16][row: u32][col: u16]
            0x3A | 0x5A | 0x7A => {
                if rgce.len().saturating_sub(i) < 8 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 8,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let row0 = u32::from_le_bytes([rgce[i + 2], rgce[i + 3], rgce[i + 4], rgce[i + 5]]);
                let col_field = u16::from_le_bytes([rgce[i + 6], rgce[i + 7]]);
                i += 8;

                let prefix = format_sheet_placeholder(ixti);
                let cell = format_cell_ref_from_field(row0, col_field);
                stack.push(ExprFragment::new(format!("{prefix}{cell}")));
            }
            // PtgArea3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3B | 0x5B | 0x7B => {
                if rgce.len().saturating_sub(i) < 14 {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed: 14,
                        remaining: rgce.len().saturating_sub(i),
                    });
                }
                let ixti = u16::from_le_bytes([rgce[i], rgce[i + 1]]);
                let row_first0 =
                    u32::from_le_bytes([rgce[i + 2], rgce[i + 3], rgce[i + 4], rgce[i + 5]]);
                let row_last0 =
                    u32::from_le_bytes([rgce[i + 6], rgce[i + 7], rgce[i + 8], rgce[i + 9]]);
                let col_first = u16::from_le_bytes([rgce[i + 10], rgce[i + 11]]);
                let col_last = u16::from_le_bytes([rgce[i + 12], rgce[i + 13]]);
                i += 14;

                let prefix = format_sheet_placeholder(ixti);
                let a = format_cell_ref_from_field(row_first0, col_first);
                let b = format_cell_ref_from_field(row_last0, col_last);

                let is_single_cell =
                    row_first0 == row_last0 && (col_first & 0x3FFF) == (col_last & 0x3FFF);
                let is_value_class = (ptg & 0x60) == 0x40;

                let mut text = String::new();
                let mut precedence = 100;
                if is_value_class && !is_single_cell {
                    text.push('@');
                    precedence = 70;
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
                    precedence,
                    contains_union: false,
                    is_missing: false,
                });
            }
            // PtgRefErr3d: [ixti: u16][row: u32][col: u16]
            0x3C | 0x5C | 0x7C => {
                let needed = 8;
                let remaining = rgce.len().saturating_sub(i);
                if remaining < needed {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed,
                        remaining,
                    });
                }
                i += needed;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            // PtgAreaErr3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
            0x3D | 0x5D | 0x7D => {
                let needed = 14;
                let remaining = rgce.len().saturating_sub(i);
                if remaining < needed {
                    return Err(DecodeRgceError::UnexpectedEof {
                        offset: ptg_offset,
                        ptg,
                        needed,
                        remaining,
                    });
                }
                i += needed;
                stack.push(ExprFragment::new("#REF!".to_string()));
            }
            _ => {
                return Err(DecodeRgceError::UnsupportedToken {
                    offset: ptg_offset,
                    ptg,
                })
            }
        }

        if stack.last().is_some_and(|s| s.text.len() > max_len) {
            return Err(DecodeRgceError::OutputTooLarge {
                offset: ptg_offset,
                ptg,
                max_len,
            });
        }
    }

    if stack.len() == 1 {
        Ok(stack.pop().expect("len checked").text)
    } else {
        Err(DecodeRgceError::StackNotSingular {
            offset: last_ptg_offset,
            ptg: last_ptg,
            stack_len: stack.len(),
        })
    }
}

fn format_sheet_placeholder(ixti: u16) -> String {
    // Best-effort placeholder: without workbook context we cannot resolve `ixti` into a real sheet
    // name, but we can still emit valid sheet-qualified formula text by quoting a stable placeholder.
    let sheet = format!("Sheet{ixti}");
    let mut out = String::new();
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
    out.push('!');
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
    push_column(col, &mut out);
    if !row_relative {
        out.push('$');
    }
    out.push_str(&row1.to_string());
    out
}

fn format_cell_ref_a1(row0: u32, col0: u32) -> String {
    let mut out = String::new();
    push_column(col0, &mut out);
    out.push_str(&(row0 + 1).to_string());
    out
}

fn push_column(mut col: u32, out: &mut String) {
    // Excel column labels are 1-based.
    col += 1;
    let mut buf = [0u8; 10];
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct PtgListDecoded {
    table_id: u32,
    flags: u32,
    col_first: u32,
    col_last: u32,
}

fn decode_ptg_list_payload_best_effort(payload: &[u8; 12]) -> PtgListDecoded {
    // There are multiple "in the wild" encodings for the 12-byte PtgList payload (table refs /
    // structured references). We try a handful of plausible layouts and prefer the one that
    // produces the most reasonable (table_id, flags, column ids) tuple.
    //
    // This logic mirrors `formula-xlsb`'s `decode_ptg_list_payload_best_effort`, but without
    // workbook context for scoring.
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

    let mut candidates = [
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

    // Default to the canonical (documented) layout when it yields a plausible result, to avoid
    // mis-decoding well-formed payloads that are ambiguous under other interpretations.
    if ptg_list_candidate_is_plausible(&candidates[0]) {
        return candidates[0];
    }

    candidates.sort_by_key(|cand| std::cmp::Reverse(score_ptg_list_candidate(cand)));

    candidates[0]
}

fn ptg_list_candidate_is_plausible(cand: &PtgListDecoded) -> bool {
    let col_first = cand.col_first;
    let col_last = cand.col_last;

    // Column id `0` is used as a sentinel for "all columns". Seeing it on only one side is
    // usually a sign we've chosen the wrong payload layout.
    if (col_first == 0) ^ (col_last == 0) {
        return false;
    }

    // Column ids should be in ascending order (except for the all-columns sentinel).
    if col_first != 0 && col_last != 0 && col_first > col_last {
        return false;
    }

    // Table column ids are bounded by Excel's max column count.
    if col_first > 16_384 || col_last > 16_384 {
        return false;
    }

    true
}

fn score_ptg_list_candidate(cand: &PtgListDecoded) -> i32 {
    const FLAG_ALL: u16 = 0x0001;
    const FLAG_HEADERS: u16 = 0x0002;
    const FLAG_DATA: u16 = 0x0004;
    const FLAG_TOTALS: u16 = 0x0008;
    const FLAG_THIS_ROW: u16 = 0x0010;
    const KNOWN_FLAGS: u16 = FLAG_ALL | FLAG_HEADERS | FLAG_DATA | FLAG_TOTALS | FLAG_THIS_ROW;

    let mut score = 0i32;

    // Prefer non-zero table ids.
    if cand.table_id != 0 {
        score += 1;
    } else {
        score -= 1;
    }

    // Prefer candidates where the low 16 bits of the flags field look like known structured-ref
    // flags. Unknown bits are allowed, but usually indicate we chose the wrong layout.
    let flags16 = (cand.flags & 0xFFFF) as u16;
    let unknown = flags16 & !KNOWN_FLAGS;
    if unknown == 0 {
        score += 10;
    } else {
        score -= 10;
    }

    // Slightly prefer flags that fit in 16 bits (canonical layout).
    if cand.flags & 0xFFFF_0000 != 0 {
        score -= 1;
    }

    let col_first = cand.col_first;
    let col_last = cand.col_last;

    // Column id `0` is treated as a sentinel for "all columns"; seeing it on only one side is
    // usually a sign we've chosen the wrong payload layout.
    if (col_first == 0) ^ (col_last == 0) {
        score -= 50;
    }

    // Prefer small, Excel-like column ids.
    if col_first <= 16_384 {
        score += 3;
    } else {
        score -= 3;
    }
    if col_last <= 16_384 {
        score += 3;
    } else {
        score -= 3;
    }

    // Prefer ascending ranges (col_first <= col_last) when both are non-zero.
    if col_first == 0 && col_last == 0 {
        score += 1;
    } else if col_first <= col_last {
        score += 2;
    } else {
        score -= 20;
    }

    // Slightly prefer single-column selections.
    if col_first == col_last {
        score += 1;
    }

    score
}

fn structured_ref_item_from_flags(flags: u16) -> Option<StructuredRefItem> {
    const FLAG_ALL: u16 = 0x0001;
    const FLAG_HEADERS: u16 = 0x0002;
    const FLAG_DATA: u16 = 0x0004;
    const FLAG_TOTALS: u16 = 0x0008;
    const FLAG_THIS_ROW: u16 = 0x0010;

    // Flags are not strictly documented as mutually exclusive. Prefer the same priority order as
    // `formula-xlsb`'s decoder.
    if flags & FLAG_THIS_ROW != 0 {
        Some(StructuredRefItem::ThisRow)
    } else if flags & FLAG_HEADERS != 0 {
        Some(StructuredRefItem::Headers)
    } else if flags & FLAG_TOTALS != 0 {
        Some(StructuredRefItem::Totals)
    } else if flags & FLAG_ALL != 0 {
        Some(StructuredRefItem::All)
    } else if flags & FLAG_DATA != 0 {
        Some(StructuredRefItem::Data)
    } else {
        None
    }
}

fn structured_columns_from_ids(col_first: u32, col_last: u32) -> StructuredColumns {
    if col_first == 0 && col_last == 0 {
        StructuredColumns::All
    } else if col_first == col_last {
        StructuredColumns::Single(format!("Column{col_first}"))
    } else {
        StructuredColumns::Range {
            start: format!("Column{col_first}"),
            end: format!("Column{col_last}"),
        }
    }
}

#[cfg(feature = "encode")]
#[derive(Debug, thiserror::Error)]
pub enum EncodeRgceError {
    #[error("formula parse error: {message} (span {start}..{end})")]
    Parse {
        message: String,
        start: usize,
        end: usize,
    },
    #[error("unsupported expression in BIFF12 encoder: {0}")]
    Unsupported(&'static str),
    #[error("unsupported function name: {0}")]
    UnknownFunction(String),
    #[error("invalid argument count for {name}: got {got}, expected {min}..={max}")]
    InvalidArgCount {
        name: String,
        got: usize,
        min: u8,
        max: u8,
    },
    #[error("invalid numeric literal: {0}")]
    InvalidNumber(String),
    #[error("unsupported error literal: {0}")]
    InvalidErrorLiteral(String),
}

#[cfg(feature = "encode")]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncodedRgce {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

#[cfg(feature = "encode")]
pub fn encode_rgce_with_rgcb(formula: &str) -> Result<EncodedRgce, EncodeRgceError> {
    use formula_engine::{parse_formula, ParseOptions};

    let ast =
        parse_formula(formula, ParseOptions::default()).map_err(|e| EncodeRgceError::Parse {
            message: e.message,
            start: e.span.start,
            end: e.span.end,
        })?;
    let mut rgce = Vec::new();
    let mut rgcb = Vec::new();
    encode_expr(&ast.expr, &mut rgce, &mut rgcb)?;
    Ok(EncodedRgce { rgce, rgcb })
}

#[cfg(feature = "encode")]
pub fn encode_rgce(formula: &str) -> Result<Vec<u8>, EncodeRgceError> {
    let encoded = encode_rgce_with_rgcb(formula)?;
    if !encoded.rgcb.is_empty() {
        return Err(EncodeRgceError::Unsupported("array literals"));
    }
    Ok(encoded.rgce)
}

#[cfg(feature = "encode")]
fn encode_expr(
    expr: &formula_engine::Expr,
    rgce: &mut Vec<u8>,
    rgcb: &mut Vec<u8>,
) -> Result<(), EncodeRgceError> {
    use formula_engine::{BinaryOp, Coord, Expr, PostfixOp, UnaryOp};

    match expr {
        Expr::Number(raw) => {
            let n: f64 = raw
                .parse()
                .map_err(|_| EncodeRgceError::InvalidNumber(raw.clone()))?;
            if n.fract() == 0.0 && n >= 0.0 && n <= u16::MAX as f64 {
                rgce.push(0x1E); // PtgInt
                rgce.extend_from_slice(&(n as u16).to_le_bytes());
            } else {
                rgce.push(0x1F); // PtgNum
                rgce.extend_from_slice(&n.to_le_bytes());
            }
        }
        Expr::String(s) => {
            rgce.push(0x17); // PtgStr
            let units: Vec<u16> = s.encode_utf16().collect();
            let cch: u16 = units
                .len()
                .try_into()
                .map_err(|_| EncodeRgceError::Unsupported("string literal too long"))?;
            rgce.extend_from_slice(&cch.to_le_bytes());
            for u in units {
                rgce.extend_from_slice(&u.to_le_bytes());
            }
        }
        Expr::Boolean(b) => {
            rgce.push(0x1D); // PtgBool
            rgce.push(if *b { 1 } else { 0 });
        }
        Expr::Error(raw) => {
            let code = match raw.to_ascii_uppercase().as_str() {
                "#NULL!" => 0x00,
                "#DIV/0!" => 0x07,
                "#VALUE!" => 0x0F,
                "#REF!" => 0x17,
                "#NAME?" => 0x1D,
                "#NUM!" => 0x24,
                "#N/A" | "#N/A!" => 0x2A,
                "#GETTING_DATA" => 0x2B,
                "#SPILL!" => 0x2C,
                "#CALC!" => 0x2D,
                "#FIELD!" => 0x2E,
                "#CONNECT!" => 0x2F,
                "#BLOCKED!" => 0x30,
                "#UNKNOWN!" => 0x31,
                _ => return Err(EncodeRgceError::InvalidErrorLiteral(raw.clone())),
            };
            rgce.push(0x1C); // PtgErr
            rgce.push(code);
        }
        Expr::CellRef(r) => {
            if r.workbook.is_some() || r.sheet.is_some() {
                return Err(EncodeRgceError::Unsupported(
                    "3D/sheet-qualified references",
                ));
            }
            let (col, col_abs) = match &r.col {
                Coord::A1 { index, abs } => (*index, *abs),
                Coord::Offset(_) => return Err(EncodeRgceError::Unsupported("relative offsets")),
            };
            let (row, row_abs) = match &r.row {
                Coord::A1 { index, abs } => (*index, *abs),
                Coord::Offset(_) => return Err(EncodeRgceError::Unsupported("relative offsets")),
            };
            rgce.push(0x24); // PtgRef
            rgce.extend_from_slice(&row.to_le_bytes());
            rgce.extend_from_slice(&encode_col_with_flags(col, col_abs, row_abs));
        }
        Expr::Binary(b) if b.op == BinaryOp::Range => {
            // Prefer encoding simple A1:A2 areas as PtgArea for Excel-compatible rgce.
            if let (Expr::CellRef(a), Expr::CellRef(bref)) = (&*b.left, &*b.right) {
                if a.workbook.is_none()
                    && a.sheet.is_none()
                    && bref.workbook.is_none()
                    && bref.sheet.is_none()
                {
                    if let (Some((c1, c1_abs)), Some((r1, r1_abs))) =
                        (coord_to_a1(&a.col), coord_to_a1(&a.row))
                    {
                        if let (Some((c2, c2_abs)), Some((r2, r2_abs))) =
                            (coord_to_a1(&bref.col), coord_to_a1(&bref.row))
                        {
                            rgce.push(0x25); // PtgArea
                            rgce.extend_from_slice(&r1.to_le_bytes());
                            rgce.extend_from_slice(&r2.to_le_bytes());
                            rgce.extend_from_slice(&encode_col_with_flags(c1, c1_abs, r1_abs));
                            rgce.extend_from_slice(&encode_col_with_flags(c2, c2_abs, r2_abs));
                            return Ok(());
                        }
                    }
                }
            }

            // Fallback: encode as operator.
            encode_expr(&b.left, rgce, rgcb)?;
            encode_expr(&b.right, rgce, rgcb)?;
            rgce.push(0x11); // PtgRange
        }
        Expr::Binary(b) => {
            encode_expr(&b.left, rgce, rgcb)?;
            encode_expr(&b.right, rgce, rgcb)?;
            let ptg = match b.op {
                BinaryOp::Add => 0x03,
                BinaryOp::Sub => 0x04,
                BinaryOp::Mul => 0x05,
                BinaryOp::Div => 0x06,
                BinaryOp::Pow => 0x07,
                BinaryOp::Concat => 0x08,
                BinaryOp::Lt => 0x09,
                BinaryOp::Le => 0x0A,
                BinaryOp::Eq => 0x0B,
                BinaryOp::Gt => 0x0C,
                BinaryOp::Ge => 0x0D,
                BinaryOp::Ne => 0x0E,
                BinaryOp::Intersect => 0x0F,
                BinaryOp::Union => 0x10,
                BinaryOp::Range => 0x11,
            };
            rgce.push(ptg);
        }
        Expr::Unary(u) if u.op == UnaryOp::ImplicitIntersection => {
            match &*u.expr {
                Expr::CellRef(r) => {
                    if r.workbook.is_some() || r.sheet.is_some() {
                        return Err(EncodeRgceError::Unsupported(
                            "3D/sheet-qualified references",
                        ));
                    }
                    let (col, col_abs) = match &r.col {
                        Coord::A1 { index, abs } => (*index, *abs),
                        Coord::Offset(_) => {
                            return Err(EncodeRgceError::Unsupported("relative offsets"))
                        }
                    };
                    let (row, row_abs) = match &r.row {
                        Coord::A1 { index, abs } => (*index, *abs),
                        Coord::Offset(_) => {
                            return Err(EncodeRgceError::Unsupported("relative offsets"))
                        }
                    };

                    // Encode `@A1` by emitting a value-class reference token (PtgRefV). Excel
                    // uses this representation for legacy implicit intersection.
                    rgce.push(0x44); // PtgRefV
                    rgce.extend_from_slice(&row.to_le_bytes());
                    rgce.extend_from_slice(&encode_col_with_flags(col, col_abs, row_abs));
                }
                Expr::StructuredRef(_) => {
                    return Err(EncodeRgceError::Unsupported(
                        "structured references require workbook table-id context",
                    ));
                }
                Expr::Binary(b) if b.op == BinaryOp::Range => {
                    // Encode `@A1:A2` as PtgAreaV.
                    if let (Expr::CellRef(a), Expr::CellRef(bref)) = (&*b.left, &*b.right) {
                        if a.workbook.is_none()
                            && a.sheet.is_none()
                            && bref.workbook.is_none()
                            && bref.sheet.is_none()
                        {
                            if let (Some((c1, c1_abs)), Some((r1, r1_abs))) =
                                (coord_to_a1(&a.col), coord_to_a1(&a.row))
                            {
                                if let (Some((c2, c2_abs)), Some((r2, r2_abs))) =
                                    (coord_to_a1(&bref.col), coord_to_a1(&bref.row))
                                {
                                    rgce.push(0x45); // PtgAreaV
                                    rgce.extend_from_slice(&r1.to_le_bytes());
                                    rgce.extend_from_slice(&r2.to_le_bytes());
                                    rgce.extend_from_slice(&encode_col_with_flags(
                                        c1, c1_abs, r1_abs,
                                    ));
                                    rgce.extend_from_slice(&encode_col_with_flags(
                                        c2, c2_abs, r2_abs,
                                    ));
                                    return Ok(());
                                }
                            }
                        }
                    }

                    return Err(EncodeRgceError::Unsupported(
                        "implicit intersection (@) on non-area range",
                    ));
                }
                _ => {
                    return Err(EncodeRgceError::Unsupported(
                        "implicit intersection (@) on non-reference",
                    ))
                }
            }
        }
        Expr::Unary(u) => {
            encode_expr(&u.expr, rgce, rgcb)?;
            match u.op {
                UnaryOp::Plus => rgce.push(0x12),
                UnaryOp::Minus => rgce.push(0x13),
                UnaryOp::ImplicitIntersection => {
                    return Err(EncodeRgceError::Unsupported("implicit intersection (@)"));
                }
            }
        }
        Expr::Postfix(p) => {
            encode_expr(&p.expr, rgce, rgcb)?;
            match p.op {
                PostfixOp::Percent => rgce.push(0x14),
                PostfixOp::SpillRange => rgce.push(0x2F),
            }
        }
        Expr::FunctionCall(call) => {
            let name = call.name.name_upper.as_str();
            let Some(func) = crate::function_ids::function_spec_from_name(name) else {
                return Err(EncodeRgceError::UnknownFunction(name.to_string()));
            };

            let argc_usize = call.args.len();
            if argc_usize < func.min_args as usize || argc_usize > func.max_args as usize {
                return Err(EncodeRgceError::InvalidArgCount {
                    name: name.to_string(),
                    got: argc_usize,
                    min: func.min_args,
                    max: func.max_args,
                });
            }

            // Encode args.
            for arg in &call.args {
                if matches!(arg, Expr::Missing) {
                    rgce.push(0x16); // PtgMissArg
                } else {
                    encode_expr(arg, rgce, rgcb)?;
                }
            }

            // Choose token form.
            if func.min_args == func.max_args {
                if argc_usize != func.min_args as usize {
                    return Err(EncodeRgceError::InvalidArgCount {
                        name: name.to_string(),
                        got: argc_usize,
                        min: func.min_args,
                        max: func.max_args,
                    });
                }
                // Fixed arity -> PtgFunc
                rgce.push(0x21);
                rgce.extend_from_slice(&func.id.to_le_bytes());
            } else {
                // Variable arity -> PtgFuncVar
                rgce.push(0x22);
                let argc: u8 = call
                    .args
                    .len()
                    .try_into()
                    .map_err(|_| EncodeRgceError::Unsupported("too many function args"))?;
                rgce.push(argc);
                rgce.extend_from_slice(&func.id.to_le_bytes());
            }
        }
        Expr::Call(_) => return Err(EncodeRgceError::Unsupported("call expressions")),
        Expr::FieldAccess(_) => return Err(EncodeRgceError::Unsupported("field access")),
        Expr::Missing => {
            rgce.push(0x16); // PtgMissArg
        }
        Expr::NameRef(_) => return Err(EncodeRgceError::Unsupported("named references")),
        Expr::ColRef(_) => return Err(EncodeRgceError::Unsupported("column references")),
        Expr::RowRef(_) => return Err(EncodeRgceError::Unsupported("row references")),
        Expr::StructuredRef(_) => {
            return Err(EncodeRgceError::Unsupported(
                "structured references require workbook table-id context",
            ))
        }
        Expr::Array(arr) => {
            // MS-XLSB 2.5.198.8 PtgArray: [unused: 7 bytes] + serialized array constant stored in
            // trailing `rgcb`.
            rgce.push(0x20); // PtgArray
            rgce.extend_from_slice(&[0u8; 7]); // unused
            encode_array_constant(arr, rgcb)?;
        }
    }

    Ok(())
}

#[cfg(feature = "encode")]
fn encode_array_constant(
    arr: &formula_engine::ArrayLiteral,
    rgcb: &mut Vec<u8>,
) -> Result<(), EncodeRgceError> {
    use formula_engine::{Expr, UnaryOp};

    let rows = arr.rows.len();
    let cols = arr.rows.first().map(|r| r.len()).unwrap_or(0);
    if rows == 0 || cols == 0 {
        return Err(EncodeRgceError::Unsupported(
            "array literal cannot be empty",
        ));
    }
    if arr.rows.iter().any(|r| r.len() != cols) {
        return Err(EncodeRgceError::Unsupported(
            "array literal rows must have the same number of columns",
        ));
    }

    let cols_minus1: u16 = (cols - 1)
        .try_into()
        .map_err(|_| EncodeRgceError::Unsupported("array literal is too wide"))?;
    let rows_minus1: u16 = (rows - 1)
        .try_into()
        .map_err(|_| EncodeRgceError::Unsupported("array literal is too tall"))?;
    rgcb.extend_from_slice(&cols_minus1.to_le_bytes());
    rgcb.extend_from_slice(&rows_minus1.to_le_bytes());

    for row in &arr.rows {
        for el in row {
            match el {
                Expr::Missing => {
                    // Empty cell in the array constant.
                    rgcb.push(0x00);
                }
                Expr::Number(raw) => {
                    let n: f64 = raw
                        .parse()
                        .map_err(|_| EncodeRgceError::InvalidNumber(raw.clone()))?;
                    rgcb.push(0x01);
                    rgcb.extend_from_slice(&n.to_le_bytes());
                }
                Expr::Unary(u) if matches!(u.op, UnaryOp::Plus | UnaryOp::Minus) => {
                    let Expr::Number(raw) = &*u.expr else {
                        return Err(EncodeRgceError::Unsupported(
                            "unary +/- in array literals is only supported on numeric literals",
                        ));
                    };
                    let mut n: f64 = raw
                        .parse()
                        .map_err(|_| EncodeRgceError::InvalidNumber(raw.clone()))?;
                    if u.op == UnaryOp::Minus {
                        n = -n;
                    }
                    rgcb.push(0x01);
                    rgcb.extend_from_slice(&n.to_le_bytes());
                }
                Expr::String(s) => {
                    rgcb.push(0x02);
                    let units: Vec<u16> = s.encode_utf16().collect();
                    let cch: u16 = units.len().try_into().map_err(|_| {
                        EncodeRgceError::Unsupported("array string literal too long")
                    })?;
                    rgcb.extend_from_slice(&cch.to_le_bytes());
                    for u in units {
                        rgcb.extend_from_slice(&u.to_le_bytes());
                    }
                }
                Expr::Boolean(b) => {
                    rgcb.push(0x04);
                    rgcb.push(if *b { 1 } else { 0 });
                }
                Expr::Error(raw) => {
                    let code = match raw.to_ascii_uppercase().as_str() {
                        "#NULL!" => 0x00,
                        "#DIV/0!" => 0x07,
                        "#VALUE!" => 0x0F,
                        "#REF!" => 0x17,
                        "#NAME?" => 0x1D,
                        "#NUM!" => 0x24,
                        "#N/A" | "#N/A!" => 0x2A,
                        "#GETTING_DATA" => 0x2B,
                        "#SPILL!" => 0x2C,
                        "#CALC!" => 0x2D,
                        "#FIELD!" => 0x2E,
                        "#CONNECT!" => 0x2F,
                        "#BLOCKED!" => 0x30,
                        "#UNKNOWN!" => 0x31,
                        _ => return Err(EncodeRgceError::InvalidErrorLiteral(raw.clone())),
                    };
                    rgcb.push(0x10);
                    rgcb.push(code);
                }
                _ => {
                    return Err(EncodeRgceError::Unsupported(
                        "only literal values are supported inside array literals",
                    ))
                }
            }
        }
    }

    Ok(())
}

#[cfg(feature = "encode")]
fn coord_to_a1(coord: &formula_engine::Coord) -> Option<(u32, bool)> {
    match coord {
        formula_engine::Coord::A1 { index, abs } => Some((*index, *abs)),
        formula_engine::Coord::Offset(_) => None,
    }
}

#[cfg(feature = "encode")]
fn encode_col_with_flags(col: u32, col_abs: bool, row_abs: bool) -> [u8; 2] {
    let col: u16 = col as u16;
    let [lo, hi] = col.to_le_bytes();
    let mut hi = hi & 0x3F;
    // In BIFF, these flags indicate *relative* (absence of '$').
    if !col_abs {
        hi |= 0x80;
    }
    if !row_abs {
        hi |= 0x40;
    }
    [lo, hi]
}

#[cfg(test)]
mod tests {
    use super::decode_rgce;

    #[test]
    fn decodes_ptg_name_to_parseable_placeholder() {
        // PtgName (ref class) + name_id=1 + reserved u16.
        let rgce = [0x23, 1, 0, 0, 0, 0, 0];
        assert_eq!(decode_rgce(&rgce).unwrap(), "Name_1");
    }

    #[test]
    fn decodes_ptg_name_value_class_preserves_implicit_intersection() {
        // PtgName (value class) should conservatively emit `@`.
        let rgce = [0x43, 1, 0, 0, 0, 0, 0];
        assert_eq!(decode_rgce(&rgce).unwrap(), "@Name_1");
    }

    #[test]
    fn decodes_ptg_namex_to_parseable_placeholder() {
        // PtgNameX (ref class) + ixti=2 + nameIndex=3.
        let rgce = [0x39, 2, 0, 3, 0];
        assert_eq!(decode_rgce(&rgce).unwrap(), "ExternName_IXTI2_N3");
    }

    #[test]
    fn decodes_ptg_namex_value_class_preserves_implicit_intersection() {
        // PtgNameX (value class) should conservatively emit `@`.
        let rgce = [0x59, 2, 0, 3, 0];
        assert_eq!(decode_rgce(&rgce).unwrap(), "@ExternName_IXTI2_N3");
    }
}
