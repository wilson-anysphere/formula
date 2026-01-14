//! BIFF8 worksheet formula helpers.
//!
//! This module contains a best-effort shared-formula fallback used by the `.xls` importer:
//! some non-standard producers emit follower-cell `FORMULA` records whose `rgce` consists of a
//! single `PtgExp` token (pointing at a "base" cell), but omit the corresponding `SHRFMLA`/`ARRAY`
//! definition record.
//!
//! In this case we attempt to recover a per-cell `rgce` by materializing the base cell's own
//! `FORMULA.rgce` token stream (when present and non-`PtgExp`) across the row/col delta.
//!
//! The materializer only needs to adjust BIFF8 reference ptgs that embed absolute coordinates plus
//! relative flags:
//! - `PtgRef` / `PtgArea` (and class variants)
//! - `PtgRef3d` / `PtgArea3d` (and class variants)
//!
//! Relative-offset ptgs (`PtgRefN` / `PtgAreaN`) are copied verbatim because they are interpreted
//! relative to the *current* formula cell at decode time.

use std::collections::HashMap;

use formula_model::CellRef;

use super::{records, rgce, worksheet_formulas};

// BIFF8 limits.
const BIFF8_MAX_ROW0: i64 = u16::MAX as i64; // 0..=65535
const BIFF8_MAX_COL0: i64 = 0x3FFF; // 14-bit field (some writers use 0x3FFF sentinels)

const COL_INDEX_MASK: u16 = 0x3FFF;
const ROW_RELATIVE_BIT: u16 = 0x4000;
const COL_RELATIVE_BIT: u16 = 0x8000;
const RELATIVE_MASK: u16 = 0xC000;

#[derive(Debug, Default)]
pub(crate) struct PtgExpFallbackResult {
    /// Recovered formulas (`CellRef` -> formula text without a leading `=`).
    pub(crate) formulas: HashMap<CellRef, String>,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort recovery of `PtgExp`-only formulas by materializing from the referenced base cell's
/// own `FORMULA.rgce`.
///
/// This is intended as a fallback for malformed BIFF8 sheets where `SHRFMLA` / `ARRAY` records are
/// missing or corrupt.
pub(crate) fn recover_ptgexp_formulas_from_base_cell(
    workbook_stream: &[u8],
    sheet_offset: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<PtgExpFallbackResult, String> {
    let allows_continuation = |id: u16| id == worksheet_formulas::RECORD_FORMULA;
    let mut iter = records::LogicalBiffRecordIter::from_offset(
        workbook_stream,
        sheet_offset,
        allows_continuation,
    )?;

    // Collect all cell formula rgce bytes first so PtgExp followers can reference bases that
    // appear later in the stream.
    let mut rgce_by_cell: HashMap<(u32, u32), Vec<u8>> = HashMap::new();
    let mut ptgexp_cells: Vec<(u32, u32, u32, u32)> = Vec::new();
    let mut warnings: Vec<String> = Vec::new();

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                warnings.push(format!("malformed BIFF record in worksheet stream: {err}"));
                break;
            }
        };

        if record.offset != sheet_offset && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }
        if record.record_id != worksheet_formulas::RECORD_FORMULA {
            continue;
        }

        let parsed = match worksheet_formulas::parse_biff8_formula_record(&record) {
            Ok(parsed) => parsed,
            Err(err) => {
                warnings.push(format!(
                    "failed to parse FORMULA record at offset {} in worksheet stream: {err}",
                    record.offset
                ));
                continue;
            }
        };
        let row = parsed.row as u32;
        let col = parsed.col as u32;
        let rgce = parsed.rgce;

        rgce_by_cell.insert((row, col), rgce.clone());

        if let Some((base_row, base_col)) = parse_ptg_exp(&rgce) {
            ptgexp_cells.push((row, col, base_row, base_col));
        }
    }

    let mut recovered: HashMap<CellRef, String> = HashMap::new();

    for (row, col, base_row, base_col) in ptgexp_cells {
        let Some(base_rgce) = rgce_by_cell.get(&(base_row, base_col)) else {
            warnings.push(format!(
                "failed to recover shared formula at {}: base cell ({},{}) has no FORMULA record",
                CellRef::new(row, col).to_a1(),
                base_row,
                base_col
            ));
            continue;
        };

        if base_rgce.first().copied() == Some(0x01) {
            // Base cell also stores PtgExp; without SHRFMLA/ARRAY we can't resolve.
            warnings.push(format!(
                "failed to recover shared formula at {}: base cell {} stores PtgExp (missing SHRFMLA/ARRAY definition)",
                CellRef::new(row, col).to_a1(),
                CellRef::new(base_row, base_col).to_a1()
            ));
            continue;
        }

        let Some(materialized) = materialize_biff8_rgce(base_rgce, base_row, base_col, row, col)
        else {
            warnings.push(format!(
                "failed to recover shared formula at {}: could not materialize base rgce from {} (unsupported or malformed tokens)",
                CellRef::new(row, col).to_a1(),
                CellRef::new(base_row, base_col).to_a1()
            ));
            continue;
        };

        let base_coord = rgce::CellCoord::new(row, col);
        let decoded = rgce::decode_biff8_rgce_with_base(&materialized, ctx, Some(base_coord));
        if !decoded.warnings.is_empty() {
            for w in decoded.warnings {
                warnings.push(format!(
                    "failed to fully decode recovered shared formula at {}: {w}",
                    CellRef::new(row, col).to_a1()
                ));
            }
        }

        if decoded.text.is_empty() {
            // Avoid replacing an existing formula with an empty string.
            warnings.push(format!(
                "failed to recover shared formula at {}: decoded rgce produced empty text",
                CellRef::new(row, col).to_a1()
            ));
            continue;
        }

        recovered.insert(CellRef::new(row, col), decoded.text);
    }

    Ok(PtgExpFallbackResult {
        formulas: recovered,
        warnings,
    })
}

fn parse_ptg_exp(rgce: &[u8]) -> Option<(u32, u32)> {
    // BIFF8 PtgExp: [0x01][rw: u16][col: u16]
    if rgce.first().copied()? != 0x01 {
        return None;
    }
    if rgce.len() < 5 {
        return None;
    }
    let row = u16::from_le_bytes([rgce[1], rgce[2]]) as u32;
    let col = u16::from_le_bytes([rgce[3], rgce[4]]) as u32;
    Some((row, col))
}

fn cell_in_bounds(row: i64, col: i64) -> bool {
    row >= 0 && row <= BIFF8_MAX_ROW0 && col >= 0 && col <= BIFF8_MAX_COL0
}

fn pack_col_with_flags(col0: u16, flags: u16) -> u16 {
    (col0 & COL_INDEX_MASK) | (flags & RELATIVE_MASK)
}

fn adjust_row_col(row0: u16, col_field: u16, delta_row: i64, delta_col: i64) -> Option<(u16, u16)> {
    let row_rel = (col_field & ROW_RELATIVE_BIT) != 0;
    let col_rel = (col_field & COL_RELATIVE_BIT) != 0;
    let col0 = (col_field & COL_INDEX_MASK) as i64;
    let row0_i64 = row0 as i64;

    let new_row = if row_rel {
        row0_i64 + delta_row
    } else {
        row0_i64
    };
    let new_col = if col_rel { col0 + delta_col } else { col0 };

    if !cell_in_bounds(new_row, new_col) {
        return None;
    }

    let new_row_u16 = new_row as u16;
    let new_col_u16 = pack_col_with_flags(new_col as u16, col_field);
    Some((new_row_u16, new_col_u16))
}

pub(crate) fn materialize_biff8_rgce(
    base: &[u8],
    base_row: u32,
    base_col: u32,
    row: u32,
    col: u32,
) -> Option<Vec<u8>> {
    let delta_row = row as i64 - base_row as i64;
    let delta_col = col as i64 - base_col as i64;

    let mut out = Vec::with_capacity(base.len());
    let mut i = 0usize;
    while i < base.len() {
        let ptg = *base.get(i)?;
        i += 1;

        match ptg {
            // PtgExp / PtgTbl: not expected in a base-cell formula for this fallback.
            0x01 | 0x02 => return None,

            // Fixed-width / no-payload tokens.
            0x03..=0x16 | 0x2F => out.push(ptg),

            // PtgStr (ShortXLUnicodeString): variable.
            0x17 => {
                let len = biff8_short_unicode_string_len(base.get(i..)?)?;
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + len)?);
                i += len;
            }

            // PtgExtend / PtgExtendV / PtgExtendA (ptg=0x18 variants).
            //
            // Structured references (tables) use an `etpg` subtype byte (`0x19` = PtgList) followed
            // by a fixed 12-byte payload. Other (unsupported) subtypes appear in the wild with an
            // opaque 5-byte payload; copy them verbatim so the rgce stream stays aligned.
            0x18 | 0x38 | 0x58 | 0x78 => {
                let etpg = *base.get(i)?;
                out.push(ptg);
                out.push(etpg);
                i += 1;

                if etpg == 0x19 {
                    out.extend_from_slice(base.get(i..i + 12)?);
                    i += 12;
                } else {
                    out.extend_from_slice(base.get(i..i + 4)?);
                    i += 4;
                }
            }

            // PtgAttr: [grbit: u8][wAttr: u16] (+ optional jump table for tAttrChoose)
            0x19 => {
                if i + 3 > base.len() {
                    return None;
                }
                out.push(ptg);
                let grbit = base[i];
                let w_attr = u16::from_le_bytes([base[i + 1], base[i + 2]]);
                out.extend_from_slice(&base[i..i + 3]);
                i += 3;

                const T_ATTR_CHOOSE: u8 = 0x04;
                if grbit & T_ATTR_CHOOSE != 0 {
                    let needed = (w_attr as usize).checked_mul(2)?;
                    out.extend_from_slice(base.get(i..i + needed)?);
                    i += needed;
                }
            }

            // PtgErr / PtgBool: 1 byte.
            0x1C | 0x1D => {
                out.push(ptg);
                out.push(*base.get(i)?);
                i += 1;
            }

            // PtgInt: 2 bytes.
            0x1E => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }

            // PtgNum: 8 bytes.
            0x1F => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }

            // PtgArray: 7 bytes of unused/reserved payload (array data stored elsewhere).
            0x20 | 0x40 | 0x60 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 7)?);
                i += 7;
            }

            // PtgFunc: [iftab: u16]
            0x21 | 0x41 | 0x61 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 2)?);
                i += 2;
            }

            // PtgFuncVar: [argc: u8][iftab: u16]
            0x22 | 0x42 | 0x62 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 3)?);
                i += 3;
            }

            // PtgName: [name_id: u32][reserved: u16]
            0x23 | 0x43 | 0x63 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }

            // PtgRef: [rw: u16][col: u16]
            0x24 | 0x44 | 0x64 => {
                let row0 = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let col_field = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let payload = base.get(i..i + 4)?;

                if let Some((new_row, new_col_field)) =
                    adjust_row_col(row0, col_field, delta_row, delta_col)
                {
                    out.push(ptg);
                    out.extend_from_slice(&new_row.to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                } else {
                    // Out-of-bounds references materialize as PtgRefErr*.
                    out.push(ptg.saturating_add(0x06));
                    out.extend_from_slice(payload);
                }

                i += 4;
            }

            // PtgArea: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x25 | 0x45 | 0x65 => {
                let row1 = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row2 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let col1 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col2 = u16::from_le_bytes(base.get(i + 6..i + 8)?.try_into().ok()?);
                let payload = base.get(i..i + 8)?;

                let adjusted1 = adjust_row_col(row1, col1, delta_row, delta_col);
                let adjusted2 = adjust_row_col(row2, col2, delta_row, delta_col);
                if let (Some((new_row1, new_col1)), Some((new_row2, new_col2))) =
                    (adjusted1, adjusted2)
                {
                    out.push(ptg);
                    out.extend_from_slice(&new_row1.to_le_bytes());
                    out.extend_from_slice(&new_row2.to_le_bytes());
                    out.extend_from_slice(&new_col1.to_le_bytes());
                    out.extend_from_slice(&new_col2.to_le_bytes());
                } else {
                    out.push(ptg.saturating_add(0x06));
                    out.extend_from_slice(payload);
                }

                i += 8;
            }

            // PtgMem* tokens: [cce: u16][rgce: cce bytes]
            0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49 | 0x69
            | 0x2E | 0x4E | 0x6E => {
                if i + 2 > base.len() {
                    return None;
                }
                let cce = u16::from_le_bytes([base[i], base[i + 1]]) as usize;
                out.push(ptg);
                out.extend_from_slice(&base[i..i + 2]);
                i += 2;
                out.extend_from_slice(base.get(i..i + cce)?);
                i += cce;
            }

            // PtgRefErr / PtgRefErrN: [rw: u16][col: u16]
            0x2A | 0x4A | 0x6A => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 4)?);
                i += 4;
            }

            // PtgAreaErr / PtgAreaErrN: [rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x2B | 0x4B | 0x6B => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }

            // PtgRefN: keep verbatim (relative offsets resolved at decode time).
            0x2C | 0x4C | 0x6C => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 4)?);
                i += 4;
            }

            // PtgAreaN: keep verbatim.
            0x2D | 0x4D | 0x6D => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 8)?);
                i += 8;
            }

            // PtgNameX: [ixti: u16][iname: u16][reserved: u16]
            0x39 | 0x59 | 0x79 => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }

            // PtgRef3d: [ixti: u16][rw: u16][col: u16]
            0x3A | 0x5A | 0x7A => {
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row0 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let col_field = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let payload = base.get(i..i + 6)?;

                if let Some((new_row, new_col_field)) =
                    adjust_row_col(row0, col_field, delta_row, delta_col)
                {
                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&new_row.to_le_bytes());
                    out.extend_from_slice(&new_col_field.to_le_bytes());
                } else {
                    // Out-of-bounds -> PtgRefErr3d*
                    out.push(ptg.saturating_add(0x02));
                    out.extend_from_slice(payload);
                }

                i += 6;
            }

            // PtgArea3d: [ixti: u16][rwFirst: u16][rwLast: u16][colFirst: u16][colLast: u16]
            0x3B | 0x5B | 0x7B => {
                let ixti = u16::from_le_bytes(base.get(i..i + 2)?.try_into().ok()?);
                let row1 = u16::from_le_bytes(base.get(i + 2..i + 4)?.try_into().ok()?);
                let row2 = u16::from_le_bytes(base.get(i + 4..i + 6)?.try_into().ok()?);
                let col1 = u16::from_le_bytes(base.get(i + 6..i + 8)?.try_into().ok()?);
                let col2 = u16::from_le_bytes(base.get(i + 8..i + 10)?.try_into().ok()?);
                let payload = base.get(i..i + 10)?;

                let adjusted1 = adjust_row_col(row1, col1, delta_row, delta_col);
                let adjusted2 = adjust_row_col(row2, col2, delta_row, delta_col);
                if let (Some((new_row1, new_col1)), Some((new_row2, new_col2))) =
                    (adjusted1, adjusted2)
                {
                    out.push(ptg);
                    out.extend_from_slice(&ixti.to_le_bytes());
                    out.extend_from_slice(&new_row1.to_le_bytes());
                    out.extend_from_slice(&new_row2.to_le_bytes());
                    out.extend_from_slice(&new_col1.to_le_bytes());
                    out.extend_from_slice(&new_col2.to_le_bytes());
                } else {
                    // Out-of-bounds -> PtgAreaErr3d*
                    out.push(ptg.saturating_add(0x02));
                    out.extend_from_slice(payload);
                }

                i += 10;
            }

            // PtgRefErr3d / PtgAreaErr3d: copy verbatim.
            0x3C | 0x5C | 0x7C => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x3D | 0x5D | 0x7D => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 10)?);
                i += 10;
            }

            // PtgRefN3d / PtgAreaN3d: keep verbatim.
            0x3E | 0x5E | 0x7E => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 6)?);
                i += 6;
            }
            0x3F | 0x5F | 0x7F => {
                out.push(ptg);
                out.extend_from_slice(base.get(i..i + 10)?);
                i += 10;
            }

            _ => return None,
        }
    }

    Some(out)
}

fn biff8_short_unicode_string_len(input: &[u8]) -> Option<usize> {
    // ShortXLUnicodeString payload:
    //   [cch: u8][flags: u8]
    //   [cRun: u16]? (if flags & 0x08)
    //   [cbExtRst: u32]? (if flags & 0x04)
    //   [chars: cch bytes or cch*2 bytes]
    //   [rgRun: 4*cRun bytes]?
    //   [ext: cbExtRst bytes]?
    if input.len() < 2 {
        return None;
    }
    let cch = input[0] as usize;
    let flags = input[1];
    let mut offset = 2usize;

    let rich_runs = if flags & 0x08 != 0 {
        let runs = u16::from_le_bytes([*input.get(offset)?, *input.get(offset + 1)?]) as usize;
        offset += 2;
        runs
    } else {
        0usize
    };

    let ext_size = if flags & 0x04 != 0 {
        let size = u32::from_le_bytes([
            *input.get(offset)?,
            *input.get(offset + 1)?,
            *input.get(offset + 2)?,
            *input.get(offset + 3)?,
        ]) as usize;
        offset += 4;
        size
    } else {
        0usize
    };

    let is_unicode = flags & 0x01 != 0;
    let char_bytes = if is_unicode { cch.checked_mul(2)? } else { cch };
    offset = offset.checked_add(char_bytes)?;

    let rich_bytes = rich_runs.checked_mul(4)?;
    offset = offset.checked_add(rich_bytes)?;
    offset = offset.checked_add(ext_size)?;

    (input.len() >= offset).then_some(offset)
}
